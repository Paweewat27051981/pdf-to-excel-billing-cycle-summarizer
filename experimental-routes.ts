// ============================================================================
// [ทดลอง] route กลุ่ม /api/experimental/* — แยก 100% ไม่แตะ business logic เดิม
// ถ้าเลิกทดลอง: ลบไฟล์นี้ + เอา registerExperimentalRoutes ออกจาก server.ts
// ============================================================================
import type { Express, Request, Response } from 'express';
import { fetchCurrentOilPrice, fetchProvincialOilPrice, fetchHistoricalProvincialOilPrice, pickDiesel, summarizeProvinceDiesel, fetchBranchDiesel, BRANCH_OIL_CONFIGS, BranchOilConfig } from './src/experimental/orOilPrice.js';
import { computeFuelTrip, DEFAULT_FUEL_POLICY, FuelTripInput, FuelTripPolicy, computeLoopTripCost } from './src/experimental/fuelTripCalc.js';
import { geocodeDistrict, nearestNeighborLoop, LoopStop } from './src/experimental/dohDistance.js';
import type { OilPriceRecord, TripDistanceRecord } from './src/types.js';

// พิกัดคลัง 6 สาขา (จุดต้นทางของลูป) — จาก Google Maps ผู้ใช้
const WAREHOUSES: Record<string, { lat: number; lon: number }> = {
  'สาย3': { lat: 13.7875785, lon: 100.3581848 },
  'นครสวรรค์': { lat: 15.7337098, lon: 100.1145263 },
  'กำแพงเพชร': { lat: 16.4405574, lon: 99.552246 },
  'พิษณุโลก': { lat: 16.8077557, lon: 100.2222959 },
  'เชียงใหม่': { lat: 18.7791288, lon: 99.0293415 },
  'แม่สอด': { lat: 16.7048519, lon: 98.5708395 },
};
// map branchId -> ชื่อสาขา (สำหรับหาคลัง + ราคาน้ำมัน)
const BRANCH_ID_NAME: Record<string, string> = {
  'br-sai3': 'สาย3', 'br-nakhonsawan': 'นครสวรรค์', 'br-kamphaengphet': 'กำแพงเพชร',
  'br-phitsanulok': 'พิษณุโลก', 'br-chiangmai': 'เชียงใหม่', 'br-maesot': 'แม่สอด',
};

export function registerExperimentalRoutes(
  app: Express,
  // inject จาก server (ใช้ getDb/saveRecord เดิม — ไม่ผูก import วนกัน)
  deps: { getDb: () => Promise<any>; saveRecord: (coll: any, rec: { id: string }) => Promise<void>; genId: (p: string) => string },
): void {
  // ราคาน้ำมันจาก OR — ?province=<ชื่อไทย> (ว่าง = กทม./ปริมณฑล), &date=DD-MM-YYYY (ย้อนหลังรายจังหวัด)
  app.get('/api/experimental/oil-price', async (req: Request, res: Response) => {
    try {
      // ?branch=<ชื่อสาขา> = ใช้ config (สาย3=กทม, แม่สอด=ตาก/อ.แม่สอด) — สะดวกสำหรับหน้าคำนวณ
      const branch = typeof req.query.branch === 'string' ? req.query.branch.trim() : '';
      if (branch) {
        const cfg = BRANCH_OIL_CONFIGS.find((c) => c.branch === branch);
        if (!cfg) return res.status(400).json({ error: `ไม่รู้จักสาขา "${branch}"` });
        const r = await fetchBranchDiesel(cfg);
        return res.json({ source: 'OR', branch, province: cfg.province || 'กทม./ปริมณฑล', fetchedAt: new Date().toISOString(),
          diesel: r.representative, dieselPriceRange: { min: r.min, max: r.max } });
      }
      const province = typeof req.query.province === 'string' ? req.query.province.trim() : '';
      const date = typeof req.query.date === 'string' ? req.query.date.trim() : '';
      let prices;
      if (province && date) {
        const [dd, mm, yyyy] = date.split(/[-/]/);
        if (!dd || !mm || !yyyy) return res.status(400).json({ error: 'date ต้องเป็น DD-MM-YYYY' });
        prices = await fetchHistoricalProvincialOilPrice(province, dd, mm, yyyy);
      } else if (province) {
        prices = await fetchProvincialOilPrice(province);
      } else {
        prices = await fetchCurrentOilPrice();
      }
      // รายจังหวัด: สรุปตัวแทน (อ.เมือง) + ช่วงราคาทุกอำเภอ | กทม.: ดีเซลตัวเดียว
      if (province) {
        const s = summarizeProvinceDiesel(prices);
        res.json({ source: 'OR', province, fetchedAt: new Date().toISOString(),
          diesel: s.representative, dieselPriceRange: { min: s.min, max: s.max }, byLocation: s.byLocation, prices });
      } else {
        res.json({ source: 'OR', province: 'กทม./ปริมณฑล', fetchedAt: new Date().toISOString(), diesel: pickDiesel(prices), prices });
      }
    } catch (err: any) {
      // ตาม prompt: OR ล่ม -> แจ้ง error ชัด (fallback ราคา verified ทำในเฟสถัดไป)
      res.status(502).json({ error: `ดึงราคา OR ไม่สำเร็จ: ${err.message}`, stale: true });
    }
  });

  // สรุปราคาดีเซล อ.เมือง หลายสาขาทีเดียว (ปัจจุบัน หรือย้อนหลังถ้าส่ง date=DD-MM-YYYY)
  // ?provinces=นครสวรรค์,พิษณุโลก,... (ว่าง = ชุด default 5 สาขา)
  app.get('/api/experimental/oil-price/branches', async (req: Request, res: Response) => {
    try {
      // ถ้าส่ง provinces เอง = ใช้ อ.เมือง; ไม่ส่ง = ใช้ config สาขาจริง (สาย3=กทม., แม่สอด=ตาก/อ.แม่สอด)
      const configs: BranchOilConfig[] = (typeof req.query.provinces === 'string' && req.query.provinces.trim())
        ? req.query.provinces.split(',').map((s) => ({ branch: s.trim(), kind: 'province' as const, province: s.trim(), repDistrict: 'เมือง' })).filter((c) => c.province)
        : BRANCH_OIL_CONFIGS;
      const date = typeof req.query.date === 'string' ? req.query.date.trim() : '';
      const [dd, mm, yyyy] = date ? date.split(/[-/]/) : [];
      const dateArg = date && dd && mm && yyyy ? { dd, mm, yyyy } : undefined;
      const rows = await Promise.all(configs.map(async (c) => {
        try {
          const r = await fetchBranchDiesel(c, dateArg);
          return { branch: c.branch, province: c.province || 'กทม./ปริมณฑล', diesel: r.representative?.price ?? null,
            location: r.representative?.location ?? (c.kind === 'bkk' ? 'กทม./ปริมณฑล' : null),
            priceDate: r.representative?.priceDate ?? null, min: r.min, max: r.max, districts: r.districts };
        } catch (e: any) { return { branch: c.branch, province: c.province || 'กทม./ปริมณฑล', error: e.message, diesel: null }; }
      }));
      res.json({ source: 'OR', date: date || 'ปัจจุบัน', fetchedAt: new Date().toISOString(), branches: rows });
    } catch (err: any) {
      res.status(502).json({ error: `ดึงราคาหลายสาขาไม่สำเร็จ: ${err.message}` });
    }
  });

  // บันทึกราคาดีเซล อ.เมือง (ตัวแทนจังหวัด) ลง Firebase ถาวร — snapshot ตรวจย้อนหลังได้
  // body: { provinces?: string[], date?: "DD-MM-YYYY" (ย้อนหลัง), savedBy?: string }
  app.post('/api/experimental/oil-price/save', async (req: Request, res: Response) => {
    try {
      const configs: BranchOilConfig[] = Array.isArray(req.body?.provinces) && req.body.provinces.length
        ? req.body.provinces.map((p: string) => ({ branch: p, kind: 'province' as const, province: p, repDistrict: 'เมือง' }))
        : BRANCH_OIL_CONFIGS;
      const date = typeof req.body?.date === 'string' ? req.body.date.trim() : '';
      const [dd, mm, yyyy] = date ? date.split(/[-/]/) : [];
      const dateArg = date && dd && mm && yyyy ? { dd, mm, yyyy } : undefined;
      const db = await deps.getDb();
      if (!Array.isArray(db.oilPrices)) db.oilPrices = [];
      const now = new Date().toISOString();
      const saved: OilPriceRecord[] = [];
      const skipped: string[] = [];
      for (const c of configs) {
        try {
          const r = await fetchBranchDiesel(c, dateArg);
          if (!r.representative) { skipped.push(c.branch); continue; }
          const priceDate = String(r.representative.priceDate || '').slice(0, 10);
          // กันซ้ำ: สาขา + วันที่ราคาเดียวกัน = อัปเดตแทนที่ (idempotent)
          const existing = db.oilPrices.find((x: OilPriceRecord) => x.branch === c.branch && x.priceDate === priceDate);
          const rec: OilPriceRecord = {
            id: existing?.id || deps.genId('oil'),
            branch: c.branch, province: c.province || 'กทม./ปริมณฑล',
            location: r.representative.location || (c.kind === 'bkk' ? 'กทม./ปริมณฑล' : ''), product: r.representative.product,
            price: r.representative.price, priceDate, minPrice: r.min ?? undefined, maxPrice: r.max ?? undefined,
            source: 'OR', fetchedAt: now, savedBy: req.body?.savedBy || 'user',
          };
          if (existing) Object.assign(existing, rec); else db.oilPrices.push(rec);
          await deps.saveRecord('oilPrices', rec); // granular เขียนแค่ record เดียว
          saved.push(rec);
        } catch (e: any) { skipped.push(`${c.branch} (${e.message})`); }
      }
      res.json({ success: true, savedCount: saved.length, saved, skipped });
    } catch (err: any) {
      res.status(502).json({ error: `บันทึกราคาไม่สำเร็จ: ${err.message}` });
    }
  });

  // อ่านราคาที่บันทึกไว้ (ถาวร) — ?province=&date=(YYYY-MM-DD)
  app.get('/api/experimental/oil-price/saved', async (req: Request, res: Response) => {
    try {
      const db = await deps.getDb();
      let list: OilPriceRecord[] = Array.isArray(db.oilPrices) ? db.oilPrices.slice() : [];
      const province = typeof req.query.province === 'string' ? req.query.province.trim() : '';
      const date = typeof req.query.date === 'string' ? req.query.date.trim() : '';
      if (province) list = list.filter((r) => r.province === province);
      if (date) list = list.filter((r) => r.priceDate === date);
      list.sort((a, b) => (b.priceDate || '').localeCompare(a.priceDate || '') || a.province.localeCompare(b.province));
      res.json({ count: list.length, records: list });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ===== รายงานเทียบต้นทุน (DOH ระยะจริง) =====
  // helper: อำเภอ distinct ของใบ (เรียงตาม normalQty มากสุดไม่สำคัญ — nearest-neighbor เรียงเอง)
  const tripDests = (t: any): { district: string; province: string }[] => {
    const seen = new Set<string>(); const out: { district: string; province: string }[] = [];
    for (const r of (t.receipts || [])) {
      const district = (r.districtRaw || '').trim(); const province = (r.provinceRaw || '').trim();
      if (!district && !province) continue;
      const k = `${district}|${province}`;
      if (!seen.has(k)) { seen.add(k); out.push({ district, province }); }
    }
    return out;
  };
  // จำนวน "จุดลงของ" จริง (ใบรับที่มีปลายทาง) — ใช้คิดเวลาขนของ ต่างจาก route stops ที่ยุบอำเภอซ้ำ
  const tripStoreCount = (t: any): number =>
    (t.receipts || []).filter((r: any) => ((r.districtRaw || '').trim() || (r.provinceRaw || '').trim())).length;
  // fingerprint ของ input ที่ใช้คิด (ปลายทาง + จำนวนจุดลง) — recalculate/แก้ใบรับ เปลี่ยนค่านี้ = cache stale ต้องคิดใหม่
  const tripFingerprint = (t: any): string => {
    const dk = tripDests(t).map((d) => `${d.district}|${d.province}`).sort().join(';');
    return `${dk}#${tripStoreCount(t)}`;
  };

  // GET รายงานต่อทะเบียน (อ่าน cache ที่คิดไว้แล้ว — ไม่ยิง DOH) ?cycleId=&branchId=
  app.get('/api/experimental/fuel-report', async (req: Request, res: Response) => {
    try {
      const cycleId = typeof req.query.cycleId === 'string' ? req.query.cycleId : '';
      const branchId = typeof req.query.branchId === 'string' ? req.query.branchId : '';
      if (!cycleId) return res.status(400).json({ error: 'ต้องระบุ cycleId' });
      const db = await deps.getDb();
      const distById = new Map<string, TripDistanceRecord>((db.tripDistances || []).map((d: TripDistanceRecord) => [d.tripId, d]));
      const trips = (db.tripDocuments || []).filter((t: any) => t.cycleId === cycleId && (!branchId || t.branchId === branchId));
      // ราคาน้ำมันอ้างอิงต่อสาขา — เลือก record "ล่าสุด" (priceDate ใหม่สุด ไม่งั้น fetchedAt)
      const oilLatest = new Map<string, OilPriceRecord>();
      for (const o of (db.oilPrices || []) as OilPriceRecord[]) {
        const cur = oilLatest.get(o.branch);
        const newer = !cur || (o.priceDate || '') > (cur.priceDate || '') ||
          ((o.priceDate || '') === (cur.priceDate || '') && (o.fetchedAt || '') > (cur.fetchedAt || ''));
        if (newer) oilLatest.set(o.branch, o);
      }
      const oilByBranch = new Map<string, number>([...oilLatest].map(([b, o]) => [b, o.price]));
      // รวมต่อทะเบียน
      const byPlate = new Map<string, any>();
      for (const t of trips) {
        const p = t.plateNo || '-';
        // key ต้องรวม branchId — ทะเบียนซ้ำข้ามสาขาได้ (ราคาน้ำมันคนละสาขา) ห้ามยุบรวมแถวเดียว
        const key = `${t.branchId}|${p}`;
        const g = byPlate.get(key) || { branchId: t.branchId, branch: BRANCH_ID_NAME[t.branchId] || '', plateNo: p, tripCount: 0, tripAmount: 0, fuelCostBand: 0, computed: 0, missing: 0 };
        g.tripCount++; g.tripAmount += t.tripAmount || 0;
        const dist = distById.get(t.id);
        // นับ computed ต่อเมื่อ: มีระยะ + ครบทุกจุด + fingerprint ตรงปัจจุบัน (ปลายทาง/จุดลงไม่เปลี่ยนหลัง recalculate)
        if (dist && dist.loopKm > 0 && !(dist.missing && dist.missing.length) && dist.destKey === tripFingerprint(t)) {
          const branchName = BRANCH_ID_NAME[t.branchId] || '';
          const oil = oilByBranch.get(branchName) ?? 0;
          if (oil > 0) { g.fuelCostBand += computeLoopTripCost(dist.loopKm, dist.storeCount, oil).totalCost; g.computed++; }
          else g.missing++;
        } else g.missing++;
        byPlate.set(key, g);
      }
      const rows = [...byPlate.values()].map((g) => ({
        ...g, tripAmount: Math.round(g.tripAmount * 100) / 100, fuelCostBand: Math.round(g.fuelCostBand * 100) / 100,
        diff: Math.round((g.tripAmount - g.fuelCostBand) * 100) / 100,
      })).sort((a, b) => (a.branch || '').localeCompare(b.branch || '') || a.plateNo.localeCompare(b.plateNo));
      res.json({ cycleId, count: rows.length, rows });
    } catch (err: any) { res.status(500).json({ error: err.message }); }
  });

  // POST คิดระยะลูป+ต้นทุน ของ "1 ทะเบียน" (ยิง DOH เฉพาะใบที่ยังไม่ cache) — ผู้ใช้กดทีละคัน
  app.post('/api/experimental/fuel-report/compute-plate', async (req: Request, res: Response) => {
    try {
      const { cycleId, plateNo, branchId } = req.body as { cycleId: string; plateNo: string; branchId?: string };
      if (!cycleId || !plateNo) return res.status(400).json({ error: 'ต้องระบุ cycleId, plateNo' });
      const db = await deps.getDb();
      if (!Array.isArray(db.tripDistances)) db.tripDistances = [];
      const cached = new Map<string, TripDistanceRecord>((db.tripDistances as TripDistanceRecord[]).map((d) => [d.tripId, d]));
      // plateNo ไม่ unique ข้ามสาขา -> filter branchId ด้วย (ตรงกับ /fuel-report) กันไปคิดใบสาขาอื่น
      const trips = (db.tripDocuments || []).filter((t: any) => t.cycleId === cycleId && t.plateNo === plateNo && (!branchId || t.branchId === branchId));
      let computed = 0, skipped = 0;
      for (const t of trips) {
        const dests = tripDests(t);
        const destKey = tripFingerprint(t); // ปลายทาง + จำนวนจุดลง
        const prev = cached.get(t.id);
        // skip เฉพาะ cache สมบูรณ์ (ไม่มี missing) + ปลายทางไม่เปลี่ยน — ล่มชั่วคราว/recalculate เปลี่ยนปลายทาง จะคิดใหม่
        if (prev && !(prev.missing && prev.missing.length) && prev.destKey === destKey) { skipped++; continue; }
        const wh = WAREHOUSES[BRANCH_ID_NAME[t.branchId] || ''];
        if (!wh) { skipped++; continue; }
        // geocode ทุกอำเภอ — อำเภอที่หาพิกัดไม่ได้ ต้องบันทึกลง missing (ไม่ปล่อยหายเงียบ)
        const stops: LoopStop[] = [];
        const geoMissing: string[] = [];
        for (const d of dests) {
          const loc = await geocodeDistrict(d.district, d.province);
          if (loc) stops.push({ ...d, loc }); else geoMissing.push(`อ.${d.district}`);
        }
        const loop = await nearestNeighborLoop(wh, stops);
        const missing = [...geoMissing, ...loop.missing];
        const rec: TripDistanceRecord = {
          id: prev ? prev.id : deps.genId('dist'), // คิดใหม่ = ทับ id เดิม (ไม่ให้ค้าง 2 record)
          tripId: t.id, branchId: t.branchId, branch: BRANCH_ID_NAME[t.branchId] || '', destKey,
          loopKm: Math.round(loop.totalKm * 10) / 10, storeCount: tripStoreCount(t), // จุดลงของจริง (ไม่ใช่ distinct อำเภอ)
          order: loop.order.map((o) => o.district), missing: missing.length ? missing : undefined,
          computedAt: new Date().toISOString(),
        };
        if (prev) { const i = db.tripDistances.findIndex((d: TripDistanceRecord) => d.id === prev.id); if (i >= 0) db.tripDistances[i] = rec; else db.tripDistances.push(rec); }
        else db.tripDistances.push(rec);
        cached.set(t.id, rec);
        await deps.saveRecord('tripDistances', rec);
        computed++;
      }
      res.json({ success: true, plateNo, computed, skipped, total: trips.length });
    } catch (err: any) { res.status(502).json({ error: `คิดระยะไม่สำเร็จ: ${err.message}` }); }
  });

  // คำนวณค่าจ้างรถร่วม (BAND) — ทดสอบสูตรผ่าน API ได้
  app.post('/api/experimental/fuel-trip/calculate', (req: Request, res: Response) => {
    try {
      const raw = req.body?.input as Record<string, unknown> | undefined;
      // ตรวจ field บังคับ + แปลงเป็น number (กันส่ง string มาแล้ว "85"+"113" ต่อสตริงตอนบวก)
      const required = ['distanceMinKm', 'distanceMaxKm', 'storeCount', 'appliedFuelPrice', 'actualBoxes'] as const;
      const optional = ['mountainAllowance', 'routeSurcharge', 'localTaxAdjustment'] as const;
      const missing = !raw ? [...required] : required.filter((k) => !Number.isFinite(Number(raw[k])));
      if (missing.length) {
        return res.status(400).json({ error: `input ไม่ครบ/ไม่ใช่ตัวเลข: ${missing.join(', ')}` });
      }
      const input = {} as FuelTripInput;
      for (const k of required) (input as any)[k] = Number(raw![k]);            // coerce -> number เสมอ
      for (const k of optional) if (raw![k] != null) (input as any)[k] = Number(raw![k]);
      const policy = { ...DEFAULT_FUEL_POLICY, ...(req.body?.policy || {}) } as FuelTripPolicy;
      res.json({ breakdown: computeFuelTrip(input, policy), policyUsed: policy });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });
}
