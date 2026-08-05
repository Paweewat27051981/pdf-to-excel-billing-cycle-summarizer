// ============================================================================
// [ทดลอง] route กลุ่ม /api/experimental/* — แยก 100% ไม่แตะ business logic เดิม
// ถ้าเลิกทดลอง: ลบไฟล์นี้ + เอา registerExperimentalRoutes ออกจาก server.ts
// ============================================================================
import type { Express, Request, Response } from 'express';
import { fetchCurrentOilPrice, fetchProvincialOilPrice, fetchHistoricalProvincialOilPrice, pickDiesel, summarizeProvinceDiesel, fetchBranchDiesel, BRANCH_OIL_CONFIGS, BranchOilConfig } from './src/experimental/orOilPrice.js';
import { computeFuelTrip, DEFAULT_FUEL_POLICY, FuelTripInput, FuelTripPolicy, computeLoopTripCost } from './src/experimental/fuelTripCalc.js';
import { geocodeDistrict, nearestNeighborLoop, LoopStop } from './src/experimental/dohDistance.js';
import { parseServiceAreas, inServiceArea } from './src/serviceArea.js';
import type { OilPriceRecord, TripDistanceRecord, MountainRouteRecord, FuelPolicyRecord } from './src/types.js';

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
  deps: { getDb: () => Promise<any>; saveRecord: (coll: any, rec: { id: string }) => Promise<void>; removeRecord: (coll: any, id: string) => Promise<void>; flushCollection: (coll: any) => Promise<void>; genId: (p: string) => string },
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

  // ===== ค่าตั้งสูตร (fuelPolicy — ชุดเดียวทั้งบริษัท) =====
  const POLICY_KEYS: (keyof FuelTripPolicy)[] = [
    'nearSpeedKmh', 'farSpeedKmh', 'speedThresholdKm', 'fuelEfficiencyKmPerL',
    'driverHourlyRate', 'unloadingMinutesPerStore', 'baseBoxes', 'lumpSumBoxThreshold', 'finalRoundingDecimals',
  ];
  // policy ปัจจุบัน = ที่ตั้งไว้ (เติม default field ที่ขาด) หรือ default ทั้งชุด
  const effectivePolicy = (db: any): FuelTripPolicy => ({ ...DEFAULT_FUEL_POLICY, ...(db.fuelPolicy || {}) });
  // signature ของ policy (ใส่ใน fingerprint cache ระยะ — แก้ policy แล้ว report คิดใหม่)
  const policySig = (p: FuelTripPolicy): string => POLICY_KEYS.map((k) => p[k]).join(',');

  app.get('/api/experimental/fuel-policy', async (_req: Request, res: Response) => {
    try {
      const db = await deps.getDb();
      res.json({ policy: effectivePolicy(db), isDefault: !db.fuelPolicy, defaults: DEFAULT_FUEL_POLICY });
    } catch (err: any) { res.status(500).json({ error: err.message }); }
  });
  // PUT ตั้งค่า (validate ทุกค่าเป็นตัวเลข > 0; speedThreshold/finalRoundingDecimals >= 0)
  app.put('/api/experimental/fuel-policy', async (req: Request, res: Response) => {
    try {
      const b = (req.body?.policy || {}) as Record<string, unknown>;
      const out = {} as FuelPolicyRecord;
      for (const k of POLICY_KEYS) {
        const v = Number(b[k]);
        if (!Number.isFinite(v)) return res.status(400).json({ error: `ค่า ${k} ต้องเป็นตัวเลข` });
        // อนุญาต 0 เฉพาะ finalRoundingDecimals (ปัดจำนวนเต็ม); ที่เหลือต้อง > 0
        if (v < 0 || (v === 0 && k !== 'finalRoundingDecimals')) return res.status(400).json({ error: `ค่า ${k} ต้อง > 0` });
        (out as any)[k] = v;
      }
      out.updatedAt = new Date().toISOString();
      const db = await deps.getDb();
      db.fuelPolicy = out;
      await deps.flushCollection('fuelPolicy'); // เขียนเฉพาะ node นี้ (เล็ก ไม่ hang)
      res.json({ success: true, policy: { ...DEFAULT_FUEL_POLICY, ...out } });
    } catch (err: any) { res.status(500).json({ error: err.message }); }
  });

  // ===== master น้ำมันขึ้นเขา (ลิตรเพิ่มต่อปลายทาง) =====
  const norm = (s: string) => (s || '').trim();
  // GET รายการทั้งหมด
  app.get('/api/experimental/mountain-routes', async (_req: Request, res: Response) => {
    try {
      const db = await deps.getDb();
      res.json({ routes: (db.mountainRoutes || []) as MountainRouteRecord[] });
    } catch (err: any) { res.status(500).json({ error: err.message }); }
  });
  // POST เพิ่ม (จังหวัดบังคับ, อำเภอว่าง=ทั้งจังหวัด, extraLiters > 0)
  app.post('/api/experimental/mountain-routes', async (req: Request, res: Response) => {
    try {
      const b = req.body as { province?: string; district?: string; extraLiters?: unknown; note?: string };
      const province = norm(b.province || '');
      const district = norm(b.district || '');
      const extraLiters = Number(b.extraLiters);
      if (!province) return res.status(400).json({ error: 'ต้องระบุจังหวัด' });
      if (!Number.isFinite(extraLiters) || extraLiters <= 0) return res.status(400).json({ error: 'ลิตรเพิ่มต้องเป็นตัวเลข > 0' });
      const db = await deps.getDb();
      if (!Array.isArray(db.mountainRoutes)) db.mountainRoutes = [];
      // กันซ้ำ: จังหวัด+อำเภอเดียวกัน (case-insensitive trim)
      const dup = (db.mountainRoutes as MountainRouteRecord[]).find((r) => norm(r.province) === province && norm(r.district || '') === district);
      if (dup) return res.status(409).json({ error: `มีปลายทางนี้แล้ว: ${province}${district ? ' / ' + district : ' (ทั้งจังหวัด)'}` });
      const rec: MountainRouteRecord = {
        id: deps.genId('mtn'), province, district: district || undefined,
        extraLiters, note: norm(b.note || '') || undefined, createdAt: new Date().toISOString(),
      };
      db.mountainRoutes.push(rec);
      await deps.saveRecord('mountainRoutes', rec);
      res.json({ success: true, route: rec });
    } catch (err: any) { res.status(500).json({ error: err.message }); }
  });
  // DELETE
  app.delete('/api/experimental/mountain-routes/:id', async (req: Request, res: Response) => {
    try {
      const db = await deps.getDb();
      const idx = (db.mountainRoutes || []).findIndex((r: MountainRouteRecord) => r.id === req.params.id);
      if (idx < 0) return res.status(404).json({ error: 'ไม่พบรายการ' });
      db.mountainRoutes.splice(idx, 1);
      await deps.removeRecord('mountainRoutes', req.params.id);
      res.json({ success: true });
    } catch (err: any) { res.status(500).json({ error: err.message }); }
  });

  // รายการจังหวัด -> อำเภอ ที่มี "จริง" ในใบกระจาย (สำหรับ dropdown master ขึ้นเขา — กันพิมพ์ผิด/ต้อง match เป๊ะ)
  app.get('/api/experimental/dest-options', async (_req: Request, res: Response) => {
    try {
      const db = await deps.getDb();
      const map = new Map<string, Set<string>>(); // province -> set(district)
      for (const t of (db.tripDocuments || [])) {
        for (const r of (t.receipts || [])) {
          const province = norm(r.provinceRaw || ''); const district = norm(r.districtRaw || '');
          if (!province) continue;
          if (!map.has(province)) map.set(province, new Set());
          if (district) map.get(province)!.add(district);
        }
      }
      const provinces = [...map.entries()]
        .map(([province, ds]) => ({ province, districts: [...ds].sort((a, b) => a.localeCompare(b, 'th')) }))
        .sort((a, b) => a.province.localeCompare(b.province, 'th'));
      res.json({ provinces });
    } catch (err: any) { res.status(500).json({ error: err.message }); }
  });

  // รวมลิตรน้ำมันขึ้นเขาของใบ = ผลรวม extraLiters ของทุกปลายทางที่ match (ต่อ 1 ปลายทาง เลือก route "เจาะจงที่สุด")
  // เจาะจง: อำเภอตรง > ทั้งจังหวัด (กันเพิ่ม route ทั้งจังหวัดก่อนแล้วบัง route อำเภอเฉพาะ)
  const mountainLitersFor = (dests: { district: string; province: string }[], routes: MountainRouteRecord[]): number => {
    let total = 0;
    for (const d of dests) {
      const dDist = norm(d.district), dProv = norm(d.province);
      let districtMatch: number | null = null, provinceMatch: number | null = null;
      for (const r of routes) {
        if (norm(r.province) !== dProv) continue;
        const rDist = norm(r.district || '');
        if (rDist && rDist === dDist) districtMatch = (districtMatch ?? 0) + (r.extraLiters || 0); // อำเภอเฉพาะ (รวมถ้ามีหลาย row)
        else if (rDist === '') provinceMatch = (provinceMatch ?? 0) + (r.extraLiters || 0);        // ทั้งจังหวัด
      }
      total += districtMatch ?? provinceMatch ?? 0; // มีอำเภอเฉพาะ = ใช้อันนั้น (บัง province default)
    }
    return total;
  };

  // ===== รายงานเทียบต้นทุน (DOH ระยะจริง) =====
  // ใช้ "พื้นที่ให้บริการต่อสาขา" (serviceAreaText ที่ HQ ตั้งไว้) กรองปลายทางที่ปนมาผิด (เช่น กทม.ในใบพิษณุโลก)
  // อ่านพื้นที่สดทุกครั้ง (ไม่ cache — HQ แก้พื้นที่แล้วมีผลทันที)
  const getAreas = (db: any, branchId: string) =>
    parseServiceAreas(((db.branches || []).find((x: any) => x.id === branchId)?.serviceAreaText) || '');
  // ใบรับนี้มีราคา (สาขาวิ่งจริง) — ตรงกับ hasRate ใน App.tsx
  const hasRate = (r: any): boolean =>
    r.flatPrice != null || r.piecePrice != null || r.collectPrice != null || r.collectFlatPrice != null || r.peatPrice != null;
  // ปลายทางนี้ต้อง "ทิ้ง" ไหม = นอกพื้นที่สาขา และ ไม่มีราคา (ปลายทางมีราคา=วิ่งจริง เก็บไว้ ตรงกับ save path)
  const dropReceipt = (areas: any, r: any): boolean =>
    !hasRate(r) && !inServiceArea(areas, (r.provinceRaw || '').trim(), (r.districtRaw || '').trim());
  // helper: อำเภอ distinct ของใบ (ทิ้งปลายทางนอกพื้นที่ที่ไม่มีราคา) — nearest-neighbor เรียงเอง
  const tripDests = (t: any, db: any): { district: string; province: string }[] => {
    const areas = getAreas(db, t.branchId);
    const seen = new Set<string>(); const out: { district: string; province: string }[] = [];
    for (const r of (t.receipts || [])) {
      const district = (r.districtRaw || '').trim(); const province = (r.provinceRaw || '').trim();
      if (!district && !province) continue;
      if (dropReceipt(areas, r)) continue; // ปลายทางนอกพื้นที่ + ไม่มีราคา = ข้อมูลผิด ข้าม
      const k = `${district}|${province}`;
      if (!seen.has(k)) { seen.add(k); out.push({ district, province }); }
    }
    return out;
  };
  // จำนวน "จุดลงของ" จริง (ใบรับที่มีปลายทาง หลังทิ้งของผิด) — ใช้คิดเวลาขนของ
  // จำนวน "จุดลงของ" = ผู้รับ distinct ต่อปลายทาง (ผู้รับ+อำเภอ+จังหวัด)
  // ร้านเดียวหลายใบรับ = แวะจุดเดียว (ไม่นับใบรับดิบ ไม่งั้นเบี้ยขับพองผิด เช่น รัตนภัณฑ์ 6 ใบ = 1 จุด)
  const tripStoreCount = (t: any, db: any): number => {
    const areas = getAreas(db, t.branchId);
    const seen = new Set<string>();
    for (const r of (t.receipts || [])) {
      const district = (r.districtRaw || '').trim(); const province = (r.provinceRaw || '').trim();
      if (!district && !province) continue;
      if (dropReceipt(areas, r)) continue;
      seen.add(`${(r.receiverName || '').trim()}|${district}|${province}`);
    }
    return seen.size;
  };
  // fingerprint ของ input ที่ใช้คิด (ปลายทาง + จำนวนจุดลง + ลิตรขึ้นเขา) — recalculate/แก้ใบรับ/แก้ master ขึ้นเขา เปลี่ยนค่านี้ = cache stale คิดใหม่
  // NOTE: ต่อ #m<n> เฉพาะเมื่อมีลิตรขึ้นเขา (>0) — คง backward-compat กับ cache เดิม (m0 = รูปแบบเดิม ไม่ให้ stale ทั้งกระดาน)
  const tripFingerprint = (t: any, routes: MountainRouteRecord[], db: any): string => {
    const dests = tripDests(t, db);
    const dk = dests.map((d) => `${d.district}|${d.province}`).sort().join(';');
    const mtn = mountainLitersFor(dests, routes);
    return `${dk}#${tripStoreCount(t, db)}${mtn > 0 ? `#m${mtn}` : ''}`;
  };

  // GET รายงานต่อทะเบียน (อ่าน cache ที่คิดไว้แล้ว — ไม่ยิง DOH) ?cycleId=&branchId=
  app.get('/api/experimental/fuel-report', async (req: Request, res: Response) => {
    try {
      const cycleId = typeof req.query.cycleId === 'string' ? req.query.cycleId : '';
      const branchId = typeof req.query.branchId === 'string' ? req.query.branchId : '';
      if (!cycleId) return res.status(400).json({ error: 'ต้องระบุ cycleId' });
      const db = await deps.getDb();
      const policy = effectivePolicy(db); // ค่าตั้งสูตรปัจจุบัน (แก้แล้วมีผลทันทีกับต้นทุน — ระยะ cache เดิมใช้ได้)
      const routes = (db.mountainRoutes || []) as MountainRouteRecord[];
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
        // นับ computed ต่อเมื่อ: มีระยะ + ครบทุกจุด + fingerprint ตรงปัจจุบัน (ปลายทาง/จุดลง/ลิตรขึ้นเขาไม่เปลี่ยน)
        if (dist && dist.loopKm > 0 && !(dist.missing && dist.missing.length) && dist.destKey === tripFingerprint(t, routes, db)) {
          const branchName = BRANCH_ID_NAME[t.branchId] || '';
          const oil = oilByBranch.get(branchName) ?? 0;
          // ลิตรขึ้นเขาใช้ค่าที่ cache ตอนคิด (fingerprint คุมความ fresh แล้ว) — คูณราคาน้ำมันปัจจุบัน
          if (oil > 0) { g.fuelCostBand += computeLoopTripCost(dist.loopKm, dist.storeCount, oil, policy, dist.mountainLiters || 0).totalCost; g.computed++; }
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

  // GET รายละเอียดต่อใบ ของ 1 ทะเบียน (สำหรับแถวขยาย) ?cycleId=&plateNo=&branchId=
  app.get('/api/experimental/fuel-report/plate-detail', async (req: Request, res: Response) => {
    try {
      const cycleId = typeof req.query.cycleId === 'string' ? req.query.cycleId : '';
      const plateNo = typeof req.query.plateNo === 'string' ? req.query.plateNo : '';
      const branchId = typeof req.query.branchId === 'string' ? req.query.branchId : '';
      if (!cycleId || !plateNo) return res.status(400).json({ error: 'ต้องระบุ cycleId, plateNo' });
      const db = await deps.getDb();
      const policy = effectivePolicy(db);
      const routes = (db.mountainRoutes || []) as MountainRouteRecord[];
      const distById = new Map<string, TripDistanceRecord>((db.tripDistances || []).map((d: TripDistanceRecord) => [d.tripId, d]));
      // ราคาน้ำมันล่าสุดต่อสาขา
      const oilLatest = new Map<string, OilPriceRecord>();
      for (const o of (db.oilPrices || []) as OilPriceRecord[]) {
        const cur = oilLatest.get(o.branch);
        if (!cur || (o.priceDate || '') > (cur.priceDate || '') || ((o.priceDate || '') === (cur.priceDate || '') && (o.fetchedAt || '') > (cur.fetchedAt || ''))) oilLatest.set(o.branch, o);
      }
      const trips = (db.tripDocuments || []).filter((t: any) => t.cycleId === cycleId && t.plateNo === plateNo && (!branchId || t.branchId === branchId));
      const items = trips.map((t: any) => {
        const dist = distById.get(t.id);
        const dests = tripDests(t, db);
        const branchName = BRANCH_ID_NAME[t.branchId] || '';
        const oil = oilLatest.get(branchName)?.price ?? 0;
        const fresh = !!dist && dist.loopKm > 0 && !(dist.missing && dist.missing.length) && dist.destKey === tripFingerprint(t, routes, db);
        const cost = fresh && oil > 0 ? computeLoopTripCost(dist!.loopKm, dist!.storeCount, oil, policy, dist!.mountainLiters || 0) : null;
        return {
          documentNo: t.documentNo, tripAmount: t.tripAmount || 0,
          dests: dests.map((d) => `${d.district}|${d.province}`),
          loopKm: dist?.loopKm ?? null, order: dist?.order ?? [], storeCount: dist?.storeCount ?? 0,
          mountainLiters: dist?.mountainLiters ?? 0, missing: dist?.missing ?? [],
          oil, fresh,
          driverAllowance: cost ? Math.round(cost.driverAllowance * 100) / 100 : null,
          fuelCost: cost ? Math.round(cost.fuelCost * 100) / 100 : null,
          mountainFuelCost: cost ? Math.round(cost.mountainFuelCost * 100) / 100 : null,
          totalCost: cost ? cost.totalCost : null,
          diff: cost ? Math.round(((t.tripAmount || 0) - cost.totalCost) * 100) / 100 : null,
        };
      });
      res.json({ plateNo, branchId, count: items.length, items, policy });
    } catch (err: any) { res.status(500).json({ error: err.message }); }
  });

  // POST คิดระยะลูป+ต้นทุน ของ "1 ทะเบียน" (ยิง DOH เฉพาะใบที่ยังไม่ cache) — ผู้ใช้กดทีละคัน
  app.post('/api/experimental/fuel-report/compute-plate', async (req: Request, res: Response) => {
    try {
      const { cycleId, plateNo, branchId } = req.body as { cycleId: string; plateNo: string; branchId?: string };
      if (!cycleId || !plateNo) return res.status(400).json({ error: 'ต้องระบุ cycleId, plateNo' });
      const db = await deps.getDb();
      if (!Array.isArray(db.tripDistances)) db.tripDistances = [];
      const routes = (db.mountainRoutes || []) as MountainRouteRecord[];
      const cached = new Map<string, TripDistanceRecord>((db.tripDistances as TripDistanceRecord[]).map((d) => [d.tripId, d]));
      // plateNo ไม่ unique ข้ามสาขา -> filter branchId ด้วย (ตรงกับ /fuel-report) กันไปคิดใบสาขาอื่น
      const trips = (db.tripDocuments || []).filter((t: any) => t.cycleId === cycleId && t.plateNo === plateNo && (!branchId || t.branchId === branchId));
      let computed = 0, skipped = 0;
      for (const t of trips) {
        const dests = tripDests(t, db);
        const destKey = tripFingerprint(t, routes, db); // ปลายทาง + จำนวนจุดลง + ลิตรขึ้นเขา
        const prev = cached.get(t.id);
        // skip เฉพาะ cache สมบูรณ์ (ไม่มี missing) + fingerprint ไม่เปลี่ยน — ล่มชั่วคราว/recalculate/แก้ master ขึ้นเขา จะคิดใหม่
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
          loopKm: Math.round(loop.totalKm * 10) / 10, storeCount: tripStoreCount(t, db), // จุดลงของจริง (ไม่ใช่ distinct อำเภอ)
          order: loop.order.map((o) => o.district), missing: missing.length ? missing : undefined,
          mountainLiters: mountainLitersFor(dests, routes) || undefined, // ลิตรขึ้นเขารวม (match ปลายทางตอนคิด)
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
  app.post('/api/experimental/fuel-trip/calculate', async (req: Request, res: Response) => {
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
      // policy: ที่ตั้งไว้ใน db (ค่าตั้งสูตรทั้งบริษัท) เป็นฐาน + ทับด้วย policy ที่ส่งมา (ถ้ามี — สำหรับทดลอง)
      const db = await deps.getDb();
      const policy = { ...effectivePolicy(db), ...(req.body?.policy || {}) } as FuelTripPolicy;
      res.json({ breakdown: computeFuelTrip(input, policy), policyUsed: policy });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });
}
