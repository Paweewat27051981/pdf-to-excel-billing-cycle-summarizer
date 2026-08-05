// ============================================================================
// [ทดลอง] route กลุ่ม /api/experimental/* — แยก 100% ไม่แตะ business logic เดิม
// ถ้าเลิกทดลอง: ลบไฟล์นี้ + เอา registerExperimentalRoutes ออกจาก server.ts
// ============================================================================
import type { Express, Request, Response } from 'express';
import { fetchCurrentOilPrice, fetchProvincialOilPrice, fetchHistoricalProvincialOilPrice, pickDiesel, summarizeProvinceDiesel, fetchBranchDiesel, BRANCH_OIL_CONFIGS, BranchOilConfig } from './src/experimental/orOilPrice.js';
import { computeFuelTrip, DEFAULT_FUEL_POLICY, FuelTripInput, FuelTripPolicy } from './src/experimental/fuelTripCalc.js';
import type { OilPriceRecord } from './src/types.js';

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
