import express from 'express';
import compression from 'compression';
import path from 'path';
import { GoogleGenAI, Type } from '@google/genai';
import dotenv from 'dotenv';
import { getDb, saveDb, saveRecord, saveRecords, removeRecord, removeRecords, flushCollection, isIdKeyed } from './server-db.js';
import {
  DatabaseState,
  BillingCycle,
  Branch,
  Vehicle,
  RateMaster,
  RateOverride,
  ReceiverGroup,
  ReceiverGroupAlias,
  ProductConversionRule,
  ManualBoxSender,
  DestinationOverride,
  MoneyCategory,
  TripDocument,
  FuelEntry,
  DeductionEntry,
  ExtractedTripDocument,
} from './src/types.js';
import { computeTripDocument, normPlate } from './src/calc.js';
import { parseDistributionExcel, parseRateExcel, parseFuelExcel } from './excel-import.js';

dotenv.config(); // โหลด .env
dotenv.config({ path: '.env.local', override: true }); // และ .env.local (ทับค่าเดิม)

function isAiEnabled(): boolean {
  const key = process.env.GEMINI_API_KEY;
  return !!key && key !== 'MY_GEMINI_API_KEY';
}

function generateId(prefix: string): string {
  return `${prefix}-${Math.random().toString(36).substring(2, 11)}`;
}

let aiClient: GoogleGenAI | null = null;
function getGeminiClient(): GoogleGenAI {
  if (!aiClient) {
    const key = process.env.GEMINI_API_KEY;
    if (!key) throw new Error('Missing GEMINI_API_KEY environment variable.');
    aiClient = new GoogleGenAI({
      apiKey: key,
      httpOptions: { headers: { 'User-Agent': 'aistudio-build' } },
    });
  }
  return aiClient;
}

// แปลง cycle เป็น context สำหรับ calc
function cycleCtx(cycle: BillingCycle) {
  return { year: cycle.year, month: cycle.month, half: cycle.half };
}

// คำนวณ trip ใหม่จาก raw extracted + master ปัจจุบัน
function recomputeTrip(
  db: DatabaseState,
  cycle: BillingCycle,
  extracted: ExtractedTripDocument,
  fileName: string,
  branchId: string
): TripDocument {
  // ราคาเฉพาะรอบ -> map (rateMasterId -> {price, threshold}) ของรอบ+สาขานี้
  const overrides = new Map<string, { price: number; pieceThreshold: number | null }>();
  db.rateOverrides
    .filter((o) => o.branchId === branchId && o.cycleId === cycle.id)
    .forEach((o) => overrides.set(o.rateMasterId, { price: o.price, pieceThreshold: o.pieceThreshold ?? null }));

  const branch = db.branches.find((b) => b.id === branchId);
  const branchVehicles = db.vehicles.filter((v) => v.branchId === branchId);
  // กลุ่มราคาของรถคันนี้ (จากทะเบียน) -> เลือกราคาของกลุ่มนั้น + ขั้นต่ำของกลุ่ม
  const vehicle = branchVehicles.find((v) => normPlate(v.plateNo) === normPlate(extracted.plateNo) && v.status === 'active');
  const group = vehicle?.rateGroup || '';
  const groupMin = branch?.rateGroups?.find((g) => g.name === group)?.minBoxes;
  const minBoxes = groupMin !== undefined ? groupMin : (branch?.minBoxes ?? null);
  // ราคาที่ใช้: ของกลุ่มเดียวกัน หรือไม่ระบุกลุ่ม (ใช้ร่วมทุกกลุ่ม)
  const branchRates = db.rateMasters.filter((r) => r.branchId === branchId && (!r.rateGroup || r.rateGroup === group));

  // ใช้รถ/ราคา/กฎของสาขานั้น + ราคาเฉพาะรอบ (ถ้ามี)
  const trip = computeTripDocument(
    extracted,
    {
      cycleId: cycle.id,
      cycle: cycleCtx(cycle),
      vehicles: branchVehicles,
      rates: branchRates,
      rateOverrides: overrides,
      groups: db.receiverGroups.filter((g) => g.branchId === branchId),
      aliases: db.receiverGroupAliases.filter((a) => a.branchId === branchId),
      rules: db.conversionRules.filter((r) => r.branchId === branchId),
      manualBoxSenders: db.manualBoxSenders.filter((m) => m.branchId === branchId),
      destOverrides: db.destinationOverrides.filter((d) => d.branchId === branchId),
      minBoxes,
      collectBackHalfPiece: branch?.collectBackHalfPiece,
      fileName,
    },
    () => generateId('rcp')
  );
  trip.branchId = branchId;
  return trip;
}

// สร้าง object รอบจาก ปี/เดือน/ครึ่งเดือน (ใช้ทั้งตอนสร้างเอง และเปิดรอบอัตโนมัติ)
function makeCycle(year: number, month: number, half: 'first' | 'second'): BillingCycle {
  const lastDay = new Date(year, month, 0).getDate();
  const startDate = half === 'first' ? `${year}-${String(month).padStart(2, '0')}-01` : `${year}-${String(month).padStart(2, '0')}-16`;
  const endDate = half === 'first' ? `${year}-${String(month).padStart(2, '0')}-15` : `${year}-${String(month).padStart(2, '0')}-${lastDay}`;
  const thaiMonth = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'][month - 1];
  return {
    id: generateId('cycle'),
    name: `${thaiMonth} ${(year + 543) % 100} รอบ ${half === 'first' ? '1-15' : '16-31'}`,
    year, month, half, startDate, endDate, status: 'open', createdAt: new Date().toISOString(),
  };
}

// หา/สร้างรอบจาก "วันที่ในใบ" (YYYY-MM-DD) — เปิดรอบอัตโนมัติตามรอบ 1-15 / 16-สิ้นเดือน
// persist=true จะสร้าง+เพิ่มลง db (ตอนบันทึกจริง), false จะสร้าง object ลอย ๆ (ตอน preview)
function resolveCycleForDate(
  db: DatabaseState,
  dateStr: string,
  persist: boolean
): { cycle: BillingCycle | null; created: boolean; closed: boolean; invalid: boolean } {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec((dateStr || '').trim());
  if (!m) return { cycle: null, created: false, closed: false, invalid: true };
  const year = +m[1], month = +m[2], day = +m[3];
  if (month < 1 || month > 12 || day < 1 || day > 31) return { cycle: null, created: false, closed: false, invalid: true };
  const half: 'first' | 'second' = day <= 15 ? 'first' : 'second';
  const existing = db.cycles.find((c) => c.year === year && c.month === month && c.half === half);
  if (existing) return { cycle: existing, created: false, closed: existing.status === 'closed', invalid: false };
  const fresh = makeCycle(year, month, half);
  if (persist) db.cycles.push(fresh);
  return { cycle: fresh, created: true, closed: false, invalid: false };
}

async function startServer() {
  const app = express();
  const PORT = Number(process.env.PORT) || 3000;
  // build marker: stripParen district match + แม่กลอง 7.5 (force fresh instance)
  // 📉 บีบอัด (gzip) ทุก response -> ลดแบนด์วิดท์ ~70% (ไฟล์ JS/CSS/JSON เล็กลงมาก)
  app.use(compression());
  app.use(express.json({ limit: '50mb' }));
  app.use(express.urlencoded({ limit: '50mb', extended: true }));

  // ===================== CONFIG =====================
  app.get('/api/config', (_req, res) => {
    res.json({ aiEnabled: isAiEnabled(), storage: 'granular-v2' });
  });

  // ===================== SETTINGS =====================
  app.put('/api/settings', async (req, res) => {
    try {
      const db = await getDb();
      db.settings = { ...db.settings, ...req.body };
      await saveDb(db);
      res.json(db.settings);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ===================== BRANCH LOGIN =====================
  // ตรวจรหัสผ่านสาขา -> คืนข้อมูลสาขา (ไม่คืน password)
  app.post('/api/branch-login', async (req, res) => {
    try {
      const { branchId, password } = req.body as { branchId: string; password: string };
      const db = await getDb();
      const branch = db.branches.find((b) => b.id === branchId && b.status === 'active');
      if (!branch) return res.status(404).json({ error: 'ไม่พบสาขา' });
      if (String(branch.password) !== String(password)) {
        return res.status(401).json({ error: 'รหัสผ่านไม่ถูกต้อง' });
      }
      res.json({ ok: true, branch: { id: branch.id, name: branch.name, isHQ: !!branch.isHQ } });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ===================== CLONE MASTER ระหว่างสาขา =====================
  // คัดลอก กฎตัวหาร/กลุ่มผู้รับ/alias/ประเภทเงิน/ผู้ส่งกล่อง จากสาขาต้นแบบ -> สาขาปลายทาง
  app.post('/api/branches/clone', async (req, res) => {
    try {
      const { sourceBranchId, targetBranchId, replace } = req.body as {
        sourceBranchId: string; targetBranchId: string; replace?: boolean;
      };
      if (!sourceBranchId || !targetBranchId || sourceBranchId === targetBranchId) {
        return res.status(400).json({ error: 'เลือกสาขาต้นแบบและปลายทางให้ถูกต้อง' });
      }
      const db = await getDb();

      // ลบของเดิมในสาขาปลายทางก่อน (ถ้าเลือกแทนที่)
      if (replace) {
        db.receiverGroups = db.receiverGroups.filter((x) => x.branchId !== targetBranchId);
        db.receiverGroupAliases = db.receiverGroupAliases.filter((x) => x.branchId !== targetBranchId);
        db.conversionRules = db.conversionRules.filter((x) => x.branchId !== targetBranchId);
        db.moneyCategories = db.moneyCategories.filter((x) => x.branchId !== targetBranchId);
        db.manualBoxSenders = db.manualBoxSenders.filter((x) => x.branchId !== targetBranchId);
      }

      // กลุ่มผู้รับ: id ใหม่ + จำ map เดิม->ใหม่ (เพื่อ remap alias/rule)
      const groupIdMap: Record<string, string> = {};
      for (const g of db.receiverGroups.filter((x) => x.branchId === sourceBranchId)) {
        const nid = generateId('grp');
        groupIdMap[g.id] = nid;
        db.receiverGroups.push({ ...g, id: nid, branchId: targetBranchId });
      }
      for (const a of db.receiverGroupAliases.filter((x) => x.branchId === sourceBranchId)) {
        db.receiverGroupAliases.push({ ...a, id: generateId('al'), branchId: targetBranchId, receiverGroupId: groupIdMap[a.receiverGroupId] || '' });
      }
      for (const r of db.conversionRules.filter((x) => x.branchId === sourceBranchId)) {
        db.conversionRules.push({ ...r, id: generateId('rule'), branchId: targetBranchId, receiverGroupId: r.receiverGroupId ? (groupIdMap[r.receiverGroupId] || '') : '' });
      }
      for (const c of db.moneyCategories.filter((x) => x.branchId === sourceBranchId)) {
        db.moneyCategories.push({ ...c, id: generateId('cat'), branchId: targetBranchId });
      }
      for (const m of db.manualBoxSenders.filter((x) => x.branchId === sourceBranchId)) {
        db.manualBoxSenders.push({ ...m, id: generateId('mbs'), branchId: targetBranchId });
      }

      await saveDb(db);
      res.json({ ok: true });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ===================== STATE =====================
  // ?branchId=xxx -> กรองข้อมูลเฉพาะสาขานั้น (ถ้าไม่ส่ง = HQ เห็นทุกสาขา)
  // ตัด password ของสาขาออกเสมอ
  app.get('/api/state', async (req, res) => {
    try {
      const db = await getDb();
      const branchId = typeof req.query.branchId === 'string' ? req.query.branchId : '';
      const inBranch = <T extends { branchId?: string }>(arr: T[]) =>
        branchId ? arr.filter((x) => x.branchId === branchId) : arr;
      const safe: DatabaseState = {
        ...db,
        branches: db.branches.map((b) => ({ ...b, password: '' })),
        vehicles: inBranch(db.vehicles),
        rateMasters: inBranch(db.rateMasters),
        rateOverrides: inBranch(db.rateOverrides),
        tripDocuments: inBranch(db.tripDocuments),
        fuelEntries: inBranch(db.fuelEntries),
        deductions: inBranch(db.deductions),
        receiverGroups: inBranch(db.receiverGroups),
        receiverGroupAliases: inBranch(db.receiverGroupAliases),
        conversionRules: inBranch(db.conversionRules),
        manualBoxSenders: inBranch(db.manualBoxSenders),
        destinationOverrides: inBranch(db.destinationOverrides),
        moneyCategories: inBranch(db.moneyCategories),
      };
      res.json(safe);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ===================== CYCLES =====================
  app.post('/api/cycles', async (req, res) => {
    try {
      const { year, month, half } = req.body as { year: number; month: number; half: 'first' | 'second' };
      if (!year || !month || !half) return res.status(400).json({ error: 'ต้องระบุ year, month, half' });

      const db = await getDb();
      const lastDay = new Date(year, month, 0).getDate();
      const startDate = half === 'first' ? `${year}-${String(month).padStart(2, '0')}-01` : `${year}-${String(month).padStart(2, '0')}-16`;
      const endDate = half === 'first' ? `${year}-${String(month).padStart(2, '0')}-15` : `${year}-${String(month).padStart(2, '0')}-${lastDay}`;
      const thaiMonth = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'][month - 1];

      const exists = db.cycles.find((c) => c.year === year && c.month === month && c.half === half);
      if (exists) return res.status(400).json({ error: 'มีรอบนี้อยู่แล้ว' });

      const newCycle: BillingCycle = {
        id: generateId('cycle'),
        name: `${thaiMonth} ${(year + 543) % 100} รอบ ${half === 'first' ? '1-15' : '16-31'}`,
        year, month, half, startDate, endDate, status: 'open', createdAt: new Date().toISOString(),
      };
      db.cycles.push(newCycle);
      await saveDb(db);
      res.status(201).json(newCycle);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.put('/api/cycles/:id', async (req, res) => {
    try {
      const db = await getDb();
      const idx = db.cycles.findIndex((c) => c.id === req.params.id);
      if (idx === -1) return res.status(404).json({ error: 'ไม่พบรอบ' });
      const { status } = req.body;
      if (status) db.cycles[idx].status = status;
      await saveDb(db);
      res.json(db.cycles[idx]);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ลบรอบ — กันลบถ้ามีข้อมูลอ้างอิง (ใบกระจาย/น้ำมัน/รายการหัก)
  app.delete('/api/cycles/:id', async (req, res) => {
    try {
      const db = await getDb();
      const id = req.params.id;
      const idx = db.cycles.findIndex((c) => c.id === id);
      if (idx === -1) return res.status(404).json({ error: 'ไม่พบรอบ' });
      const nTrip = db.tripDocuments.filter((t) => t.cycleId === id).length;
      const nFuel = db.fuelEntries.filter((f) => f.cycleId === id).length;
      const nDed = db.deductions.filter((d) => d.cycleId === id).length;
      if (nTrip || nFuel || nDed) {
        return res.status(409).json({ error: `รอบนี้มีข้อมูลอยู่ (ใบกระจาย ${nTrip} · น้ำมัน ${nFuel} · รายการหัก ${nDed}) — ลบข้อมูลในรอบก่อนจึงจะลบรอบได้` });
      }
      const name = db.cycles[idx].name;
      db.cycles.splice(idx, 1);
      await saveDb(db);
      res.json({ success: true, name });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ===================== Generic master CRUD helper =====================
  function masterRoutes<T extends { id: string }>(
    name: string,
    key: keyof DatabaseState,
    idPrefix: string
  ) {
    app.post(`/api/${name}`, async (req, res) => {
      try {
        const db = await getDb();
        const item = { ...req.body, id: generateId(idPrefix) } as T;
        (db[key] as unknown as T[]).push(item);
        if (isIdKeyed(key)) await saveRecord(key, item as any); else await saveDb(db);
        res.status(201).json(item);
      } catch (err: any) {
        res.status(500).json({ error: err.message });
      }
    });
    app.put(`/api/${name}/:id`, async (req, res) => {
      try {
        const db = await getDb();
        const arr = db[key] as unknown as T[];
        const idx = arr.findIndex((x) => x.id === req.params.id);
        if (idx === -1) return res.status(404).json({ error: 'ไม่พบรายการ' });
        arr[idx] = { ...arr[idx], ...req.body, id: req.params.id };
        if (isIdKeyed(key)) await saveRecord(key, arr[idx] as any); else await saveDb(db);
        res.json(arr[idx]);
      } catch (err: any) {
        res.status(500).json({ error: err.message });
      }
    });
    app.delete(`/api/${name}/:id`, async (req, res) => {
      try {
        const db = await getDb();
        (db[key] as unknown as T[]) = (db[key] as unknown as T[]).filter((x) => x.id !== req.params.id) as any;
        if (isIdKeyed(key)) await removeRecord(key, req.params.id); else await saveDb(db);
        res.json({ success: true });
      } catch (err: any) {
        res.status(500).json({ error: err.message });
      }
    });
    // ลบเป็นกลุ่ม (ติ๊กหลายรายการแล้วลบทีเดียว)
    app.post(`/api/${name}/bulk-delete`, async (req, res) => {
      try {
        const ids: string[] = req.body?.ids || [];
        const db = await getDb();
        const idset = new Set(ids);
        (db[key] as unknown as T[]) = (db[key] as unknown as T[]).filter((x) => !idset.has(x.id)) as any;
        if (isIdKeyed(key)) await removeRecords(key, ids); else await saveDb(db);
        res.json({ success: true, deleted: ids.length });
      } catch (err: any) {
        res.status(500).json({ error: err.message });
      }
    });
  }

  // 🔒 ค่าน้ำมัน: เลขใบสั่งเติมห้ามซ้ำในสาขา (ลงทะเบียน POST ก่อน masterRoutes เพื่อ override)
  app.post('/api/fuel', async (req, res) => {
    try {
      const db = await getDb();
      const body = req.body as FuelEntry;
      const refNo = (body.refNo || '').trim();
      if (refNo) {
        const dup = db.fuelEntries.find((f) => f.branchId === body.branchId && (f.refNo || '').trim() === refNo);
        if (dup) return res.status(409).json({ error: `เลขใบสั่งเติมน้ำมัน "${refNo}" ซ้ำ — มีอยู่แล้วในระบบ (ห้ามบันทึกซ้ำ)` });
      }
      const item = { ...body, id: generateId('fuel') } as FuelEntry;
      db.fuelEntries.push(item);
      await saveRecord('fuelEntries', item);
      res.status(201).json(item);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  masterRoutes<Branch>('branches', 'branches', 'br');
  masterRoutes<RateOverride>('rate-overrides', 'rateOverrides', 'rov');
  masterRoutes<MoneyCategory>('money-categories', 'moneyCategories', 'cat');
  masterRoutes<ManualBoxSender>('manual-box-senders', 'manualBoxSenders', 'mbs');
  masterRoutes<Vehicle>('vehicles', 'vehicles', 'veh');
  masterRoutes<ReceiverGroup>('receiver-groups', 'receiverGroups', 'grp');
  masterRoutes<ReceiverGroupAlias>('receiver-aliases', 'receiverGroupAliases', 'al');
  masterRoutes<ProductConversionRule>('conversion-rules', 'conversionRules', 'rule');
  masterRoutes<DestinationOverride>('destination-overrides', 'destinationOverrides', 'do');
  masterRoutes<FuelEntry>('fuel', 'fuelEntries', 'fuel');
  masterRoutes<DeductionEntry>('deductions', 'deductions', 'ded');

  // ===================== RATE MASTER (มีประวัติราคา) =====================
  app.post('/api/rate-masters', async (req, res) => {
    try {
      const db = await getDb();
      const item: RateMaster = {
        ...req.body,
        id: generateId('rate'),
        createdBy: req.body.createdBy || 'user',
        createdAt: new Date().toISOString(),
      };
      db.rateMasters.push(item);
      await saveDb(db);
      res.status(201).json(item);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // นำเข้า "ค่าน้ำมัน" จาก Excel -> FuelEntry (จัดเข้ารอบตามวันที่อัตโนมัติ)
  app.post('/api/import-fuel', async (req, res) => {
    try {
      const { branchId, fileBase64 } = req.body as { branchId: string; fileBase64: string };
      if (!branchId) return res.status(400).json({ error: 'ต้องระบุสาขา' });
      if (!fileBase64) return res.status(400).json({ error: 'ต้องส่งไฟล์ Excel' });
      const buffer = Buffer.from(fileBase64.replace(/^data:[^;]+;base64,/, ''), 'base64');
      const { fuel, summary } = parseFuelExcel(buffer);
      if (!fuel.length) return res.status(422).json({ error: 'อ่านไฟล์ไม่พบใบสั่งเติมน้ำมัน — ตรวจสอบหัวคอลัมน์ ทะเบียน/วันที่/จำนวนเงิน' });
      const db = await getDb();
      let created = 0, skippedDup = 0;
      const closedCycles = new Set<string>();
      const createdCycles = new Set<string>();
      const createdEntries: FuelEntry[] = [];
      // เลขใบสั่งเติมห้ามซ้ำในสาขา (เทียบกับของเดิม + ในไฟล์เดียวกัน)
      const seenRef = new Set(db.fuelEntries.filter((f) => f.branchId === branchId).map((f) => (f.refNo || '').trim()).filter(Boolean));
      for (const f of fuel) {
        const rv = resolveCycleForDate(db, f.date, true);
        if (rv.invalid || !rv.cycle) continue;
        if (rv.closed) { closedCycles.add(rv.cycle.name); continue; }
        const rn = (f.refNo || '').trim();
        if (rn && seenRef.has(rn)) { skippedDup++; continue; }
        if (rn) seenRef.add(rn);
        if (rv.created) createdCycles.add(rv.cycle.name);
        const entry: FuelEntry = { id: generateId('fuel'), branchId, cycleId: rv.cycle.id, plateNo: f.plateNo, refNo: f.refNo, date: f.date, amount: f.amount };
        db.fuelEntries.push(entry);
        createdEntries.push(entry);
        created++;
      }
      await saveRecords('fuelEntries', createdEntries);
      if (createdCycles.size) await flushCollection('cycles');
      if (createdCycles.size) summary.push(`เปิดรอบใหม่อัตโนมัติ: ${[...createdCycles].join(', ')}`);
      if (skippedDup) summary.push(`⚠️ ข้ามเลขใบสั่งเติมที่ซ้ำ ${skippedDup} รายการ`);
      if (closedCycles.size) summary.push(`⚠️ ข้ามรายการของรอบที่ปิดอยู่: ${[...closedCycles].join(', ')} (ให้ HQ เปิดรอบก่อน)`);
      res.status(201).json({ success: true, created, summary });
    } catch (err: any) {
      console.error('import-fuel error:', err);
      res.status(500).json({ error: `นำเข้าค่าน้ำมันไม่สำเร็จ: ${err.message}` });
    }
  });

  // นำเข้า "ตารางราคา" จาก Excel (2 ชีต เหมาคัน/รายชิ้น) -> สร้าง rate masters
  app.post('/api/import-rates', async (req, res) => {
    try {
      const { branchId, fileBase64, replaceExisting } = req.body as { branchId: string; fileBase64: string; replaceExisting?: boolean };
      if (!branchId) return res.status(400).json({ error: 'ต้องระบุสาขา' });
      if (!fileBase64) return res.status(400).json({ error: 'ต้องส่งไฟล์ Excel' });
      const buffer = Buffer.from(fileBase64.replace(/^data:[^;]+;base64,/, ''), 'base64');
      const { rates, summary } = parseRateExcel(buffer);
      if (!rates.length) return res.status(422).json({ error: 'อ่านไฟล์ไม่พบราคา — ตรวจสอบหัวคอลัมน์ จังหวัด/อำเภอ/ราคา' });
      const db = await getDb();
      let removed = 0;
      if (replaceExisting) {
        // ลบเฉพาะหมวดที่อยู่ในไฟล์นำเข้า (งานปกติ/เก็บคืน/Peat/บวกเพิ่ม) แล้วแทนที่ด้วยไฟล์
        const cats = new Set(rates.map((r) => r.productCategory || 'normal'));
        const before = db.rateMasters.length;
        db.rateMasters = db.rateMasters.filter((r) => !(r.branchId === branchId && cats.has(r.productCategory || 'normal')));
        removed = before - db.rateMasters.length;
      }
      const now = new Date().toISOString();
      for (const r of rates) db.rateMasters.push({ ...r, branchId, id: generateId('rate'), createdBy: 'import', createdAt: now } as RateMaster);
      await saveDb(db);
      res.status(201).json({ success: true, created: rates.length, removed, summary });
    } catch (err: any) {
      console.error('import-rates error:', err);
      res.status(500).json({ error: `นำเข้าราคาไม่สำเร็จ: ${err.message}` });
    }
  });

  // นำเข้าราคาแบบชุด (เขียน DB ครั้งเดียว) — ใช้ตอน import ตารางราคาทั้งสาขา
  app.post('/api/rate-masters/bulk-create', async (req, res) => {
    try {
      const { rates } = req.body as { rates: Partial<RateMaster>[] };
      if (!Array.isArray(rates) || !rates.length) return res.status(400).json({ error: 'ต้องส่ง rates เป็น array' });
      const db = await getDb();
      const now = new Date().toISOString();
      for (const r of rates) {
        db.rateMasters.push({
          productCategory: 'normal', effectiveFrom: '2020-01-01', effectiveTo: null, status: 'active',
          ...r, id: generateId('rate'), createdBy: r.createdBy || 'import', createdAt: now,
        } as RateMaster);
      }
      await saveDb(db);
      res.status(201).json({ success: true, count: rates.length });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.put('/api/rate-masters/:id', async (req, res) => {
    try {
      const db = await getDb();
      const idx = db.rateMasters.findIndex((r) => r.id === req.params.id);
      if (idx === -1) return res.status(404).json({ error: 'ไม่พบราคา' });
      const old = db.rateMasters[idx];
      // เก็บประวัติถ้าราคาเปลี่ยน
      if (typeof req.body.price === 'number' && req.body.price !== old.price) {
        db.rateMasterHistory.push({
          id: generateId('rhist'),
          rateMasterId: old.id,
          oldPrice: old.price,
          newPrice: req.body.price,
          changedBy: req.body.updatedBy || 'user',
          changedAt: new Date().toISOString(),
          changeReason: req.body.changeReason || 'แก้ไขราคา',
        });
      }
      db.rateMasters[idx] = { ...old, ...req.body, id: old.id, updatedAt: new Date().toISOString() };
      await saveDb(db);
      res.json(db.rateMasters[idx]);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.delete('/api/rate-masters/:id', async (req, res) => {
    try {
      const db = await getDb();
      db.rateMasters = db.rateMasters.filter((r) => r.id !== req.params.id);
      await saveDb(db);
      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/rate-masters/bulk-delete', async (req, res) => {
    try {
      const ids: string[] = req.body?.ids || [];
      const idset = new Set(ids);
      const db = await getDb();
      db.rateMasters = db.rateMasters.filter((r) => !idset.has(r.id));
      await saveDb(db);
      res.json({ success: true, deleted: ids.length });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ราคาเฉพาะรอบ: สร้างหรืออัปเดต (1 รอบ + 1 ราคาหลัก = 1 override)
  app.post('/api/rate-overrides/upsert', async (req, res) => {
    try {
      const { branchId, cycleId, rateMasterId, price, pieceThreshold } = req.body as RateOverride;
      if (!branchId || !cycleId || !rateMasterId) return res.status(400).json({ error: 'ข้อมูลไม่ครบ' });
      const db = await getDb();
      let o = db.rateOverrides.find((x) => x.branchId === branchId && x.cycleId === cycleId && x.rateMasterId === rateMasterId);
      if (o) { o.price = price; o.pieceThreshold = pieceThreshold ?? null; }
      else {
        o = { id: generateId('rov'), branchId, cycleId, rateMasterId, price, pieceThreshold: pieceThreshold ?? null };
        db.rateOverrides.push(o);
      }
      await saveDb(db);
      res.json(o);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ===================== TRIP DOCUMENTS =====================
  // บันทึก trip ที่ผ่าน Review (รับ extracted + คำนวณใหม่ฝั่ง server)
  app.post('/api/trips', async (req, res) => {
    try {
      const { extracted, fileName, branchId } = req.body as {
        extracted: ExtractedTripDocument; fileName: string; branchId: string;
      };
      if (!branchId) return res.status(400).json({ error: 'ต้องระบุสาขา' });
      const db = await getDb();

      // 📅 เปิดรอบอัตโนมัติ: จัดใบเข้ารอบตาม "วันที่ในใบ" (1-15 / 16-สิ้นเดือน) สร้างรอบให้ถ้ายังไม่มี
      const resolved = resolveCycleForDate(db, extracted.documentDate, false);
      if (resolved.invalid || !resolved.cycle) {
        return res.status(400).json({ error: `วันที่ในใบไม่ถูกต้อง (${extracted.documentDate || 'ว่าง'}) — ระบุรอบอัตโนมัติไม่ได้ กรุณาแก้วันที่ออกให้ถูกต้อง` });
      }
      if (resolved.closed) {
        return res.status(400).json({ error: `รอบ "${resolved.cycle.name}" ถูกปิดอยู่ — ต้องให้ HQ เปิดรอบก่อนจึงบันทึกได้` });
      }
      const cycle = resolved.cycle;

      const trip = recomputeTrip(db, cycle, extracted, fileName || 'manual.pdf', branchId);

      // 🔒 กฎเหล็ก: เลขใบกระจายห้ามซ้ำภายในสาขา (ทุกรอบ) — ซ้ำ = การเงินผิดเพี้ยน
      const docNo = (trip.documentNo || '').trim();
      if (docNo) {
        const dup = db.tripDocuments.find(
          (t) => t.branchId === branchId && (t.documentNo || '').trim() === docNo
        );
        if (dup) {
          const dupCycle = db.cycles.find((c) => c.id === dup.cycleId);
          const where = dup.cycleId === cycle.id ? 'ในรอบนี้' : `ในรอบ "${dupCycle?.name || dup.cycleId}"`;
          return res.status(409).json({ error: `เลขใบกระจาย ${docNo} ซ้ำ — มีอยู่แล้ว${where} (ห้ามบันทึกซ้ำ ถ้าต้องการแก้ ให้ลบใบเดิมก่อน)` });
        }
      }

      // บังคับ: ผู้ส่งที่ต้องกรอกกล่อง แต่ยังไม่กรอก -> บันทึกไม่ได้
      const missingBox = trip.receipts.find((r) => r.requiresManualBox && (r.manualBoxQty == null || r.manualBoxQty <= 0));
      if (missingBox) {
        return res.status(400).json({ error: `ใบรับ ${missingBox.receiptNo}: ต้องกรอกจำนวนกล่องก่อนบันทึก (ผู้ส่งส่งเป็นชิ้น)` });
      }
      // ผ่านทุกด่าน -> ถ้าเป็นรอบใหม่ ค่อยเพิ่มลงระบบตอนนี้ (กันสร้างรอบเปล่าเวลาบันทึกไม่ผ่าน)
      if (resolved.created) db.cycles.push(cycle);
      trip.isVerified = true;
      db.tripDocuments.push(trip);
      await saveRecord('tripDocuments', trip);           // เขียนแค่ใบเดียว (เร็ว)
      if (resolved.created) await flushCollection('cycles'); // รอบใหม่ -> เขียน cycles (เล็ก)
      res.status(201).json({ ...trip, _cycle: cycle, _cycleCreated: resolved.created });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.delete('/api/trips/:id', async (req, res) => {
    try {
      const db = await getDb();
      db.tripDocuments = db.tripDocuments.filter((t) => t.id !== req.params.id);
      await removeRecord('tripDocuments', req.params.id);
      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // Recalculate: คำนวณ trip เดิมทั้งรอบใหม่ด้วย master ปัจจุบัน
  app.post('/api/cycles/:id/recalculate', async (req, res) => {
    try {
      const db = await getDb();
      const cycle = db.cycles.find((c) => c.id === req.params.id);
      if (!cycle) return res.status(404).json({ error: 'ไม่พบรอบ' });

      const changed: TripDocument[] = [];
      db.tripDocuments = db.tripDocuments.map((t) => {
        if (t.cycleId !== cycle.id) return t;
        // สร้าง extracted กลับจาก trip เดิม
        const extracted: ExtractedTripDocument = {
          documentNo: t.documentNo,
          documentDate: t.documentDate,
          plateNo: t.plateNo,
          provinceRaw: t.provinceRaw,
          districtRaw: t.districtRaw,
          rateChoice: t.rateType ?? undefined,
          docNote: t.docNote,
          receipts: t.receipts.map((r) => ({
            receiptNo: r.receiptNo,
            receiverName: r.receiverName,
            senderName: r.senderName,
            items: r.items,
            provinceRaw: r.provinceRaw,
            districtRaw: r.districtRaw,
            manualBoxQty: r.manualBoxQty ?? undefined,
          })),
        };
        const recomputed = recomputeTrip(db, cycle, extracted, t.fileName, t.branchId);
        const out = { ...recomputed, id: t.id, isVerified: t.isVerified, createdAt: t.createdAt };
        changed.push(out);
        return out;
      });
      await saveRecords('tripDocuments', changed);
      res.json({ success: true, count: db.tripDocuments.filter((t) => t.cycleId === cycle.id).length });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ===================== PREVIEW (คำนวณก่อนบันทึก โดยไม่เซฟ) =====================
  app.post('/api/trips/preview', async (req, res) => {
    try {
      const { cycleId, extracted, fileName, branchId } = req.body as {
        cycleId: string; extracted: ExtractedTripDocument; fileName: string; branchId: string;
      };
      const db = await getDb();
      // คำนวณตามรอบของ "วันที่ในใบ" (ไม่เซฟรอบ) — ถ้าวันที่ไม่ถูกต้อง fallback ใช้รอบที่เลือก
      const resolved = resolveCycleForDate(db, extracted.documentDate, false);
      const cycle = resolved.cycle || db.cycles.find((c) => c.id === cycleId);
      if (!cycle) return res.status(404).json({ error: 'ระบุรอบไม่ได้ — ตรวจสอบวันที่ในใบ' });
      const preview = recomputeTrip(db, cycle, extracted, fileName || 'manual.pdf', branchId || '');
      res.json({ ...preview, _cycleName: cycle.name, _cycleClosed: resolved.closed, _cycleCreated: resolved.created });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ===================== AI PDF EXTRACTION (ใบกระจาย) =====================
  app.post('/api/extract-pdf', async (req, res) => {
    try {
      const { pdfBase64 } = req.body;
      if (!pdfBase64) return res.status(400).json({ error: 'ต้องส่ง pdfBase64' });

      if (!isAiEnabled()) {
        return res.status(400).json({ error: 'ยังไม่ได้ตั้งค่า GEMINI_API_KEY ใน .env.local — ใช้ปุ่ม "กรอกเอง" เพื่อทดสอบได้' });
      }

      const ai = getGeminiClient();
      // เลือกรุ่นโมเดล: จากการตั้งค่าในแอป -> env -> ค่าเริ่มต้น
      const settingsDb = await getDb();
      const geminiModel = settingsDb.settings?.geminiModel || process.env.GEMINI_MODEL || 'gemini-3.5-flash';
      const rawBase64 = pdfBase64.replace(/^data:application\/pdf;base64,/, '');
      const docPart = { inlineData: { mimeType: 'application/pdf', data: rawBase64 } };
      const promptPart = {
        text: `วิเคราะห์ไฟล์ PDF "ใบกระจายสินค้า" ภาษาไทยนี้ ดึงข้อมูลออกมาให้ครบ:
- ระดับใบกระจาย: เลขที่ใบกระจาย (documentNo), วันที่ออก (documentDate รูปแบบ YYYY-MM-DD), ทะเบียนรถ (plateNo), จังหวัด (provinceRaw), อำเภอ (districtRaw)
- ระดับใบรับสินค้า (receipts): เลขที่ใบรับสินค้า (receiptNo), ผู้รับสินค้า (receiverName), ผู้ส่งสินค้า (senderName), ปลายทางของจุดส่งนี้ (provinceRaw=จังหวัด, districtRaw=อำเภอ ถ้าระบุในที่อยู่ผู้รับ), และรายการสินค้า (items) แต่ละชิ้นมี ชื่อสินค้า (productName), จำนวน (quantity), และหน่วยนับ (unit เช่น กล่อง/หีบ/ลัง จากคอลัมน์หน่วย)
ให้ดึงทุกบรรทัดที่มีจำนวนในคอลัมน์จำนวน รวมบรรทัดที่ชื่อสินค้าเป็น "*** โปรดระบุ ***" ด้วย (ใส่ productName ตามที่เห็น) เพื่อให้ยอดรวมตรงกับเอกสาร.

สำคัญมาก — อ่านชื่อสินค้า/แบรนด์ภาษาไทยให้แม่นยำ เพราะตัวอักษรไทยคล้ายกันมาก ให้เทียบกับรายชื่อแบรนด์ที่ถูกต้องด้านล่าง ถ้าอ่านได้ใกล้เคียงให้สะกดตามนี้:
- "ยูปี้" (ห้ามอ่านเป็น ยูบี/ยูจี/ยูปิ) — เช่น ยูปี้ ฟรุตคอกเทล, ยูปี้ กัมมี่พิชซ่า, ยูปี้ เบอร์เกอร์, ยูปี้ มิกซ์
- "พริงเกิลส" (ห้ามอ่านเป็น ทริงเก็ตส์/พริงเกิ้ล) — เช่น พริงเกิลส ออริจินอล PIL 42
- "มอนเด", "บอน โอ บอน", "ซีโน-แปซิฟิค เทรดดิ้ง (ไทยแลนด์)"
- ชื่อผู้รับ/บริษัทให้คงตามต้นฉบับ อย่าสลับตัวอักษร
ถ้าปีเป็น พ.ศ. ให้แปลงเป็น ค.ศ. (ลบ 543). คงชื่อภาษาไทยไว้. ตอบเป็น JSON ตาม schema เท่านั้น.`,
      };

      const response = await ai.models.generateContent({
        model: geminiModel,
        contents: [docPart, promptPart],
        config: {
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              documentNo: { type: Type.STRING },
              documentDate: { type: Type.STRING, description: 'YYYY-MM-DD' },
              plateNo: { type: Type.STRING },
              provinceRaw: { type: Type.STRING },
              districtRaw: { type: Type.STRING },
              receipts: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    receiptNo: { type: Type.STRING },
                    receiverName: { type: Type.STRING },
                    senderName: { type: Type.STRING },
                    provinceRaw: { type: Type.STRING, description: 'จังหวัดปลายทางของใบรับนี้ (ถ้ามี)' },
                    districtRaw: { type: Type.STRING, description: 'อำเภอปลายทางของใบรับนี้ (ถ้ามี)' },
                    items: {
                      type: Type.ARRAY,
                      items: {
                        type: Type.OBJECT,
                        properties: {
                          productName: { type: Type.STRING },
                          quantity: { type: Type.NUMBER },
                          unit: { type: Type.STRING, description: 'หน่วยนับจากคอลัมน์หน่วย เช่น กล่อง หีบ ลัง' },
                        },
                        required: ['productName', 'quantity'],
                      },
                    },
                  },
                  required: ['receiptNo', 'receiverName', 'senderName', 'items'],
                },
              },
            },
            required: ['documentNo', 'documentDate', 'plateNo', 'provinceRaw', 'districtRaw', 'receipts'],
          },
        },
      });

      const textOutput = response.text;
      if (!textOutput) throw new Error('AI ตอบกลับว่าง');
      res.json({ result: JSON.parse(textOutput.trim()) });
    } catch (err: any) {
      console.error('Gemini extraction error:', err);
      res.status(500).json({ error: `อ่าน PDF ไม่สำเร็จ: ${err.message}` });
    }
  });

  // ===================== IMPORT EXCEL ใบกระจาย (ไม่ใช้ AI) =====================
  app.post('/api/import-excel', async (req, res) => {
    try {
      const { fileBase64 } = req.body as { fileBase64: string };
      if (!fileBase64) return res.status(400).json({ error: 'ต้องส่ง fileBase64' });
      const raw = fileBase64.replace(/^data:[^;]+;base64,/, '');
      const buffer = Buffer.from(raw, 'base64');
      const results = parseDistributionExcel(buffer);
      if (!results.length) {
        return res.status(422).json({ error: 'อ่านไฟล์ Excel ไม่พบใบกระจาย — ตรวจสอบว่าเป็นไฟล์ใบกระจายที่ถูกต้อง' });
      }
      res.json({ results });
    } catch (err: any) {
      console.error('Excel import error:', err);
      res.status(500).json({ error: `อ่าน Excel ไม่สำเร็จ: ${err.message}` });
    }
  });

  // ===================== Static / Vite =====================
  if (process.env.NODE_ENV !== 'production') {
    // โหลด vite เฉพาะตอน dev (production ไม่มี vite ใน deps)
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({ server: { middlewareMode: true }, appType: 'spa' });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    // ไฟล์ asset (js/css) ของ Vite มีรหัส hash ในชื่อ -> cache ถาวรได้ (1 ปี, immutable)
    // เปิดเว็บซ้ำ = ไม่ต้องโหลดไฟล์ใหม่เลย (แบนด์วิดท์ ~0) จนกว่าจะ deploy เวอร์ชันใหม่
    // index:false -> ไม่ให้ static เสิร์ฟ index.html (กันถูก cache ถาวรจนอัปเดตไม่เห็น)
    app.use(express.static(distPath, { index: false, maxAge: '1y', immutable: true }));
    // index.html ต้องไม่ cache (อ้างชื่อไฟล์ asset ใหม่ทุก deploy) -> โหลดสดเสมอ แต่เล็กมาก
    app.get('*', (_req, res) => {
      res.setHeader('Cache-Control', 'no-cache');
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => console.log(`Server running at http://0.0.0.0:${PORT}`));
}

startServer().catch((error) => console.error('Failed to start server:', error));
