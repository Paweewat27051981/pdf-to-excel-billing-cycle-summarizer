import express from 'express';
import path from 'path';
import { GoogleGenAI, Type } from '@google/genai';
import dotenv from 'dotenv';
import { getDb, saveDb } from './server-db.js';
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
  MoneyCategory,
  TripDocument,
  FuelEntry,
  DeductionEntry,
  ExtractedTripDocument,
} from './src/types.js';
import { computeTripDocument } from './src/calc.js';
import { parseDistributionExcel } from './excel-import.js';

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

  // ใช้รถ/ราคา/กฎของสาขานั้น + ราคาเฉพาะรอบ (ถ้ามี)
  const trip = computeTripDocument(
    extracted,
    {
      cycleId: cycle.id,
      cycle: cycleCtx(cycle),
      vehicles: db.vehicles.filter((v) => v.branchId === branchId),
      rates: db.rateMasters.filter((r) => r.branchId === branchId),
      rateOverrides: overrides,
      groups: db.receiverGroups.filter((g) => g.branchId === branchId),
      aliases: db.receiverGroupAliases.filter((a) => a.branchId === branchId),
      rules: db.conversionRules.filter((r) => r.branchId === branchId),
      manualBoxSenders: db.manualBoxSenders.filter((m) => m.branchId === branchId),
      fileName,
    },
    () => generateId('rcp')
  );
  trip.branchId = branchId;
  return trip;
}

async function startServer() {
  const app = express();
  const PORT = Number(process.env.PORT) || 3000;
  app.use(express.json({ limit: '50mb' }));
  app.use(express.urlencoded({ limit: '50mb', extended: true }));

  // ===================== CONFIG =====================
  app.get('/api/config', (_req, res) => {
    res.json({ aiEnabled: isAiEnabled() });
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
        await saveDb(db);
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
        await saveDb(db);
        res.json(arr[idx]);
      } catch (err: any) {
        res.status(500).json({ error: err.message });
      }
    });
    app.delete(`/api/${name}/:id`, async (req, res) => {
      try {
        const db = await getDb();
        (db[key] as unknown as T[]) = (db[key] as unknown as T[]).filter((x) => x.id !== req.params.id) as any;
        await saveDb(db);
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
        await saveDb(db);
        res.json({ success: true, deleted: ids.length });
      } catch (err: any) {
        res.status(500).json({ error: err.message });
      }
    });
  }

  masterRoutes<Branch>('branches', 'branches', 'br');
  masterRoutes<RateOverride>('rate-overrides', 'rateOverrides', 'rov');
  masterRoutes<MoneyCategory>('money-categories', 'moneyCategories', 'cat');
  masterRoutes<ManualBoxSender>('manual-box-senders', 'manualBoxSenders', 'mbs');
  masterRoutes<Vehicle>('vehicles', 'vehicles', 'veh');
  masterRoutes<ReceiverGroup>('receiver-groups', 'receiverGroups', 'grp');
  masterRoutes<ReceiverGroupAlias>('receiver-aliases', 'receiverGroupAliases', 'al');
  masterRoutes<ProductConversionRule>('conversion-rules', 'conversionRules', 'rule');
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
      const { cycleId, extracted, fileName, branchId } = req.body as {
        cycleId: string; extracted: ExtractedTripDocument; fileName: string; branchId: string;
      };
      if (!branchId) return res.status(400).json({ error: 'ต้องระบุสาขา' });
      const db = await getDb();
      const cycle = db.cycles.find((c) => c.id === cycleId);
      if (!cycle) return res.status(404).json({ error: 'ไม่พบรอบ' });
      if (cycle.status === 'closed') return res.status(400).json({ error: 'รอบนี้ถูกปิดล็อกแล้ว' });

      const trip = recomputeTrip(db, cycle, extracted, fileName || 'manual.pdf', branchId);
      // บังคับ: ผู้ส่งที่ต้องกรอกกล่อง แต่ยังไม่กรอก -> บันทึกไม่ได้
      const missingBox = trip.receipts.find((r) => r.requiresManualBox && (r.manualBoxQty == null || r.manualBoxQty <= 0));
      if (missingBox) {
        return res.status(400).json({ error: `ใบรับ ${missingBox.receiptNo}: ต้องกรอกจำนวนกล่องก่อนบันทึก (ผู้ส่งส่งเป็นชิ้น)` });
      }
      trip.isVerified = true;
      db.tripDocuments.push(trip);
      await saveDb(db);
      res.status(201).json(trip);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.delete('/api/trips/:id', async (req, res) => {
    try {
      const db = await getDb();
      db.tripDocuments = db.tripDocuments.filter((t) => t.id !== req.params.id);
      await saveDb(db);
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
        return { ...recomputed, id: t.id, isVerified: t.isVerified, createdAt: t.createdAt };
      });
      await saveDb(db);
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
      const cycle = db.cycles.find((c) => c.id === cycleId);
      if (!cycle) return res.status(404).json({ error: 'ไม่พบรอบ' });
      res.json(recomputeTrip(db, cycle, extracted, fileName || 'manual.pdf', branchId || ''));
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
    app.use(express.static(distPath));
    app.get('*', (_req, res) => res.sendFile(path.join(distPath, 'index.html')));
  }

  app.listen(PORT, '0.0.0.0', () => console.log(`Server running at http://0.0.0.0:${PORT}`));
}

startServer().catch((error) => console.error('Failed to start server:', error));
