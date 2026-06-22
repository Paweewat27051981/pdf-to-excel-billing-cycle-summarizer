import express from 'express';
import path from 'path';
import { GoogleGenAI, Type } from '@google/genai';
import dotenv from 'dotenv';
import { getDb, saveDb } from './server-db.js';
import {
  DatabaseState,
  BillingCycle,
  Vehicle,
  RateMaster,
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
  fileName: string
): TripDocument {
  return computeTripDocument(
    extracted,
    {
      cycleId: cycle.id,
      cycle: cycleCtx(cycle),
      vehicles: db.vehicles,
      rates: db.rateMasters,
      groups: db.receiverGroups,
      aliases: db.receiverGroupAliases,
      rules: db.conversionRules,
      manualBoxSenders: db.manualBoxSenders,
      fileName,
    },
    () => generateId('rcp')
  );
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

  // ===================== STATE =====================
  app.get('/api/state', async (_req, res) => {
    try {
      res.json(await getDb());
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
  }

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

  // ===================== TRIP DOCUMENTS =====================
  // บันทึก trip ที่ผ่าน Review (รับ extracted + คำนวณใหม่ฝั่ง server)
  app.post('/api/trips', async (req, res) => {
    try {
      const { cycleId, extracted, fileName } = req.body as {
        cycleId: string; extracted: ExtractedTripDocument; fileName: string;
      };
      const db = await getDb();
      const cycle = db.cycles.find((c) => c.id === cycleId);
      if (!cycle) return res.status(404).json({ error: 'ไม่พบรอบ' });
      if (cycle.status === 'closed') return res.status(400).json({ error: 'รอบนี้ถูกปิดล็อกแล้ว' });

      const trip = recomputeTrip(db, cycle, extracted, fileName || 'manual.pdf');
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
        const recomputed = recomputeTrip(db, cycle, extracted, t.fileName);
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
      const { cycleId, extracted, fileName } = req.body as {
        cycleId: string; extracted: ExtractedTripDocument; fileName: string;
      };
      const db = await getDb();
      const cycle = db.cycles.find((c) => c.id === cycleId);
      if (!cycle) return res.status(404).json({ error: 'ไม่พบรอบ' });
      res.json(recomputeTrip(db, cycle, extracted, fileName || 'manual.pdf'));
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
        model: 'gemini-3.5-flash',
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
