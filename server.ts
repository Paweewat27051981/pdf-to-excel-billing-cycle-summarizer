import express from 'express';
import compression from 'compression';
import path from 'path';
import fs from 'fs';
import { randomUUID, timingSafeEqual } from 'crypto';
import { GoogleGenAI, Type } from '@google/genai';
import dotenv from 'dotenv';
import { getDb, peekCache, peekCacheAsync, warmCacheOnBoot, saveDb, saveRecord, saveRecords, removeRecord, removeRecords, flushCollection, isIdKeyed } from './server-db.js';
import {
  DatabaseState,
  BillingCycle,
  Branch,
  Vehicle,
  RateMaster,
  RateMasterHistory,
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
import { computeTripDocument, normPlate, normDoc, round2, textContains, isDateInCycle } from './src/calc.js';
import { parseDistributionExcel, parseRateExcel, parseFuelExcel } from './excel-import.js';
import { registerExperimentalRoutes } from './experimental-routes.js'; // [ทดลอง] แยก 100%
import { startOilPriceScheduler } from './src/experimental/oilPriceScheduler.js'; // [ทดลอง] auto 05:30 ไทย

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

  // รองรับรันใต้ subpath (เช่น /neosiam หลัง reverse proxy บน NAS)
  // ตัด prefix ออกก่อนเข้า routes -> โค้ด routes ทั้งหมด + static ยังทำงานที่ root เหมือนเดิม
  const BASE_PATH = (process.env.BASE_PATH || '').replace(/\/$/, '');
  if (BASE_PATH) {
    app.use((req, _res, next) => {
      if (req.url === BASE_PATH) req.url = '/';
      else if (req.url.startsWith(BASE_PATH + '/')) req.url = req.url.slice(BASE_PATH.length);
      next();
    });
  }

  // ===================== CONFIG =====================
  app.get('/api/config', (_req, res) => {
    res.json({ aiEnabled: isAiEnabled(), storage: 'granular-v2' });
  });

  // ===================== รูปแนบ (เก็บไฟล์บน NAS โฟลเดอร์ uploads/ ไม่เก็บใน Firebase) =====================
  const UPLOADS_DIR = path.join(process.cwd(), 'uploads');
  fs.mkdirSync(UPLOADS_DIR, { recursive: true }); // สร้างโฟลเดอร์ถ้ายังไม่มี (mount เป็น volume บน NAS)
  // เสิร์ฟรูป: ดูได้ที่ /api/uploads/<ชื่อไฟล์>
  // 🔒 กันไฟล์ลับหลุด: uploads/ เป็น volume เดียวที่เขียนได้บน NAS จึงมีไฟล์ตั้งค่า
  //    (เช่น jastran-data/agent-token.txt) มาอยู่ด้วย — ต้องไม่ให้ดึงผ่าน URL ได้
  //    เสิร์ฟเฉพาะไฟล์รูปที่แบนราบในโฟลเดอร์นี้เท่านั้น
  app.use('/api/uploads', (req, res, next) => {
    // decodeURIComponent โยน URIError ถ้า % ไม่ครบคู่ (เช่น /api/uploads/%)
    // ไม่ดักไว้ = ตอบ 500 + รก log (เคยเห็น URIError: Failed to decode param ในล็อก prod)
    let p: string;
    try { p = decodeURIComponent(req.path || ''); }
    catch { return res.status(404).end(); }   // URL เพี้ยน = ไม่มีไฟล์นี้อยู่แล้ว
    // เช็คทั้งก่อนและหลัง decode — กัน %2f ที่ decode แล้วกลายเป็น / (หลบด่านชั้นเดียวได้)
    const raw = req.path || '';
    const nested = (s: string) => s.includes('/', 1) || s.includes('\\');
    if (nested(p) || nested(raw) || !/\.(png|jpe?g|gif|webp|avif)$/i.test(p)) {
      return res.status(404).end();   // ไม่ใช่รูปในระดับบนสุด -> ไม่มีให้ดู
    }
    next();
  }, express.static(UPLOADS_DIR, { maxAge: '7d' }));
  // อัปโหลด: รับ base64 ที่ "ย่อขนาดจากเบราว์เซอร์แล้ว" -> เขียนไฟล์ -> คืนชื่อไฟล์
  app.post('/api/upload-image', async (req, res) => {
    try {
      const b64: string = req.body?.imageBase64 || '';
      const m = /^data:image\/(png|jpe?g|webp);base64,(.+)$/i.exec(b64);
      if (!m) return res.status(400).json({ error: 'รูปไม่ถูกต้อง (รองรับ jpg/png/webp)' });
      const ext = m[1].toLowerCase() === 'jpeg' ? 'jpg' : m[1].toLowerCase();
      const buf = Buffer.from(m[2], 'base64');
      if (buf.length > 5 * 1024 * 1024) return res.status(400).json({ error: 'รูปใหญ่เกิน 5MB (ควรย่อก่อน)' });
      const name = `img-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.${ext}`;
      await fs.promises.writeFile(path.join(UPLOADS_DIR, name), buf);
      res.json({ filename: name });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ===================== SETTINGS =====================
  app.put('/api/settings', async (req, res) => {
    try {
      const db = await getDb();
      db.settings = { ...db.settings, ...req.body };
      // เขียนเฉพาะ node ที่แก้ (เดิม saveDb เขียนทั้ง tree -> Firebase "Write too large" เมื่อ DB โต)
      await flushCollection('settings');
      res.json(db.settings);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ===================== SESSION (ตรวจสิทธิ์ฝั่ง server) =====================
  // เก็บใน memory: restart แล้วหาย -> ผู้ใช้ล็อกอินใหม่ (ยอมรับได้ ปลอดภัยกว่าเก็บถาวร)
  // ใช้ตัดสิน "ใครแก้ราคาได้" ไม่ให้ client โกหกได้ (เดิมไม่มี session เลย ใครยิง API ก็แก้ได้)
  type Session = { branchId: string; name: string; isHQ: boolean; canEditRates: boolean; at: number };
  const sessions = new Map<string, Session>();
  const SESSION_TTL_MS = 30 * 24 * 60 * 60 * 1000; // 30 วัน
  const newToken = () => `${Date.now().toString(36)}-${randomUUID()}`;
  const getSession = (req: any): Session | null => {
    const raw = String(req.headers['authorization'] || '');
    const token = raw.startsWith('Bearer ') ? raw.slice(7).trim() : '';
    if (!token) return null;
    const s = sessions.get(token);
    if (!s) return null;
    if (Date.now() - s.at > SESSION_TTL_MS) { sessions.delete(token); return null; }
    return s;
  };
  // ด่านสิทธิ์ "แก้ราคา" — ใช้กับทุก endpoint ที่แตะ Master ราคา/ราคาเฉพาะรอบ/นำเข้า Excel
  // อ่านสิทธิ์จาก DB ปัจจุบันทุกครั้ง (ไม่เชื่อค่าที่ติดมากับ token ตอน login):
  // ถอนสิทธิ์/ปิดบัญชีแล้วต้องมีผลทันที ไม่ต้องรอ token หมดอายุ (Codex P2)
  const requireRateEditor = async (req: any, res: any, next: any) => {
    try {
      const sess = getSession(req);
      if (!sess) return res.status(401).json({ error: 'กรุณาเข้าสู่ระบบใหม่ (เซสชันหมดอายุ)' });
      const db = await getDb();
      const me = db.branches.find((b) => b.id === sess.branchId);
      if (!me || me.status !== 'active') {
        return res.status(403).json({ error: 'บัญชีนี้ถูกปิดใช้งานแล้ว — ติดต่อผู้ดูแลระบบ' });
      }
      if (me.canEditRates) return next();
      // ยังไม่มีบัญชีผู้ดูแลราคาในระบบเลย (เช่น ติดตั้งใหม่) -> ให้ HQ ทำแทนได้ชั่วคราว
      // กันสถานการณ์ "ล็อกตาย" ที่ไม่มีใครแก้ราคาได้เลยและสร้าง admin ไม่ได้ (Codex P1)
      const hasAnyEditor = db.branches.some((b) => b.canEditRates && b.status === 'active');
      if (!hasAnyEditor && me.isHQ) {
        console.warn('[rate-guard] ยังไม่มีบัญชีผู้ดูแลราคา — อนุญาต HQ ชั่วคราว (ควรสร้างบัญชี admin)');
        return next();
      }
      return res.status(403).json({ error: `บัญชี "${me.name}" ไม่มีสิทธิ์แก้ไขราคา — ใช้บัญชีผู้ดูแลราคา (admin) เท่านั้น` });
    } catch (err: any) {
      return res.status(500).json({ error: err.message });
    }
  };

  // ===================== ใบกระจายจากจัสทราน (เก็บบน NAS ไม่เก็บใน Firebase) =====================
  // เจ้าของกำหนด: หน้าคำนวณค่าเที่ยวทำงานเหมือนเดิมทุกอย่าง เปลี่ยนแค่ "วิธีนำเข้า"
  // agent บนเครื่องจัสทรานส่งใบกระจายมาที่นี่ -> เก็บเป็นไฟล์ JSON รายวันบน NAS
  // -> หน้าเว็บดึงไปเข้า flow เดิม (preview -> ตรวจ -> กดบันทึกทีละใบ)
  // ⚠️ ห้ามวางใต้ uploads/ — โฟลเดอร์นั้นถูกเสิร์ฟสาธารณะด้วย express.static
  //    (ใครก็โหลด /api/uploads/jastran/jb-2026-08-18.json ได้ = ข้อมูลรั่ว แม้ POST จะมี token)
  //    ใช้โฟลเดอร์แยกที่ mount บน NAS เหมือนกันแต่ไม่ถูกเสิร์ฟ
  const JASTRAN_DIR = path.join(process.cwd(), 'jastran-data');
  fs.mkdirSync(JASTRAN_DIR, { recursive: true });

  const jastranFile = (date: string) => path.join(JASTRAN_DIR, `jb-${date}.json`);

  // "ยังไม่ส่งเสร็จสักจุด" = ห้ามคิดค่าเที่ยว (กฎเหล็ก)
  // ตัดสินจากตัวเลขเป็นหลัก แล้วค่อยดู flag จาก agent — ข้อมูลเก่าที่ไม่มี flag จึงยังถูกกัน
  const isNotDelivered = (d: any) => {
    const done = Number(d?._delivered);
    const total = Number(d?._totalReceipts);
    if (Number.isFinite(done) && Number.isFinite(total) && total > 0) return done === 0;
    return d?._notDelivered === true;   // ไม่มีตัวเลข -> เชื่อ flag (ถ้ามี)
  };
  const isYmd = (d: any) => /^\d{4}-\d{2}-\d{2}$/.test(String(d || ''));

  // "สาขานี้วิ่งปลายทางนี้ไหม" — ตัดสินจากราคาที่สาขามีใน Master
  // ใช้ร่วมกันทั้ง /available (นับ) และ /trips (รายการ) เพื่อให้ตัวเลขบนปุ่มตรงกับที่เปิดได้จริง
  // เทียบชื่อด้วย textContains ตัวเดียวกับ matchRate -> ผลตรงกับตอนคิดราคา
  // ปลายทาง "จริง" ของใบ — ต้องผ่านกฎแก้ปลายทางก่อน ไม่งั้นกรองผิดสาขา
  // ใบกระจายบางใบเขียนจังหวัดผิด (เคสจริง JB0626076941 เขียน "อุตรดิตถ์" แต่ส่งพิษณุโลก)
  // กฎจับจาก ชื่อผู้รับ/ผู้ส่ง/ชื่อสินค้า เหมือน computeReceipt ใน calc.ts (บรรทัด ~409)
  // ถ้ากรองด้วยจังหวัดดิบ ใบจะไปโผล่ผิดสาขา หรือหายถ้าสาขาไม่มีราคาจังหวัดที่เขียนผิด
  const realProvinceOf = (d: any, overrides: any[]) => {
    const hay = [d?.receiverName, d?.senderName]
      .concat((d?.receipts || []).map((r: any) => r?.receiverName))
      .concat((d?.receipts || []).map((r: any) => r?.senderName))
      .concat((d?.receipts || []).flatMap((r: any) => (r?.items || []).map((it: any) => it?.productName)))
      .filter(Boolean).join('  ');
    const ov = overrides.find((o) => o.status === 'active' && o.keyword && textContains(hay, o.keyword));
    return (ov && ov.province) ? ov.province : d?.provinceRaw;
  };

  const makeAreaCheck = (rates: any[]) => {
    const active = rates.filter((r) => r.status === 'active');
    // ⚠️ ห้ามใช้ provinceShort เทียบแบบ substring เด็ดขาด
    //    ชื่อย่อ 2 ตัวอักษรไปโผล่กลางชื่อจังหวัดอื่นได้ เช่น
    //      "กำแ[พช]ร"    ตรงกับ เพชรบูรณ์ (พช)  -> นครสวรรค์/พิษณุโลกเห็นใบกำแพงเพชร
    //      "กำแพ[งเพชร]" มี "ชร" ตรงกับ เชียงราย (ชร) -> เชียงใหม่เห็นด้วย
    //    พิสูจน์กับข้อมูลจริง: ใบกำแพงเพชร 1 ใบ ถูก 5 สาขาเห็นพร้อมกัน
    //    ที่นี่ใช้เทียบ "ชื่อเต็มเท่านั้น" (ยังยอมให้มี อ./จ. นำหน้าได้ผ่าน textContains)
    return (prov: string) => {
      if (!active.length) return false;         // สาขาไม่มีราคาเลย = ไม่รู้ว่าวิ่งที่ไหน -> ไม่เดา
      const p = String(prov || '').trim();
      if (!p) return true;                      // ใบไม่มีจังหวัด -> ให้เห็น (ไม่ตกหล่น)
      // ตรงระดับจังหวัด = สาขานี้วิ่งจังหวัดนั้น -> ให้เห็น
      // ไม่เช็คลึกถึงอำเภอ เพราะปลายทางใหม่ที่ยังไม่มีราคาจะตกหล่น
      // (เคสจริง: เชียงราย อ.เมือง — เชียงใหม่วิ่งเชียงรายแต่ยังไม่มีราคาอำเภอนี้
      //  ถ้ากรองด้วยอำเภอจะไม่มีใครเห็นใบนี้เลย = ตกหล่น ไม่ได้เงิน)
      // อำเภอปล่อยให้ ReviewBoard เตือน "ไม่เจอราคา" ตอนตรวจแทน
      return active.some((r) =>
        textContains(p, r.provinceName) || textContains(r.provinceName, p)
      );
    };
  };

  // อ่าน token ได้ 2 ทาง: env หรือไฟล์ agent-token.txt ใน jastran-data/
  //
  // ทำไมต้องมีทางที่ 2: env_file ของ docker-compose ถูกอ่านตอน "สร้าง" container เท่านั้น
  // และ .env ไม่ได้เป็น volume -> แก้ .env แล้ว Stop/Start ค่าไม่เข้า ต้อง recreate container
  // ซึ่งบน NAS ที่ container อยู่ใน Project จะกดยาก + เสี่ยงเว็บล่ม
  // ส่วน jastran-data/ เป็น volume อยู่แล้ว -> วางไฟล์แล้ว restart ธรรมดาก็มีผลทันที
  //
  // ปลอดภัย: jastran-data/ ไม่ได้ถูกเสิร์ฟด้วย express.static (ต่างจาก uploads/)
  // อ่านสดทุกครั้ง ไม่ cache -> เปลี่ยน token แล้วมีผลทันทีโดยไม่ต้อง restart ด้วยซ้ำ
  // หา token ได้หลายที่ เพราะบน NAS มี volume ที่เขียนได้จริงแค่ uploads/
  // (jastran-data/ ต้องเพิ่ม volume ใน compose + recreate container ซึ่งเสี่ยงกับ prod)
  // -> รองรับ uploads/jastran-data/ ด้วย จะได้วางไฟล์ได้เลยโดยไม่ต้องแตะ container
  // ⚠️ uploads/ ถูกเสิร์ฟผ่าน /api/uploads แต่มี guard กรองเฉพาะไฟล์รูประดับบนสุดแล้ว
  //    ไฟล์ในโฟลเดอร์ย่อยจึงดึงผ่าน URL ไม่ได้ (ดูด้านบน + เทสต์ test-token-leak)
  const TOKEN_FILES = [
    path.join(JASTRAN_DIR, 'agent-token.txt'),
    path.join(UPLOADS_DIR, 'jastran-data', 'agent-token.txt'),
  ];
  const readAgentToken = () => {
    // trim ทั้งสองทาง: ก๊อปวาง token แล้วติดช่องว่าง/ขึ้นบรรทัดใหม่มาด้วยเป็นเรื่องปกติ
    // ถ้าไม่ trim จะได้ 403 โดยหาสาเหตุไม่เจอ (ตาเปล่ามองไม่เห็นช่องว่าง)
    // ผลข้างเคียง: token ที่ "ตั้งใจ" ให้มีช่องว่างหัวท้ายจะใช้ไม่ได้ — ซึ่งไม่ควรมีใครทำ
    const fromEnv = (process.env.AGENT_TOKEN || '').trim();
    if (fromEnv) return fromEnv;
    for (const f of TOKEN_FILES) {
      try {
        const t = fs.readFileSync(f, 'utf8').trim();
        if (t) return t;      // ไฟล์ว่าง/มีแต่ช่องว่าง -> ข้ามไปหาที่อื่นต่อ
      } catch (e: any) {
        // ไม่มีไฟล์ = ปกติ (ยังไม่ได้วาง) แต่ permission ผิด/เป็นโฟลเดอร์ = ต้องรู้
        // ไม่งั้นเจอ 503 แล้วหาสาเหตุไม่เจอ (ไฟล์อยู่ตรงนั้นแต่อ่านไม่ได้)
        if (e?.code && e.code !== 'ENOENT') console.warn(`[jastran] อ่าน ${f} ไม่ได้: ${e.code}`);
      }
    }
    return '';                // ไม่เจอที่ไหนเลย = ยังไม่ตั้ง token (ปิดช่องทางไว้)
  };

  // agent ส่งใบกระจายเข้ามา (แทนที่ไฟล์ของวันนั้น — ส่งซ้ำได้ ข้อมูลล่าสุดชนะ)
  app.post('/api/jastran/trips', async (req, res) => {
    try {
      const token = String(req.headers['x-agent-token'] || '');
      const expect = readAgentToken();
      // ต้องตั้ง token ก่อนใช้ — ไม่ตั้ง = ปิดช่องทางนี้ (กันคนอื่นยิงเข้ามา)
      if (!expect) return res.status(503).json({ error: 'ยังไม่ได้ตั้ง AGENT_TOKEN บนเซิร์ฟเวอร์ (วางไฟล์ agent-token.txt ที่ jastran-data/ หรือ uploads/jastran-data/)' });
      // เทียบแบบ timing-safe — `!==` หยุดทันทีที่ตัวอักษรต่างกัน
      // ทำให้เดา token ทีละตัวได้จากเวลาตอบกลับ (timing attack)
      const a = Buffer.from(token), b = Buffer.from(expect);
      if (a.length !== b.length || !timingSafeEqual(a, b)) {
        return res.status(403).json({ error: 'token ไม่ถูกต้อง' });
      }

      const { date, docs } = req.body as { date: string; docs: any[] };
      if (!isYmd(date)) return res.status(400).json({ error: 'ต้องระบุ date รูปแบบ YYYY-MM-DD' });
      if (!Array.isArray(docs)) return res.status(400).json({ error: 'ต้องส่ง docs เป็น array' });
      // กันไฟล์ใหญ่ผิดปกติ (ใบกระจายจริง ~130 ใบ/วัน)
      if (docs.length > 2000) return res.status(413).json({ error: `ส่งมา ${docs.length} ใบ มากผิดปกติ` });

      const payload = { date, receivedAt: new Date().toISOString(), count: docs.length, docs };
      // เขียนแบบ atomic (tmp + rename) — agent ส่งซ้ำวันเดิมได้ ถ้าเขียนทับตรงๆ
      // คนที่อ่านอยู่พร้อมกันจะได้ JSON ครึ่งไฟล์ (พัง) หรือไฟล์เสียถาวรถ้า crash กลางคัน
      // แพทเทิร์นเดียวกับ snapshot ใน server-db.ts
      const target = jastranFile(date);
      // tmp ต้องไม่ซ้ำต่อ request — ถ้าสองคำขอของวันเดียวกันมาพร้อมกัน จะ rename ทับ tmp ของกันและกัน
      // (ตัวหนึ่ง rename ไปแล้ว อีกตัว rename ต่อ -> ENOENT หรือได้ข้อมูลของคำขอผิดตัว)
      const tmp = `${target}.${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.tmp`;
      await fs.promises.writeFile(tmp, JSON.stringify(payload), 'utf8');
      // Windows คืน EPERM/EBUSY ถ้า rename ทับไฟล์ที่มีคนเปิดอ่านอยู่พอดี (Linux/NAS ไม่มีปัญหานี้)
      // agent ส่งซ้ำวันเดียวกันพร้อมกันได้ -> ลองใหม่สั้นๆ แทนที่จะโยน 500 ทิ้งข้อมูล
      let renamed = false;
      for (let i = 0; i < 5 && !renamed; i++) {
        try {
          await fs.promises.rename(tmp, target);
          renamed = true;
        } catch (e: any) {
          if (i === 4 || !['EPERM', 'EBUSY', 'EACCES'].includes(e?.code)) {
            await fs.promises.unlink(tmp).catch(() => {});   // ไม่ทิ้งไฟล์ขยะไว้
            throw e;
          }
          await new Promise((r) => setTimeout(r, 50 * (i + 1)));
        }
      }
      console.log(`[jastran] รับใบกระจาย ${docs.length} ใบ วันที่ ${date}`);
      res.json({ success: true, date, count: docs.length });
    } catch (err: any) {
      console.error('[jastran/trips] error:', err);
      res.status(500).json({ error: err.message });
    }
  });

  // หน้าเว็บถามว่ามีใบของวันไหนบ้าง (ไว้โชว์ตัวเลขข้างปุ่ม)
  app.get('/api/jastran/available', async (req, res) => {
    try {
      // ต้องล็อกอินก่อน — ข้อมูลใบกระจายเป็นข้อมูลธุรกิจ ไม่ควรเปิดสาธารณะ
      // (ย้ายไฟล์ออกจาก uploads/ แล้ว แต่ถ้า API ยังเปิด ก็เท่ากับประตูหลังยังเปิดอยู่ — Codex P1)
      if (!getSession(req)) return res.status(401).json({ error: 'กรุณาเข้าสู่ระบบก่อน' });
      // นับเฉพาะใบที่สาขานั้น "เปิดดูได้จริง" — ไม่งั้นตัวเลขบนปุ่มไม่ตรงกับรายการที่เปิดได้
      // และสาขาหนึ่งจะเห็นปริมาณงานของสาขาอื่น
      const dbA = await getDb();
      const sessA = getSession(req)!;
      const meA = dbA.branches.find((b) => b.id === sessA.branchId);
      // HQ เลือกสาขาทำงานได้ -> ใช้ค่าที่ส่งมา · สาขาปกติบังคับเป็นของตัวเอง (ไม่เชื่อ query)
      const isHQLikeA = !!(meA?.isHQ || meA?.isSystemUser);
      const scopeBranch = isHQLikeA ? String(req.query.branchId || '') : sessA.branchId;
      // เก็บ "ทุกสาขา" ที่มีทะเบียนนี้ ไม่ใช่เจ้าของคนแรกที่เจอ
      // เคสจริง: 3ฒษ-2509 อยู่ทั้งนครสวรรค์และพิษณุโลก -> เดิมเดาเป็นนครสวรรค์
      // ทำให้ใบพิษณุโลกไปโผล่ที่นครสวรรค์ (เจ้าของเห็นแล้วทัก)
      const ownerOf = new Map<string, Set<string>>();
      for (const v of dbA.vehicles) {
        const k = normPlate(v.plateNo || '');
        if (!k) continue;
        if (!ownerOf.has(k)) ownerOf.set(k, new Set());
        ownerOf.get(k)!.add(v.branchId);
      }
      // ต้องกรองเหมือน /trips เป๊ะ ไม่งั้นตัวเลขบนปุ่มไม่ตรงกับรายการที่เปิดได้
      const inAreaA = makeAreaCheck(scopeBranch ? dbA.rateMasters.filter((r) => r.branchId === scopeBranch) : []);
      const ovA = (dbA.destinationOverrides || []).filter((o: any) => !scopeBranch || o.branchId === scopeBranch);
      // ต้องกรองเหมือน /trips ทุกโหมด ไม่งั้นตัวเลขบนปุ่มไม่ตรงกับรายการ
      const unknownOnlyA = String(req.query.unknownPlate || '') === '1';
      // ต้องนิยามเหมือน /trips เป๊ะ ไม่งั้นตัวเลขบนปุ่มไม่ตรงกับรายการ
      const isUnknownPlateA = (d: any) => {
        const owners = ownerOf.get(normPlate(String(d?.plateNo || '')));
        if (!owners || owners.size === 0) return true;
        if (owners.size === 1) return false;
        return !Array.from(owners).some((b) => {
          const chk = makeAreaCheck(dbA.rateMasters.filter((r) => r.branchId === b));
          const ov = (dbA.destinationOverrides || []).filter((o: any) => o.branchId === b);
          return chk(realProvinceOf(d, ov));
        });
      };
      const visibleTo = (d: any) => {
        if (unknownOnlyA) return isUnknownPlateA(d);      // หน้า "ทะเบียนไม่รู้จัก" — ทุกสาขาเห็นเท่ากัน
        if (!scopeBranch) return !isUnknownPlateA(d);
        const owners = ownerOf.get(normPlate(String(d?.plateNo || '')));
        // ทะเบียนอยู่สาขาเดียว -> ชี้ขาดได้
        if (owners && owners.size === 1) return owners.has(scopeBranch);
        // ทะเบียนอยู่หลายสาขา -> ชี้ขาดไม่ได้ ต้องเป็นหนึ่งในนั้น + ผ่านพื้นที่ด้วย
        if (owners && owners.size > 1) {
          if (!owners.has(scopeBranch)) return false;
          // ใช้ปลายทาง "จริง" (ผ่านกฎแก้ปลายทางของสาขานี้) ไม่ใช่จังหวัดดิบในใบ
          return inAreaA(realProvinceOf(d, ovA));
        }
        return false;   // ทะเบียนไม่รู้จัก -> ไปอยู่หน้า "ทะเบียนไม่รู้จัก" แทน
      };

      // กรองตามงวดที่เลือกอยู่ — ใบวันที่ 28/08 ต้องไม่โผล่ตอนเปิดงวด มิ.ย. รอบ 1-15
      // ใช้ isDateInCycle ตัวเดียวกับตอนบันทึกใบจริง -> วันไหนเข้างวดไหน ตรงกันเสมอ
      const cycIdA = String(req.query.cycleId || '');
      const cyc = cycIdA ? dbA.cycles.find((c) => c.id === cycIdA) : null;
      if (cycIdA && !cyc) return res.status(400).json({ error: 'ไม่พบรอบที่ระบุ' });

      const files = await fs.promises.readdir(JASTRAN_DIR).catch(() => [] as string[]);
      const days: { date: string; count: number; receivedAt: string }[] = [];
      for (const f of files) {
        const m = /^jb-(\d{4}-\d{2}-\d{2})\.json$/.exec(f);
        if (!m) continue;
        try {
          const j = JSON.parse(await fs.promises.readFile(path.join(JASTRAN_DIR, f), 'utf8'));
          // กรองงวดที่ "ระดับใบ" ด้วย documentDate — ตัวเดียวกับที่ตอนบันทึกใช้เลือกงวด
          // ไม่ใช้วันที่ของไฟล์ เพราะไฟล์ jb-16 อาจมีใบลงวันที่ 15 (คนละงวด) ปนมาได้
          const inCyc = (d: any) => !cyc || isDateInCycle(String(d?.documentDate || ''), cyc);
          const n = Array.isArray(j.docs) ? j.docs.filter((d: any) => inCyc(d) && visibleTo(d)).length : 0;
          if (!n) continue;   // วันที่สาขานี้ไม่มีใบเลย -> ไม่ต้องโชว์
          days.push({ date: m[1], count: n, receivedAt: j.receivedAt || '' });
        } catch { /* ไฟล์เสีย -> ข้าม ไม่ให้ล้มทั้ง endpoint */ }
      }
      days.sort((a, b) => b.date.localeCompare(a.date));
      res.json({ days });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ดึงใบของวันที่ระบุ + บอกว่าใบไหน "บันทึกไปแล้ว" (เทียบเลขใบกับที่มีในระบบ)
  app.get('/api/jastran/trips', async (req, res) => {
    try {
      if (!getSession(req)) return res.status(401).json({ error: 'กรุณาเข้าสู่ระบบก่อน' });
      const date = String(req.query.date || '');
      if (!isYmd(date)) return res.status(400).json({ error: 'ต้องระบุ date=YYYY-MM-DD' });
      const f = jastranFile(date);
      if (!fs.existsSync(f)) return res.json({ date, count: 0, docs: [] });
      const j = JSON.parse(await fs.promises.readFile(f, 'utf8'));
      const docs: any[] = Array.isArray(j.docs) ? j.docs : [];

      // ใบที่บันทึกแล้ว — ต้องเทียบด้วย "กติกาเดียวกับกฎเหล็ก" ที่ใช้ตอนบันทึกจริง:
      //   ต่อสาขา + เทียบด้วย .trim() (ไม่ใช่ normDoc ที่ตัดอักขระพิเศษ) + เลขว่างไม่นับซ้ำ
      // ถ้าเทียบคนละแบบ จะบล็อกใบที่บันทึกได้จริง หรือปล่อยใบที่ซ้ำจริงผ่าน
      const db = await getDb();
      const sess = getSession(req)!;   // ผ่านด่านข้างบนมาแล้ว
      // กันเปิดใบข้ามงวด — ต้องตรงกับที่ /available กรองไว้ ไม่งั้นกดจาก URL ตรงๆ ก็ยังเห็น
      // ส่ง cycleId มาแต่หาไม่เจอ = พารามิเตอร์เพี้ยน -> ไม่ยอมปล่อยผ่านแบบไม่กรอง
      const cycIdQ = String(req.query.cycleId || '');
      const cycQ = cycIdQ ? db.cycles.find((c) => c.id === cycIdQ) : null;
      if (cycIdQ && !cycQ) return res.status(400).json({ error: 'ไม่พบรอบที่ระบุ' });
      // สาขาที่ดูได้: HQ ดูได้ทุกสาขา · สาขาปกติดูได้เฉพาะของตัวเอง (ไม่เชื่อ query จากผู้ใช้)
      const asked = String(req.query.branchId || '');
      // อ่านสิทธิ์จาก DB ปัจจุบัน ไม่เชื่อค่าที่ติดมากับ token (แพทเทิร์นเดียวกับ requireRateEditor)
      const me = db.branches.find((b) => b.id === sess.branchId);
      const isHQLike = !!(me?.isHQ || me?.isSystemUser);
      const branchId = isHQLike ? asked : sess.branchId;

      // ⭐ กรองใบตามสาขา — ใบจากจัสทรานไม่มี branchId (จัสทรานไม่รู้จักสาขาของเรา)
      //    แต่ Master รถผูกสาขาไว้แล้ว -> ใช้ "ทะเบียนรถ" เป็นตัวแยก
      //    ได้ทั้งความปลอดภัย (สาขาอื่นไม่เห็นใบเรา) และใช้งานง่ายขึ้น (ไม่ต้องไล่หาในใบทั้งวัน)
      //    ใบที่ทะเบียนไม่อยู่ใน Master เลย -> แสดงให้ทุกสาขาเห็น (ไม่งั้นตกหล่นไปเฉยๆ)
      // เก็บ "ทุกสาขา" ที่มีทะเบียนนี้ (ทะเบียนเดียวอาจลงไว้หลายสาขา — ดู /available)
      const plateOwner = new Map<string, Set<string>>();
      for (const v of db.vehicles) {
        const k = normPlate(v.plateNo || '');
        if (!k) continue;
        if (!plateOwner.has(k)) plateOwner.set(k, new Set());
        plateOwner.get(k)!.add(v.branchId);
      }
      // กรอง 2 ชั้น: ทะเบียนรถ + พื้นที่ (จากราคาที่สาขานั้นมีใน Master)
      //
      // ทำไมต้องมีชั้นพื้นที่: รถในจัสทรานมี 1,712 คัน แต่ Master รถเรามีแค่ 75 คัน
      // ทะเบียนส่วนใหญ่จึง "ไม่รู้จัก" -> ชั้นทะเบียนอย่างเดียวปล่อยผ่านหมด
      // ผลคือเชียงใหม่เห็นใบของนครสวรรค์/พิจิตร/ตาก ปนกันทั้งตาราง
      //
      // ใช้ Master ราคาเป็นตัวบอกพื้นที่: สาขาไหนมีราคาที่ปลายทางไหน = สาขานั้นวิ่งที่นั่น
      // (แหล่งความจริงเดียวกับที่ใช้คิดเงิน ไม่ต้อง sync รายการพื้นที่แยกอีกชุด)
      // เทียบชื่อด้วย textContains ตัวเดียวกับ matchRate -> ผลตรงกับตอนคิดราคา
      const inServiceArea = makeAreaCheck(branchId ? db.rateMasters.filter((r) => r.branchId === branchId) : []);
      const ovT = (db.destinationOverrides || []).filter((o) => !branchId || o.branchId === branchId);
      // โหมด "ทะเบียนไม่รู้จัก" — ใบที่ทะเบียนไม่มีใน Master รถ (หรือไม่มีทะเบียนเลย)
      // แยกออกมาหน้าต่างหาก เพราะเดิมมันกระจายไปปนกับใบปกติของทุกสาขา -> คนดูสับสน
      // ใบกลุ่มนี้ระบบไม่รู้ว่าเป็นของสาขาไหน จึงให้ทุกสาขาเห็นเหมือนกัน แล้วให้คนตัดสิน
      const unknownOnly = String(req.query.unknownPlate || '') === '1';
      // "ชี้ขาดสาขาไม่ได้" = ไม่รู้จักทะเบียน  หรือ  ทะเบียนอยู่หลายสาขาแต่ไม่มีสาขาไหนตรงพื้นที่
      // เคสหลัง (Codex ชี้): ถ้าไม่ดัก ใบจะหายจากทั้งหน้าปกติและหน้านี้ = ไม่ได้เงิน
      const isUnknownPlate = (d: any) => {
        const owners = plateOwner.get(normPlate(String(d?.plateNo || '')));
        if (!owners || owners.size === 0) return true;      // ไม่มีใน Master
        if (owners.size === 1) return false;                // ชี้ขาดได้ -> ไปหน้าปกติ
        // หลายสาขา: ถ้ามีสาขาเจ้าของสักรายที่พื้นที่ตรง -> ชี้ขาดได้
        return !Array.from(owners).some((b) => {
          const chk = makeAreaCheck(db.rateMasters.filter((r) => r.branchId === b));
          const ov = (db.destinationOverrides || []).filter((o) => o.branchId === b);
          return chk(realProvinceOf(d, ov));
        });
      };

      // กรองงวดที่ระดับใบ (documentDate) ไม่ใช่วันที่ของไฟล์ — ตรงกับตอนบันทึก
      const docs2 = cycQ ? docs.filter((d) => isDateInCycle(String(d?.documentDate || ''), cycQ)) : docs;
      const visible = unknownOnly
        ? docs2.filter(isUnknownPlate)
        : branchId
        ? docs2.filter((d) => {
            const owners = plateOwner.get(normPlate(String(d.plateNo || '')));
            // ทะเบียนอยู่สาขาเดียว -> ชี้ขาดได้
            if (owners && owners.size === 1) return owners.has(branchId);
            // ทะเบียนอยู่หลายสาขา -> ต้องเป็นหนึ่งในนั้น + พื้นที่ตรง
            // (ถ้าไม่มีสาขาไหนพื้นที่ตรงเลย ใบจะถูกจัดเป็น "ชี้ขาดไม่ได้" ไปหน้าทะเบียนไม่รู้จัก)
            if (owners && owners.size > 1) {
              if (!owners.has(branchId)) return false;
              // ใช้ปลายทาง "จริง" (ผ่านกฎแก้ปลายทางของสาขานี้) ไม่ใช่จังหวัดดิบในใบ
              return inServiceArea(realProvinceOf(d, ovT));
            }
            // ทะเบียนไม่รู้จัก -> ไม่แสดงในหน้าปกติ ไปอยู่หน้า "ทะเบียนไม่รู้จัก" แทน
            return false;
          })
        : docs2.filter((d) => !isUnknownPlate(d));   // HQ ไม่เลือกสาขา: ใบปกติทั้งหมด

      // ปกติเช็คซ้ำ "ต่อสาขา" ตามกฎเหล็ก
      // แต่หน้าทะเบียนไม่รู้จักทุกสาขาเห็นใบเดียวกัน -> ถ้าเช็คแค่สาขาตัวเอง
      // สาขา B จะเห็นใบที่สาขา A บันทึกไปแล้วเป็น "ตรวจใบนี้" กดเข้าไปตรวจเสียเวลา
      // แล้วโดน 409 ตอนกดบันทึก -> โหมดนี้จึงเช็คข้ามสาขา
      const savedByBranch = new Set(
        db.tripDocuments
          .filter((t) => unknownOnly || !branchId || t.branchId === branchId)
          .map((t) => (t.documentNo || '').trim())
          .filter(Boolean)
      );
      const withFlag = visible.map((d) => {
        const dn = String(d.documentNo || '').trim();
        // คำนวณ "ยังไม่ส่งเสร็จ" จากตัวเลขเสมอ ไม่พึ่ง flag จาก agent อย่างเดียว
        // (agent เก่า/ข้อมูลเก่าไม่มี _notDelivered -> ถ้าเชื่อ flag อย่างเดียวจะปล่อยใบ 0/N ผ่าน)
        return { ...d, _alreadySaved: dn ? savedByBranch.has(dn) : false, _notDelivered: isNotDelivered(d) };
      });
      res.json({ date, receivedAt: j.receivedAt || '', count: withFlag.length, total: docs.length, docs: withFlag });
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
      // สิทธิ์แก้ราคา: ธงในฐานข้อมูล หรือ HQ ตอนที่ยังไม่มีบัญชี admin เลย (bootstrap — Codex P1)
      const hasAnyEditor = db.branches.some((b) => b.canEditRates && b.status === 'active');
      const canEditRates = !!branch.canEditRates || (!hasAnyEditor && !!branch.isHQ);
      const token = newToken();
      sessions.set(token, { branchId: branch.id, name: branch.name, isHQ: !!branch.isHQ, canEditRates, at: Date.now() });
      res.json({
        ok: true, token,
        branch: { id: branch.id, name: branch.name, isHQ: !!branch.isHQ, canEditRates, isSystemUser: !!branch.isSystemUser },
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ด่านจัดการบัญชี/สาขา — ต้องเป็น HQ หรือ admin เท่านั้น
  // สำคัญด้านความปลอดภัย: ถ้าปล่อยเปิด ใครก็ยิง PUT /api/branches/:id ตั้ง canEditRates:true
  // ให้ตัวเอง แล้วข้ามด่านแก้ราคาทั้งหมดได้ (Codex P1)
  const requireBranchAdmin = async (req: any, res: any, next: any) => {
    try {
      const sess = getSession(req);
      if (!sess) return res.status(401).json({ error: 'กรุณาเข้าสู่ระบบใหม่ (เซสชันหมดอายุ)' });
      const db = await getDb();
      const me = db.branches.find((b) => b.id === sess.branchId);
      if (!me || me.status !== 'active') return res.status(403).json({ error: 'บัญชีนี้ถูกปิดใช้งานแล้ว' });
      if (!me.isHQ && !me.canEditRates) {
        return res.status(403).json({ error: 'จัดการสาขา/บัญชีได้เฉพาะสำนักงานใหญ่ (HQ) หรือผู้ดูแลราคา (admin)' });
      }
      // เฉพาะ HQ เท่านั้นที่แตะ "ฟิลด์ความปลอดภัย" ได้ — กันยกระดับสิทธิ์ตัวเอง (Codex P1)
      // ถ้าไม่ล็อกครบ: admin ตั้ง isHQ:true ให้ตัวเอง หรือแก้รหัสผ่านบัญชี HQ แล้วสวมสิทธิ์ HQ ได้
      const body = req.body || {};
      if (!me.isHQ) {
        const isCreate = req.method === 'POST' && !req.params?.id;
        // สร้างสาขาใหม่: ต้องตั้งรหัสผ่าน/สถานะได้ (จำเป็นต่อการใช้งาน ไม่ใช่การยกระดับสิทธิ์)
        // แต่ห้ามติดธงสิทธิ์มาตั้งแต่แรก — มอบสิทธิ์เป็นเรื่องของ HQ เท่านั้น (Codex P2)
        const SECURITY_FIELDS = isCreate
          ? ['canEditRates', 'isSystemUser', 'isHQ', 'id']
          : ['canEditRates', 'isSystemUser', 'isHQ', 'password', 'id', 'status'];
        const touched = SECURITY_FIELDS.filter((f) => Object.prototype.hasOwnProperty.call(body, f));
        if (touched.length) {
          return res.status(403).json({ error: `ตั้งค่า/แก้ไขข้อมูลด้านสิทธิ์ (${touched.join(', ')}) ได้เฉพาะสำนักงานใหญ่ (HQ)` });
        }
        // ห้ามแตะบัญชี HQ / บัญชีระบบอื่น (สร้าง/ลบ/แก้) แม้ไม่ได้ส่งฟิลด์ความปลอดภัย
        const targetId = String(req.params?.id || '');
        if (targetId) {
          const target = db.branches.find((b) => b.id === targetId);
          if (target && (target.isHQ || (target.isSystemUser && target.id !== me.id))) {
            return res.status(403).json({ error: 'แก้ไข/ลบบัญชีสำนักงานใหญ่หรือบัญชีระบบได้เฉพาะ HQ' });
          }
        }
        // สร้างสาขาใหม่: กันตั้งค่าสิทธิ์มาตั้งแต่แรก (ครอบคลุมโดย SECURITY_FIELDS ด้านบนแล้ว)
      }
      next();
    } catch (err: any) {
      return res.status(500).json({ error: err.message });
    }
  };

  // ต่ออายุ/ตรวจเซสชัน — client เรียกตอนเปิดแอปเพื่อรู้ว่า token ยังใช้ได้ + สิทธิ์ปัจจุบัน
  app.get('/api/session', async (req, res) => {
    const s = getSession(req);
    if (!s) return res.status(401).json({ error: 'เซสชันหมดอายุ' });
    try {
      const db = await getDb();
      const me = db.branches.find((b) => b.id === s.branchId);
      // บัญชีถูกลบ/ปิดใช้งานระหว่างที่ยังถือ token อยู่ -> ให้ client เด้งกลับหน้า login (Codex P2)
      if (!me || me.status !== 'active') {
        return res.status(401).json({ error: 'บัญชีนี้ถูกปิดใช้งานหรือถูกลบแล้ว — กรุณาเข้าสู่ระบบใหม่' });
      }
      const hasAnyEditor = db.branches.some((b) => b.canEditRates && b.status === 'active');
      // สิทธิ์ที่ใช้ได้จริง = ธงในฐานข้อมูล หรือ HQ ตอนที่ยังไม่มีบัญชี admin (bootstrap)
      const canEditRates = !!me?.canEditRates || (!hasAnyEditor && !!me?.isHQ);
      res.json({ ok: true, branchId: s.branchId, name: me?.name || s.name, isHQ: !!me?.isHQ, canEditRates, hasAnyEditor });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ===================== CLONE MASTER ระหว่างสาขา =====================
  // คัดลอก กฎตัวหาร/กลุ่มผู้รับ/alias/ประเภทเงิน/ผู้ส่งกล่อง จากสาขาต้นแบบ -> สาขาปลายทาง
  app.post('/api/branches/clone', requireBranchAdmin, async (req, res) => {
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

      // เขียนเฉพาะ node ที่แตะ (ทั้ง 5 เป็น id-keyed) — เดิม saveDb เขียนทั้ง tree 36MB
      // ทำให้ Firebase ตอบ "Write too large" เมื่อ DB โต -> คัดลอกกฎไม่สำเร็จ
      for (const key of ['receiverGroups', 'receiverGroupAliases', 'conversionRules', 'moneyCategories', 'manualBoxSenders'] as const) {
        await flushCollection(key);
      }
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
      // read เร็ว: รอแค่ snapshot โหลด (จาก disk เร็ว) ไม่รอ Firebase verify ที่อาจช้า
      //   ถ้าไม่มี snapshot -> await getDb() ปกติ; write path ยังใช้ getDb() ที่ verify แล้วเสมอ
      const db = (await peekCacheAsync()) || await getDb();
      if (peekCache()) void getDb().catch(() => {}); // เริ่ม verify Firebase พื้นหลัง (ครั้งถัดไปได้ของ verified)
      const branchId = typeof req.query.branchId === 'string' ? req.query.branchId : '';
      // lazy load ต่องวด: tripsCycleId=<id> -> ส่งใบเฉพาะงวดนั้น (state เล็กลงมาก ไม่โตตามจำนวนงวด)
      // tripsCycleId=current -> server เลือกงวดที่เปิดอยู่ให้ (logic เดียวกับ client auto-pick)
      // ไม่ส่ง param = ส่งเต็มเหมือนเดิม (Dashboard/รายงานที่ใช้ข้ามงวด + client เก่า)
      let tripsCycleId = typeof req.query.tripsCycleId === 'string' ? req.query.tripsCycleId : '';
      if (tripsCycleId === 'current') {
        const open = db.cycles.find((c) => c.status === 'open');
        tripsCycleId = open ? open.id : (db.cycles[db.cycles.length - 1]?.id || '');
      }
      const inBranch = <T extends { branchId?: string }>(arr: T[]) =>
        branchId ? arr.filter((x) => x.branchId === branchId) : arr;
      const trips = tripsCycleId
        ? db.tripDocuments.filter((t) => t.cycleId === tripsCycleId)
        : db.tripDocuments;
      // เฟส 2: น้ำมัน+รายการหัก ก็โตตามงวดเช่นกัน -> กรองต่องวดด้วย param เดียวกัน
      const fuelScoped = tripsCycleId ? db.fuelEntries.filter((f) => f.cycleId === tripsCycleId) : db.fuelEntries;
      const dedScoped = tripsCycleId ? db.deductions.filter((d) => d.cycleId === tripsCycleId) : db.deductions;
      const safe: DatabaseState = {
        ...db,
        branches: db.branches.map((b) => ({ ...b, password: '' })),
        vehicles: inBranch(db.vehicles),
        rateMasters: inBranch(db.rateMasters),
        rateOverrides: inBranch(db.rateOverrides),
        tripDocuments: inBranch(trips),
        fuelEntries: inBranch(fuelScoped),
        deductions: inBranch(dedScoped),
        receiverGroups: inBranch(db.receiverGroups),
        receiverGroupAliases: inBranch(db.receiverGroupAliases),
        conversionRules: inBranch(db.conversionRules),
        manualBoxSenders: inBranch(db.manualBoxSenders),
        destinationOverrides: inBranch(db.destinationOverrides),
        moneyCategories: inBranch(db.moneyCategories),
        tripDistances: inBranch(db.tripDistances || []),
      };
      // บอก client ว่า tripDocuments ในก้อนนี้ครอบคลุมแค่งวดไหน ('' = ทุกงวด)
      res.json({ ...safe, _tripsCycleId: tripsCycleId });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // เลขใบกระจายทุกงวดของสาขา (เบามาก ~ไม่กี่ร้อย KB) — ให้หน้า Review เช็คเลขซ้ำข้ามงวดได้
  // แม้ state จะโหลดใบแค่งวดเดียว (lazy) — กฎเหล็ก docNo ห้ามซ้ำทุกรอบยังเตือนก่อนบันทึกได้
  app.get('/api/doc-numbers', async (req, res) => {
    try {
      const db = (await peekCacheAsync()) || await getDb();
      const branchId = typeof req.query.branchId === 'string' ? req.query.branchId : '';
      const list = db.tripDocuments
        .filter((t) => !branchId || t.branchId === branchId)
        .map((t) => ({ documentNo: t.documentNo, cycleId: t.cycleId }));
      res.json(list);
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
      // ---- คัดลอก "ราคาเฉพาะรอบ" จากรอบก่อนหน้า 1 รอบ (ทุกสาขา) ----
      // เดิมขึ้นรอบใหม่แล้วราคาเฉพาะรอบหายหมด ต้องตั้งใหม่ทุกรอบ (สาย3 ~311 รายการ/รอบ)
      // กติกา (เจ้าของกำหนด): คัดลอกจาก "รอบล่าสุดที่มีอยู่จริง" ก่อนหน้ารอบใหม่
      //   - ปกติคือรอบติดกัน (ก.ย.1-15 <- ส.ค.16-31)
      //   - ถ้ารอบติดกันไม่เคยถูกสร้าง ให้ข้ามไปเอารอบเก่าสุดที่มีข้อมูลแทน (ยืนยันแล้ว: ดีกว่าได้ 0 รายการ)
      //   - คัดลอกทุกรายการ ไม่สนว่าราคาซ้ำกับราคาหลักหรือไม่
      //
      // ⚠️ ลำดับสำคัญ: ต้องคัดลอกราคา "ให้เสร็จก่อน" แล้วค่อยประกาศรอบใหม่ (flushCollection('cycles'))
      // ถ้าประกาศรอบก่อน แล้วมีคนสร้างรอบถัดไปจังหวะนั้นพอดี รอบใหม่จะมองเห็นรอบนี้เป็น "รอบก่อนหน้า"
      // ทั้งที่ยังไม่มีราคาเลย -> คัดลอกได้ 0 รายการ แล้วราคาหายเงียบๆ (Codex P2)
      let copiedOverrides = 0;
      let copiedFrom: string | null = null;
      try {
        // หา "รอบก่อนหน้า" = รอบที่เรียงตามเวลาแล้วอยู่ติดกันก่อนรอบใหม่ (ข้ามรอบที่ยังไม่มีอยู่จริง)
        const ord = (c: { year: number; month: number; half: string }) =>
          c.year * 100 + c.month * 2 + (c.half === 'first' ? 0 : 1);
        const target = ord(newCycle);
        const prev = db.cycles
          .filter((c) => c.id !== newCycle.id && ord(c) < target)
          .sort((a, b) => ord(b) - ord(a))[0];
        if (prev) {
          const src = db.rateOverrides.filter((o) => o.cycleId === prev.id);
          if (src.length) {
            const copies: RateOverride[] = src.map((o) => ({
              id: generateId('rov'), branchId: o.branchId, cycleId: newCycle.id,
              rateMasterId: o.rateMasterId, price: o.price, pieceThreshold: o.pieceThreshold ?? null,
            }));
            for (const c of copies) db.rateOverrides.push(c);
            await saveRecords('rateOverrides', copies);
            copiedOverrides = copies.length;
            copiedFrom = prev.name;
          }
        }
      } catch (e: any) {
        // คัดลอกพลาดไม่ทำให้สร้างรอบล้มเหลว — รอบยังเปิดได้ แค่ต้องตั้งราคาเอง
        console.error('[cycles] คัดลอกราคาเฉพาะรอบไม่สำเร็จ:', e.message);
      }

      // ประกาศรอบใหม่เป็นลำดับสุดท้าย (ราคาพร้อมแล้ว)
      db.cycles.push(newCycle);
      // เขียนเฉพาะ node ที่แก้ (เดิม saveDb เขียนทั้ง tree -> Firebase "Write too large" เมื่อ DB โต)
      await flushCollection('cycles');

      res.status(201).json({ ...newCycle, copiedOverrides, copiedFrom });
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
      // เขียนเฉพาะ node ที่แก้ (เดิม saveDb เขียนทั้ง tree -> Firebase "Write too large" เมื่อ DB โต)
      await flushCollection('cycles');
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
      // ลำดับสำคัญ: ลบ "รอบ" ให้ลงดิสก์ก่อน แล้วค่อยลบราคาเฉพาะรอบที่ผูกอยู่
      // ถ้าลบราคาก่อนแล้วเขียนรอบพลาด -> ราคาหายถาวรทั้งที่รอบยังอยู่ = ข้อมูลเงินหาย กู้ไม่ได้ (Codex P2)
      // สลับลำดับแล้ว กรณีแย่สุดคือเหลือ override กำพร้า ซึ่งไม่กระทบการคำนวณและเก็บกวาดทีหลังได้
      // เขียนเฉพาะ node ที่แก้ (เดิม saveDb เขียนทั้ง tree -> Firebase "Write too large" เมื่อ DB โต)
      await flushCollection('cycles');

      // ลบราคาเฉพาะรอบของรอบนี้ทิ้งด้วย ไม่งั้นกลายเป็นขยะกำพร้าถาวร
      // (สำคัญขึ้นมากตั้งแต่มีการคัดลอกอัตโนมัติ — รอบใหม่พกมาหลายร้อยรายการ)
      const orphans = db.rateOverrides.filter((o) => o.cycleId === id);
      if (orphans.length) {
        db.rateOverrides = db.rateOverrides.filter((o) => o.cycleId !== id);
        // ลบเฉพาะ id ที่เกี่ยว (multi-path update) — ห้าม flush ทั้ง node
        // เพราะ rateOverrides จะโตเร็วมากตั้งแต่มีคัดลอกอัตโนมัติ (~87+/รอบ) แล้วชน Firebase "Write too large"
        try {
          await removeRecords('rateOverrides', orphans.map((o) => o.id));
        } catch (e: any) {
          // รอบถูกลบไปแล้ว (สำเร็จ) — ลบราคาไม่สำเร็จไม่ควรทำให้ทั้ง request พัง
          console.error('[cycles] ลบราคาเฉพาะรอบที่กำพร้าไม่สำเร็จ:', e.message);
        }
      }
      res.json({ success: true, name, removedOverrides: orphans.length });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ลบ "ราคาเฉพาะรอบ" ที่ผูกกับราคาหลักที่ถูกลบไป
  // ถ้าไม่ลบตาม override จะกลายเป็นขยะกำพร้าที่ชี้ไปยัง rateMasterId ที่ไม่มีแล้ว
  // (เจอจริง 2026-08-22: 616 จาก 1,162 รายการเป็นกำพร้า) — ไม่ทำให้คิดเงินผิดเพราะจับคู่ไม่ติด
  // แต่กินที่ + ถูกคัดลอกต่อไปทุกรอบตั้งแต่มีคัดลอกอัตโนมัติ ทำให้พอกขึ้นเรื่อยๆ
  async function cascadeDeleteOverrides(db: DatabaseState, rateIds: string[]): Promise<number> {
    if (!rateIds.length) return 0;
    const set = new Set(rateIds);
    const doomed = db.rateOverrides.filter((o) => set.has(o.rateMasterId));
    if (!doomed.length) return 0;
    db.rateOverrides = db.rateOverrides.filter((o) => !set.has(o.rateMasterId));
    try {
      // ลบเฉพาะ id ที่เกี่ยว (multi-path) — ห้าม flush ทั้ง node กัน Firebase "Write too large"
      await removeRecords('rateOverrides', doomed.map((o) => o.id));
    } catch (e: any) {
      // ราคาหลักถูกลบไปแล้ว (สำเร็จ) — ลบ override ไม่สำเร็จไม่ควรทำให้ request พัง
      console.error('[rates] ลบราคาเฉพาะรอบที่ผูกกับราคาที่ลบไม่สำเร็จ:', e.message);
    }
    return doomed.length;
  }

  // ===================== Generic master CRUD helper =====================
  function masterRoutes<T extends { id: string }>(
    name: string,
    key: keyof DatabaseState,
    idPrefix: string,
    guard?: (req: any, res: any, next: any) => void   // ด่านสิทธิ์ (ถ้ามี) เช่น ราคาเฉพาะรอบต้องเป็น admin
  ) {
    const g = guard ?? ((_req: any, _res: any, next: any) => next());
    app.post(`/api/${name}`, g, async (req, res) => {
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
    app.put(`/api/${name}/:id`, g, async (req, res) => {
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
    app.delete(`/api/${name}/:id`, g, async (req, res) => {
      try {
        const db = await getDb();
        (db[key] as unknown as T[]) = (db[key] as unknown as T[]).filter((x) => x.id !== req.params.id) as any;
        if (isIdKeyed(key)) await removeRecord(key, req.params.id); else await saveDb(db);
        // ลบราคาหลัก -> ลบราคาเฉพาะรอบที่ผูกอยู่ด้วย (ทำหลังลบตัวหลักสำเร็จแล้ว)
        const removedOverrides = key === 'rateMasters' ? await cascadeDeleteOverrides(db, [req.params.id]) : 0;
        res.json({ success: true, ...(removedOverrides ? { removedOverrides } : {}) });
      } catch (err: any) {
        res.status(500).json({ error: err.message });
      }
    });
    // ลบเป็นกลุ่ม (ติ๊กหลายรายการแล้วลบทีเดียว)
    app.post(`/api/${name}/bulk-delete`, g, async (req, res) => {
      try {
        const ids: string[] = req.body?.ids || [];
        const db = await getDb();
        const idset = new Set(ids);
        (db[key] as unknown as T[]) = (db[key] as unknown as T[]).filter((x) => !idset.has(x.id)) as any;
        if (isIdKeyed(key)) await removeRecords(key, ids); else await saveDb(db);
        // ลบราคาหลัก -> ลบราคาเฉพาะรอบที่ผูกอยู่ด้วย (ทำหลังลบตัวหลักสำเร็จแล้ว)
        const removedOverrides = key === 'rateMasters' ? await cascadeDeleteOverrides(db, ids) : 0;
        res.json({ success: true, deleted: ids.length, ...(removedOverrides ? { removedOverrides } : {}) });
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

  masterRoutes<Branch>('branches', 'branches', 'br', requireBranchAdmin); // จัดการสาขา/บัญชี = HQ หรือ admin
  masterRoutes<RateOverride>('rate-overrides', 'rateOverrides', 'rov', requireRateEditor); // ราคาเฉพาะรอบ = admin เท่านั้น
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
  app.post('/api/rate-masters', requireRateEditor, async (req, res) => {
    try {
      const db = await getDb();
      const item: RateMaster = {
        ...req.body,
        id: generateId('rate'),
        createdBy: req.body.createdBy || 'user',
        createdAt: new Date().toISOString(),
      };
      db.rateMasters.push(item);
      await saveRecord('rateMasters', item); // เขียนแค่ record เดียว (ไม่เขียนทั้ง DB)
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
      // guard เพดานหลวม: ไฟล์เทมเพลตจริง ~ไม่กี่ร้อยแถว; เกิน 5,000 = ไฟล์ผิด/parse เพี้ยน ไม่ใช่งานจริง
      if (fuel.length > 5000) return res.status(422).json({ error: `ไฟล์มีรายการมากผิดปกติ (${fuel.length} แถว) — ตรวจว่าเลือกไฟล์เทมเพลตค่าน้ำมันถูกไฟล์ไหม` });
      const db = await getDb();
      let created = 0, skippedDup = 0;
      const closedCycles = new Set<string>();
      const createdCycles = new Set<string>();
      const createdEntries: FuelEntry[] = [];
      // นับว่าใบไปเข้ารอบไหนบ้าง — ไฟล์เทมเพลตปกติควรเป็นรอบเดียว; แตกหลายรอบ = วันที่ในไฟล์
      // น่าจะพิมพ์ผิด (เคสจริง: ใบ 3242 พิมพ์ 17/6 แทน 17/7 -> เข้ารอบ มิ.ย. เงียบๆ ยอดรอบขาด 1,000)
      const perCycle = new Map<string, { name: string; refs: string[] }>();
      // เลขใบสั่งเติมห้ามซ้ำในสาขา (เทียบกับของเดิม + ในไฟล์เดียวกัน)
      const seenRef = new Set(db.fuelEntries.filter((f) => f.branchId === branchId).map((f) => (f.refNo || '').trim()).filter(Boolean));
      const thDate = (iso: string) => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); return m ? `${+m[3]}/${+m[2]}/${m[1]}` : iso; };
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
        const pc = perCycle.get(rv.cycle.id) || { name: rv.cycle.name, refs: [] };
        pc.refs.push(`${rn || '(ไม่มีเลข)'} (${thDate(f.date)}, ${f.plateNo})`);
        perCycle.set(rv.cycle.id, pc);
      }
      await saveRecords('fuelEntries', createdEntries);
      if (createdCycles.size) await flushCollection('cycles');
      if (createdCycles.size) summary.push(`เปิดรอบใหม่อัตโนมัติ: ${[...createdCycles].join(', ')}`);
      // ⚠️ เตือนเมื่อใบในไฟล์เดียวกระจายเข้าหลายรอบ — ให้ตรวจวันที่ใบส่วนน้อยว่าพิมพ์ผิดเดือน/ปีไหม
      if (perCycle.size > 1) {
        const groups = [...perCycle.values()].sort((a, b) => b.refs.length - a.refs.length);
        const minority = groups.slice(1); // ทุกกลุ่มที่ไม่ใช่กลุ่มใหญ่สุด
        const detail = minority.map((g) => `${g.refs.join(', ')} -> เข้ารอบ "${g.name}"`).join(' | ');
        summary.push(`⚠️ ใบในไฟล์กระจายเข้า ${perCycle.size} รอบ (ส่วนใหญ่เข้า "${groups[0].name}") — โปรดตรวจวันที่ของ: ${detail} ว่าพิมพ์เดือน/ปีถูกต้องไหม (ถ้าผิด ให้ลบใบนั้นในรอบที่เข้าไป แล้วแก้วันที่ในไฟล์และนำเข้าใหม่)`);
      }
      if (skippedDup) summary.push(`⚠️ ข้ามเลขใบสั่งเติมที่ซ้ำ ${skippedDup} รายการ`);
      if (closedCycles.size) summary.push(`⚠️ ข้ามรายการของรอบที่ปิดอยู่: ${[...closedCycles].join(', ')} (ให้ HQ เปิดรอบก่อน)`);
      res.status(201).json({ success: true, created, summary });
    } catch (err: any) {
      console.error('import-fuel error:', err);
      res.status(500).json({ error: `นำเข้าค่าน้ำมันไม่สำเร็จ: ${err.message}` });
    }
  });

  // นำเข้า "ตารางราคา" จาก Excel (2 ชีต เหมาคัน/รายชิ้น) -> สร้าง rate masters
  // นำเข้าราคาแบบ "ผสาน (merge)": เทียบไฟล์กับราคาเดิมของสาขา แล้วบอกว่าอะไร ใหม่/อัปเดต/เท่าเดิม
  // - dryRun=1 -> ตรวจอย่างเดียว ไม่เขียน DB (ให้ผู้ใช้ดู diff ก่อนยืนยัน)
  // - ราคาที่มีในระบบแต่ไม่มีในไฟล์ -> คงไว้เฉยๆ (เจ้าของสั่ง: ไม่ลบอัตโนมัติ กันใบใหม่หาราคาไม่เจอ)
  // - อัปเดตราคาเก็บ rateMasterHistory เหมือนแก้ทีละช่อง เพื่อตรวจย้อนหลังได้
  app.post('/api/import-rates', requireRateEditor, async (req, res) => {
    try {
      const { branchId, fileBase64 } = req.body as { branchId: string; fileBase64: string };
      const dryRun = req.query.dryRun === '1' || req.query.dryRun === 'true' || req.body?.dryRun === true;
      if (!branchId) return res.status(400).json({ error: 'ต้องระบุสาขา' });
      if (!fileBase64) return res.status(400).json({ error: 'ต้องส่งไฟล์ Excel' });
      const buffer = Buffer.from(fileBase64.replace(/^data:[^;]+;base64,/, ''), 'base64');
      const { rates, summary } = parseRateExcel(buffer);
      if (!rates.length) return res.status(422).json({ error: 'อ่านไฟล์ไม่พบราคา — ตรวจสอบหัวคอลัมน์ จังหวัด/อำเภอ/ราคา' });
      // ด่านสุดท้ายก่อนเขียน DB (กันแม้ parser พลาด): ต้องมีจังหวัด + ราคา > 0 เสมอ
      // ส่วน "อำเภอ" ต้องมี ยกเว้นราคาที่ตั้งใจครอบทั้งจังหวัด (หมวดพิเศษ/มีเงื่อนไขจำกัด
      // เช่น เก็บคืนทั้งจังหวัด, CP All ขั้นบันได, คูห์เน่ตาม keyword) — ตรงกับกติกาใน parser
      const scoped = (r: any) => (r.productCategory || 'normal') !== 'normal' || !!r.rateGroup ||
        !!r.receiverKeyword || !!r.senderKeyword || !!r.productKeyword ||
        r.minQty != null || r.maxQty != null || r.pieceThreshold != null;
      const bad = rates.filter((r) => !String(r.provinceName || '').trim() || !(Number(r.price) > 0) ||
        (!String(r.districtName || '').trim() && !scoped(r)));
      if (bad.length) {
        const ex = bad.slice(0, 5).map((r) => `${r.provinceName || '(ไม่มีจังหวัด)'}/${r.districtName || '(ไม่มีอำเภอ)'} ฿${r.price}`).join(', ');
        return res.status(422).json({ error: `ไฟล์มี ${bad.length} แถวที่กรอกไม่ครบ (ต้องมีจังหวัด+ราคา>0 และงานปกติต้องระบุอำเภอ): ${ex}${bad.length > 5 ? ' ...' : ''}` });
      }
      // เพดานหลวมกันไฟล์ผิด/parse เพี้ยน (ตารางราคาจริงต่อสาขา ~ไม่กี่ร้อยแถว)
      if (rates.length > 5000) return res.status(422).json({ error: `ไฟล์มีราคามากผิดปกติ (${rates.length} แถว) — ตรวจว่าเลือกไฟล์ตารางราคาถูกไฟล์ไหม` });

      const db = await getDb();
      // กุญแจจับคู่ "ราคาเดียวกัน" — ต้องระบุเงื่อนไขให้ครบ ไม่งั้นราคาคนละเงื่อนไขจะทับกัน
      // *ไม่* รวมวันที่มีผล (effectiveFrom/To) เพราะไฟล์นำเข้าไม่มีคอลัมน์วันที่ (parser ตั้ง 2020-01-01 เสมอ)
      // ถ้าเอาวันที่มาเป็นกุญแจ ราคาที่ขึ้นรอบใหม่ (เช่น นครสวรรค์ effectiveFrom=2026-05-01) จะไม่ match
      // แล้วกลายเป็น "สร้างใหม่" ทั้งหมด = ราคาซ้ำซ้อนเต็ม DB และคิดเงินผิด
      const norm = (v: any) => String(v ?? '').trim().toLowerCase().replace(/\s+/g, '');
      const keyOf = (r: any) => [
        norm(r.provinceName), norm(r.districtName), r.productCategory || 'normal', r.priceType,
        norm(r.rateGroup), norm(r.receiverKeyword), norm(r.senderKeyword), norm(r.productKeyword),
        r.minQty ?? '', r.maxQty ?? '',
      ].join('|');

      // เทียบเฉพาะราคาที่ "ใช้งานอยู่ ณ วันนี้" — ราคาเก่าที่หมดอายุแล้วต้องคงไว้เป็นประวัติ ห้ามถูกทับ
      const today = new Date().toISOString().slice(0, 10);
      const inEffect = (r: RateMaster) =>
        (!r.effectiveFrom || r.effectiveFrom <= today) && (!r.effectiveTo || r.effectiveTo >= today);
      const all = db.rateMasters.filter((r) => r.branchId === branchId);
      const mine = all.filter((r) => r.status !== 'inactive' && inEffect(r));
      const expiredCount = all.length - mine.length;
      // ถ้ามีหลายแถวที่กุญแจเดียวกัน แปลว่าแยกจากกันไม่ได้ด้วยข้อมูลในไฟล์ -> ไม่แตะ กันอัปเดตผิดตัว
      const byKey = new Map<string, RateMaster>();
      const ambiguous = new Set<string>();
      for (const r of mine) {
        const k = keyOf(r);
        if (byKey.has(k)) ambiguous.add(k); else byKey.set(k, r);
      }

      const label = (r: any) => `${r.districtName || 'ทั้งจังหวัด'} จ.${r.provinceName}` +
        `${(r.productCategory || 'normal') !== 'normal' ? ` [${r.productCategory}]` : ''}` +
        ` (${r.priceType === 'flat' ? 'เหมา' : 'ชิ้น'})`;

      const created: any[] = [], updated: any[] = [], same: any[] = [];
      const seenKeys = new Set<string>();
      const dupInFile: string[] = [];
      const ambiguousRows: string[] = [];
      for (const r of rates) {
        const k = keyOf(r);
        if (seenKeys.has(k)) { dupInFile.push(label(r)); continue; } // แถวซ้ำในไฟล์เอง -> ใช้ตัวแรก
        seenKeys.add(k);
        if (ambiguous.has(k)) { ambiguousRows.push(label(r)); continue; } // ซ้ำในระบบ -> ข้าม ให้แก้มือ
        const old = byKey.get(k);
        if (!old) { created.push({ key: k, row: r, label: label(r), price: r.price }); continue; }
        // จุดตัดชิ้นไม่อยู่ในกุญแจ (แก้จุดตัด = แก้ราคาเดิม ไม่ใช่สร้างใหม่) จึงต้องเทียบตรงนี้ด้วย
        const thrOf = (x: any) => (x.pieceThreshold == null || x.pieceThreshold === '' ? null : Number(x.pieceThreshold));
        const oldThr = thrOf(old), newThr = thrOf(r);
        if (Number(old.price) !== Number(r.price) || oldThr !== newThr) {
          updated.push({ key: k, row: r, old, label: label(r), oldPrice: old.price, newPrice: r.price,
            oldThreshold: oldThr, newThreshold: newThr });
        } else same.push({ key: k, label: label(r) });
      }
      // ราคาที่ระบบมีแต่ไม่มีในไฟล์ — รายงานให้เห็น แต่ "ไม่แตะ" ตามที่เจ้าของกำหนด
      // ไม่นับกลุ่มที่กุญแจซ้ำ (ambiguous) ซ้ำเข้ามาอีก เพราะรายงานแยกไว้แล้ว ไม่งั้นผู้ใช้เห็นตัวเลขซ้ำซ้อน
      const missing = mine
        .filter((r) => !seenKeys.has(keyOf(r)) && !ambiguous.has(keyOf(r)))
        .map((r) => ({ label: label(r), price: r.price }));

      const preview = {
        branchId,
        createdCount: created.length, updatedCount: updated.length,
        sameCount: same.length, missingCount: missing.length,
        created: created.slice(0, 100).map((c) => ({ label: c.label, price: c.price })),
        updated: updated.slice(0, 100).map((u) => ({ label: u.label, oldPrice: u.oldPrice, newPrice: u.newPrice,
          oldThreshold: u.oldThreshold, newThreshold: u.newThreshold })),
        missing: missing.slice(0, 100),
        duplicateInFile: dupInFile.slice(0, 50),
        ambiguousCount: ambiguousRows.length,
        ambiguous: ambiguousRows.slice(0, 50),
        expiredCount,
        summary: [
          ...summary,
          ...(dupInFile.length ? [`⚠️ ในไฟล์มีแถวซ้ำกันเอง ${dupInFile.length} แถว — ใช้แถวแรก`] : []),
          ...(ambiguousRows.length ? [`⚠️ ข้าม ${ambiguousRows.length} แถว: ในระบบมีราคาเงื่อนไขเหมือนกันหลายรายการ แยกไม่ออกว่าจะอัปเดตตัวไหน — แก้ในหน้า Master โดยตรง (${ambiguousRows.slice(0, 3).join(', ')}${ambiguousRows.length > 3 ? ' ...' : ''})`] : []),
          ...(expiredCount ? [`ℹ️ ไม่นับราคาที่หมดอายุ/ปิดใช้แล้ว ${expiredCount} รายการ — คงไว้เป็นประวัติ`] : []),
          ...(missing.length ? [`ℹ️ มีในระบบแต่ไม่มีในไฟล์ ${missing.length} รายการ — คงไว้ตามเดิม (ไม่ลบ)`] : []),
        ],
      };
      if (dryRun) return res.json({ dryRun: true, ...preview });

      // ---- เขียนจริง: เพิ่มของใหม่ + อัปเดตที่ราคาเปลี่ยน (พร้อมบันทึกประวัติ) ----
      const now = new Date().toISOString();
      const newRows: RateMaster[] = created.map((c) => ({
        ...c.row, branchId, id: generateId('rate'), status: 'active',
        effectiveFrom: c.row.effectiveFrom || '2020-01-01', effectiveTo: c.row.effectiveTo ?? null,
        createdBy: 'import', createdAt: now,
      } as RateMaster));
      const histories: RateMasterHistory[] = [];
      const changedRows: RateMaster[] = [];
      for (const u of updated) {
        const idx = db.rateMasters.findIndex((x) => x.id === u.old.id);
        if (idx === -1) continue; // ถูกลบไประหว่างรอยืนยัน -> ข้าม
        const cur = db.rateMasters[idx];
        const curThr = cur.pieceThreshold == null ? null : Number(cur.pieceThreshold);
        const priceSame = Number(cur.price) === Number(u.row.price);
        if (priceSame && curThr === u.newThreshold) continue; // มีคนแก้ให้ตรงแล้ว -> ไม่ต้องเขียนซ้ำ
        // ประวัติบันทึกเฉพาะตอน "ราคา" เปลี่ยน (โครงสร้าง history เก็บ old/newPrice)
        if (!priceSame) {
          histories.push({
            id: generateId('rhist'), rateMasterId: cur.id, oldPrice: cur.price, newPrice: u.row.price,
            changedBy: 'import', changedAt: now, changeReason: 'นำเข้าราคาจาก Excel',
          });
        }
        // อัปเดตราคา + จุดตัดชิ้น — คงวันที่มีผล/เงื่อนไข/ผู้สร้างเดิมไว้ทั้งหมด
        db.rateMasters[idx] = { ...cur, price: u.row.price, pieceThreshold: u.newThreshold, updatedBy: 'import', updatedAt: now };
        changedRows.push(db.rateMasters[idx]);
      }
      for (const r of newRows) db.rateMasters.push(r);
      // อัปเดต cache ให้ครบก่อนเขียน — โหมด db.json เขียนไฟล์จาก cache ปัจจุบัน
      // ถ้า push ประวัติทีหลัง ประวัติจะหายเมื่อ process ดับก่อนการเขียนครั้งถัดไป
      for (const h of histories) db.rateMasterHistory.push(h);
      // เขียนแบบ granular (id-keyed) — ไม่แตะรายการอื่น ไม่ต้องเขียนทั้ง node
      if (newRows.length || changedRows.length) await saveRecords('rateMasters', [...newRows, ...changedRows]);
      if (histories.length) await saveRecords('rateMasterHistory', histories);

      res.status(201).json({
        success: true, ...preview,
        // ids ของราคาที่เปลี่ยน -> client เอาไปเช็คใบที่กระทบต่อได้ทันที (rate-impact)
        changedRateIds: [...changedRows.map((r) => r.id), ...newRows.map((r) => r.id)],
      });
    } catch (err: any) {
      console.error('import-rates error:', err);
      res.status(500).json({ error: `นำเข้าราคาไม่สำเร็จ: ${err.message}` });
    }
  });

  // นำเข้าราคาแบบชุด (เขียน DB ครั้งเดียว) — ใช้ตอน import ตารางราคาทั้งสาขา
  app.post('/api/rate-masters/bulk-create', requireRateEditor, async (req, res) => {
    try {
      const { rates } = req.body as { rates: Partial<RateMaster>[] };
      if (!Array.isArray(rates) || !rates.length) return res.status(400).json({ error: 'ต้องส่ง rates เป็น array' });
      const db = await getDb();
      const now = new Date().toISOString();
      const created: RateMaster[] = [];
      for (const r of rates) {
        const item = {
          productCategory: 'normal', effectiveFrom: '2020-01-01', effectiveTo: null, status: 'active',
          ...r, id: generateId('rate'), createdBy: r.createdBy || 'import', createdAt: now,
        } as RateMaster;
        db.rateMasters.push(item);
        created.push(item);
      }
      await saveRecords('rateMasters', created); // multi-path update = 1 round-trip
      res.status(201).json({ success: true, count: rates.length });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.put('/api/rate-masters/:id', requireRateEditor, async (req, res) => {
    try {
      const db = await getDb();
      const idx = db.rateMasters.findIndex((r) => r.id === req.params.id);
      if (idx === -1) return res.status(404).json({ error: 'ไม่พบราคา' });
      const old = db.rateMasters[idx];
      // เก็บประวัติถ้าราคาเปลี่ยน
      let hist: RateMasterHistory | null = null;
      if (typeof req.body.price === 'number' && req.body.price !== old.price) {
        hist = {
          id: generateId('rhist'),
          rateMasterId: old.id,
          oldPrice: old.price,
          newPrice: req.body.price,
          changedBy: req.body.updatedBy || 'user',
          changedAt: new Date().toISOString(),
          changeReason: req.body.changeReason || 'แก้ไขราคา',
        };
        db.rateMasterHistory.push(hist);
      }
      db.rateMasters[idx] = { ...old, ...req.body, id: old.id, updatedAt: new Date().toISOString() };
      await saveRecord('rateMasters', db.rateMasters[idx]); // เขียนแค่ราคาที่แก้
      if (hist) await saveRecord('rateMasterHistory', hist);
      res.json(db.rateMasters[idx]);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // แก้ราคาหลักหลายช่องทีเดียว = เขียน DB ครั้งเดียว (คงบันทึกประวัติราคาไว้)
  app.post('/api/rate-masters/bulk-update', requireRateEditor, async (req, res) => {
    try {
      const updates = (req.body?.updates || []) as { id: string; price?: number; pieceThreshold?: number | null; changeReason?: string; updatedBy?: string }[];
      if (!Array.isArray(updates) || !updates.length) return res.status(400).json({ error: 'ต้องส่ง updates เป็น array' });
      const byId = new Map(updates.map((u) => [u.id, u]));
      const db = await getDb();
      const now = new Date().toISOString();
      let n = 0;
      const changed: RateMaster[] = [];
      const hists: RateMasterHistory[] = [];
      db.rateMasters = db.rateMasters.map((r) => {
        const u = byId.get(r.id);
        if (!u) return r;
        n++;
        if (typeof u.price === 'number' && u.price !== r.price) {
          hists.push({
            id: generateId('rhist'), rateMasterId: r.id, oldPrice: r.price, newPrice: u.price,
            changedBy: u.updatedBy || 'user', changedAt: now, changeReason: u.changeReason || 'แก้ไขราคา (หลายช่อง)',
          });
        }
        const patch: any = {};
        if (u.price !== undefined) patch.price = u.price;
        if (u.pieceThreshold !== undefined) patch.pieceThreshold = u.pieceThreshold;
        const next = { ...r, ...patch, updatedAt: now };
        changed.push(next);
        return next;
      });
      db.rateMasterHistory.push(...hists);
      await saveRecords('rateMasters', changed); // เขียนเฉพาะรายการที่แก้ (1 round-trip)
      if (hists.length) await saveRecords('rateMasterHistory', hists);
      res.json({ success: true, count: n });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.delete('/api/rate-masters/:id', requireRateEditor, async (req, res) => {
    try {
      const db = await getDb();
      db.rateMasters = db.rateMasters.filter((r) => r.id !== req.params.id);
      await removeRecord('rateMasters', req.params.id); // ลบแค่ node เดียว
      // ลบราคาเฉพาะรอบที่ผูกกับราคานี้ด้วย ไม่งั้นเหลือเป็นขยะกำพร้าที่ชี้ไปยัง id ที่ไม่มีแล้ว
      const removedOverrides = await cascadeDeleteOverrides(db, [req.params.id]);
      res.json({ success: true, ...(removedOverrides ? { removedOverrides } : {}) });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/rate-masters/bulk-delete', requireRateEditor, async (req, res) => {
    try {
      const ids: string[] = req.body?.ids || [];
      const idset = new Set(ids);
      const db = await getDb();
      db.rateMasters = db.rateMasters.filter((r) => !idset.has(r.id));
      await removeRecords('rateMasters', ids); // ลบเฉพาะ node ที่ระบุ (1 round-trip)
      // ลบราคาเฉพาะรอบที่ผูกกับราคาเหล่านี้ด้วย (กันขยะกำพร้า)
      const removedOverrides = await cascadeDeleteOverrides(db, ids);
      res.json({ success: true, deleted: ids.length, ...(removedOverrides ? { removedOverrides } : {}) });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ตั้งวันเริ่ม/สิ้นสุดมีผลของหลายราคาพร้อมกัน (เขียนครั้งเดียว) — ใช้ตอนอัปเดตราคาทั้งสาขา (ปิดเก่า/เปิดใหม่)
  app.post('/api/rate-masters/bulk-set-effective', requireRateEditor, async (req, res) => {
    try {
      const { ids, effectiveTo, effectiveFrom } = req.body as { ids: string[]; effectiveTo?: string | null; effectiveFrom?: string };
      if (!Array.isArray(ids) || !ids.length) return res.status(400).json({ error: 'ต้องส่ง ids เป็น array' });
      const idset = new Set(ids);
      const db = await getDb();
      let n = 0;
      const now = new Date().toISOString();
      const changed: RateMaster[] = [];
      db.rateMasters = db.rateMasters.map((r) => {
        if (!idset.has(r.id)) return r;
        n++;
        const next = {
          ...r,
          ...(effectiveTo !== undefined ? { effectiveTo } : {}),
          ...(effectiveFrom !== undefined ? { effectiveFrom } : {}),
          updatedAt: now,
        };
        changed.push(next);
        return next;
      });
      await saveRecords('rateMasters', changed); // เขียนเฉพาะรายการที่แก้
      res.json({ success: true, count: n });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ราคาเฉพาะรอบ: สร้างหรืออัปเดต (1 รอบ + 1 ราคาหลัก = 1 override)
  app.post('/api/rate-overrides/upsert', requireRateEditor, async (req, res) => {
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
      await saveRecord('rateOverrides', o); // เขียนแค่ record เดียว (ไม่เขียนทั้ง DB) เบามากแม้ server ฟรี
      res.json(o);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ราคาเฉพาะรอบเป็นชุด (บันทึกหลายช่องทีเดียว = เขียน DB ครั้งเดียว กัน server ฟรีล่มตอนบันทึกเยอะ)
  app.post('/api/rate-overrides/bulk-upsert', requireRateEditor, async (req, res) => {
    try {
      const items = (req.body?.items || []) as RateOverride[];
      if (!Array.isArray(items) || !items.length) return res.status(400).json({ error: 'ต้องส่ง items เป็น array' });
      const db = await getDb();
      const changed: RateOverride[] = [];
      for (const it of items) {
        const { branchId, cycleId, rateMasterId, price, pieceThreshold } = it;
        if (!branchId || !cycleId || !rateMasterId) continue;
        let o = db.rateOverrides.find((x) => x.branchId === branchId && x.cycleId === cycleId && x.rateMasterId === rateMasterId);
        if (o) { o.price = price; o.pieceThreshold = pieceThreshold ?? null; }
        else { o = { id: generateId('rov'), branchId, cycleId, rateMasterId, price, pieceThreshold: pieceThreshold ?? null }; db.rateOverrides.push(o); }
        changed.push(o);
      }
      await saveRecords('rateOverrides', changed); // multi-path update = 1 round-trip (ไม่เขียนทั้ง DB)
      res.json({ success: true, count: changed.length });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ===================== TRIP DOCUMENTS =====================
  // บันทึก trip ที่ผ่าน Review (รับ extracted + คำนวณใหม่ฝั่ง server)
  app.post('/api/trips', async (req, res) => {
    try {
      const { extracted, fileName, branchId, overwrite } = req.body as {
        extracted: ExtractedTripDocument; fileName: string; branchId: string; overwrite?: boolean;
      };
      if (!branchId) return res.status(400).json({ error: 'ต้องระบุสาขา' });
      // guard เพดานหลวม: ใบจริงมี ~10-40 จุด; เกิน 1,000 = ข้อมูล parse เพี้ยน ไม่ใช่ใบจริง
      if ((extracted?.receipts?.length || 0) > 1000) {
        return res.status(422).json({ error: `ใบนี้มีใบรับมากผิดปกติ (${extracted.receipts.length} จุด) — ตรวจไฟล์ว่าถูกต้องไหม` });
      }
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

      // 🔒 กฎเหล็ก: ใบกระจายต้องส่งเสร็จอย่างน้อย 1 จุด ถึงคิดค่าเที่ยวได้
      //    UI ซ่อนปุ่มไว้แล้ว แต่ flag _notDelivered ถูกตัดทิ้งก่อนส่งมา (เป็น field แสดงผล)
      //    -> ตรวจย้อนกับไฟล์ต้นทางจากจัสทรานแทน ด่านนี้บายพาสจากฝั่ง client ไม่ได้
      //    ใบที่คีย์เอง/นำเข้าไฟล์ ไม่มีในไฟล์จัสทราน -> ไม่โดนด่านนี้ (ทำงานเหมือนเดิม)
      const srcDate = /^จัสทราน\s+(\d{4}-\d{2}-\d{2})$/.exec(String(fileName || ''))?.[1];
      if (srcDate) {
        try {
          const raw = JSON.parse(await fs.promises.readFile(jastranFile(srcDate), 'utf8'));
          const no = (trip.documentNo || '').trim();
          // ⚠️ ห้ามหาด้วยเลขใบอย่างเดียว — คนแก้เลขใบตอนตรวจได้ แก้แล้วจะหาไม่เจอ = หลุดด่าน
          //    จึงหาเผื่อด้วย "เลขใบรับ" ซึ่งไม่ได้อยู่ในช่องที่แก้กันตามปกติ
          const rcpKeys = new Set(
            (extracted?.receipts || []).map((r: any) => String(r?.receiptNo || '').trim()).filter(Boolean)
          );
          const src = (raw.docs || []).find((x: any) => {
            if (String(x?.documentNo || '').trim() === no) return true;
            if (!rcpKeys.size) return false;
            return (x?.receipts || []).some((r: any) => rcpKeys.has(String(r?.receiptNo || '').trim()));
          });
          if (src && isNotDelivered(src)) {
            return res.status(400).json({
              error: `ใบ ${no} ยังไม่ส่งเสร็จสักจุด (0/${src._totalReceipts ?? '?'}) — ยังคิดค่าเที่ยวไม่ได้ ต้องรอส่งของเสร็จก่อน`,
            });
          }
        } catch { /* ไม่มีไฟล์/อ่านไม่ได้ -> ปล่อยผ่าน (ไม่บล็อกงานเพราะไฟล์หาย) */ }
      }

      // 🔒 กฎเหล็ก: เลขใบกระจายห้ามซ้ำภายในสาขา (ทุกรอบ) — ซ้ำ = การเงินผิดเพี้ยน
      // overwrite:true = ผู้ใช้ยืนยันทับใบเดิม (ไฟล์แก้/พิมพ์ใหม่) -> ลบใบเดิมแล้วบันทึกใหม่แทน 409
      // NOTE: ยังไม่ลบตรงนี้ (Codex P2) — ลบหลัง validation ครบ กันข้อมูลหายเมื่อ validate ล้มเหลว
      const docNo = (trip.documentNo || '').trim();
      const dupsToRemove: TripDocument[] = [];
      if (docNo) {
        const dups = db.tripDocuments.filter(
          (t) => t.branchId === branchId && (t.documentNo || '').trim() === docNo
        );
        // 🔒 ด่านที่ 2: ซ้ำ "ข้ามสาขา" — หน้าดึงจากจัสทรานทำให้ใบทะเบียนไม่รู้จัก
        //    ปรากฏกับทุกสาขา ถ้าเช็คแค่สาขาตัวเอง สองสาขาจะบันทึกใบเดียวกันได้ = จ่ายซ้ำ
        //    ตรวจข้อมูลจริงแล้ว 1,012 ใบไม่มีเลขซ้ำข้ามสาขาเลย -> เพิ่มด่านนี้ไม่กระทบของเดิม
        //    ไม่ให้ overwrite ข้ามสาขา (ใบของสาขาอื่น ห้ามลบทับ)
        if (!dups.length) {
          const other = db.tripDocuments.find(
            (t) => t.branchId !== branchId && (t.documentNo || '').trim() === docNo
          );
          if (other) {
            const b = db.branches.find((x) => x.id === other.branchId);
            return res.status(409).json({
              error: `เลขใบกระจาย ${docNo} ถูกบันทึกไปแล้วโดยสาขา "${b?.name || other.branchId}" — ห้ามบันทึกซ้ำ (ถ้าเป็นใบของสาขาคุณจริง ให้สาขานั้นลบออกก่อน)`,
            });
          }
        }
        if (dups.length) {
          if (!overwrite) {
            const dup = dups[0];
            const dupCycle = db.cycles.find((c) => c.id === dup.cycleId);
            const where = dup.cycleId === cycle.id ? 'ในรอบนี้' : `ในรอบ "${dupCycle?.name || dup.cycleId}"`;
            return res.status(409).json({ error: `เลขใบกระจาย ${docNo} ซ้ำ — มีอยู่แล้ว${where} (ห้ามบันทึกซ้ำ ถ้าต้องการแก้ ให้ลบใบเดิมก่อน)`, canOverwrite: true, dupCycleName: dupCycle?.name || dup.cycleId });
          }
          // 🔒 ห้ามทับใบที่อยู่ในรอบปิด (ประวัติการเงินถูกล็อก) — กันแก้ย้อนหลังผ่านการทับ
          const closedDup = dups.find((d) => db.cycles.find((c) => c.id === d.cycleId)?.status === 'closed');
          if (closedDup) {
            const cn = db.cycles.find((c) => c.id === closedDup.cycleId)?.name || closedDup.cycleId;
            return res.status(400).json({ error: `ทับไม่ได้ — ใบเดิมอยู่ในรอบ "${cn}" ที่ปิดแล้ว (ประวัติล็อก) ให้ HQ เปิดรอบก่อนจึงแก้ได้` });
          }
          dupsToRemove.push(...dups); // ทับ: จำใบเดิมไว้ ลบทีหลัง (หลัง validate ครบ)
        }
      }

      // บังคับ: ผู้ส่งที่ต้องกรอกกล่อง แต่ยังไม่กรอก -> บันทึกไม่ได้
      const missingBox = trip.receipts.find((r) => r.requiresManualBox && (r.manualBoxQty == null || r.manualBoxQty <= 0));
      if (missingBox) {
        return res.status(400).json({ error: `ใบรับ ${missingBox.receiptNo}: ต้องกรอกจำนวนกล่องก่อนบันทึก (ผู้ส่งส่งเป็นชิ้น)` });
      }
      // ✅ ผ่านทุก validation แล้ว -> ค่อยลบใบเดิม (overwrite) กันข้อมูลหายเมื่อ validate ล้มเหลว
      const removedDupIds = dupsToRemove.map((d) => d.id);
      if (removedDupIds.length) db.tripDocuments = db.tripDocuments.filter((t) => !removedDupIds.includes(t.id));
      // ถ้าเป็นรอบใหม่ ค่อยเพิ่มลงระบบตอนนี้ (กันสร้างรอบเปล่าเวลาบันทึกไม่ผ่าน)
      if (resolved.created) db.cycles.push(cycle);
      trip.isVerified = true;
      db.tripDocuments.push(trip);
      // เขียนใบใหม่ก่อน แล้วค่อยลบใบเดิม (Codex P2) — ถ้า crash กลางทาง จะเหลือ 2 ใบ (กู้ได้)
      // ดีกว่าลบก่อน-เขียนไม่ทัน = ข้อมูลหายถาวร
      await saveRecord('tripDocuments', trip);           // เขียนใบใหม่ก่อน
      for (const id of removedDupIds) if (id !== trip.id) await removeRecord('tripDocuments', id); // แล้วลบเก่า
      if (resolved.created) await flushCollection('cycles'); // รอบใหม่ -> เขียน cycles (เล็ก)
      res.status(201).json({ ...trip, _cycle: cycle, _cycleCreated: resolved.created, _overwritten: removedDupIds.length });
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
  // ?dryRun=1 -> คำนวณเทียบ old vs new แล้วคืน diff โดยไม่บันทึก (ดูก่อนแตะข้อมูลจริง)
  app.post('/api/cycles/:id/recalculate', async (req, res) => {
    try {
      const dryRun = req.query.dryRun === '1' || req.query.dryRun === 'true';
      const db = await getDb();
      const cycle = db.cycles.find((c) => c.id === req.params.id);
      if (!cycle) return res.status(404).json({ error: 'ไม่พบรอบ' });

      // คำนวณใหม่จาก trip เดิม (ไม่แก้ cache จนกว่าจะ apply จริง)
      const recompute = (t: TripDocument): TripDocument => {
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
        return { ...recomputed, id: t.id, isVerified: t.isVerified, createdAt: t.createdAt };
      };

      // เลือกเฉพาะใบที่ระบุ (docNos) — สำหรับแก้ราคาเจาะจงหลัง master เปลี่ยน โดยไม่ให้
      // ราคา master ตัวอื่นที่แก้หลังใบบันทึกรั่วเข้าใบทั้งรอบ (เคยเกือบพลาด: recalc ทั้งรอบ
      // จะเปลี่ยน 126 ใบ +2,889 ทั้งที่ตั้งใจแก้ 6 ใบชัยนาท -274)
      const docNosRaw = Array.isArray(req.body?.docNos) ? req.body.docNos : null;
      const docNos: string[] = docNosRaw ? docNosRaw.map((s: any) => String(s).trim()).filter(Boolean) : [];
      // ส่ง docNos มาแต่กรองแล้วว่างหมด (เช่นเลขใบว่าง) -> ห้าม fallback เป็นทั้งรอบ (อันตราย) — 400 ชัดๆ
      if (docNosRaw && !docNos.length) {
        return res.status(400).json({ error: 'docNos ที่ส่งมาว่างทั้งหมด — ระบุเลขใบที่ต้องการ หรือไม่ส่ง docNos ถ้าต้องการทั้งรอบ' });
      }
      // เลขใบซ้ำได้ข้ามสาขา (unique ต่อสาขาเท่านั้น) — ระบุ scopeBranchId เมื่อใช้ docNos
      // เพื่อไม่ให้ recompute ใบสาขาอื่นที่บังเอิญเลขตรงกัน (Codex P2)
      const scopeBranchId: string = typeof req.body?.branchId === 'string' ? req.body.branchId : '';
      const docNoSet = new Set(docNos);
      let inCycle = db.tripDocuments.filter((t) => t.cycleId === cycle.id && (!scopeBranchId || t.branchId === scopeBranchId));
      let notFound: string[] = [];
      if (docNoSet.size) {
        const found = new Set(inCycle.map((t) => (t.documentNo || '').trim()));
        notFound = docNos.filter((n) => !found.has(n));
        inCycle = inCycle.filter((t) => docNoSet.has((t.documentNo || '').trim()));
        if (!inCycle.length) return res.status(404).json({ error: `ไม่พบใบที่ระบุในรอบนี้: ${docNos.join(', ')}` });
      }

      if (dryRun) {
        // จำแนกผลต่าง: <=0.10 บาท = ปัดเศษล้วน, >0.10 = ราคา/logic เปลี่ยน
        const ROUND_EPS = 0.10;
        const diffs = inCycle.map((t) => {
          const nw = recompute(t);
          const delta = round2(nw.tripAmount - t.tripAmount);
          return { docNo: t.documentNo, plate: t.plateNo, old: t.tripAmount, new: nw.tripAmount, delta };
        });
        const changed = diffs.filter((d) => Math.abs(d.delta) > 0.0001);
        const roundingOnly = changed.filter((d) => Math.abs(d.delta) <= ROUND_EPS);
        const priceChanged = changed.filter((d) => Math.abs(d.delta) > ROUND_EPS);
        const oldSum = round2(diffs.reduce((s, d) => s + d.old, 0));
        const newSum = round2(diffs.reduce((s, d) => s + d.new, 0));
        return res.json({
          dryRun: true,
          cycleId: cycle.id,
          cycleName: cycle.name,
          total: inCycle.length,
          changedCount: changed.length,
          roundingOnlyCount: roundingOnly.length,
          priceChangedCount: priceChanged.length,
          oldSum,
          newSum,
          totalDelta: round2(newSum - oldSum),
          // ใบที่เปลี่ยนราคา/logic (ต่าง > 0.10) เรียงตามผลต่างมากสุด — ต้องดูก่อน apply
          priceChanges: priceChanged.sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta)).slice(0, 50),
          ...(docNoSet.size ? { scopedTo: docNos, notFound } : {}),
        });
      }

      const changed: TripDocument[] = [];
      db.tripDocuments = db.tripDocuments.map((t) => {
        if (t.cycleId !== cycle.id) return t;
        if (scopeBranchId && t.branchId !== scopeBranchId) return t;                // จำกัดสาขา (กันเลขซ้ำข้ามสาขา)
        if (docNoSet.size && !docNoSet.has((t.documentNo || '').trim())) return t; // จำกัดเฉพาะใบที่ระบุ
        const out = recompute(t);
        changed.push(out);
        return out;
      });
      await saveRecords('tripDocuments', changed);
      res.json({ success: true, count: changed.length, ...(docNoSet.size ? { scopedTo: docNos, notFound } : {}) });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ตรวจผลกระทบหลังแก้ราคา: ใบที่บันทึกแล้วในรอบนี้ ที่ "ราคาที่เพิ่งแก้" match และยอดจะเปลี่ยน
  // -> UI แสดงแถบเตือน + ปุ่มอัปเดต (เรียก recalculate?docNos) — แก้ปัญหาถาวรเคสชัยนาท:
  // แก้ราคา/override หลังใบถูกบันทึก ใบเก่าค้างราคาเดิมเงียบๆ ไม่มีใครรู้จนยอดไม่ตรง
  app.post('/api/cycles/:id/rate-impact', async (req, res) => {
    try {
      const { rateMasterIds, branchId, extraProvinces, broad } = req.body as {
        rateMasterIds: string[]; branchId?: string; extraProvinces?: string[]; broad?: boolean;
      };
      if (!Array.isArray(rateMasterIds) || !rateMasterIds.length) {
        return res.status(400).json({ error: 'ต้องระบุ rateMasterIds' });
      }
      // จังหวัดเพิ่มเติมที่ต้องตรวจด้วย — เช่นแก้ "จังหวัด" ของ rate ผ่านฟอร์ม: ใบจังหวัดเดิม
      // เคย match ตัวเก่าแต่ไม่ match ตัวใหม่ -> client ส่งจังหวัดเดิมมาให้ตรวจ (Codex P2)
      const extraProvs: string[] = Array.isArray(extraProvinces)
        ? extraProvinces.map((s) => String(s).trim()).filter(Boolean) : [];
      const db = await getDb();
      const cycle = db.cycles.find((c) => c.id === req.params.id);
      if (!cycle) return res.status(404).json({ error: 'ไม่พบรอบ' });
      if (cycle.status === 'closed') return res.json({ closed: true, checked: 0, affected: [], totalDelta: 0 });
      const rates = db.rateMasters.filter((r) => rateMasterIds.includes(r.id));
      if (!rates.length) return res.status(404).json({ error: 'ไม่พบราคาที่ระบุ' });
      const bId = branchId || rates[0].branchId || '';

      // ใบในรอบ+สาขา ที่ "อาจ" เกี่ยวกับราคาที่แก้ — จับกว้างระดับจังหวัด (ไม่ใช้ matchRate แคบๆ)
      // เหตุผล (Codex P2 สองรอบ): แก้อำเภอ/วันที่มีผล/สถานะ/เงื่อนไขพิเศษ ทำให้ใบที่เคย match
      // ตัวเก่าไม่ match ตัวใหม่ -> ถ้ากรองด้วย rate ใหม่จะหลุด. จังหวัดกว้างพอครอบทั้งเก่า/ใหม่
      // (recompute + delta จริงเป็นตัวกรองชั้นสุดท้าย จึงไม่รายงานเกินจริง; recompute เพิ่มไม่กี่ใบ)
      // limitation: ถ้าแก้ "จังหวัด" ของ rate เอง ใบจังหวัดเดิมจะไม่ถูกตรวจ (พบยาก — เท่ากับสร้างเส้นทางใหม่)
      const provinceHit = (prov: string, r: RateMaster) =>
        textContains(prov, r.provinceName) || textContains(prov, (r as any).provinceShort || '') || textContains(r.provinceName || '', prov);
      // rate แบบ keyword/หมวดพิเศษ (เช่น collect_back จับด้วยชื่อผู้ส่ง ไม่เช็คจังหวัด — Codex P2)
      // -> จำกัดจังหวัดไม่ได้ ตรวจทั้งรอบ+สาขาไปเลย (แก้ rate พวกนี้นานๆ ครั้ง ยอม recompute เพิ่ม)
      // broad=true จาก client: rate "ก่อนแก้" เคยเป็น keyword-based (client เห็นค่าเดิม server ไม่เห็น
      // หลัง update แล้ว) เช่นแปลง collect_back->normal + ล้าง keyword -> ตรวจทั้งรอบเช่นกัน
      const keywordBased = broad === true || rates.some((r) =>
        (r.productCategory || 'normal') !== 'normal' || !!r.receiverKeyword || !!r.senderKeyword || !!r.productKeyword);
      // ใบที่ปลายทางถูกแก้ด้วย DestinationOverride (keyword -> จังหวัดจริง): provinceRaw ในใบ
      // เป็นค่าผิดจาก PDF แต่ recompute ใช้จังหวัดจริง -> ถ้าจังหวัดจริงของ override match rate
      // ที่แก้ ให้รวมใบที่มี keyword นั้นด้วย (Codex P2)
      const relevantOv = (db.destinationOverrides || []).filter((o) =>
        o.status === 'active' && (!bId || o.branchId === bId) && rates.some((r) => provinceHit(o.province || '', r)));
      // เช็คทุก field ที่ computeReceipt ใช้จับ override: docNote, ผู้รับ, ผู้ส่ง, ชื่อสินค้า (Codex P2)
      const hitOverride = (t: TripDocument) => relevantOv.some((o) =>
        textContains(t.docNote || '', o.keyword) ||
        (t.receipts || []).some((rc) =>
          textContains(rc.receiverName || '', o.keyword) ||
          textContains(rc.senderName || '', o.keyword) ||
          (rc.items || []).some((it) => textContains(it.productName || '', o.keyword))));
      const inCycle = db.tripDocuments.filter((t) => t.cycleId === cycle.id && (!bId || t.branchId === bId));
      const related = keywordBased ? inCycle : inCycle.filter((t) =>
        (t.receipts || []).some((rc) => {
          const prov = rc.provinceRaw || t.provinceRaw || '';
          return rates.some((r) => provinceHit(prov, r)) ||
            extraProvs.some((p) => textContains(prov, p) || textContains(p, prov));
        }) || hitOverride(t)
      );

      // recompute เฉพาะใบที่เกี่ยว -> เทียบยอด (ไม่บันทึกอะไร — read only)
      const affected: { docNo: string; plate: string; old: number; new: number; delta: number }[] = [];
      for (const t of related) {
        const extracted: ExtractedTripDocument = {
          documentNo: t.documentNo,
          documentDate: t.documentDate,
          plateNo: t.plateNo,
          provinceRaw: t.provinceRaw,
          districtRaw: t.districtRaw,
          rateChoice: t.rateType ?? undefined,
          docNote: t.docNote,
          receipts: t.receipts.map((r) => ({
            receiptNo: r.receiptNo, receiverName: r.receiverName, senderName: r.senderName,
            items: r.items, provinceRaw: r.provinceRaw, districtRaw: r.districtRaw,
            manualBoxQty: r.manualBoxQty ?? undefined,
          })),
        };
        const nw = recomputeTrip(db, cycle, extracted, t.fileName, t.branchId);
        const delta = round2(nw.tripAmount - t.tripAmount);
        if (Math.abs(delta) > 0.0001) {
          affected.push({ docNo: t.documentNo, plate: t.plateNo, old: t.tripAmount, new: nw.tripAmount, delta });
        }
      }
      affected.sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));
      res.json({
        cycleId: cycle.id, cycleName: cycle.name, checked: related.length,
        affected, totalDelta: round2(affected.reduce((s, a) => s + a.delta, 0)),
      });
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
- โน้ตท้ายใบ (docNote): ข้อความหมายเหตุเส้นทาง/เงื่อนไขราคาที่อยู่ท้ายเอกสาร เช่น "วิ่งย่อยไม่เกิน 13 จุด" หรือ "วิ่งย่อยไม่เกิน 13 จุด (กลุ่ม2)" — ถ้ามีให้ดึงมาทั้งบรรทัดตามที่เห็น, ถ้าไม่มีเว้นว่าง
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
              docNote: { type: Type.STRING, description: 'โน้ตท้ายใบ เช่น "วิ่งย่อยไม่เกิน 13 จุด" (ถ้ามี)' },
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

  // ===================== [ทดลอง] route แยก (ลบได้ ไม่กระทบของเดิม) =====================
  registerExperimentalRoutes(app, { getDb, saveRecord, removeRecord, flushCollection, genId: generateId });
  startOilPriceScheduler({ getDb, saveRecord, genId: generateId }); // บันทึกราคาน้ำมันอัตโนมัติ 05:30 น. เวลาไทย

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

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running at http://0.0.0.0:${PORT}`);
    // warm cache ตอน boot: โหลด snapshot ก่อน (peekCache พร้อม -> /api/state ตอบได้แม้ Firebase ช้า)
    // แล้ว verify กับ Firebase พื้นหลัง; ไม่ block listen (พอร์ตเปิดแล้ว config ตอบได้ทันที)
    warmCacheOnBoot().then(() => console.log('[boot] cache พร้อมใช้งาน'))
      .catch((e) => console.error('[boot] warm cache ล้มเหลว (จะ retry ตอน request):', e.message));
  });
}

startServer().catch((error) => console.error('Failed to start server:', error));
