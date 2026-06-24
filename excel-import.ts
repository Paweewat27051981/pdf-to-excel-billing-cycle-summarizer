// ============================================================================
// Import ใบกระจายจากไฟล์ Excel (.xls/.xlsx) — parse ตรง ๆ ไม่ใช้ AI
// ข้อความใน Excel เป๊ะ 100% -> ชื่อสินค้าถูกต้อง -> ตัวหารจับอัตโนมัติ + ฟรี
// คืนโครงสร้าง ExtractedTripDocument[] เหมือนที่ Gemini สร้าง (เข้า flow Review เดิม)
// ============================================================================
import * as XLSX from 'xlsx';
import { ExtractedTripDocument, ExtractedReceipt } from './src/types.js';

type Cols = { seq: number[]; receiver: number[]; sender: number[]; qty: number[]; unit: number[]; product: number[]; receipt: number[] };

// ค่าเริ่มต้น (รูปแบบ 11111.xls) — ใช้เมื่อหาหัวตารางไม่เจอ
const COL: Cols = {
  seq: [0, 2], receiver: [3, 8], sender: [9, 18], qty: [18, 20], unit: [21, 23], product: [24, 32], receipt: [33, 49],
};

type Row = any[];

// จับตำแหน่งคอลัมน์จาก "หัวตาราง" อัตโนมัติ (รองรับ layout ที่คอลัมน์เลื่อน เช่น ใบปุ๋ย)
function detectCols(rows: Row[]): Cols | null {
  for (const row of rows) {
    const find = (label: string) => row.findIndex((c) => String(c).includes(label));
    const seqH = find('ลำดับ');
    const qtyH = find('จำนวน');
    const prodH = find('รายการ');
    const recvH = find('ผู้รับ');
    const sendH = find('ผู้ส่ง');
    const unitH = find('หน่วย');
    const receiptH = find('เลขที่ใบรับ');
    if ([seqH, qtyH, prodH, recvH, sendH, unitH, receiptH].some((x) => x < 0)) continue;
    // ช่วงข้อมูลอ้างอิงตำแหน่งหัวตาราง (ข้อมูลมักอยู่ซ้ายของ label เล็กน้อยเพราะ merge cell)
    return {
      seq: [seqH, Math.max(seqH, recvH - 1)],
      receiver: [seqH + 1, sendH - 1],
      sender: [recvH + 1, qtyH - 1],
      qty: [qtyH, unitH - 1],
      unit: [unitH, prodH - 1],
      product: [prodH, receiptH - 1],
      receipt: [receiptH - 1, receiptH + 18],
    };
  }
  return null;
}

function txt(row: Row, range: number[]): string {
  for (let i = range[0]; i <= range[1] && i < row.length; i++) {
    const v = row[i];
    if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

function num(row: Row, range: number[]): number | null {
  for (let i = range[0]; i <= range[1] && i < row.length; i++) {
    const v = row[i];
    if (typeof v === 'number' && !Number.isNaN(v)) return v;
    if (typeof v === 'string' && v.trim() !== '' && !Number.isNaN(Number(v))) return Number(v);
  }
  return null;
}

// คืน index คอลัมน์ของตัวเลขตัวแรกในช่วง (ไว้ยึดตำแหน่ง "จำนวน" จริง)
function numCol(row: Row, range: number[]): number {
  for (let i = range[0]; i <= range[1] && i < row.length; i++) {
    const v = row[i];
    if (typeof v === 'number' && !Number.isNaN(v)) return i;
    if (typeof v === 'string' && v.trim() !== '' && !Number.isNaN(Number(v))) return i;
  }
  return -1;
}

// คำที่เป็น "หน่วยนับ" — ใช้แยกหน่วยออกจากชื่อสินค้า เวลาคอลัมน์เลื่อน
const UNIT_RE = /^(กล่อง|ชิ้น|ลัง|หีบ|แพ็?ก|แพ็?ค|ถุง|กระสอบ|อัน|ห่อ|แผง|ขวด|มัด|ตัว|ใบ|ชุด|โหล|กก\.?|กิโล)\.?$/;

// อ่าน "หน่วย" และ "ชื่อสินค้า" จากเนื้อหาหลังคอลัมน์จำนวน (กันคอลัมน์เลื่อนจากหัวตาราง)
// หน่วย = คำที่ตรง UNIT_RE ตัวแรก, ชื่อสินค้า = ข้อความตัวแรกที่ไม่ใช่หน่วย/ตัวเลข
function extractUnitProduct(row: Row, qtyCol: number, endCol: number): { unit: string; product: string } {
  let unit = '', product = '';
  const start = (qtyCol >= 0 ? qtyCol : -1) + 1;
  for (let i = start; i <= endCol && i < row.length; i++) {
    const v = String(row[i] ?? '').trim();
    if (!v) continue;
    if (!Number.isNaN(Number(v))) continue; // ข้ามตัวเลข
    if (UNIT_RE.test(v)) { if (!unit) unit = v; continue; }
    if (!product) product = v;
  }
  return { unit, product };
}

// หาค่าที่อยู่ถัดจาก label ในแถวเดียวกัน (เช่น "ใบกระจายเลขที่" -> "JB...")
function labelValue(row: Row, label: string): string {
  for (let i = 0; i < row.length; i++) {
    if (String(row[i]).includes(label)) {
      for (let j = i + 1; j < row.length; j++) {
        const v = row[j];
        if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
      }
    }
  }
  return '';
}

function labelNumber(row: Row, label: string): number | null {
  for (let i = 0; i < row.length; i++) {
    if (String(row[i]).includes(label)) {
      for (let j = i + 1; j < row.length; j++) {
        const v = row[j];
        if (typeof v === 'number') return v;
        if (typeof v === 'string' && v.trim() !== '' && !Number.isNaN(Number(v))) return Number(v);
      }
    }
  }
  return null;
}

// Excel serial -> YYYY-MM-DD (ระบบวันที่ 1900: serial 1 = 1900-01-01)
function excelDate(serial: number | null): string {
  if (serial == null || Number.isNaN(serial)) return new Date().toISOString().slice(0, 10);
  // 1899-12-30 เป็นฐาน (ชดเชย bug ปี 1900 ของ Excel)
  const ms = Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000;
  return new Date(ms).toISOString().slice(0, 10);
}

// เลขที่ใบรับ เช่น B0926159948 (ตัวอักษร 1 ตัว + ตัวเลขยาว)
function findReceiptNo(row: Row, cols: Cols): string {
  const raw = txt(row, cols.receipt);
  if (/^[A-Za-z]\d{5,}/.test(raw)) return raw;
  // เผื่อ shift: สแกนทั้งแถว
  for (const v of row) {
    const s = String(v).trim();
    if (/^[A-Za-z]\d{5,}$/.test(s)) return s;
  }
  return raw;
}

function blankDoc(): ExtractedTripDocument {
  return { documentNo: '', documentDate: '', plateNo: '', provinceRaw: '', districtRaw: '', receipts: [] };
}

/**
 * แปลงไฟล์ Excel ใบกระจาย -> ExtractedTripDocument[]
 * (1 ไฟล์อาจมีหลายชีต/หลายใบกระจาย)
 */
export function parseDistributionExcel(buffer: Buffer): ExtractedTripDocument[] {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false });
  const docs: ExtractedTripDocument[] = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rows: Row[] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
    const cols = detectCols(rows) || COL; // จับคอลัมน์จากหัวตาราง ไม่งั้นใช้ค่าเริ่มต้น

    let doc: ExtractedTripDocument | null = null;
    let curProvince = '';
    let curDistrict = '';
    let curReceipt: ExtractedReceipt | null = null;

    const pushDoc = () => {
      if (doc && doc.receipts.length) {
        if (!doc.provinceRaw && doc.receipts[0]) doc.provinceRaw = doc.receipts[0].provinceRaw || '';
        if (!doc.districtRaw && doc.receipts[0]) doc.districtRaw = doc.receipts[0].districtRaw || '';
        docs.push(doc);
      }
    };

    for (const row of rows) {
      const joined = row.map((c) => String(c)).join(' ');

      // --- หัวเอกสาร: เริ่มใบกระจายใหม่ ---
      if (joined.includes('ใบกระจายเลขที่')) {
        pushDoc();
        doc = blankDoc();
        doc.documentNo = labelValue(row, 'ใบกระจายเลขที่');
        doc.documentDate = excelDate(labelNumber(row, 'วันที่ออก'));
        // บางรูปแบบ ทะเบียนรถอยู่แถวเดียวกับใบกระจายเลขที่ -> ดึงเลย
        if (joined.includes('ทะเบียนรถ')) doc.plateNo = labelValue(row, 'ทะเบียนรถ') || doc.plateNo;
        curReceipt = null;
        continue;
      }
      if (!doc) continue;

      // --- ทะเบียนรถ ---
      if (joined.includes('ทะเบียนรถ')) {
        doc.plateNo = labelValue(row, 'ทะเบียนรถ');
        continue;
      }

      // --- จังหวัด/อำเภอ (กำหนดปลายทางของใบรับถัด ๆ ไป) ---
      if (joined.includes('จังหวัด') && joined.includes('อำเภอ')) {
        curProvince = labelValue(row, 'จังหวัด') || curProvince;
        curDistrict = labelValue(row, 'อำเภอ') || curDistrict;
        if (!doc.provinceRaw) doc.provinceRaw = curProvince;
        if (!doc.districtRaw) doc.districtRaw = curDistrict;
        continue;
      }

      // --- จบรายการสินค้า ---
      if (joined.includes('ยอดรวมสินค้า')) {
        curReceipt = null;
        continue;
      }

      // --- ข้ามแถวหัวตาราง ---
      if (joined.includes('ผู้รับสินค้า') && joined.includes('ผู้ส่งสินค้า')) continue;

      // บางรูปแบบ เลขที่ใบรับอยู่ "แถวถัดจาก" แถวสินค้า -> เติมให้ใบรับล่าสุดที่ยังว่าง
      if (curReceipt && !curReceipt.receiptNo) {
        const rn = findReceiptNo(row, cols);
        if (rn) curReceipt.receiptNo = rn;
      }

      // --- แถวรายการสินค้า: ต้องมีจำนวน(ตัวเลข) + ชื่อสินค้า ---
      // จับ "จำนวน" จากตำแหน่งเลขจริง แล้วอ่าน หน่วย/ชื่อสินค้า ตามเนื้อหา (กันคอลัมน์เลื่อนจากหัวตาราง)
      const qtyCol = numCol(row, cols.qty);
      // ปัดทศนิยม 2 ตำแหน่ง กัน float error จาก Excel (เช่น 1.9999999999 -> 2)
      const qty = qtyCol >= 0 ? Math.round((Number(row[qtyCol]) + Number.EPSILON) * 100) / 100 : null;
      const { unit, product } = extractUnitProduct(row, qtyCol, cols.receipt[0]);
      if (qty == null || !product) continue;

      const seq = num(row, cols.seq);
      if (seq != null && Number.isInteger(seq)) {
        // เริ่มใบรับใหม่
        curReceipt = {
          receiptNo: findReceiptNo(row, cols),
          receiverName: txt(row, cols.receiver),
          senderName: txt(row, cols.sender),
          provinceRaw: curProvince,
          districtRaw: curDistrict,
          items: [],
        };
        doc.receipts.push(curReceipt);
      }
      if (!curReceipt) {
        // รายการต่อเนื่องแต่ยังไม่มีใบรับ -> สร้างใบรับชั่วคราว
        curReceipt = { receiptNo: '', receiverName: '', senderName: '', provinceRaw: curProvince, districtRaw: curDistrict, items: [] };
        doc.receipts.push(curReceipt);
      }
      curReceipt.items.push({ productName: product, quantity: qty, ...(unit ? { unit } : {}) });
    }

    pushDoc();
  }

  return docs;
}

// ============================================================================
// Import "ตารางราคาขนส่ง" จาก Excel -> rate masters (เหมาต่ออำเภอ + ชิ้นต่อจังหวัด/อำเภอ)
// 2 ชีต: "เหมาคัน" (จังหวัด/อำเภอ/ราคาเหมา) + "รายชิ้น" (จังหวัด/อำเภอ-เว้นว่าง=ทั้งจังหวัด/ราคาชิ้น)
// ตรรกะ max mode: ระบบเทียบ (จำนวน×ชิ้น) กับ เหมา จ่ายอันสูงกว่า (ไม่ต้องตั้ง threshold)
// ============================================================================
export interface ParsedRate {
  provinceName: string; provinceShort: string; districtName: string; destinationName: string;
  priceType: 'flat' | 'piece'; price: number; productCategory: 'normal'; status: 'active';
  effectiveFrom: string; effectiveTo: null;
}
const normTxt = (s: any) => String(s ?? '').toLowerCase().replace(/\s+/g, '').trim();
const stripAmphoe = (s: any) => String(s ?? '').replace(/^\s*อำเภอ\s*/, '').trim();
function findHeaderRow(rows: Row[], kw: string): number {
  return rows.findIndex((r) => r.some((c) => String(c).includes(kw)));
}
function colOf(header: Row, ...kws: string[]): number {
  return header.findIndex((c) => kws.some((k) => String(c).includes(k)));
}

export function parseRateExcel(buffer: Buffer): { rates: ParsedRate[]; summary: string[] } {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false });
  const findSheet = (kw: string) => wb.SheetNames.find((n) => n.includes(kw));
  const flatName = findSheet('เหมา') || wb.SheetNames[0];
  const pieceName = findSheet('ชิ้น');
  const summary: string[] = [];

  // ---- ราคาชิ้น: default ต่อจังหวัด + override ต่ออำเภอ ----
  const pieceDefault: Record<string, number> = {};
  const pieceOverride: { prov: string; distKw: string; price: number }[] = [];
  if (pieceName) {
    const rows: Row[] = XLSX.utils.sheet_to_json(wb.Sheets[pieceName], { header: 1, defval: '' });
    const h = findHeaderRow(rows, 'จังหวัด');
    if (h >= 0) {
      const header = rows[h];
      const pc = colOf(header, 'จังหวัด'), dc = colOf(header, 'อำเภอ'), rc = colOf(header, 'ราคา', 'ชิ้น');
      for (const row of rows.slice(h + 1)) {
        const prov = String(row[pc] ?? '').trim();
        const dist = dc >= 0 ? stripAmphoe(row[dc]) : '';
        const price = Number(row[rc]);
        if (!prov || !(price > 0)) continue;
        if (dist) pieceOverride.push({ prov: normTxt(prov), distKw: dist, price });
        else pieceDefault[normTxt(prov)] = price;
      }
    }
  }
  const pieceFor = (prov: string, dist: string): number | null => {
    const np = normTxt(prov);
    for (const o of pieceOverride) if (o.prov === np && dist.includes(o.distKw)) return o.price;
    return pieceDefault[np] ?? null;
  };

  // ---- ราคาเหมา ต่ออำเภอ (+ สร้างราคาชิ้นคู่กัน ถ้าหาเจอ) ----
  const flatRows: Row[] = XLSX.utils.sheet_to_json(wb.Sheets[flatName], { header: 1, defval: '' });
  const fh = findHeaderRow(flatRows, 'จังหวัด');
  if (fh < 0) throw new Error('ชีตราคาเหมาไม่พบหัวคอลัมน์ "จังหวัด"');
  const fheader = flatRows[fh];
  const fp = colOf(fheader, 'จังหวัด'), fd = colOf(fheader, 'อำเภอ'), fr = colOf(fheader, 'ราคา', 'เหมา');
  if (fp < 0 || fd < 0 || fr < 0) throw new Error('ชีตราคาเหมาต้องมีคอลัมน์ จังหวัด / อำเภอ / ราคา');

  const rates: ParsedRate[] = [];
  let nFlat = 0, nPiece = 0, nCombined = 0;
  const provNoPiece = new Set<string>();
  for (const row of flatRows.slice(fh + 1)) {
    const prov = String(row[fp] ?? '').trim();
    const dist = stripAmphoe(row[fd]);
    const flatPrice = Number(row[fr]);
    if (!prov || !dist || !(flatPrice > 0)) continue;
    const base = {
      provinceName: prov, provinceShort: '', districtName: dist, destinationName: `${dist} จ.${prov}`,
      productCategory: 'normal' as const, status: 'active' as const, effectiveFrom: '2020-01-01', effectiveTo: null,
    };
    rates.push({ ...base, priceType: 'flat', price: flatPrice });
    nFlat++;
    // ส่งหลายอำเภอ (มี +) = ราคาเหมารวม ไม่มีราคาชิ้น
    if (dist.includes('+')) { nCombined++; continue; }
    const pc = pieceFor(prov, dist);
    if (pc != null) { rates.push({ ...base, priceType: 'piece', price: pc }); nPiece++; }
    else provNoPiece.add(prov);
  }
  summary.push(`ราคาเหมา ${nFlat} อำเภอ (ในนี้เป็นราคาชุดส่งหลายอำเภอ ${nCombined})`);
  summary.push(`ราคาชิ้น ${nPiece} อำเภอ`);
  if (provNoPiece.size) summary.push(`จังหวัดที่ไม่มีราคาชิ้น (คิดเหมาล้วน): ${[...provNoPiece].join(', ')}`);
  return { rates, summary };
}
