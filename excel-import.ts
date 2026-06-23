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

      // --- แถวรายการสินค้า: ต้องมีจำนวน(ตัวเลข) + ชื่อสินค้า ---
      const qty = num(row, cols.qty);
      const product = txt(row, cols.product);
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
      const unit = txt(row, cols.unit);
      curReceipt.items.push({ productName: product, quantity: qty, ...(unit ? { unit } : {}) });
    }

    pushDoc();
  }

  return docs;
}
