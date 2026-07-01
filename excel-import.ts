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
  const d = new Date(ms);
  let y = d.getUTCFullYear();
  // 🛠️ แก้ปีเพี้ยนจาก Excel แปลงวันที่ "พ.ศ." ที่พิมพ์ ให้เป็น date serial เอง:
  //   - พ.ศ. เต็ม 4 หลัก (2569) -> ค.ศ. (ลบ 543) = 2026
  //   - Excel ตีความเลขท้าย 2 หลัก "69" เป็น ค.ศ.1969 -> คืนเป็น พ.ศ.2569 -> ค.ศ.2026
  //   (แอปใช้งานช่วงปี 2025+ จึงไม่มีวันที่จริงก่อน ค.ศ.2000)
  if (y >= 2400) y -= 543;
  else if (y < 2000) y = (y % 100) + 1957;
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${mm}-${dd}`;
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
  return { documentNo: '', documentDate: '', plateNo: '', provinceRaw: '', districtRaw: '', receipts: [], docNote: '' };
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

      // --- โน้ตเส้นทาง/เงื่อนไขราคา (เช่น "วิ่งย่อยไม่เกิน 13 จุด") — เก็บเป็น keyword จับราคาแบบมีเงื่อนไข ---
      if (joined.includes('วิ่ง')) {
        const noteCell = row.map((c) => String(c).trim()).find((s) => s.includes('วิ่ง')) || joined.trim();
        doc.docNote = (doc.docNote ? doc.docNote + ' ' : '') + noteCell;
      }

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
  priceType: 'flat' | 'piece'; price: number;
  productCategory: 'normal' | 'collect_back' | 'peat_mass' | 'fixed_addon'; status: 'active';
  effectiveFrom: string; effectiveTo: null;
  // เงื่อนไขพิเศษ (จากชีต "พิเศษ") — เว้นว่างได้ทั้งหมด
  rateGroup?: string; pieceThreshold?: number | null;
  minQty?: number | null; maxQty?: number | null;
  receiverKeyword?: string; senderKeyword?: string; productKeyword?: string; remark?: string;
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
  const pieceName = findSheet('ชิ้น');
  const advName = findSheet('พิเศษ');
  // ชีตราคาเหมา: เลือกชีตที่ชื่อมี "เหมา" ก่อน ไม่งั้นชีตแรกที่ไม่ใช่ ชิ้น/พิเศษ/วิธีใช้
  const flatName = findSheet('เหมา') || wb.SheetNames.find((n) => n !== pieceName && n !== advName && !String(n).includes('วิธีใช้')) || '';
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
  const rates: ParsedRate[] = [];
  let nFlat = 0, nPiece = 0, nCombined = 0;
  const provNoPiece = new Set<string>();
  if (flatName && flatName !== advName) {
    const flatRows: Row[] = XLSX.utils.sheet_to_json(wb.Sheets[flatName], { header: 1, defval: '' });
    const fh = findHeaderRow(flatRows, 'จังหวัด');
    if (fh < 0) throw new Error('ชีตราคาเหมาไม่พบหัวคอลัมน์ "จังหวัด"');
    const fheader = flatRows[fh];
    const fp = colOf(fheader, 'จังหวัด'), fd = colOf(fheader, 'อำเภอ'), fr = colOf(fheader, 'ราคา', 'เหมา');
    if (fp < 0 || fd < 0 || fr < 0) throw new Error('ชีตราคาเหมาต้องมีคอลัมน์ จังหวัด / อำเภอ / ราคา');
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
  }

  // ---- ชีต "พิเศษ": ทุกหมวด/เงื่อนไข (เก็บคืน·Peat·บวกเพิ่ม·กลุ่มราคา·จุดตัด·ขั้นบันได·keyword ผู้รับ/ผู้ส่ง/สินค้า) ----
  if (advName) {
    const rows: Row[] = XLSX.utils.sheet_to_json(wb.Sheets[advName], { header: 1, defval: '' });
    const h = findHeaderRow(rows, 'จังหวัด');
    if (h >= 0) {
      const H = rows[h];
      const ci = {
        cat: colOf(H, 'หมวด'), prov: colOf(H, 'จังหวัด'), dist: colOf(H, 'อำเภอ'),
        type: colOf(H, 'ประเภท'),
        price: H.findIndex((c) => { const s = String(c); return s.includes('ราคา') && !s.includes('ประเภท'); }), // กัน match "ประเภทราคา"
        group: colOf(H, 'กลุ่ม'),
        thr: colOf(H, 'จุดตัด'), min: colOf(H, 'ตั้งแต่', 'กล่อง≥'), max: colOf(H, 'ถึง', 'กล่อง≤'),
        recv: colOf(H, 'ผู้รับ'), send: colOf(H, 'ผู้ส่ง'), prod: colOf(H, 'สินค้า'), note: colOf(H, 'หมายเหตุ'),
      };
      const cellS = (row: Row, i: number) => (i >= 0 ? String(row[i] ?? '').trim() : '');
      const numN = (row: Row, i: number) => { if (i < 0) return null; const v = row[i]; const x = Number(v); return v !== '' && v != null && !Number.isNaN(x) ? x : null; };
      const catOf = (s: string): ParsedRate['productCategory'] => {
        const n = s.toLowerCase();
        if (s.includes('เก็บ')) return 'collect_back';
        if (n.includes('peat') || s.includes('พีท')) return 'peat_mass';
        if (s.includes('บวก') || n.includes('addon')) return 'fixed_addon';
        return 'normal';
      };
      let nAdv = 0;
      for (const row of rows.slice(h + 1)) {
        const prov = cellS(row, ci.prov);
        const price = numN(row, ci.price);
        if (!prov || !(price != null && price > 0)) continue;
        const cat = catOf(cellS(row, ci.cat));
        // เก็บคืน/งานปกติ เลือกเหมา/ชิ้นได้ · Peat = ชิ้นเสมอ · บวกเพิ่ม = เหมาเสมอ
        let pt: 'flat' | 'piece' = cellS(row, ci.type).includes('ชิ้น') ? 'piece' : 'flat';
        if (cat === 'peat_mass') pt = 'piece';
        if (cat === 'fixed_addon') pt = 'flat';
        const dist = stripAmphoe(cellS(row, ci.dist));
        rates.push({
          provinceName: prov, provinceShort: '', districtName: dist,
          destinationName: `${dist || 'ทั้งจังหวัด'} จ.${prov}`,
          productCategory: cat, priceType: pt, price, status: 'active', effectiveFrom: '2020-01-01', effectiveTo: null,
          rateGroup: cellS(row, ci.group) || undefined,
          pieceThreshold: numN(row, ci.thr),
          minQty: numN(row, ci.min), maxQty: numN(row, ci.max),
          receiverKeyword: cellS(row, ci.recv) || undefined,
          senderKeyword: cellS(row, ci.send) || undefined,
          productKeyword: cellS(row, ci.prod) || undefined,
          remark: cellS(row, ci.note) || undefined,
        });
        nAdv++;
      }
      if (nAdv) summary.push(`ราคาพิเศษ (เก็บคืน/Peat/บวกเพิ่ม/กลุ่ม/เงื่อนไข) ${nAdv} แถว`);
    }
  }

  return { rates, summary };
}

// ============================================================================
// Import "ค่าน้ำมัน" (ใบสั่งเติมน้ำมัน) จาก Excel -> FuelEntry
// คอลัมน์: ทะเบียน | วัน/เดือน/ปี | ใบสั่งเติมน้ำมัน | จำนวนเงิน
// ============================================================================
export interface ParsedFuel { plateNo: string; date: string; refNo: string; amount: number; }

const THAI_MON = ['มค', 'กพ', 'มีค', 'เมย', 'พค', 'มิย', 'กค', 'สค', 'กย', 'ตค', 'พย', 'ธค'];
const pad2 = (n: number) => String(n).padStart(2, '0');
const normYear = (y: number) => (y >= 2400 ? y - 543 : y < 100 ? 2000 + y : y); // พ.ศ.->ค.ศ., 2หลัก->20xx

// แปลงค่าวันที่หลายรูปแบบ -> YYYY-MM-DD (เลข Excel, ค.ศ./พ.ศ., ไทย "2 มิ.ย. 2569")
function parseAnyDate(v: any): string {
  if (v == null || v === '') return '';
  if (typeof v === 'number') return excelDate(v);
  const s = String(v).trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (m) return `${normYear(+m[1])}-${pad2(+m[2])}-${pad2(+m[3])}`; // แปลง พ.ศ.->ค.ศ. ด้วย (2569-06-05 -> 2026-06-05)
  m = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/.exec(s); // D/M/Y
  if (m) return `${normYear(+m[3])}-${pad2(+m[2])}-${pad2(+m[1])}`;
  const tm = /^(\d{1,2})\s*([฀-๿.]+)\s*(\d{2,4})$/.exec(s); // ไทย: 2 มิ.ย. 2569
  if (tm) {
    const monKey = tm[2].replace(/[.\s]/g, ''); // คงสระไว้ (มิ.ย. -> มิย)
    const mi = THAI_MON.findIndex((mn) => mn === monKey);
    if (mi >= 0) return `${normYear(+tm[3])}-${pad2(mi + 1)}-${pad2(+tm[1])}`;
  }
  return '';
}

// รองรับ 2 รูปแบบ:
//  (1) ตารางแบน 1 ชีต: มีคอลัมน์ "ทะเบียน" ในแต่ละแถว
//  (2) แยก sheet ต่อทะเบียน: ชื่อ sheet = ทะเบียน (ไม่มีคอลัมน์ทะเบียนในแถว)
export function parseFuelExcel(buffer: Buffer): { fuel: ParsedFuel[]; summary: string[] } {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false });
  const fuel: ParsedFuel[] = [];
  const badDates: string[] = [];
  let usedSheets = 0;

  for (const sheetName of wb.SheetNames) {
    if (sheetName.includes('วิธีใช้')) continue;
    const rows: Row[] = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });
    // หาหัวตาราง: ต้องมี "วัน" และ ("จำนวนเงิน/เงิน/บาท")
    const h = rows.findIndex((r) => r.some((c) => String(c).includes('วัน')) && r.some((c) => /จำนวน|เงิน|บาท/.test(String(c))));
    if (h < 0) continue;
    const header = rows[h];
    const pc = colOf(header, 'ทะเบียน'); // -1 = ไม่มีคอลัมน์ทะเบียน -> ใช้ชื่อ sheet
    const dc = colOf(header, 'วัน', 'เดือน'), rc = colOf(header, 'ใบสั่ง', 'เลข'), ac = colOf(header, 'จำนวนเงิน', 'จำนวน', 'เงิน', 'บาท');
    if (dc < 0 || ac < 0) continue;
    const sheetPlate = sheetName.trim();
    let added = 0;
    for (const row of rows.slice(h + 1)) {
      const plate = (pc >= 0 ? String(row[pc] ?? '').trim() : sheetPlate);
      const amount = Number(row[ac]);
      if (!plate || !(amount > 0)) continue;
      const date = parseAnyDate(row[dc]);
      if (!date) { badDates.push(`${plate} (${String(row[dc] ?? '').trim() || 'ว่าง'})`); continue; }
      fuel.push({ plateNo: plate, date, refNo: rc >= 0 ? String(row[rc] ?? '').trim() : '', amount });
      added++;
    }
    if (added > 0) usedSheets++;
  }
  if (!fuel.length && !badDates.length) throw new Error('ไม่พบใบสั่งเติมน้ำมัน — ตรวจสอบหัวคอลัมน์ วันที่/จำนวนเงิน');
  const summary: string[] = [`อ่านได้ ${fuel.length} ใบสั่งเติมน้ำมัน จาก ${usedSheets} ทะเบียน/ชีต`];
  if (badDates.length) summary.push(`ข้ามแถววันที่อ่านไม่ได้ ${badDates.length} แถว: ${badDates.slice(0, 5).join(', ')}${badDates.length > 5 ? ' ...' : ''}`);
  return { fuel, summary };
}
