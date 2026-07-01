// ============================================================================
// Excel Export ด้วย ExcelJS — รองรับฟอนต์ Cordia New, สี, Border, หลาย Sheet
// ตาม Requirement section 8-9
// ============================================================================
// type เท่านั้น -> ถูกลบตอน build ไม่ฝัง ExcelJS (~1MB) ในบันเดิลหน้าเว็บ
// ตัวจริงโหลดแบบ dynamic ใน newWorkbook() เฉพาะตอนผู้ใช้กด export
import type ExcelJS from 'exceljs';
import {
  BillingCycle,
  TripDocument,
  FuelEntry,
  DeductionEntry,
  Vehicle,
  RateMaster,
} from './types';
import { summarizeByVehicle, normPlate, normDoc } from './calc';

const FONT = 'Cordia New';

// ---- Palette ตาม Requirement ----
const C = {
  headerBg: 'FF1B365D',   // น้ำเงินเข้ม
  headerText: 'FFFFFFFF',
  title: 'FF1B365D',
  sub: 'FF2C3E50',
  subBg: 'FFEAF2F8',
  totalBg: 'FFE6EEF8',
  zebra: 'FFF9FAFC',
  dividerBg: 'FFFFF2CC',   // รายการมีตัวหาร
  dividerText: 'FFC65911',
  billingText: 'FFC00000',  // จำนวนคิดค่าเที่ยว เด่น
  warnBg: 'FFFCE4D6',       // รายการผิดปกติ
  warnText: 'FF9C0006',
  border: 'FFCCCCCC',
};

const thinBorder: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: C.border } },
  left: { style: 'thin', color: { argb: C.border } },
  right: { style: 'thin', color: { argb: C.border } },
  bottom: { style: 'thin', color: { argb: C.border } },
};

function solid(argb: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function styleHeaderRow(row: ExcelJS.Row) {
  row.eachCell((cell) => {
    cell.font = { name: FONT, size: 14, bold: true, color: { argb: C.headerText } };
    cell.fill = solid(C.headerBg);
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = thinBorder;
  });
  row.height = 22;
}

function styleTitle(ws: ExcelJS.Worksheet, text: string, span: number, sub?: string) {
  ws.mergeCells(1, 1, 1, span);
  const t = ws.getCell(1, 1);
  t.value = text;
  t.font = { name: FONT, size: 18, bold: true, color: { argb: C.title } };
  t.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(1).height = 28;
  if (sub) {
    ws.mergeCells(2, 1, 2, span);
    const s = ws.getCell(2, 1);
    s.value = sub;
    s.font = { name: FONT, size: 13, color: { argb: C.sub } };
  }
}

function bodyCell(cell: ExcelJS.Cell, opts: { bold?: boolean; color?: string; align?: 'left' | 'right' | 'center'; bg?: string } = {}) {
  cell.font = { name: FONT, size: 13, bold: opts.bold, color: opts.color ? { argb: opts.color } : undefined };
  cell.alignment = { vertical: 'middle', horizontal: opts.align || 'left' };
  cell.border = thinBorder;
  if (opts.bg) cell.fill = solid(opts.bg);
}

const NUM = '#,##0.00';
const round2 = (n: number) => Math.round((n + Number.EPSILON) * 100) / 100;

// โหลด ExcelJS แบบ dynamic เฉพาะตอนผู้ใช้กด export จริง (code-splitting)
// -> ExcelJS (~1MB) แยกเป็นไฟล์ย่อย ไม่อยู่ในบันเดิลหน้าแรก โหลดครั้งแรกที่ export
async function newWorkbook(): Promise<ExcelJS.Workbook> {
  const mod = await import('exceljs');
  return new mod.default.Workbook();
}

// ===========================================================================
export async function exportCycleToExcel(
  cycle: BillingCycle,
  trips: TripDocument[],
  fuel: FuelEntry[],
  deductions: DeductionEntry[],
  vehicles: Vehicle[],
  rateMasters: RateMaster[]
) {
  const wb = await newWorkbook();
  wb.creator = 'ระบบค่าเที่ยว+ค่าน้ำมันรถร่วม';
  wb.created = new Date();

  const cycleTrips = trips.filter((t) => t.cycleId === cycle.id);
  const cycleFuel = fuel.filter((f) => f.cycleId === cycle.id);
  const cycleDed = deductions.filter((d) => d.cycleId === cycle.id);
  const summaries = summarizeByVehicle(cycle.id, trips, fuel, deductions, vehicles);

  // ---------------- SHEET: Dashboard / สรุปรวม ----------------
  {
    const ws = wb.addWorksheet('Dashboard');
    styleTitle(ws, `สรุปภาพรวม — ${cycle.name}`, 8, `ช่วงรอบ ${cycle.startDate} ถึง ${cycle.endDate}`);
    ws.addRow([]);
    const headers = ['ทะเบียนรถ', 'คนขับ', 'รายได้ทั้งหมด', 'หัก 1%', 'ค่าน้ำมัน', '+ รายได้เพิ่ม', 'รวมรายการหัก', 'รับสุทธิ'];
    const hr = ws.addRow(headers);
    styleHeaderRow(hr);

    let zebra = false;
    let gTrip = 0, gNet = 0, gFuel = 0;
    summaries.forEach((s) => {
      const r = ws.addRow([
        s.plateNo, s.driverName, s.totalTripAmount, s.deduction1Percent, s.fuelTotal,
        s.incomeAdd, s.deductionTotal, s.netReceive,
      ]);
      r.eachCell((cell, col) => bodyCell(cell, { align: col <= 2 ? 'left' : 'right', bg: zebra ? C.zebra : undefined }));
      [3, 4, 5, 6, 7, 8].forEach((c) => (r.getCell(c).numFmt = NUM));
      r.getCell(8).font = { name: FONT, size: 14, bold: true, color: { argb: C.billingText } };
      zebra = !zebra;
      gTrip += s.totalTripAmount; gNet += s.netReceive; gFuel += s.fuelTotal;
    });
    const tr = ws.addRow(['รวมทั้งหมด', '', gTrip, '', gFuel, '', '', gNet]);
    tr.eachCell((cell) => bodyCell(cell, { bold: true, align: 'right', bg: C.totalBg }));
    tr.getCell(1).alignment = { horizontal: 'left' };
    [3, 5, 8].forEach((c) => (tr.getCell(c).numFmt = NUM));
    ws.columns.forEach((c) => (c.width = 16));
    ws.getColumn(2).width = 18;
  }

  // ---------------- SHEET: ค่าเที่ยวทั้งหมด ----------------
  {
    const ws = wb.addWorksheet('ค่าเที่ยวทั้งหมด');
    styleTitle(ws, `ค่าเที่ยวทั้งหมด — ${cycle.name}`, 9);
    ws.addRow([]);
    const headers = ['เลขที่ใบกระจาย', 'วันที่', 'ทะเบียน', 'ปลายทาง', 'เลขที่ใบรับสินค้า', 'ผู้รับสินค้า', 'จำนวนจริง', 'คิดค่าเที่ยว', 'ค่าเที่ยว (บาท)'];
    styleHeaderRow(ws.addRow(headers));

    let zebra = false;
    cycleTrips.forEach((t) => {
      // แถบหัวใบกระจาย (Section)
      const sec = ws.addRow([t.documentNo, t.documentDate, t.plateNo, `${t.provinceRaw} ${t.districtRaw}`, '', '', t.totalQty, t.billingQty, t.tripAmount]);
      sec.eachCell((cell) => bodyCell(cell, { bold: true, bg: C.subBg, align: 'left' }));
      [7, 8, 9].forEach((c) => { sec.getCell(c).alignment = { horizontal: 'right' }; sec.getCell(c).numFmt = c === 9 ? NUM : '#,##0'; });
      if (t.warnings.length) {
        sec.getCell(1).font = { name: FONT, size: 13, bold: true, color: { argb: C.warnText } };
      }
      // แถวใบรับ (มีปลายทาง + ค่าเที่ยวต่อจุด)
      t.receipts.forEach((rcp) => {
        const ptAmount = t.rateType === 'piece' ? rcp.receiptAmount : (rcp.flatPrice ?? '');
        const r = ws.addRow(['', '', '', `${rcp.districtRaw} ${rcp.provinceRaw}`.trim(), rcp.receiptNo, rcp.receiverName, rcp.totalQty, rcp.billingQty, ptAmount]);
        const isDiv = rcp.hasAdjustment;
        r.eachCell((cell, col) => bodyCell(cell, {
          align: col >= 7 ? 'right' : 'left',
          bg: isDiv ? C.dividerBg : (zebra ? C.zebra : undefined),
        }));
        r.getCell(7).numFmt = '#,##0';
        r.getCell(8).numFmt = '#,##0';
        r.getCell(9).numFmt = NUM;
        if (isDiv) {
          const div = rcp.adjustments[0];
          r.getCell(5).value = `🟧÷${div.divisor}  ${rcp.receiptNo}`;
          r.getCell(5).font = { name: FONT, size: 13, bold: true, color: { argb: C.dividerText } };
          r.getCell(8).font = { name: FONT, size: 14, bold: true, color: { argb: C.billingText } };
          r.getCell(8).note = rcp.adjustments.map((a) => a.note).join(' | ');
        }
        zebra = !zebra;
      });
    });
    // รวมยอด
    const total = cycleTrips.reduce((s, t) => s + t.tripAmount, 0);
    const tr = ws.addRow(['', '', '', '', '', 'รวมค่าเที่ยวทั้งหมด', '', '', total]);
    tr.eachCell((cell) => bodyCell(cell, { bold: true, bg: C.totalBg, align: 'right' }));
    tr.getCell(9).numFmt = NUM;
    ws.columns.forEach((c) => (c.width = 16));
    ws.getColumn(4).width = 22; ws.getColumn(6).width = 24; ws.getColumn(5).width = 22;
  }

  // ---------------- SHEET: ปรับจำนวน/ตัวหาร ----------------
  {
    const ws = wb.addWorksheet('ปรับจำนวนตัวหาร');
    styleTitle(ws, `รายการปรับจำนวน (ตัวหาร) — ${cycle.name}`, 8);
    ws.addRow([]);
    styleHeaderRow(ws.addRow(['เลขที่ใบรับสินค้า', 'ผู้รับสินค้า', 'รายการสินค้า', 'จำนวนจริง', 'ตัวหาร', 'หลังหาร', 'สูตร', 'ใบกระจาย']));
    cycleTrips.forEach((t) => {
      t.receipts.forEach((rcp) => {
        rcp.adjustments.forEach((a) => {
          const r = ws.addRow([rcp.receiptNo, rcp.receiverName, a.productName, a.specialQty, `÷${a.divisor}`, a.convertedQty, a.note, t.documentNo]);
          r.eachCell((cell, col) => bodyCell(cell, { align: col >= 4 && col <= 6 ? 'center' : 'left', bg: C.dividerBg, color: C.dividerText, bold: col === 6 }));
        });
      });
    });
    ws.columns.forEach((c) => (c.width = 16));
    ws.getColumn(2).width = 22; ws.getColumn(3).width = 20; ws.getColumn(7).width = 18;
  }

  // ---------------- SHEET: ค่าน้ำมันทั้งหมด ----------------
  {
    const ws = wb.addWorksheet('ค่าน้ำมันทั้งหมด');
    styleTitle(ws, `ค่าน้ำมันทั้งหมด — ${cycle.name}`, 5);
    ws.addRow([]);
    styleHeaderRow(ws.addRow(['ทะเบียน', 'เลขใบสั่งเติม', 'วันที่', 'ค่าน้ำมัน', 'หมายเหตุ']));
    let zebra = false;
    cycleFuel.forEach((f) => {
      const r = ws.addRow([f.plateNo, f.refNo, f.date, f.amount, f.note || '']);
      r.eachCell((cell, col) => bodyCell(cell, { align: col === 4 ? 'right' : 'left', bg: zebra ? C.zebra : undefined }));
      r.getCell(4).numFmt = NUM;
      zebra = !zebra;
    });
    const total = cycleFuel.reduce((s, f) => s + f.amount, 0);
    const tr = ws.addRow(['', '', 'รวมค่าน้ำมัน', total, '']);
    tr.eachCell((cell) => bodyCell(cell, { bold: true, bg: C.totalBg, align: 'right' }));
    tr.getCell(4).numFmt = NUM;
    ws.columns.forEach((c) => (c.width = 16));
    ws.getColumn(5).width = 24;
  }

  // ---------------- SHEET: รายการหัก ----------------
  {
    const ws = wb.addWorksheet('รายการหัก');
    styleTitle(ws, `รายการหัก — ${cycle.name}`, 6);
    ws.addRow([]);
    styleHeaderRow(ws.addRow(['ทะเบียน', 'ทิศทาง', 'รายการ', 'จำนวนเงิน', 'ใบกระจาย', 'หมายเหตุ']));
    let zebra = false;
    cycleDed.forEach((d) => {
      const r = ws.addRow([d.plateNo, d.kind === 'income' ? 'รายได้เพิ่ม (+)' : 'หักออก (-)', d.label, d.amount, d.docNo || '', d.note || '']);
      r.eachCell((cell, col) => bodyCell(cell, { align: col === 4 ? 'right' : 'left', bg: zebra ? C.zebra : undefined }));
      r.getCell(4).numFmt = NUM;
      zebra = !zebra;
    });
    ws.columns.forEach((c) => (c.width = 16));
    ws.getColumn(3).width = 22; ws.getColumn(5).width = 18; ws.getColumn(6).width = 24;
  }

  // ---------------- SHEET: ราคาขนส่ง (Master ที่ใช้) ----------------
  {
    const ws = wb.addWorksheet('ราคาขนส่ง');
    styleTitle(ws, 'Master ราคาขนส่ง', 7);
    ws.addRow([]);
    styleHeaderRow(ws.addRow(['ปลายทาง', 'จังหวัด', 'อำเภอ', 'ประเภท', 'ราคา', 'เริ่มใช้', 'สถานะ']));
    let zebra = false;
    rateMasters.forEach((rm) => {
      const r = ws.addRow([rm.destinationName, rm.provinceName, rm.districtName, rm.priceType === 'flat' ? 'ราคาเหมา' : 'ราคาชิ้น', rm.price, rm.effectiveFrom, rm.status === 'active' ? 'ใช้งาน' : 'ปิด']);
      r.eachCell((cell, col) => bodyCell(cell, { align: col === 5 ? 'right' : 'left', bg: zebra ? C.zebra : undefined }));
      r.getCell(5).numFmt = NUM;
      zebra = !zebra;
    });
    ws.columns.forEach((c) => (c.width = 16));
    ws.getColumn(1).width = 20;
  }

  // ---------------- SHEET: รายชื่อ (รถ+คนขับ) ----------------
  {
    const ws = wb.addWorksheet('รายชื่อรถ');
    styleTitle(ws, 'Master ทะเบียนรถและคนขับ', 4);
    ws.addRow([]);
    styleHeaderRow(ws.addRow(['ทะเบียนรถ', 'คนขับ', 'ประเภทรถ', 'สถานะ']));
    let zebra = false;
    vehicles.forEach((v) => {
      const r = ws.addRow([v.plateNo, v.driverName, v.vehicleType, v.status === 'active' ? 'ใช้งาน' : 'ปิด']);
      r.eachCell((cell) => bodyCell(cell, { bg: zebra ? C.zebra : undefined }));
      zebra = !zebra;
    });
    ws.columns.forEach((c) => (c.width = 18));
  }

  // ---- Download ----
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `รายงานค่าเที่ยว_${cycle.name.replace(/[\s/]/g, '_')}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ===========================================================================
// รายงานต่อทะเบียน — 1 ไฟล์, sheet ต่อทะเบียน (ค่าบรรทุก + สรุป + ใบสั่งน้ำมัน)
// ===========================================================================
function fmtDate(s: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec((s || '').trim());
  return m ? `${+m[3]}/${+m[2]}/${m[1]}` : (s || '');
}
function safeSheetName(name: string, used: Set<string>): string {
  let n = (name || 'รถ').replace(/[\\/?*[\]:]/g, '-').slice(0, 28);
  let base = n, i = 2;
  while (used.has(n)) { n = `${base} (${i++})`; }
  used.add(n);
  return n;
}

// 1 ใบ -> หลายบรรทัด: แยกตามปลายทาง (เฉพาะใบราคาชิ้น งานปกติล้วน หลายอำเภอ) ไม่งั้น 1 บรรทัด
export interface TripSubRow { date: string; dest: string; docNo: string; qty: number; rateType: string; price: number | null; amount: number; hasDiv: boolean; first: boolean; }
export function tripSubRows(t: TripDocument): TripSubRow[] {
  const splittable = t.rateType === 'piece' && (t.breakdown?.collect || 0) === 0 && (t.breakdown?.peat || 0) === 0;
  if (splittable) {
    const groups = new Map<string, { dist: string; prov: string; qty: number; amount: number; price: number | null; hasDiv: boolean; addon: number }>();
    const getG = (dist: string, prov: string) => {
      const key = (dist || '') + '|' + (prov || '');
      const g = groups.get(key) || { dist: dist || '', prov: prov || '', qty: 0, amount: 0, price: null, hasDiv: false, addon: 0 };
      groups.set(key, g); return g;
    };
    for (const r of t.receipts || []) {
      if ((r.normalQty || 0) <= 0) continue;
      const g = getG(r.districtRaw || '', r.provinceRaw || '');
      g.qty += r.billingQty || 0;
      g.amount += r.receiptAmount || 0;
      if (r.piecePrice != null) g.price = r.piecePrice;
      if (r.hasAdjustment) g.hasDiv = true;
    }
    // ค่าเหมาบวกเพิ่มตายตัว (เช่น ท่าสองยาง +700) -> ลงบรรทัดปลายทางนั้นเลย (เหมา 700)
    for (const ad of t.addonByDest || []) {
      const g = getG(ad.dist, ad.prov);
      g.amount += ad.amount; g.addon += ad.amount;
    }
    const arr = [...groups.values()];
    if (arr.length > 1) {
      return arr.map((g, i) => {
        const pureAddon = g.addon > 0 && round2(g.amount) === round2(g.addon); // ยอดมาจาก addon ล้วน -> แสดงเป็นเหมา
        return {
          date: i === 0 ? t.documentDate : '', dest: `${g.dist ? 'อ.' + g.dist : ''}${g.prov ? ' จ.' + g.prov : ''}`.trim(),
          docNo: i === 0 ? t.documentNo : '', qty: g.qty, rateType: pureAddon ? 'flat' : 'piece', price: pureAddon ? round2(g.addon) : g.price, amount: round2(g.amount), hasDiv: g.hasDiv, first: i === 0,
        };
      });
    }
  }
  return [{ date: t.documentDate, dest: tripDestinations(t), docNo: t.documentNo, qty: t.billingQty, rateType: t.rateType || '', price: tripUnitRate(t), amount: t.tripAmount, hasDiv: !!t.receipts?.some((r) => r.hasAdjustment), first: true }];
}

// รายชื่อปลายทางทุกอำเภอของใบกระจาย (กลุ่มตามจังหวัด) เช่น "อ.ชนแดน, อ.เมือง จ.เพชรบูรณ์"
export function tripDestinations(t: TripDocument): string {
  const byProv = new Map<string, Set<string>>();
  for (const r of t.receipts || []) {
    const prov = (r.provinceRaw || '').trim();
    const dist = (r.districtRaw || '').trim();
    if (!dist && !prov) continue;
    if (!byProv.has(prov)) byProv.set(prov, new Set());
    if (dist) byProv.get(prov)!.add(dist);
  }
  if (byProv.size === 0) return `${t.districtRaw ? 'อ.' + t.districtRaw : ''}${t.provinceRaw ? ' จ.' + t.provinceRaw : ''}`.trim();
  return [...byProv.entries()].map(([prov, dists]) => `${[...dists].map((d) => 'อ.' + d).join(', ')}${prov ? ' จ.' + prov : ''}`).join(' · ');
}

// ราคา/ชิ้น ของใบ: เหมา=ราคาเหมา, ชิ้นราคาเดียว=ราคานั้น, ชิ้นหลายราคา=null (แสดง "หลายราคา")
export function tripUnitRate(t: TripDocument): number | null {
  if (t.rateType === 'flat') return t.rateValue ?? null;
  if (t.rateType === 'piece') {
    const prices = [...new Set((t.receipts || []).filter((r) => r.piecePrice != null).map((r) => r.piecePrice as number))];
    return prices.length === 1 ? prices[0] : null;
  }
  return null;
}

export async function exportPerVehicleReport(
  cycle: BillingCycle,
  branchName: string,
  trips: TripDocument[],
  fuel: FuelEntry[],
  deductions: DeductionEntry[],
  vehicles: Vehicle[]
) {
  const wb = await newWorkbook();
  wb.creator = 'ระบบค่าเที่ยว+ค่าน้ำมันรถร่วม';
  wb.created = new Date();

  const cycleTrips = trips.filter((t) => t.cycleId === cycle.id);
  const summaries = summarizeByVehicle(cycle.id, trips, fuel, deductions, vehicles);
  const usedNames = new Set<string>();

  if (!summaries.length) {
    const ws = wb.addWorksheet('ไม่มีข้อมูล');
    ws.getCell(1, 1).value = 'ยังไม่มีข้อมูลในรอบนี้';
  }

  const cycleFuel = fuel.filter((f) => f.cycleId === cycle.id);
  const driverOf = (plate: string) =>
    summaries.find((s) => normPlate(s.plateNo) === normPlate(plate))?.driverName ||
    vehicles.find((v) => normPlate(v.plateNo) === normPlate(plate))?.driverName || '';

  // ===== Sheet สรุปรวมต่อทะเบียน (แบบ Dashboard) =====
  if (summaries.length) {
    const ws = wb.addWorksheet('สรุปรวม');
    styleTitle(ws, `สรุปรวมต่อทะเบียน — สาขา${branchName}`, 8, `รอบ ${cycle.name}`);
    ws.addRow([]);
    styleHeaderRow(ws.addRow(['ทะเบียน', 'คนขับ', 'รายได้', 'หัก 1%', 'ค่าน้ำมัน', '+ รายได้เพิ่ม', 'รวมรายการหัก', 'รับสุทธิ']));
    let z = false; const g = { trip: 0, d1: 0, fuel: 0, inc: 0, ded: 0, net: 0 };
    for (const s of summaries) {
      const r = ws.addRow([s.plateNo, s.driverName, s.totalTripAmount, s.deduction1Percent, s.fuelTotal, s.incomeAdd, s.deductionTotal, s.netReceive]);
      r.eachCell((cell, col) => bodyCell(cell, { align: col <= 2 ? 'left' : 'right', bg: z ? C.zebra : undefined, bold: col === 8, color: col === 8 ? C.billingText : undefined }));
      [3, 4, 5, 6, 7, 8].forEach((c) => (r.getCell(c).numFmt = NUM));
      z = !z; g.trip += s.totalTripAmount; g.d1 += s.deduction1Percent; g.fuel += s.fuelTotal; g.inc += s.incomeAdd; g.ded += s.deductionTotal; g.net += s.netReceive;
    }
    const tr = ws.addRow(['รวมทุกคัน', '', round2(g.trip), round2(g.d1), round2(g.fuel), round2(g.inc), round2(g.ded), round2(g.net)]);
    tr.eachCell((cell, col) => bodyCell(cell, { bold: true, align: col <= 2 ? 'left' : 'right', bg: C.totalBg, color: C.title }));
    [3, 4, 5, 6, 7, 8].forEach((c) => (tr.getCell(c).numFmt = NUM));
    [13, 20, 13, 11, 13, 14, 14, 13].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  }

  // ===== Sheet ค่าน้ำมันรวมทุกคัน — สรุปยอดรวมต่อทะเบียน (ไม่ลงรายละเอียดทุกใบ) =====
  if (summaries.length) {
    const ws = wb.addWorksheet('น้ำมันรวมทุกคัน');
    styleTitle(ws, `สรุปค่าน้ำมันรวมต่อทะเบียน — สาขา${branchName}`, 3, `รอบ ${cycle.name}`);
    ws.addRow([]);
    styleHeaderRow(ws.addRow(['ทะเบียน', 'คนขับ', 'จำนวนเงินรวม (บาท)']));
    // รวมยอดน้ำมันต่อทะเบียน
    const byPlate = new Map<string, { plate: string; total: number }>();
    for (const f of cycleFuel) {
      const key = normPlate(f.plateNo);
      const cur = byPlate.get(key) || { plate: f.plateNo, total: 0 };
      cur.total += f.amount;
      byPlate.set(key, cur);
    }
    const rows = [...byPlate.values()].sort((a, b) => (a.plate > b.plate ? 1 : a.plate < b.plate ? -1 : 0));
    let z = false, sum = 0;
    for (const v of rows) {
      const r = ws.addRow([v.plate, driverOf(v.plate), round2(v.total)]);
      r.eachCell((cell, col) => bodyCell(cell, { align: col === 3 ? 'right' : 'left', bg: z ? C.zebra : undefined }));
      r.getCell(3).numFmt = NUM; z = !z; sum += v.total;
    }
    const tr = ws.addRow(['', 'ผลรวมทุกคัน', round2(sum)]);
    tr.eachCell((cell) => bodyCell(cell, { bold: true, align: 'right', bg: C.totalBg, color: C.title }));
    tr.getCell(3).numFmt = NUM;
    [16, 24, 18].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  }

  // ===== Sheet สรุปภาษีหัก ณ ที่จ่าย 1% (รายได้ก่อน/หลังหัก) =====
  if (summaries.length) {
    const ws = wb.addWorksheet('ภาษีหัก ณ ที่จ่าย 1%');
    styleTitle(ws, `สรุปภาษีหัก ณ ที่จ่าย 1% — สาขา${branchName}`, 5, `รอบ ${cycle.name} · คิด 1% จากค่าเที่ยว + รายได้เพิ่ม (ยกเว้นค่าอัพเดทบิล)`);
    ws.addRow([]);
    styleHeaderRow(ws.addRow(['ทะเบียน', 'คนขับ', 'รายได้ก่อนหัก (ค่าเที่ยว)', 'ภาษีหัก ณ ที่จ่าย 1%', 'รายได้หลังหัก 1%']));
    let z = false; const g = { gross: 0, tax: 0, net: 0 };
    for (const s of summaries) {
      const gross = s.totalTripAmount, tax = s.deduction1Percent, net = round2(gross - tax);
      const r = ws.addRow([s.plateNo, s.driverName, round2(gross), round2(tax), net]);
      r.eachCell((cell, col) => bodyCell(cell, { align: col <= 2 ? 'left' : 'right', bg: z ? C.zebra : undefined, bold: col === 5, color: col === 4 ? C.billingText : col === 5 ? C.title : undefined }));
      [3, 4, 5].forEach((c) => (r.getCell(c).numFmt = NUM));
      z = !z; g.gross += gross; g.tax += tax; g.net += net;
    }
    const tr = ws.addRow(['รวมทุกคัน', '', round2(g.gross), round2(g.tax), round2(g.net)]);
    tr.eachCell((cell, col) => bodyCell(cell, { bold: true, align: col <= 2 ? 'left' : 'right', bg: C.totalBg, color: C.title }));
    [3, 4, 5].forEach((c) => (tr.getCell(c).numFmt = NUM));
    [16, 24, 22, 20, 20].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  }

  for (const s of summaries) {
    const plate = s.plateNo;
    const np = normPlate(plate);
    const vTrips = cycleTrips.filter((t) => normPlate(t.plateNo) === np).sort((a, b) => (a.documentDate || '').localeCompare(b.documentDate || '') || (a.documentNo || '').localeCompare(b.documentNo || ''));
    const vFuel = fuel.filter((f) => f.cycleId === cycle.id && normPlate(f.plateNo) === np).sort((a, b) => (a.date || '').localeCompare(b.date || ''));
    const vIncome = deductions.filter((d) => d.cycleId === cycle.id && d.kind === 'income' && normPlate(d.plateNo) === np);
    const vDeduct = deductions.filter((d) => d.cycleId === cycle.id && d.kind === 'deduction' && normPlate(d.plateNo) === np);

    const ws = wb.addWorksheet(safeSheetName(plate, usedNames));
    styleTitle(ws, `สรุปค่าบรรทุกรถร่วม — สาขา${branchName}`, 10,
      `รอบ ${cycle.name}  ·  ทะเบียน ${plate}  ·  ${s.driverName || '-'}`);
    ws.addRow([]);

    // ---- ตารางค่าบรรทุก ----
    styleHeaderRow(ws.addRow(['วันที่', 'ปลายทาง', 'เลขที่ใบกระจาย', 'จำนวนชิ้น', 'แบบ', 'ราคา', 'เป็นเงิน', 'รายได้เพิ่ม (พิเศษ)', 'ราคารวม', 'หมายเหตุ']));
    let zebra = false;
    let sumQty = 0, sumMoney = 0, sumExtra = 0, sumTotal = 0;
    const fmtB = (n: number) => round2(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    for (const t of vTrips) {
      const docInc = vIncome.filter((d) => normDoc(d.docNo || '') && normDoc(d.docNo || '') === normDoc(t.documentNo || ''));
      const docDed = vDeduct.filter((d) => normDoc(d.docNo || '') && normDoc(d.docNo || '') === normDoc(t.documentNo || ''));
      const extra = docInc.reduce((a, d) => a + d.amount, 0);
      const destFixed = (t.receipts || []).filter((r) => r.destCorrected);
      const destNoted = (t.receipts || []).filter((r) => r.destNote && !r.destCorrected);
      const subs = tripSubRows(t);
      subs.forEach((sub) => {
        const rowExtra = sub.first ? extra : 0;
        const rowTotal = round2(sub.amount + rowExtra);
        // หมายเหตุ: มีหาร + แก้ปลายทาง + รายการ +/- ที่ผูกเลขใบกระจายนี้ (แสดงในแถวแรกของใบ)
        const parts: { text: string; color: string }[] = [];
        if (sub.hasDiv) parts.push({ text: 'มีหาร', color: C.dividerText });
        if (sub.first) {
          if ((t.breakdown?.collect || 0) > 0) parts.push({ text: `🔄 เก็บคืน ฿${fmtB(t.breakdown.collect)}`, color: 'FF0E9F6E' });
          if ((t.breakdown?.peat || 0) > 0) parts.push({ text: `🌱 Peat ฿${fmtB(t.breakdown.peat)}`, color: 'FF0D9488' });
          if ((t.breakdown?.addon || 0) > 0) parts.push({ text: `📌 เหมาเพิ่ม ฿${fmtB(t.breakdown.addon)}${(t.addonByDest || []).length ? ` (${(t.addonByDest || []).map((a) => a.dist || a.prov).filter(Boolean).join(', ')})` : ''}`, color: 'FFB7791F' });
          for (const r of destFixed) parts.push({ text: `📍แก้ปลายทางจาก ${r.origDistrict || '-'} ${r.origProvince || ''}`.trim(), color: 'FF7C3AED' });
          for (const r of destNoted) parts.push({ text: `⚠️โน้ต ${r.destNote}`, color: 'FFB7791F' });
          for (const d of docInc) parts.push({ text: `➕${d.label} ฿${fmtB(d.amount)}`, color: 'FF0E9F6E' });
          for (const d of docDed) parts.push({ text: `➖${d.label} ฿${fmtB(d.amount)}`, color: 'FFE10026' });
        }
        const r = ws.addRow([
          fmtDate(sub.date), sub.dest, sub.docNo, sub.qty,
          sub.rateType === 'piece' ? 'ราคาชิ้น' : sub.rateType === 'flat' ? 'ราคาเหมา' : '-',
          sub.price != null ? sub.price : (sub.rateType === 'piece' ? 'หลายราคา' : ''), sub.amount, rowExtra || '', rowTotal, '',
        ]);
        r.eachCell((cell, col) => bodyCell(cell, { align: col >= 4 && col <= 9 ? 'right' : 'left', bg: zebra ? C.zebra : undefined, color: col === 9 ? C.title : undefined, bold: col === 9 }));
        [7, 8, 9].forEach((c) => (r.getCell(c).numFmt = NUM));
        if (sub.price != null) r.getCell(6).numFmt = NUM;
        if (parts.length) {
          const rt: ExcelJS.RichText[] = [];
          parts.forEach((p, i) => {
            if (i) rt.push({ font: { name: FONT, size: 12, color: { argb: 'FF999999' } }, text: ' · ' });
            rt.push({ font: { name: FONT, size: 13, bold: true, color: { argb: p.color } }, text: p.text });
          });
          r.getCell(10).value = { richText: rt };
        }
        zebra = !zebra;
        sumQty += sub.qty; sumMoney += sub.amount; sumTotal += rowTotal;
      });
      sumExtra += extra;
    }
    // แถวยอดรวม
    const tr = ws.addRow(['ยอดรวม', '', '', round2(sumQty), '', '', round2(sumMoney), round2(sumExtra), round2(sumTotal), '']);
    tr.eachCell((cell, col) => bodyCell(cell, { bold: true, align: col >= 4 ? 'right' : 'left', bg: C.totalBg, color: C.title }));
    [7, 8, 9].forEach((c) => (tr.getCell(c).numFmt = NUM));

    // ---- สรุปหัก/รับสุทธิ ----
    ws.addRow([]);
    const addSummary = (label: string, val: number, opts: { bold?: boolean; color?: string } = {}) => {
      const row = ws.addRow(['', '', '', '', '', label, val]);
      const lc = row.getCell(6); lc.value = label;
      lc.font = { name: FONT, size: 13, bold: opts.bold, color: opts.color ? { argb: opts.color } : { argb: C.sub } };
      lc.alignment = { horizontal: 'right' };
      const vc = row.getCell(7); vc.numFmt = NUM;
      vc.font = { name: FONT, size: 13, bold: opts.bold, color: opts.color ? { argb: opts.color } : undefined };
      vc.alignment = { horizontal: 'right' };
      return row;
    };
    // แยกรายได้เพิ่ม: มีเลขใบกระจาย = อยู่ในใบ (แยกตามชื่อ), ไม่มี = ประจำงวด (ค่าอัพบิล)
    const byLabel = (list: typeof vIncome) => Object.values(list.reduce((m: any, d) => {
      const k = d.label || 'รายการ'; (m[k] = m[k] || { label: k, amount: 0 }).amount += d.amount; return m;
    }, {})) as { label: string; amount: number }[];
    const inDocInc = byLabel(vIncome.filter((d) => normDoc(d.docNo || '')));
    const perCycleInc = byLabel(vIncome.filter((d) => !normDoc(d.docNo || '')));
    addSummary('รายได้ค่าเที่ยวทั้งหมด', round2(s.totalTripAmount), { bold: true });
    for (const l of inDocInc) addSummary(`+ ${l.label} (ในใบ)`, round2(l.amount));
    for (const l of perCycleInc) addSummary(`+ ${l.label}`, round2(l.amount));
    addSummary('หัก 1%', -round2(s.deduction1Percent));
    addSummary('หักค่าน้ำมัน', -round2(s.fuelTotal));
    for (const ln of s.lines.filter((l: any) => l.kind === 'deduction')) addSummary(`หัก ${ln.label}`, -round2(ln.amount));
    const net = addSummary('รวมรับสุทธิ', round2(s.netReceive), { bold: true, color: C.billingText });
    net.getCell(6).fill = solid(C.totalBg); net.getCell(7).fill = solid(C.totalBg);

    // ความกว้างคอลัมน์ (sheet ค่าบรรทุก)
    [13, 22, 18, 11, 11, 11, 13, 16, 13, 12].forEach((w, i) => (ws.getColumn(i + 1).width = w));

    // ---- sheet ใบสั่งเติมน้ำมัน (แยกต่างหากต่อทะเบียน) ----
    const wf = wb.addWorksheet(safeSheetName(`${plate} (น้ำมัน)`, usedNames));
    styleTitle(wf, `สรุปใบสั่งเติมน้ำมัน — สาขา${branchName}`, 4,
      `รอบ ${cycle.name}  ·  ทะเบียน ${plate}  ·  ${s.driverName || '-'}`);
    wf.addRow([]);
    styleHeaderRow(wf.addRow(['ลำดับ', 'วัน/เดือน/ปี', 'ใบสั่งเติมน้ำมัน', 'จำนวนเงิน (บาท)']));
    let fz = false, fSum = 0;
    vFuel.forEach((f, i) => {
      const r = wf.addRow([i + 1, fmtDate(f.date), f.refNo, f.amount]);
      r.eachCell((cell, col) => bodyCell(cell, { align: col === 1 || col === 4 ? 'right' : 'left', bg: fz ? C.zebra : undefined }));
      r.getCell(4).numFmt = NUM; fz = !fz; fSum += f.amount;
    });
    if (!vFuel.length) {
      const er = wf.addRow(['', 'ไม่มีใบสั่งเติมน้ำมัน', '', '']);
      er.eachCell((cell) => bodyCell(cell, { align: 'left' }));
    }
    const fr = wf.addRow(['', '', 'ผลรวม', round2(fSum)]);
    fr.eachCell((cell) => bodyCell(cell, { bold: true, align: 'right', bg: C.totalBg, color: C.title }));
    fr.getCell(4).numFmt = NUM;
    [10, 18, 24, 18].forEach((w, i) => (wf.getColumn(i + 1).width = w));
  }

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `รายงานสรุปค่าเที่ยว_${branchName}_${cycle.name.replace(/[\s/]/g, '_')}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ===========================================================================
// เทมเพลตนำเข้าราคาขนส่ง — 2 ชีต (เหมาคัน / รายชิ้น) พร้อมตัวอย่าง
// ===========================================================================
export async function downloadRateTemplate() {
  const wb = await newWorkbook();
  // ชีต เหมาคัน
  const f = wb.addWorksheet('เหมาคัน');
  styleHeaderRow(f.addRow(['ลำดับ', 'จังหวัด', 'อำเภอ', 'ราคา เหมา']));
  ([[1, 'สมุทรสาคร', 'อำเภอเมืองสมุทรสาคร', 1000],
    [2, 'สมุทรสาคร', 'อำเภอกระทุ่มแบน', 950],
    [3, 'ชลบุรี', 'อำเภอบางละมุง', 1500],
    [4, 'สุโขทัย', 'เมือง+กงไกรลาศ', 1400]] as any[][]).forEach((r) => {
    const row = f.addRow(r);
    row.eachCell((c, col) => bodyCell(c, { align: col >= 1 ? (col === 1 || col === 4 ? 'right' : 'left') : 'left' }));
  });
  [8, 18, 26, 12].forEach((w, i) => (f.getColumn(i + 1).width = w));

  // ชีต รายชิ้น
  const p = wb.addWorksheet('รายชิ้น');
  styleHeaderRow(p.addRow(['จังหวัด', 'อำเภอ', 'ราคา รายชิ้น']));
  ([['สมุทรสาคร', '', 7.5],
    ['ชลบุรี', '', 9.6],
    ['ชลบุรี', 'บางละมุง', 11.7],
    ['ชลบุรี', 'สัตหีบ', 11.7]] as any[][]).forEach((r) => {
    const row = p.addRow(r);
    row.eachCell((c, col) => bodyCell(c, { align: col === 3 ? 'right' : 'left' }));
  });
  [18, 20, 14].forEach((w, i) => (p.getColumn(i + 1).width = w));

  // ชีต พิเศษ — ทุกหมวด/เงื่อนไข (เก็บคืน/Peat/บวกเพิ่ม/กลุ่มราคา/จุดตัด/ขั้นบันได/เงื่อนไขผู้รับ-ผู้ส่ง-สินค้า)
  const s = wb.addWorksheet('พิเศษ');
  styleHeaderRow(s.addRow(['หมวด', 'จังหวัด', 'อำเภอ', 'ประเภทราคา', 'ราคา', 'กลุ่ม', 'จุดตัดชิ้น', 'กล่องตั้งแต่', 'ถึงกล่อง', 'ผู้รับ', 'ผู้ส่ง', 'สินค้า', 'หมายเหตุ']));
  ([
    ['เก็บคืน', 'สุโขทัย', 'ศรีสำโรง', 'เหมา', 1450, '', '', '', '', '', '', '', 'เก็บคืนอย่างเดียว = เลือกสูงกว่า เหมา/ชิ้น'],
    ['เก็บคืน', 'สุโขทัย', 'ศรีสำโรง', 'ชิ้น', 3.15, '', '', '', '', '', '', '', ''],
    ['Peat', 'ตาก', 'แม่สอด', 'ชิ้น', 26, '', '', '', '', '', '', '', 'Peat mass ราคาตามอำเภอ'],
    ['บวกเพิ่ม', 'ตาก', 'ท่าสองยาง', 'เหมา', 700, '', '', '', '', '', '', '', 'บวกเพิ่มตายตัวต่อปลายทาง'],
    ['งานปกติ', 'แม่ฮ่องสอน', 'ปาย', 'เหมา', 4100, 'กลุ่ม 1', '', '', '', '', '', '', 'ราคาเฉพาะรถกลุ่ม 1'],
    ['งานปกติ', 'ลำพูน', '', 'เหมา', 1200, '', '', 151, '', 'ซีพี ออลล์', '', '', 'ราคาขั้นบันได 151 กล่องขึ้นไป'],
  ] as any[][]).forEach((r) => {
    const row = s.addRow(r);
    row.eachCell((c, col) => bodyCell(c, { align: (col === 5 || col === 7 || col === 8 || col === 9) ? 'right' : 'left' }));
  });
  [10, 12, 14, 10, 8, 10, 9, 11, 9, 14, 14, 14, 26].forEach((w, i) => (s.getColumn(i + 1).width = w));

  // ชีต คำอธิบาย
  const g = wb.addWorksheet('วิธีใช้');
  ['วิธีกรอกเทมเพลตราคาขนส่ง',
    '',
    'ชีต "เหมาคัน": จังหวัด | อำเภอ | ราคาเหมา (1 แถวต่อ 1 อำเภอ)',
    '  - ส่งหลายอำเภอ ใช้เครื่องหมาย + เช่น "เมือง+กงไกรลาศ" (เป็นราคาเหมารวม ไม่มีราคาชิ้น)',
    '',
    'ชีต "รายชิ้น": จังหวัด | อำเภอ | ราคาชิ้น',
    '  - อำเภอเว้นว่าง = ใช้ทั้งจังหวัด (default)',
    '  - ใส่อำเภอ = ราคาเฉพาะอำเภอนั้น (override default) เช่น ชลบุรี/บางละมุง = 11.7',
    '',
    'ระบบจะเทียบอัตโนมัติ (max mode): จำนวน×ราคาชิ้น สูงกว่าเหมา → จ่ายชิ้น, ไม่งั้น → จ่ายเหมา',
    'ไม่ต้องตั้งจุดตัด/threshold เอง',
    '',
    'ชีต "พิเศษ" (เว้นว่างได้ ใช้เฉพาะที่ต้องการ): หมวด | จังหวัด | อำเภอ | ประเภทราคา | ราคา | กลุ่ม | จุดตัดชิ้น | กล่องตั้งแต่ | ถึงกล่อง | ผู้รับ | ผู้ส่ง | สินค้า | หมายเหตุ',
    '  - หมวด: งานปกติ / เก็บคืน / Peat / บวกเพิ่ม',
    '  - ประเภทราคา: เหมา หรือ ชิ้น (Peat=ชิ้นเสมอ, บวกเพิ่ม=เหมาเสมอ)',
    '  - เก็บคืน: ใส่ได้ทั้งเหมาและชิ้น — ถ้าเที่ยวนั้นเก็บคืนอย่างเดียว ระบบเลือกยอดสูงกว่า',
    '  - กลุ่ม: ใส่ชื่อกลุ่มราคา (เช่น "กลุ่ม 1") = ราคาเฉพาะรถกลุ่มนั้น (เว้นว่าง=ทุกกลุ่ม)',
    '  - จุดตัดชิ้น: ≤จุดตัด=เหมา, >จุดตัด=ชิ้น',
    '  - กล่องตั้งแต่/ถึงกล่อง: ราคาขั้นบันไดตามจำนวนกล่อง',
    '  - ผู้รับ/ผู้ส่ง/สินค้า: คำที่ต้องเจอ จึงใช้ราคานี้ (เช่น "ซีพี ออลล์")',
  ].forEach((t, i) => {
    const r = g.addRow([t]);
    if (i === 0) r.getCell(1).font = { name: FONT, size: 16, bold: true, color: { argb: C.title } };
    else r.getCell(1).font = { name: FONT, size: 13 };
  });
  g.getColumn(1).width = 90;

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'เทมเพลตราคาขนส่ง.xlsx';
  a.click();
  URL.revokeObjectURL(url);
}

// ===========================================================================
// เทมเพลตนำเข้าค่าน้ำมัน — แยก sheet ต่อทะเบียน (ตามรถของสาขา) + ชีตวิธีใช้
// ===========================================================================
export async function downloadFuelTemplate(
  branchName: string,
  vehicles: { plateNo: string; driverName?: string }[]
) {
  const wb = await newWorkbook();
  const used = new Set<string>();
  const sheetName = (p: string) => {
    let n = (p || 'รถ').replace(/[\/?*[\]:]/g, '-').slice(0, 28).trim();
    let base = n, i = 2;
    while (used.has(n)) n = `${base} (${i++})`;
    used.add(n);
    return n;
  };

  const list = vehicles.length ? vehicles : [{ plateNo: 'ตัวอย่าง', driverName: '' }];
  for (const v of list) {
    const ws = wb.addWorksheet(sheetName(v.plateNo));
    styleTitle(ws, `ใบสั่งเติมน้ำมัน — ${v.plateNo}${v.driverName ? ' · ' + v.driverName : ''}`, 4, `สาขา ${branchName}`);
    ws.addRow([]);
    styleHeaderRow(ws.addRow(['ลำดับที่', 'วัน/เดือน/ปี', 'ใบสั่งเติมน้ำมัน', 'เป็นจำนวนเงิน (บาท)']));
    // แถวว่างให้กรอก (ใส่เลขลำดับไว้ให้ ที่เหลือเว้นว่าง)
    for (let i = 1; i <= 15; i++) {
      const row = ws.addRow([i, '', '', '']);
      row.eachCell((c, col) => bodyCell(c, { align: col === 1 || col === 4 ? 'right' : 'left' }));
      row.getCell(4).numFmt = NUM;
    }
    [10, 18, 22, 18].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  }

  const g = wb.addWorksheet('วิธีใช้');
  ['วิธีกรอกเทมเพลตค่าน้ำมัน',
    '',
    '• แต่ละ sheet = 1 ทะเบียนรถ (ระบบสร้างให้ตามรถของสาขาแล้ว)',
    '• กรอกแต่ละแถว: วัน/เดือน/ปี · เลขใบสั่งเติมน้ำมัน · จำนวนเงิน',
    '• วันที่ใส่ได้หลายแบบ: 2/6/2026 · 2026-06-01 · 2 มิ.ย. 2569 · 2 มิ.ย. 26',
    '• ระบบจะจัดเข้ารอบ (1-15 / 16-31) ตามวันที่อัตโนมัติ — ใส่ได้หลายรอบในไฟล์เดียว',
    '• เลขลำดับมีให้แล้ว แถวไหนไม่ใช้เว้นว่างได้ (ระบบข้ามแถวที่ไม่มีจำนวนเงิน)',
    '• เพิ่มทะเบียนใหม่: เพิ่ม sheet แล้วตั้งชื่อ sheet เป็นทะเบียนรถ',
  ].forEach((t, i) => {
    const r = g.addRow([t]);
    r.getCell(1).font = i === 0 ? { name: FONT, size: 16, bold: true, color: { argb: C.title } } : { name: FONT, size: 13 };
  });
  g.getColumn(1).width = 80;

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `เทมเพลตค่าน้ำมัน_${branchName}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ===========================================================================
// Export สรุปภาพรวมทุกสาขา (HQ Dashboard) -> Excel
// ===========================================================================
export async function exportBranchSummary(
  cycleName: string,
  rows: { branchName: string; docs: number; trucks: number; trip: number; fuel: number; income: number; deduct: number; net: number }[]
) {
  const wb = await newWorkbook();
  const ws = wb.addWorksheet('สรุปทุกสาขา');
  styleTitle(ws, 'สรุปภาพรวมทุกสาขา — NEOSIAM', 8, `รอบ ${cycleName}`);
  ws.addRow([]);
  styleHeaderRow(ws.addRow(['สาขา', 'ใบกระจาย', 'รถ', 'ค่าเที่ยว', 'ค่าน้ำมัน', '+ รายได้เพิ่ม', 'รวมหัก', 'รับสุทธิ']));
  let z = false;
  const g = { docs: 0, trip: 0, fuel: 0, income: 0, deduct: 0, net: 0 };
  for (const r of rows) {
    const row = ws.addRow([r.branchName, r.docs, r.trucks, round2(r.trip), round2(r.fuel), round2(r.income), round2(r.deduct), round2(r.net)]);
    row.eachCell((cell, col) => bodyCell(cell, { align: col === 1 ? 'left' : 'right', bg: z ? C.zebra : undefined, bold: col === 8, color: col === 8 ? C.billingText : undefined }));
    [4, 5, 6, 7, 8].forEach((c) => (row.getCell(c).numFmt = NUM));
    z = !z;
    g.docs += r.docs; g.trip += r.trip; g.fuel += r.fuel; g.income += r.income; g.deduct += r.deduct; g.net += r.net;
  }
  const tr = ws.addRow(['รวมทุกสาขา', g.docs, '', round2(g.trip), round2(g.fuel), round2(g.income), round2(g.deduct), round2(g.net)]);
  tr.eachCell((cell, col) => bodyCell(cell, { bold: true, align: col === 1 ? 'left' : 'right', bg: C.totalBg, color: C.title }));
  [4, 5, 6, 7, 8].forEach((c) => (tr.getCell(c).numFmt = NUM));
  [18, 12, 8, 15, 15, 15, 13, 15].forEach((w, i) => (ws.getColumn(i + 1).width = w));

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `สรุปทุกสาขา_${cycleName.replace(/[\s/]/g, '_')}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ===========================================================================
// Export วิเคราะห์รายได้ พขร (Driver KPI) -> Excel (เป้าหมาย / รายได้คนรถ / KPI สาขา)
// ===========================================================================
export async function exportDriverKpi(
  cycleName: string,
  t: { totalExpense: number; boxesPerMonth: number; targetPerBox: number; targetPerCycle: number; targetBoxCycle: number; items?: { label: string; amount: number }[] },
  drivers: { plate: string; driver: string; branch: string; boxes: number; net: number; trip: number }[],
  branchRows: { branch: string; boxes: number; trip: number; perBox: number; nHit: number; nDrivers: number; nTrips?: number; boxPerTrip?: number; boxPerDriver?: number; pctHit?: number }[]
) {
  const wb = await newWorkbook();
  const INT = '#,##0';

  // ---- เป้าหมาย ----
  {
    const ws = wb.addWorksheet('เป้าหมาย');
    styleTitle(ws, 'เป้าหมายรายได้ พขร (KPI)', 2, `รอบ ${cycleName}`);
    ws.addRow([]);
    const add = (k: string, v: number, money?: boolean) => {
      const r = ws.addRow([k, round2(v)]);
      bodyCell(r.getCell(1), { bold: true });
      bodyCell(r.getCell(2), { align: 'right' });
      r.getCell(2).numFmt = money ? NUM : INT;
    };
    for (const it of t.items || []) {
      const r = ws.addRow([`• ${it.label}`, round2(+it.amount || 0)]);
      bodyCell(r.getCell(1), {}); bodyCell(r.getCell(2), { align: 'right' }); r.getCell(2).numFmt = NUM;
    }
    add('รายจ่ายเป้า/เดือน (รวม)', t.totalExpense, true);
    add('กล่อง/เดือน (capacity)', t.boxesPerMonth);
    add('เป้า บาท/กล่อง', t.targetPerBox, true);
    add('เป้ารายรับ/รอบ', t.targetPerCycle, true);
    add('เป้ากล่อง/รอบ', t.targetBoxCycle);
    [28, 18].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  }

  // ---- รายได้ต่อคนรถ ----
  {
    const ws = wb.addWorksheet('รายได้ พขร');
    styleTitle(ws, 'รายได้ต่อคนรถ', 8, `รอบ ${cycleName} · เป้ากล่อง/รอบ ${Math.round(t.targetBoxCycle)} · เป้า ${t.targetPerBox.toFixed(2)} บ/กล่อง`);
    ws.addRow([]);
    styleHeaderRow(ws.addRow(['ทะเบียน', 'คนขับ', 'สาขา', 'กล่องวิ่ง', '%เป้า', 'รับสุทธิ', 'บาท/กล่อง', 'สถานะ']));
    let z = false;
    for (const d of drivers) {
      const pct = t.targetBoxCycle > 0 ? d.boxes / t.targetBoxCycle : 0;
      const perBox = d.boxes > 0 ? d.net / d.boxes : 0;
      const col = pct >= 1 ? 'FF0E9F6E' : pct >= 0.8 ? 'FFB7791F' : 'FFE10026';
      const text = pct >= 1 ? 'ถึงเป้า' : pct >= 0.8 ? 'ใกล้เป้า' : 'ต่ำกว่าเป้า';
      const r = ws.addRow([d.plate, d.driver, d.branch, Math.round(d.boxes), pct, round2(d.net), round2(perBox), text]);
      r.eachCell((cell, c) => bodyCell(cell, { align: c <= 3 ? 'left' : 'right', bg: z ? C.zebra : undefined }));
      r.getCell(4).numFmt = INT; r.getCell(5).numFmt = '0%'; r.getCell(6).numFmt = NUM; r.getCell(7).numFmt = NUM;
      r.getCell(5).font = { name: FONT, size: 13, bold: true, color: { argb: col } };
      r.getCell(8).font = { name: FONT, size: 13, bold: true, color: { argb: col } };
      z = !z;
    }
    [14, 20, 16, 12, 10, 14, 12, 14].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  }

  // ---- KPI สาขา ----
  {
    const ws = wb.addWorksheet('KPI สาขา');
    styleTitle(ws, 'KPI สาขา — การจัดงานคนรถ (ค่ากระจาย บาท/กล่อง ต่ำสุด = ดีสุด)', 9, `รอบ ${cycleName}`);
    ws.addRow([]);
    styleHeaderRow(ws.addRow(['อันดับ', 'สาขา', 'กล่องรวม', 'เที่ยว', 'กล่อง/เที่ยว', 'กล่อง/คนรถ', 'ค่าเที่ยวรวม', 'ค่ากระจาย บาท/กล่อง', '%ถึงเป้า']));
    let z = false;
    // Excel: ใช้ค่าเที่ยว (ก่อนหัก) เสมอ + เรียงตาม ค่ากระจาย/กล่อง (gross) ให้คงที่ ไม่ขึ้นกับสวิตช์บนจอ
    const gRows = [...branchRows].map((g) => ({ ...g, _per: g.boxes > 0 ? g.trip / g.boxes : 0 })).sort((a, b) => a._per - b._per);
    gRows.forEach((g, i) => {
      const r = ws.addRow([i + 1, g.branch, Math.round(g.boxes), Math.round(g.nTrips || 0), Math.round(g.boxPerTrip || 0), Math.round(g.boxPerDriver || 0), round2(g.trip), round2(g._per), (g.pctHit || 0) / 100]);
      r.eachCell((cell, c) => bodyCell(cell, { align: c <= 1 ? 'left' : 'right', bg: z ? C.zebra : undefined, bold: c === 8, color: c === 8 ? (i === 0 ? 'FF0E9F6E' : C.title) : undefined }));
      [3, 4, 5, 6].forEach((c) => (r.getCell(c).numFmt = INT));
      r.getCell(7).numFmt = NUM; r.getCell(8).numFmt = NUM; r.getCell(9).numFmt = '0%';
      z = !z;
    });
    [10, 20, 12, 10, 14, 14, 16, 20, 12].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  }

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `วิเคราะห์รายได้พขร_${cycleName.replace(/[\s/]/g, '_')}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ===========================================================================
// Export พื้นที่ต้นทุนสูง (ต่อสาขา -> จังหวัด) -> Excel
// ===========================================================================
export async function exportCostAreas(
  cycleName: string,
  data: { branch: string; total: number; provs: { prov: string; boxes: number; cost: number; docs: number; perBox: number; dists?: { dist: string; boxes: number; cost: number; docs: number; perBox: number }[] }[] }[],
  thr = 0
) {
  const wb = await newWorkbook();
  const INT = '#,##0';
  const ws = wb.addWorksheet('พื้นที่ต้นทุนสูง');
  styleTitle(ws, 'พื้นที่ต้นทุนสูง — ต่อสาขา', 5, `รอบ ${cycleName} · เรียงตาม บาท/กล่อง มาก→น้อย${thr > 0 ? ` · กรองเฉพาะ ≥ ${thr} บ/กล่อง` : ''}`);
  ws.addRow([]);
  styleHeaderRow(ws.addRow(['สาขา / จังหวัด / อำเภอ', 'กล่อง', 'จำนวนใบ', 'ค่าขนส่งรวม', 'บาท/กล่อง']));
  for (const d of data) {
    const hr = ws.addRow([`🏢 ${d.branch} — ต้นทุนรวม`, '', '', round2(d.total), '']);
    hr.eachCell((cell, col) => bodyCell(cell, { bold: true, bg: C.subBg, color: C.title, align: col === 1 ? 'left' : 'right' }));
    hr.getCell(4).numFmt = NUM;
    let z = false;
    d.provs.forEach((p, i) => {
      const r = ws.addRow([`   จ.${p.prov}`, Math.round(p.boxes), p.docs, round2(p.cost), round2(p.perBox)]);
      r.eachCell((cell, col) => bodyCell(cell, { bold: true, align: col === 1 ? 'left' : 'right', bg: z ? C.zebra : undefined }));
      r.getCell(2).numFmt = INT; r.getCell(3).numFmt = INT; r.getCell(4).numFmt = NUM; r.getCell(5).numFmt = NUM;
      const hi = i === 0, lo = i === d.provs.length - 1 && d.provs.length > 1;
      if (hi || lo) r.getCell(5).font = { name: FONT, size: 13, bold: true, color: { argb: hi ? 'FFE10026' : 'FF0E9F6E' } };
      z = !z;
      // อำเภอย่อย
      (p.dists || []).forEach((dd, k) => {
        const dr = ws.addRow([`        อ.${dd.dist}`, Math.round(dd.boxes), dd.docs, round2(dd.cost), round2(dd.perBox)]);
        dr.eachCell((cell, col) => bodyCell(cell, { align: col === 1 ? 'left' : 'right', color: 'FF666666' }));
        dr.getCell(2).numFmt = INT; dr.getCell(3).numFmt = INT; dr.getCell(4).numFmt = NUM; dr.getCell(5).numFmt = NUM;
        const dhi = k === 0, dlo = k === (p.dists || []).length - 1 && (p.dists || []).length > 1;
        if (dhi || dlo) dr.getCell(5).font = { name: FONT, size: 13, color: { argb: dhi ? 'FFE10026' : 'FF0E9F6E' } };
      });
    });
  }
  [34, 12, 11, 16, 14].forEach((w, i) => (ws.getColumn(i + 1).width = w));

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `พื้นที่ต้นทุนสูง_${cycleName.replace(/[\s/]/g, '_')}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}
