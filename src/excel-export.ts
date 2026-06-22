// ============================================================================
// Excel Export ด้วย ExcelJS — รองรับฟอนต์ Cordia New, สี, Border, หลาย Sheet
// ตาม Requirement section 8-9
// ============================================================================
import ExcelJS from 'exceljs';
import {
  BillingCycle,
  TripDocument,
  FuelEntry,
  DeductionEntry,
  Vehicle,
  RateMaster,
} from './types';
import { summarizeByVehicle } from './calc';

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

// ===========================================================================
export async function exportCycleToExcel(
  cycle: BillingCycle,
  trips: TripDocument[],
  fuel: FuelEntry[],
  deductions: DeductionEntry[],
  vehicles: Vehicle[],
  rateMasters: RateMaster[]
) {
  const wb = new ExcelJS.Workbook();
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
    styleTitle(ws, `รายการหัก — ${cycle.name}`, 5);
    ws.addRow([]);
    styleHeaderRow(ws.addRow(['ทะเบียน', 'ทิศทาง', 'รายการ', 'จำนวนเงิน', 'หมายเหตุ']));
    let zebra = false;
    cycleDed.forEach((d) => {
      const r = ws.addRow([d.plateNo, d.kind === 'income' ? 'รายได้เพิ่ม (+)' : 'หักออก (-)', d.label, d.amount, d.note || '']);
      r.eachCell((cell, col) => bodyCell(cell, { align: col === 4 ? 'right' : 'left', bg: zebra ? C.zebra : undefined }));
      r.getCell(4).numFmt = NUM;
      zebra = !zebra;
    });
    ws.columns.forEach((c) => (c.width = 16));
    ws.getColumn(3).width = 22; ws.getColumn(5).width = 24;
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
