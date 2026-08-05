// ============================================================================
// [ทดลอง] Export รายงานเทียบต้นทุนต่อทะเบียน -> Excel (.xlsx) 2 sheet
//   Sheet1 สรุปต่อทะเบียน · Sheet2 รายละเอียดต่อใบ
// ตัวเงินเป็น number จริง (คำนวณต่อใน Excel ได้) · จัดรูปแบบหัวสี/zebra/สรุปยอด
// แยก 100% ไม่แตะ excel-export.ts เดิม — โหลด exceljs แบบ dynamic เฉพาะตอนกด export
// ============================================================================
import type ExcelJS from 'exceljs';

const FONT = 'Tahoma';
const C = {
  headerBg: 'FF1F3864', headerText: 'FFFFFFFF', title: 'FF1F3864', sub: 'FF7F7F7F',
  zebra: 'FFF2F5FA', pos: 'FF2E7D32', neg: 'FFC00000', border: 'FFCCCCCC',
  detailHdr: 'FF375623', totalBg: 'FFFFF2CC',
};
const NUM = '#,##0.00';
const solid = (argb: string): ExcelJS.Fill => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const thin: ExcelJS.Borders = {
  top: { style: 'thin', color: { argb: C.border } }, left: { style: 'thin', color: { argb: C.border } },
  bottom: { style: 'thin', color: { argb: C.border } }, right: { style: 'thin', color: { argb: C.border } },
} as ExcelJS.Borders;

function title(ws: ExcelJS.Worksheet, text: string, span: number, sub?: string) {
  ws.mergeCells(1, 1, 1, span);
  const t = ws.getCell(1, 1); t.value = text;
  t.font = { name: FONT, size: 18, bold: true, color: { argb: C.title } };
  ws.getRow(1).height = 28;
  if (sub) { ws.mergeCells(2, 1, 2, span); const s = ws.getCell(2, 1); s.value = sub; s.font = { name: FONT, size: 12, color: { argb: C.sub } }; }
}
function headerRow(row: ExcelJS.Row, bg = C.headerBg) {
  row.eachCell((cell) => {
    cell.font = { name: FONT, size: 13, bold: true, color: { argb: C.headerText } };
    cell.fill = solid(bg); cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; cell.border = thin;
  });
  row.height = 22;
}

export interface ReportRow {
  branchId: string; branch: string; plateNo: string;
  tripCount: number; tripAmount: number; fuelCostBand: number; diff: number; computed: number; missing: number;
}
export interface DetailItem {
  documentNo: string; dests: string[]; loopKm: number | null; storeCount: number;
  speedKmh: number | null; travelHours: number | null; unloadingHours: number | null;
  driverAllowance: number | null; fuelCost: number | null; mountainLiters: number;
  totalCost: number | null; tripAmount: number; diff: number | null; missing: string[];
}

/** สร้าง + ดาวน์โหลดไฟล์ Excel รายงานเทียบต้นทุน */
export async function exportFuelReportToExcel(
  cycleName: string, branchLabel: string, rows: ReportRow[],
  detailByPlate: Record<string, DetailItem[]>, // key = "branchId|plateNo"
) {
  const mod = await import('exceljs');
  const wb = new mod.default.Workbook();
  wb.creator = 'ระบบค่าจ้างรถร่วม (น้ำมัน)'; wb.created = new Date();
  const isComplete = (r: ReportRow) => r.computed > 0 && r.missing === 0;

  // ---------- Sheet 1: สรุปต่อทะเบียน ----------
  {
    const ws = wb.addWorksheet('สรุปต่อทะเบียน');
    title(ws, `เทียบต้นทุนขนส่งต่อทะเบียน — ${cycleName}`, 7, `สาขา: ${branchLabel} · ค่าเที่ยวจ่ายจริง เทียบ ต้นทุนตามน้ำมัน+ระยะจริง (DOH)`);
    ws.addRow([]);
    const hdr = ws.addRow(['ทะเบียน', 'สาขา', 'เที่ยว', 'ค่าเที่ยว (จ่ายจริง)', 'ต้นทุนน้ำมัน (คิด)', 'ส่วนต่าง', 'สถานะ']);
    headerRow(hdr);
    let zebra = false, tTrip = 0, tFuel = 0, tDiff = 0, done = 0;
    for (const r of rows) {
      const ok = isComplete(r);
      const row = ws.addRow([
        r.plateNo, r.branch, r.tripCount,
        Math.round(r.tripAmount * 100) / 100,
        ok ? Math.round(r.fuelCostBand * 100) / 100 : null,
        ok ? Math.round(r.diff * 100) / 100 : null,
        ok ? 'คิดครบ' : (r.computed > 0 ? `ขาด ${r.missing}` : 'ยังไม่คิด'),
      ]);
      row.eachCell((c) => { c.border = thin; c.font = { name: FONT, size: 12 }; });
      if (zebra) row.eachCell((c) => (c.fill = solid(C.zebra)));
      [3, 4, 5, 6].forEach((c) => { row.getCell(c).alignment = { horizontal: 'right' }; if (c >= 4) row.getCell(c).numFmt = NUM; });
      if (ok) { row.getCell(6).font = { name: FONT, size: 12, bold: true, color: { argb: r.diff >= 0 ? C.pos : C.neg } }; tTrip += r.tripAmount; tFuel += r.fuelCostBand; tDiff += r.diff; done++; }
      zebra = !zebra;
    }
    // แถวรวม (เฉพาะคิดครบ)
    const tr = ws.addRow([`รวมเฉพาะคิดครบ (${done}/${rows.length} ทะเบียน)`, '', '', Math.round(tTrip * 100) / 100, Math.round(tFuel * 100) / 100, Math.round(tDiff * 100) / 100, '']);
    ws.mergeCells(tr.number, 1, tr.number, 3);
    tr.eachCell((c) => { c.border = thin; c.fill = solid(C.totalBg); c.font = { name: FONT, size: 12, bold: true }; });
    [4, 5, 6].forEach((c) => { tr.getCell(c).numFmt = NUM; tr.getCell(c).alignment = { horizontal: 'right' }; });
    tr.getCell(6).font = { name: FONT, size: 12, bold: true, color: { argb: tDiff >= 0 ? C.pos : C.neg } };
    ws.columns = [{ width: 16 }, { width: 14 }, { width: 8 }, { width: 18 }, { width: 18 }, { width: 16 }, { width: 12 }];
    ws.views = [{ state: 'frozen', ySplit: 3 }];
  }

  // ---------- Sheet 2: รายละเอียดต่อใบ ----------
  {
    const ws = wb.addWorksheet('รายละเอียดต่อใบ');
    title(ws, `รายละเอียดต่อใบ — ${cycleName}`, 9, `สาขา: ${branchLabel} · เบี้ยขับ = (เวลาเดินทาง + จุด×นาที) × ค่าแรง · ค่าน้ำมัน = ระยะ÷อัตราสิ้นเปลือง × ราคาดีเซล`);
    ws.addRow([]);
    const hdr = ws.addRow(['ทะเบียน', 'ใบกระจาย', 'ปลายทาง (ลำดับวิ่ง)', 'ระยะลูป (กม.)', 'จุดลง', 'เบี้ยขับ', 'ค่าน้ำมัน', 'ต้นทุนรวม', 'ค่าเที่ยว', 'ส่วนต่าง']);
    headerRow(hdr, C.detailHdr);
    let zebra = false;
    for (const r of rows) {
      const items = detailByPlate[`${r.branchId}|${r.plateNo}`] || [];
      for (const it of items) {
        const dests = it.dests.join(' → ') + (it.missing?.length ? ` (หาไม่เจอ: ${it.missing.join(', ')})` : '') + (it.mountainLiters > 0 ? ` ⛰️+${it.mountainLiters}ล.` : '');
        const row = ws.addRow([
          r.plateNo, it.documentNo, dests,
          it.loopKm, it.storeCount,
          it.driverAllowance, it.fuelCost, it.totalCost,
          Math.round(it.tripAmount * 100) / 100, it.diff,
        ]);
        row.eachCell((c) => { c.border = thin; c.font = { name: FONT, size: 11 }; c.alignment = { vertical: 'top' }; });
        if (zebra) row.eachCell((c) => (c.fill = solid(C.zebra)));
        [4, 6, 7, 8, 9, 10].forEach((c) => { row.getCell(c).alignment = { horizontal: 'right', vertical: 'top' }; if (c !== 4) row.getCell(c).numFmt = NUM; });
        row.getCell(3).alignment = { wrapText: true, vertical: 'top' };
        if (it.diff != null) row.getCell(10).font = { name: FONT, size: 11, bold: true, color: { argb: it.diff >= 0 ? C.pos : C.neg } };
        zebra = !zebra;
      }
    }
    ws.columns = [{ width: 14 }, { width: 16 }, { width: 42 }, { width: 12 }, { width: 8 }, { width: 12 }, { width: 12 }, { width: 13 }, { width: 13 }, { width: 13 }];
    ws.views = [{ state: 'frozen', ySplit: 3 }];
  }

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `เทียบต้นทุนรถร่วม_${branchLabel}_${cycleName.replace(/[\s/]/g, '_')}.xlsx`;
  a.click(); URL.revokeObjectURL(a.href);
}
