// อ่านไฟล์ใบปะ (Excel) -> แถว ทะเบียน/คนขับ/เลขบัญชี/หมายเหตุ  (ไม่ตัดสินถูกผิด — planBankImport เป็นคนตัดสิน)
import * as XLSX from 'xlsx';
import type { BankSheetRow } from './src/bankAccounts.js';

const txt = (c: XLSX.CellObject | undefined): string => (c == null || c.v == null ? '' : String(c.w ?? c.v).trim());

export function parseBankSheets(buffer: Buffer): BankSheetRow[] {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false, cellText: true });
  const out: BankSheetRow[] = [];
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws?.['!ref']) continue;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const merges = ws['!merges'] || [];
    // เซลล์ผสาน (เชียงใหม่: 1 บัญชีคร่อม 4 ทะเบียน) -> อ่านค่าจากมุมซ้ายบน
    const cell = (r: number, c: number): XLSX.CellObject | undefined => {
      const m = merges.find((x) => r >= x.s.r && r <= x.e.r && c >= x.s.c && c <= x.e.c);
      return ws[XLSX.utils.encode_cell(m ? m.s : { r, c })];
    };
    let hr = -1, cPlate = -1, cAcc = -1, cDrv = -1, cNote = -1;
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 30) && hr < 0; r++) {
      let p = -1, a = -1, d = -1, n = -1, last = -1;
      for (let c = range.s.c; c <= range.e.c; c++) {
        const t = txt(ws[XLSX.utils.encode_cell({ r, c })]).replace(/\s+/g, '');
        if (!t) continue;
        last = c;
        if (p < 0 && /^ทะเบียน/.test(t)) p = c;
        else if (a < 0 && /^เลข(ที่)?บัญชี/.test(t)) a = c;
        else if (d < 0 && /^(คนขับ|ชื่อ)/.test(t)) d = c;
        else if (n < 0 && /^หมายเหตุ/.test(t)) n = c;
      }
      if (p >= 0 && a >= 0) { hr = r; cPlate = p; cAcc = a; cDrv = d; cNote = n >= 0 ? n : last + 1; }
    }
    if (hr < 0) continue;
    for (let r = hr + 1; r <= range.e.r; r++) {
      const plate = txt(cell(r, cPlate));
      if (/^รวม/.test(plate)) break;              // จบตารางโอน (ตารางถัดไป = ยังไม่มีเลขบัญชี)
      const ac = cell(r, cAcc);
      let account = txt(ac);
      let numeric = false;
      if (ac && ac.t === 'n') {
        // เซลล์ตัวเลข: ถ้ารูปแบบเซลล์เติม 0 นำหน้าให้แล้ว ใช้ข้อความที่เห็น ไม่งั้นใช้ค่าจริง (กัน 1.12E+09)
        const raw = Number.isInteger(ac.v) ? String(ac.v) : '';
        const shown = account.replace(/\D/g, '');
        if (!(shown.length > raw.length && shown.endsWith(raw))) { account = raw; numeric = true; }
      }
      if (!plate || !account || account === '-') continue; // ว่าง = ข้าม (ไม่ลบของเดิม)
      const nm = merges.find((x) => r >= x.s.r && r <= x.e.r && x.s.c === cNote); // หมายเหตุผสานหลายแถว
      const note = txt(ws[XLSX.utils.encode_cell(nm ? nm.s : { r, c: cNote })]);
      out.push({ sheet: name, row: r + 1, plate, driver: cDrv >= 0 ? txt(cell(r, cDrv)) : '', account, note: note.replace(/\s+/g, ' ').slice(0, 200), numeric });
    }
  }
  return out;
}
