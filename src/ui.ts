// ============================================================================
// UI helpers — SweetAlert2 ปรับธีมเข้ากับระบบ (น้ำเงิน 1B365D)
// ใช้แทน window.alert / confirm / prompt และ toast
// ============================================================================
import Swal from 'sweetalert2';

const NAVY = '#1B365D';
const RED = '#C00000';

// instance หลัก: ปุ่มโค้งมน สีเข้าธีม ฟอนต์เดียวกับแอป
const Themed = Swal.mixin({
  buttonsStyling: true,
  confirmButtonColor: NAVY,
  cancelButtonColor: '#9ca3af',
  reverseButtons: true,
  customClass: {
    popup: 'rounded-2xl',
    title: 'text-natural-brand',
    confirmButton: 'rounded-full px-5 py-2 text-sm font-semibold',
    cancelButton: 'rounded-full px-5 py-2 text-sm font-semibold',
  },
});

const DELETE_PASSWORD = '1234';

// ยืนยันการลบ + กรอกรหัสผ่าน (กันเผลอกด)
export async function confirmDelete(itemLabel?: string): Promise<boolean> {
  const res = await Themed.fire({
    title: 'ยืนยันการลบ',
    html: itemLabel
      ? `ต้องการลบ <b>"${escapeHtml(itemLabel)}"</b> หรือไม่?<br><span style="font-size:12px;color:${RED}">ลบแล้วกู้คืนไม่ได้</span>`
      : `กรอกรหัสผ่านเพื่อยืนยันการลบ<br><span style="font-size:12px;color:${RED}">ลบแล้วกู้คืนไม่ได้</span>`,
    icon: 'warning',
    iconColor: RED,
    input: 'password',
    inputPlaceholder: 'รหัสผ่าน',
    inputAttributes: { autocapitalize: 'off', autocorrect: 'off', autocomplete: 'off' },
    showCancelButton: true,
    confirmButtonText: '🗑 ลบ',
    confirmButtonColor: RED,
    cancelButtonText: 'ยกเลิก',
    focusCancel: true,
    inputValidator: (v) => (v === DELETE_PASSWORD ? undefined : 'รหัสผ่านไม่ถูกต้อง'),
  });
  return res.isConfirmed;
}

// ยืนยันด้วยรหัสผ่าน (สำหรับการแก้ไข) — คืน true ถ้ารหัสถูก
export async function confirmPassword(actionLabel = 'แก้ไข'): Promise<boolean> {
  const res = await Themed.fire({
    title: `ยืนยันการ${actionLabel}`,
    html: `กรอกรหัสผ่านเพื่อ${actionLabel}`,
    icon: 'warning',
    iconColor: NAVY,
    input: 'password',
    inputPlaceholder: 'รหัสผ่าน',
    inputAttributes: { autocapitalize: 'off', autocorrect: 'off', autocomplete: 'off' },
    showCancelButton: true,
    confirmButtonText: `✎ ${actionLabel}`,
    cancelButtonText: 'ยกเลิก',
    inputValidator: (v) => (v === DELETE_PASSWORD ? undefined : 'รหัสผ่านไม่ถูกต้อง'),
  });
  return res.isConfirmed;
}

// ยืนยันทั่วไป (ใช่/ไม่ใช่) — คืน true ถ้ากดยืนยัน
export async function confirmAction(opts: {
  title: string; text?: string; html?: string; confirmText?: string; danger?: boolean;
}): Promise<boolean> {
  const res = await Themed.fire({
    title: opts.title,
    text: opts.text,
    html: opts.html,
    icon: 'question',
    showCancelButton: true,
    confirmButtonText: opts.confirmText || 'ยืนยัน',
    confirmButtonColor: opts.danger ? RED : NAVY,
    cancelButtonText: 'ยกเลิก',
  });
  return res.isConfirmed;
}

// แจ้งเตือนแบบ toast มุมจอ
export function notify(type: 'success' | 'error' | 'warning' | 'info', message: string) {
  Swal.fire({
    toast: true,
    position: 'top-end',
    icon: type,
    title: message,
    showConfirmButton: false,
    timer: 3500,
    timerProgressBar: true,
    customClass: { popup: 'rounded-xl' },
  });
}

// กล่องข้อความสั้น (แทน window.alert)
export function alertBox(title: string, text?: string, icon: 'success' | 'error' | 'warning' | 'info' = 'info') {
  Themed.fire({ title, text, icon, confirmButtonText: 'ตกลง' });
}

// กล่องรายงานแบบมีระดับ (ok/info/warn/error) — ใช้กับสรุปการนำเข้า
//   เดิม alertBox ยัดทุกอย่างเป็นข้อความก้อนเดียว อ่านไม่ออกเมื่อมี 45 รายการ (เจ้าของสั่งจัดรูปแบบ 19 ก.ย.69)
//   แต่ละหัวข้อเป็นการ์ดสีตามระดับ รายการย่อยเป็น bullet เลื่อนดูได้ ไอคอนใหญ่เลือกตามระดับที่แย่ที่สุด
export type ReportLevel = 'ok' | 'info' | 'warn' | 'error';
export interface ReportSection { level: ReportLevel; title: string; items?: string[] }
const LEVEL_STYLE: Record<ReportLevel, { bg: string; border: string; fg: string; badge: string; label: string }> = {
  ok:    { bg: '#ecfdf5', border: '#6ee7b7', fg: '#065f46', badge: '#059669', label: 'สำเร็จ' },
  info:  { bg: '#eff6ff', border: '#93c5fd', fg: '#1e3a8a', badge: '#2563eb', label: 'ข้อมูล' },
  warn:  { bg: '#fffbeb', border: '#fcd34d', fg: '#92400e', badge: '#d97706', label: 'ควรตรวจ' },
  error: { bg: '#fef2f2', border: '#fca5a5', fg: '#991b1b', badge: '#dc2626', label: 'ไม่ได้บันทึก' },
};
const LEVEL_RANK: Record<ReportLevel, number> = { ok: 0, info: 1, warn: 2, error: 3 };
export function reportBox(title: string, sections: ReportSection[]) {
  const worst = sections.reduce<ReportLevel>((w, s) => (LEVEL_RANK[s.level] > LEVEL_RANK[w] ? s.level : w), 'ok');
  const icon = worst === 'error' ? 'error' : worst === 'warn' ? 'warning' : worst === 'ok' ? 'success' : 'info';
  const html = sections.map((s) => {
    const st = LEVEL_STYLE[s.level] || LEVEL_STYLE.info;
    const items = (s.items || []).length
      ? `<ul style="margin:6px 0 0;padding-left:18px;max-height:220px;overflow:auto;font-size:12px;line-height:1.5">${s.items!.map((i) => `<li>${escapeHtml(i)}</li>`).join('')}</ul>`
      : '';
    return `<div style="background:${st.bg};border:1px solid ${st.border};border-left:5px solid ${st.badge};color:${st.fg};border-radius:10px;padding:8px 10px;margin:6px 0;text-align:left">
      <div style="font-weight:600;font-size:13px;display:flex;gap:8px;align-items:flex-start">
        <span style="background:${st.badge};color:#fff;border-radius:999px;padding:1px 8px;font-size:11px;white-space:nowrap">${st.label}</span>
        <span>${escapeHtml(s.title)}${s.items?.length ? ` <span style="opacity:.7;font-weight:400">(${s.items.length})</span>` : ''}</span>
      </div>${items}</div>`;
  }).join('');
  return Themed.fire({ title, icon, html, confirmButtonText: 'ตกลง', width: 720, customClass: { popup: 'rounded-2xl', htmlContainer: 'text-left' } });
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] as string));
}
