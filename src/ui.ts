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

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] as string));
}
