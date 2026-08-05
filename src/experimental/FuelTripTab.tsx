// ============================================================================
// [ทดลอง] แท็บ "คำนวณค่าจ้างรถร่วม (น้ำมัน)" — แยก 100% ไม่แตะ component เดิม
// 2 ส่วน: (1) คำนวณค่าจ้าง 1 เที่ยว  (2) สรุปราคาน้ำมันย้อนหลังแต่ละสาขา
// ============================================================================
import { useState, useEffect, useRef } from 'react';

const API = (import.meta as any).env?.BASE_URL || '/';
const api = (p: string) => `${API.replace(/\/$/, '')}${p}`;
const baht = (n: number | null | undefined) => (n == null ? '-' : n.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));

// สาขาจริง (สำหรับ dropdown คำนวณ) — สาย3=กทม., แม่สอด=ตาก/อ.แม่สอด (backend map ให้)
const BRANCHES = ['สาย3', 'นครสวรรค์', 'กำแพงเพชร', 'พิษณุโลก', 'แม่สอด', 'เชียงใหม่'];

interface Breakdown {
  driverAllowance: number; fuelCost: number; extraCharges: number;
  tripRateBeforeRounding: number; boxRateUnrounded: number; boxRateDisplay: number;
  travelHours: number; unloadingHours: number; speedKmh: number; estimatedRoundTripKm: number;
  paymentMode: string; finalPayment: number;
}

export default function FuelTripTab() {
  const [sub, setSub] = useState<'calc' | 'history'>('calc');
  return (
    <div className="space-y-4">
      <div className="rounded-xl bg-amber-50 border border-amber-200 px-4 py-2 text-sm text-amber-800">
        🧪 <b>หน้าทดลอง</b> — คำนวณค่าจ้างรถร่วมจากราคาน้ำมัน OR (แยกจากระบบค่าเที่ยวเดิม 100% ยังไม่ใช้จริง)
      </div>
      <div className="flex gap-2">
        <button onClick={() => setSub('calc')} className={`px-4 py-2 rounded-lg text-sm font-semibold ${sub === 'calc' ? 'bg-brand-navy text-white' : 'bg-natural-100 text-natural-muted'}`}>คำนวณค่าจ้าง 1 เที่ยว</button>
        <button onClick={() => setSub('history')} className={`px-4 py-2 rounded-lg text-sm font-semibold ${sub === 'history' ? 'bg-brand-navy text-white' : 'bg-natural-100 text-natural-muted'}`}>ราคาน้ำมันย้อนหลัง (ทุกสาขา)</button>
      </div>
      {sub === 'calc' ? <CalcSection /> : <HistorySection />}
    </div>
  );
}

function CalcSection() {
  const [branch, setBranch] = useState(BRANCHES[0]);
  const [oil, setOil] = useState<{ price: number; location: string } | null>(null);
  const [oilBusy, setOilBusy] = useState(false);
  const [oilErr, setOilErr] = useState('');
  const [form, setForm] = useState({ distanceMinKm: 85, distanceMaxKm: 113, storeCount: 3, actualBoxes: 250, fuelPrice: 0 });
  const [result, setResult] = useState<Breakdown | null>(null);
  const reqRef = useRef(0); // กัน response เก่ามาทับตอนเปลี่ยนสาขาเร็วๆ (race)

  // ดึงราคาน้ำมันตามสาขา (backend map: สาย3=กทม, แม่สอด=ตาก/อ.แม่สอด, อื่น=อ.เมือง)
  const loadOil = async (br: string) => {
    const myReq = ++reqRef.current;
    setOilBusy(true); setOilErr(''); setOil(null);
    try {
      const r = await fetch(api(`/api/experimental/oil-price?branch=${encodeURIComponent(br)}`));
      const j = await r.json();
      if (myReq !== reqRef.current) return; // มี request ใหม่กว่าแล้ว -> ทิ้งผลเก่า
      if (j.diesel) { setOil({ price: j.diesel.price, location: j.diesel.location }); setForm((f) => ({ ...f, fuelPrice: j.diesel.price })); }
      else setOilErr(j.error || 'ไม่พบราคาดีเซล');
    } catch (e: any) { if (myReq === reqRef.current) setOilErr(e.message); }
    finally { if (myReq === reqRef.current) setOilBusy(false); }
  };
  useEffect(() => { loadOil(branch); }, [branch]);

  const calc = async () => {
    const r = await fetch(api('/api/experimental/fuel-trip/calculate'), {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ input: { distanceMinKm: form.distanceMinKm, distanceMaxKm: form.distanceMaxKm, storeCount: form.storeCount, appliedFuelPrice: form.fuelPrice, actualBoxes: form.actualBoxes } }),
    });
    setResult((await r.json()).breakdown);
  };

  const num = (v: string) => (v === '' ? 0 : Number(v));
  return (
    <div className="grid md:grid-cols-2 gap-4">
      {/* ฟอร์ม */}
      <div className="rounded-xl bg-white border p-4 space-y-3">
        <h3 className="font-bold text-brand-navy">ข้อมูลเที่ยว</h3>
        <label className="block text-sm">สาขา
          <select value={branch} onChange={(e) => setBranch(e.target.value)} className="mt-1 w-full border rounded-lg px-3 py-2">
            {BRANCHES.map((b) => <option key={b} value={b}>{b}</option>)}
          </select>
        </label>
        <div className="text-sm rounded-lg bg-natural-50 px-3 py-2">
          ราคาดีเซล อ.เมือง: {oilBusy ? 'กำลังดึง...' : oil ? <b className="text-emerald-700">฿{baht(oil.price)}/ลิตร</b> : <span className="text-brand-red">{oilErr || '-'}</span>}
          {oil && <span className="text-natural-muted"> ({oil.location} · จาก OR)</span>}
        </div>
        <div className="grid grid-cols-2 gap-2 text-sm">
          <label>ระยะต่ำสุด (กม.)<input type="number" value={form.distanceMinKm} onChange={(e) => setForm({ ...form, distanceMinKm: num(e.target.value) })} className="mt-1 w-full border rounded-lg px-2 py-1.5" /></label>
          <label>ระยะสูงสุด (กม.)<input type="number" value={form.distanceMaxKm} onChange={(e) => setForm({ ...form, distanceMaxKm: num(e.target.value) })} className="mt-1 w-full border rounded-lg px-2 py-1.5" /></label>
          <label>จำนวนร้าน<input type="number" value={form.storeCount} onChange={(e) => setForm({ ...form, storeCount: num(e.target.value) })} className="mt-1 w-full border rounded-lg px-2 py-1.5" /></label>
          <label>จำนวนกล่อง<input type="number" value={form.actualBoxes} onChange={(e) => setForm({ ...form, actualBoxes: num(e.target.value) })} className="mt-1 w-full border rounded-lg px-2 py-1.5" /></label>
          <label className="col-span-2">ราคาน้ำมันที่ใช้ (แก้เองได้)<input type="number" step="0.01" value={form.fuelPrice} onChange={(e) => setForm({ ...form, fuelPrice: num(e.target.value) })} className="mt-1 w-full border rounded-lg px-2 py-1.5" /></label>
        </div>
        <button onClick={calc} className="w-full bg-brand-red text-white rounded-lg py-2 font-semibold">คำนวณ</button>
      </div>
      {/* ผลลัพธ์ */}
      <div className="rounded-xl bg-white border p-4">
        <h3 className="font-bold text-brand-navy mb-2">รายละเอียดการคำนวณ</h3>
        {!result ? <p className="text-natural-muted text-sm">กรอกข้อมูลแล้วกดคำนวณ</p> : (
          <table className="w-full text-sm">
            <tbody>
              <Row k="ระยะไป-กลับ (ประมาณ)" v={`${result.estimatedRoundTripKm.toFixed(1)} กม. @ ${result.speedKmh} กม./ชม.`} />
              <Row k="เวลาเดินทาง + ลงสินค้า" v={`${result.travelHours.toFixed(2)} + ${result.unloadingHours.toFixed(2)} ชม.`} />
              <Row k="เบี้ยขับ (คนขับ)" v={`฿${baht(result.driverAllowance)}`} />
              <Row k="ค่าน้ำมัน" v={`฿${baht(result.fuelCost)}`} />
              {result.extraCharges > 0 && <Row k="ค่าพิเศษ (เขา/เส้นทาง/ภาษี)" v={`฿${baht(result.extraCharges)}`} />}
              <Row k="ค่าเที่ยว (ก่อนปัด)" v={`฿${baht(result.tripRateBeforeRounding)}`} bold />
              <Row k="เรทต่อกล่อง" v={`฿${baht(result.boxRateDisplay)}`} />
              <tr><td colSpan={2} className="pt-2"><div className="rounded-lg bg-emerald-50 px-3 py-2 flex justify-between items-center">
                <span className="font-bold text-emerald-800">จ่ายจริง ({result.paymentMode === 'LUMP_SUM' ? 'เหมาทั้งเที่ยว' : 'ตามกล่อง'})</span>
                <span className="font-bold text-lg text-emerald-800">฿{baht(result.finalPayment)}</span>
              </div></td></tr>
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function Row({ k, v, bold }: { k: string; v: string; bold?: boolean }) {
  return <tr className="border-b last:border-0"><td className="py-1.5 text-natural-muted">{k}</td><td className={`py-1.5 text-right ${bold ? 'font-bold text-brand-navy' : ''}`}>{v}</td></tr>;
}

function HistorySection() {
  const [saveDate, setSaveDate] = useState(''); // วันที่จะบันทึก (ว่าง = วันนี้)
  const [preview, setPreview] = useState<any[]>([]); // ดึงสดเพื่อดูก่อนบันทึก
  const [saved, setSaved] = useState<any[]>([]);     // ที่บันทึกถาวรแล้ว
  const [busy, setBusy] = useState('');
  const [msg, setMsg] = useState('');

  const toOrDate = (d: string) => (d ? d.split('-').reverse().join('-') : ''); // YYYY-MM-DD -> DD-MM-YYYY

  // ดึงสด (ยังไม่บันทึก) เพื่อดูก่อน — ใช้ config สาขาจริง (ไม่ส่ง provinces)
  const loadPreview = async () => {
    setBusy('preview'); setMsg('');
    try {
      const q = saveDate ? `?date=${toOrDate(saveDate)}` : '';
      const r = await fetch(api(`/api/experimental/oil-price/branches${q}`));
      setPreview((await r.json()).branches || []);
    } finally { setBusy(''); }
  };

  // อ่านที่บันทึกถาวรแล้ว
  const loadSaved = async () => {
    const r = await fetch(api('/api/experimental/oil-price/saved'));
    setSaved((await r.json()).records || []);
  };
  useEffect(() => { loadSaved(); }, []); // eslint-disable-line

  // บันทึกลง Firebase (ถาวร -> Hyper Backup ลง NAS)
  const save = async () => {
    setBusy('save'); setMsg('');
    try {
      const r = await fetch(api('/api/experimental/oil-price/save'), {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ date: saveDate ? toOrDate(saveDate) : undefined }), // ไม่ส่ง provinces = ใช้ config สาขาจริง
      });
      const j = await r.json();
      setMsg(j.success ? `✅ บันทึกถาวรแล้ว ${j.savedCount} สาขา${j.skipped?.length ? ` (ข้าม: ${j.skipped.join(', ')})` : ''}` : `❌ ${j.error}`);
      await loadSaved();
    } catch (e: any) { setMsg('❌ ' + e.message); } finally { setBusy(''); }
  };

  return (
    <div className="space-y-4">
      {/* ส่วนบันทึก */}
      <div className="rounded-xl bg-white border p-4 space-y-3">
        <h3 className="font-bold text-brand-navy">บันทึกราคาน้ำมัน (ถาวร)</h3>
        <div className="flex items-end gap-2 flex-wrap">
          <label className="text-sm">วันที่ (เว้นว่าง = วันนี้)
            <input type="date" value={saveDate} onChange={(e) => setSaveDate(e.target.value)} className="mt-1 block border rounded-lg px-3 py-2" />
          </label>
          <button onClick={loadPreview} disabled={!!busy} className="bg-natural-100 text-brand-navy rounded-lg px-4 py-2 text-sm font-semibold">ดูราคาก่อน</button>
          <button onClick={save} disabled={!!busy} className="bg-brand-red text-white rounded-lg px-4 py-2 text-sm font-semibold">💾 บันทึกถาวร</button>
          {busy && <span className="text-sm text-natural-muted">{busy === 'save' ? 'กำลังบันทึก...' : 'กำลังดึงจาก OR...'}</span>}
        </div>
        {msg && <div className="text-sm rounded-lg bg-natural-50 px-3 py-2">{msg}</div>}
        {preview.length > 0 && (
          <table className="w-full text-sm">
            <thead><tr className="bg-natural-100 text-left"><th className="px-3 py-2 rounded-l-lg">สาขา</th><th className="px-3 py-2">ราคาจาก (อ.เมือง)</th><th className="px-3 py-2 text-right">ดีเซล</th><th className="px-3 py-2 text-right">ช่วงจังหวัด</th><th className="px-3 py-2 rounded-r-lg text-right">วันที่</th></tr></thead>
            <tbody>{preview.map((b) => (
              <tr key={b.branch} className="border-b last:border-0">
                <td className="px-3 py-2 font-semibold">{b.branch}</td>
                <td className="px-3 py-2 text-natural-muted">{b.location || b.province}</td>
                <td className="px-3 py-2 text-right font-bold">{b.diesel != null ? `฿${baht(b.diesel)}` : <span className="text-brand-red">ไม่เจอ</span>}</td>
                <td className="px-3 py-2 text-right text-natural-muted">{b.min != null && b.min !== b.max ? `฿${baht(b.min)}–${baht(b.max)}` : '-'}</td>
                <td className="px-3 py-2 text-right text-natural-muted">{b.priceDate ? String(b.priceDate).slice(0, 10) : '-'}</td>
              </tr>
            ))}</tbody>
          </table>
        )}
      </div>

      {/* ส่วนที่บันทึกถาวรแล้ว */}
      <div className="rounded-xl bg-white border p-4 space-y-2">
        <h3 className="font-bold text-brand-navy">ประวัติราคาที่บันทึกไว้ ({saved.length})</h3>
        <table className="w-full text-sm">
          <thead><tr className="bg-natural-100 text-left"><th className="px-3 py-2 rounded-l-lg">วันที่ราคา</th><th className="px-3 py-2">สาขา</th><th className="px-3 py-2">ราคาจาก</th><th className="px-3 py-2 text-right">ดีเซล</th><th className="px-3 py-2 rounded-r-lg text-right">ช่วงจังหวัด</th></tr></thead>
          <tbody>
            {saved.map((r) => (
              <tr key={r.id} className="border-b last:border-0">
                <td className="px-3 py-2">{r.priceDate}</td>
                <td className="px-3 py-2 font-semibold">{r.branch}</td>
                <td className="px-3 py-2 text-natural-muted">{r.location}</td>
                <td className="px-3 py-2 text-right font-bold text-brand-navy">฿{baht(r.price)}</td>
                <td className="px-3 py-2 text-right text-natural-muted">{r.minPrice != null && r.minPrice !== r.maxPrice ? `฿${baht(r.minPrice)}–${baht(r.maxPrice)}` : '-'}</td>
              </tr>
            ))}
            {!saved.length && <tr><td colSpan={5} className="px-3 py-4 text-center text-natural-muted">ยังไม่มีราคาที่บันทึก — กด "บันทึกถาวร"</td></tr>}
          </tbody>
        </table>
        <p className="text-xs text-natural-muted">เก็บใน Firebase → Hyper Backup ลง NAS อัตโนมัติ · ตรวจย้อนหลังได้ · ที่มา OR</p>
      </div>
    </div>
  );
}
