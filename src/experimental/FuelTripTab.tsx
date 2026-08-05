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
  const [sub, setSub] = useState<'calc' | 'history' | 'report'>('report');
  const tabBtn = (k: typeof sub, label: string) => (
    <button onClick={() => setSub(k)} className={`px-4 py-2 rounded-lg text-sm font-semibold ${sub === k ? 'bg-brand-navy text-white' : 'bg-natural-100 text-natural-muted'}`}>{label}</button>
  );
  return (
    <div className="space-y-4">
      <div className="rounded-xl bg-amber-50 border border-amber-200 px-4 py-2 text-sm text-amber-800">
        🧪 <b>หน้าทดลอง</b> — คำนวณค่าจ้างรถร่วมจากราคาน้ำมัน OR (แยกจากระบบค่าเที่ยวเดิม 100% ยังไม่ใช้จริง)
      </div>
      <div className="flex gap-2 flex-wrap">
        {tabBtn('report', 'เทียบต้นทุน (ต่อทะเบียน)')}
        {tabBtn('calc', 'คำนวณค่าจ้าง 1 เที่ยว')}
        {tabBtn('history', 'ราคาน้ำมันย้อนหลัง (ทุกสาขา)')}
      </div>
      {sub === 'calc' ? <CalcSection /> : sub === 'history' ? <HistorySection /> : <ReportSection />}
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

// ===== เทียบต้นทุนต่อทะเบียน (ค่าเที่ยวจ่ายจริง vs ต้นทุนตามน้ำมัน+ระยะจริง DOH) =====
interface ReportRow {
  branchId: string; branch: string; plateNo: string;
  tripCount: number; tripAmount: number; fuelCostBand: number; diff: number;
  computed: number; missing: number;
}
function ReportSection() {
  const [cycles, setCycles] = useState<{ id: string; name: string; status: string }[]>([]);
  const [branches, setBranches] = useState<{ id: string; name: string }[]>([]);
  const [cycleId, setCycleId] = useState('');
  const [branchId, setBranchId] = useState(''); // '' = ทุกสาขา
  const [rows, setRows] = useState<ReportRow[]>([]);
  const [loading, setLoading] = useState(false);
  const [computing, setComputing] = useState(''); // key branchId|plate ที่กำลังคิด (คิดทีละคัน)
  const [bulk, setBulk] = useState<{ done: number; total: number } | null>(null); // progress "คำนวณทั้งหมด"
  const [msg, setMsg] = useState('');
  const reqRef = useRef(0); // กัน response เก่ามาทับตอนสลับรอบ/สาขาเร็วๆ (race)
  const cancelBulkRef = useRef(false); // ธงยกเลิก "คำนวณทั้งหมด" กลางคัน

  // โหลดรอบ + สาขา จาก /api/state (HQ เห็นทุกสาขา)
  useEffect(() => {
    (async () => {
      try {
        const r = await fetch(api('/api/state'));
        const s = await r.json();
        const cs = (s.cycles || []) as any[];
        setCycles(cs.map((c) => ({ id: c.id, name: c.name, status: c.status })));
        setBranches(((s.branches || []) as any[]).filter((b) => !b.isHQ).map((b) => ({ id: b.id, name: b.name })));
        if (cs.length) { const open = cs.find((c) => c.status === 'open'); setCycleId(open ? open.id : cs[cs.length - 1].id); }
      } catch { /* เงียบ */ }
    })();
  }, []);

  const loadReport = async (cid = cycleId, bid = branchId) => {
    if (!cid) return;
    const myReq = ++reqRef.current;
    setLoading(true); setMsg('');
    try {
      const q = `?cycleId=${encodeURIComponent(cid)}${bid ? `&branchId=${encodeURIComponent(bid)}` : ''}`;
      const r = await fetch(api(`/api/experimental/fuel-report${q}`));
      const j = await r.json();
      if (myReq !== reqRef.current) return; // มี request ใหม่กว่าแล้ว (สลับรอบ/สาขา) -> ทิ้งผลเก่า
      if (j.error) { setMsg('❌ ' + j.error); setRows([]); }
      else setRows(j.rows || []);
    } catch (e: any) { if (myReq === reqRef.current) setMsg('❌ ' + e.message); }
    finally { if (myReq === reqRef.current) setLoading(false); }
  };
  useEffect(() => { if (cycleId) loadReport(cycleId, branchId); }, [cycleId, branchId]); // eslint-disable-line

  // ยิง compute-plate 1 คัน (ไม่แตะ UI msg/refresh) — ใช้ร่วมทั้งกดทีละคันและคำนวณทั้งหมด
  const computeOne = async (row: ReportRow): Promise<{ ok: boolean; computed?: number; error?: string }> => {
    try {
      const r = await fetch(api('/api/experimental/fuel-report/compute-plate'), {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ cycleId, plateNo: row.plateNo, branchId: row.branchId }),
      });
      const j = await r.json();
      return j.success ? { ok: true, computed: j.computed } : { ok: false, error: j.error };
    } catch (e: any) { return { ok: false, error: e.message }; }
  };

  // กดคิดทีละคัน แล้วรีเฟรชรายงาน
  const computePlate = async (row: ReportRow) => {
    setComputing(`${row.branchId}|${row.plateNo}`); setMsg('');
    const res = await computeOne(row);
    await loadReport(); // รีเฟรชก่อน (loadReport ล้าง msg) แล้วค่อยตั้งข้อความ ผู้ใช้จะได้เห็น
    setMsg(res.ok ? `✅ ${row.plateNo}: คิดเสร็จ (${res.computed} ใบ)` : `❌ ${row.plateNo}: ${res.error}`);
    setComputing('');
  };

  // คำนวณทุกคันที่ยังไม่ครบ (ไล่ทีละคัน กัน DOH ถล่ม) — หยุดกลางคันได้ | filter ถูก disable ตอนนี้ (cid/bid คงที่)
  const computeAll = async () => {
    const cid = cycleId, bid = branchId; // ล็อกรอบ/สาขาที่กำลังคำนวณ (filter disable ระหว่างนี้)
    const todo = rows.filter((r) => !isComplete(r)); // ข้ามคันที่คิดครบแล้ว (cache)
    if (!todo.length) { setMsg('✅ ทุกคันคิดครบแล้ว'); return; }
    cancelBulkRef.current = false;
    setBulk({ done: 0, total: todo.length }); setMsg('');
    let ok = 0, fail = 0, stopped = false;
    for (let i = 0; i < todo.length; i++) {
      if (cancelBulkRef.current) { stopped = true; break; }
      const res = await computeOne(todo[i]);
      res.ok ? ok++ : fail++;
      setBulk({ done: i + 1, total: todo.length });
    }
    await loadReport(cid, bid); // รีเฟรชด้วยรอบ/สาขาเดิม แล้วค่อยตั้งข้อความ (loadReport ล้าง msg)
    setMsg(stopped
      ? `⏹ หยุดแล้ว — คิดสำเร็จ ${ok} คัน${fail ? ` · ล้มเหลว ${fail} คัน` : ''} (ยังเหลือ ${todo.length - ok - fail} คัน)`
      : `✅ คำนวณทั้งหมดเสร็จ — สำเร็จ ${ok} คัน${fail ? ` · ล้มเหลว ${fail} คัน` : ''}`);
    setBulk(null);
  };

  // รวมเฉพาะแถวที่คิดครบ (computed>0 && ไม่มี missing) — แถวคิดไม่ครบ diff เพี้ยน (ค่าเที่ยวเต็ม แต่ต้นทุนบางส่วน)
  const isComplete = (r: ReportRow) => r.computed > 0 && r.missing === 0;
  const completeRows = rows.filter(isComplete);
  const incompleteCount = rows.length - completeRows.length;
  const totals = completeRows.reduce((a, r) => ({ trip: a.trip + r.tripAmount, fuel: a.fuel + r.fuelCostBand, diff: a.diff + r.diff }), { trip: 0, fuel: 0, diff: 0 });

  return (
    <div className="space-y-4">
      <div className="rounded-xl bg-white border p-4 space-y-3">
        <h3 className="font-bold text-brand-navy">เทียบต้นทุนขนส่งต่อทะเบียน</h3>
        <p className="text-sm text-natural-muted">เทียบ <b>ค่าเที่ยวที่จ่ายจริง</b> กับ <b>ต้นทุนตามน้ำมัน + ระยะถนนจริง (DOH)</b> — กด "คำนวณ" ต่อคันเพื่อยิงระยะทาง (cache ไว้ ไม่ยิงซ้ำ)</p>
        <div className="flex items-end gap-2 flex-wrap">
          <label className="text-sm">รอบบิล
            <select value={cycleId} onChange={(e) => setCycleId(e.target.value)} disabled={!!bulk} className="mt-1 block border rounded-lg px-3 py-2 min-w-[220px] disabled:opacity-60">
              {cycles.map((c) => <option key={c.id} value={c.id}>{c.name}{c.status === 'closed' ? ' 🔒' : ''}</option>)}
            </select>
          </label>
          <label className="text-sm">สาขา
            <select value={branchId} onChange={(e) => setBranchId(e.target.value)} disabled={!!bulk} className="mt-1 block border rounded-lg px-3 py-2 min-w-[160px] disabled:opacity-60">
              <option value="">ทุกสาขา</option>
              {branches.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}
            </select>
          </label>
          <button onClick={() => loadReport()} disabled={loading || !!bulk} className="bg-natural-100 text-brand-navy rounded-lg px-4 py-2 text-sm font-semibold">รีเฟรช</button>
          {/* คำนวณทั้งหมด: ไล่ทีละคันที่ยังไม่ครบ (กัน DOH ถล่ม) — ระหว่างทำ เปลี่ยนเป็นปุ่มหยุด */}
          {!bulk ? (
            <button onClick={computeAll} disabled={loading || !!computing || incompleteCount === 0}
              className="bg-brand-red text-white rounded-lg px-4 py-2 text-sm font-semibold disabled:opacity-50">
              คำนวณทั้งหมด{incompleteCount > 0 ? ` (${incompleteCount} คัน)` : ''}
            </button>
          ) : (
            <button onClick={() => { cancelBulkRef.current = true; }} className="bg-natural-700 text-white rounded-lg px-4 py-2 text-sm font-semibold">
              ⏹ หยุด ({bulk.done}/{bulk.total})
            </button>
          )}
          {loading && <span className="text-sm text-natural-muted">กำลังโหลด...</span>}
        </div>
        {/* progress bar ตอนคำนวณทั้งหมด */}
        {bulk && (
          <div className="space-y-1">
            <div className="text-sm text-natural-muted">กำลังคำนวณ {bulk.done}/{bulk.total} คัน... (ยิงระยะจริงทีละคัน อาจใช้เวลาสักครู่)</div>
            <div className="h-2 rounded-full bg-natural-100 overflow-hidden">
              <div className="h-full bg-brand-red transition-all" style={{ width: `${bulk.total ? (bulk.done / bulk.total) * 100 : 0}%` }} />
            </div>
          </div>
        )}
        {msg && <div className="text-sm rounded-lg bg-natural-50 px-3 py-2">{msg}</div>}
      </div>

      <div className="rounded-xl bg-white border p-4 overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-natural-100 text-left">
              <th className="px-3 py-2 rounded-l-lg">ทะเบียน</th>
              {!branchId && <th className="px-3 py-2">สาขา</th>}
              <th className="px-3 py-2 text-right">เที่ยว</th>
              <th className="px-3 py-2 text-right">ค่าเที่ยว (จ่ายจริง)</th>
              <th className="px-3 py-2 text-right">ต้นทุนน้ำมัน (คิด)</th>
              <th className="px-3 py-2 text-right">ส่วนต่าง</th>
              <th className="px-3 py-2 rounded-r-lg text-right">คิดระยะ</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => {
              const key = `${r.branchId}|${r.plateNo}`;
              const done = isComplete(r);
              return (
                <tr key={key} className="border-b last:border-0">
                  <td className="px-3 py-2 font-semibold">{r.plateNo}</td>
                  {!branchId && <td className="px-3 py-2 text-natural-muted">{r.branch}</td>}
                  <td className="px-3 py-2 text-right">{r.tripCount}</td>
                  <td className="px-3 py-2 text-right font-bold text-brand-navy">฿{baht(r.tripAmount)}</td>
                  {/* แสดงต้นทุน/ส่วนต่าง เฉพาะแถวที่คิดครบ — บางส่วน diff เพี้ยน (ค่าเที่ยวเต็ม vs ต้นทุนบางใบ) */}
                  <td className="px-3 py-2 text-right">{done ? `฿${baht(r.fuelCostBand)}` : <span className="text-natural-muted">{r.missing > 0 && r.computed > 0 ? 'บางส่วน' : '—'}</span>}</td>
                  <td className={`px-3 py-2 text-right font-bold ${done ? (r.diff >= 0 ? 'text-emerald-700' : 'text-brand-red') : 'text-natural-muted'}`}>
                    {done ? `${r.diff >= 0 ? '+' : ''}฿${baht(r.diff)}` : '—'}
                  </td>
                  <td className="px-3 py-2 text-right">
                    <button onClick={() => computePlate(r)} disabled={computing === key || !!bulk}
                      className={`rounded-lg px-3 py-1 text-xs font-semibold disabled:opacity-50 ${done ? 'bg-emerald-50 text-emerald-700' : 'bg-brand-red text-white'}`}>
                      {computing === key ? 'กำลังคิด...' : done ? '✓ คิดแล้ว' : r.missing ? `คำนวณ (ขาด ${r.missing})` : 'คำนวณ'}
                    </button>
                  </td>
                </tr>
              );
            })}
            {!rows.length && <tr><td colSpan={branchId ? 6 : 7} className="px-3 py-6 text-center text-natural-muted">{loading ? 'กำลังโหลด...' : 'ไม่มีข้อมูลรถร่วมในรอบนี้ (หรือยังไม่เลือกรอบ)'}</td></tr>}
          </tbody>
          {rows.length > 0 && (
            <tfoot>
              <tr className="border-t-2 font-bold">
                <td className="px-3 py-2" colSpan={branchId ? 2 : 3}>
                  รวมเฉพาะคิดครบ ({completeRows.length}/{rows.length} ทะเบียน)
                  {incompleteCount > 0 && <span className="ml-1 text-xs font-normal text-brand-red">· ยังไม่ครบ {incompleteCount} คัน (ไม่รวมในยอด)</span>}
                </td>
                <td className="px-3 py-2 text-right text-brand-navy">฿{baht(totals.trip)}</td>
                <td className="px-3 py-2 text-right">฿{baht(totals.fuel)}</td>
                <td className={`px-3 py-2 text-right ${totals.diff >= 0 ? 'text-emerald-700' : 'text-brand-red'}`}>{totals.diff >= 0 ? '+' : ''}฿{baht(totals.diff)}</td>
                <td className="px-3 py-2"></td>
              </tr>
            </tfoot>
          )}
        </table>
        <p className="text-xs text-natural-muted mt-2">
          ส่วนต่าง = ค่าเที่ยวจ่ายจริง − ต้นทุนน้ำมันที่คิด · <span className="text-emerald-700">บวก</span> = จ่ายจริงแพงกว่าต้นทุน · <span className="text-brand-red">ลบ</span> = จ่ายจริงถูกกว่าต้นทุน ·
          ต้นทุนน้ำมันคิดจาก DOH ระยะถนนจริง (loop nearest-neighbor) × ราคาดีเซลล่าสุดของสาขา
        </p>
      </div>
    </div>
  );
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
