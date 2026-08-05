// ============================================================================
// [ทดลอง] แท็บ "คำนวณค่าจ้างรถร่วม (น้ำมัน)" — แยก 100% ไม่แตะ component เดิม
// 2 ส่วน: (1) คำนวณค่าจ้าง 1 เที่ยว  (2) สรุปราคาน้ำมันย้อนหลังแต่ละสาขา
// ============================================================================
import { useState, useEffect, useRef, Fragment } from 'react';

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
  const [sub, setSub] = useState<'calc' | 'history' | 'report' | 'mountain' | 'policy'>('report');
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
        {tabBtn('mountain', '⛰️ เส้นทางขึ้นเขา')}
        {tabBtn('policy', '⚙️ ตั้งค่าสูตร')}
      </div>
      {sub === 'calc' ? <CalcSection /> : sub === 'history' ? <HistorySection /> : sub === 'mountain' ? <MountainSection /> : sub === 'policy' ? <PolicySection /> : <ReportSection />}
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
  const [policy, setPolicy] = useState<{ driverHourlyRate: number; unloadingMinutesPerStore: number; fuelEfficiencyKmPerL: number } | null>(null);
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
    const j = await r.json();
    setResult(j.breakdown);
    if (j.policyUsed) setPolicy(j.policyUsed); // เก็บค่าแรง/นาทีต่อจุด ไว้แสดงวิธีคิด
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
          <div className="space-y-3 text-sm">
            {/* ===== วิธีคิดเบี้ยขับ (step-by-step) — เวลาใช้ "ระยะสูงสุด × 2" ===== */}
            <div className="rounded-lg border border-natural-200 p-3 space-y-1.5">
              <div className="font-semibold text-brand-navy">🧑‍✈️ เบี้ยขับ (คนขับ) = (เวลาเดินทาง + เวลาลงของ) × ค่าแรง/ชม.</div>
              <Step label="① ระยะไป-กลับ (คิดเวลา)" calc={`${result.distanceMaxKm} × 2`} val={`${(result.distanceMaxKm * 2).toFixed(0)} กม.`} />
              <Step label={`② ความเร็ว (${policy && result.distanceMaxKm <= policy.speedThresholdKm ? 'ทางใกล้' : result.speedKmh === (policy?.nearSpeedKmh ?? -1) ? 'ทางใกล้' : 'ทางไกล'})`} val={`${result.speedKmh} กม./ชม.`} />
              <Step label="③ เวลาเดินทาง" calc={`${(result.distanceMaxKm * 2).toFixed(0)} ÷ ${result.speedKmh}`} val={`${result.travelHours.toFixed(3)} ชม.`} />
              <Step label={`④ เวลาลงของ (${result.storeCount} จุด${policy ? ` × ${policy.unloadingMinutesPerStore} นาที` : ''})`} val={`${result.unloadingHours.toFixed(2)} ชม.`} />
              <Step label="⑤ รวมเวลาทำงาน" calc={`${result.travelHours.toFixed(3)} + ${result.unloadingHours.toFixed(2)}`} val={`${(result.travelHours + result.unloadingHours).toFixed(3)} ชม.`} />
              <div className="border-t pt-1.5 flex justify-between font-bold text-brand-navy">
                <span>= เบี้ยขับ{policy ? ` (${(result.travelHours + result.unloadingHours).toFixed(3)} × ฿${policy.driverHourlyRate})` : ''}</span>
                <span>฿{baht(result.driverAllowance)}</span>
              </div>
            </div>

            {/* ===== วิธีคิดค่าน้ำมัน — ใช้ "ค่ากลางของช่วง × 2" (ต่างจากเวลาที่ใช้ระยะสูงสุด) ===== */}
            <div className="rounded-lg border border-natural-200 p-3 space-y-1.5">
              <div className="font-semibold text-brand-navy">⛽ ค่าน้ำมัน = (ระยะ ÷ อัตราสิ้นเปลือง) × ราคา</div>
              <Step label="ระยะไป-กลับ (คิดน้ำมัน = ค่ากลางช่วง)" calc={`(${result.distanceMinKm}+${result.distanceMaxKm})÷2 × 2`} val={`${result.estimatedRoundTripKm.toFixed(1)} กม.`} />
              <div className="border-t pt-1.5 flex justify-between font-bold text-brand-navy">
                <span>= ค่าน้ำมัน ({result.estimatedRoundTripKm.toFixed(1)} ÷ {policy ? policy.fuelEfficiencyKmPerL : 10} × ฿{baht(result.appliedFuelPrice)})</span>
                <span>฿{baht(result.fuelCost)}</span>
              </div>
            </div>

            {/* ===== รวม ===== */}
            <table className="w-full">
              <tbody>
                {result.extraCharges > 0 && <Row k="ค่าพิเศษ (เขา/เส้นทาง/ภาษี)" v={`฿${baht(result.extraCharges)}`} />}
                <Row k="ค่าเที่ยว (เบี้ยขับ + น้ำมัน)" v={`฿${baht(result.tripRateBeforeRounding)}`} bold />
                <Row k="เรทต่อกล่อง" v={`฿${baht(result.boxRateDisplay)}`} />
                <tr><td colSpan={2} className="pt-2"><div className="rounded-lg bg-emerald-50 px-3 py-2 flex justify-between items-center">
                  <span className="font-bold text-emerald-800">จ่ายจริง ({result.paymentMode === 'LUMP_SUM' ? 'เหมาทั้งเที่ยว' : 'ตามกล่อง'})</span>
                  <span className="font-bold text-lg text-emerald-800">฿{baht(result.finalPayment)}</span>
                </div></td></tr>
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

function Row({ k, v, bold }: { k: string; v: string; bold?: boolean }) {
  return <tr className="border-b last:border-0"><td className="py-1.5 text-natural-muted">{k}</td><td className={`py-1.5 text-right ${bold ? 'font-bold text-brand-navy' : ''}`}>{v}</td></tr>;
}

// 1 ขั้นในวิธีคิด: ป้าย + (สูตรย่อย) + ค่าที่ได้
function Step({ label, calc, val }: { label: string; calc?: string; val: string }) {
  return (
    <div className="flex justify-between items-baseline gap-2">
      <span className="text-natural-muted">{label}{calc && <span className="text-xs text-natural-400"> = {calc}</span>}</span>
      <span className="tabular-nums">{val}</span>
    </div>
  );
}

// รายละเอียดต่อใบ (แถวขยายในรายงานเทียบต้นทุน) — แต่ละใบ: ปลายทาง ระยะลูป ต้นทุน ส่วนต่าง
function PlateDetail({ busy, items }: { busy: boolean; items: any[] | null }) {
  if (busy || !items) return <div className="text-sm text-natural-muted py-2">กำลังโหลดรายละเอียด...</div>;
  if (!items.length) return <div className="text-sm text-natural-muted py-2">ไม่มีใบ</div>;
  return (
    <div className="overflow-x-auto">
      <table className="w-full text-xs">
        <thead><tr className="text-left text-natural-muted border-b">
          <th className="px-2 py-1">ใบกระจาย</th><th className="px-2 py-1">ปลายทาง (ลำดับวิ่ง)</th>
          <th className="px-2 py-1 text-right">ระยะลูป</th><th className="px-2 py-1 text-right">เบี้ยขับ</th>
          <th className="px-2 py-1 text-right">ค่าน้ำมัน</th><th className="px-2 py-1 text-right">ต้นทุนรวม</th>
          <th className="px-2 py-1 text-right">ค่าเที่ยว</th><th className="px-2 py-1 text-right">ส่วนต่าง</th>
        </tr></thead>
        <tbody>
          {items.map((it, i) => {
            const hasMissing = (it.missing || []).length > 0;
            const dests = (it.order && it.order.length ? it.order : it.dests.map((d: string) => d.split('|')[0])).join(' → ');
            return (
              <tr key={i} className="border-b last:border-0">
                <td className="px-2 py-1 font-mono">{it.documentNo}</td>
                <td className="px-2 py-1">
                  {dests || '-'}
                  {hasMissing && <span className="text-brand-red"> · หาไม่เจอ: {it.missing.join(', ')}</span>}
                  {it.mountainLiters > 0 && <span className="text-amber-600"> · ⛰️+{it.mountainLiters}ล.</span>}
                </td>
                <td className="px-2 py-1 text-right">{it.loopKm != null ? `${it.loopKm} กม.` : '-'}</td>
                <td className="px-2 py-1 text-right">{it.driverAllowance != null ? `฿${baht(it.driverAllowance)}` : '-'}</td>
                <td className="px-2 py-1 text-right">{it.fuelCost != null ? `฿${baht(it.fuelCost)}` : '-'}</td>
                <td className="px-2 py-1 text-right font-semibold">{it.totalCost != null ? `฿${baht(it.totalCost)}` : <span className="text-brand-red">ยังไม่คิด</span>}</td>
                <td className="px-2 py-1 text-right">฿{baht(it.tripAmount)}</td>
                <td className={`px-2 py-1 text-right font-semibold ${it.diff == null ? 'text-natural-muted' : it.diff >= 0 ? 'text-emerald-700' : 'text-brand-red'}`}>
                  {it.diff != null ? `${it.diff >= 0 ? '+' : ''}฿${baht(it.diff)}` : '—'}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
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
  const [expanded, setExpanded] = useState<string | null>(null); // key ทะเบียนที่ขยายดูรายละเอียด
  const [detail, setDetail] = useState<{ key: string; items: any[] } | null>(null); // รายละเอียดใบของทะเบียนที่ขยาย
  const [detailBusy, setDetailBusy] = useState(false);
  const reqRef = useRef(0); // กัน response เก่ามาทับตอนสลับรอบ/สาขาเร็วๆ (race)
  const cancelBulkRef = useRef(false); // ธงยกเลิก "คำนวณทั้งหมด" กลางคัน

  // ย่อ/ขยาย ดูรายละเอียดต่อใบของ 1 ทะเบียน
  const toggleExpand = async (row: ReportRow) => {
    const key = `${row.branchId}|${row.plateNo}`;
    if (expanded === key) { setExpanded(null); return; } // ปิด
    setExpanded(key); setDetail(null); setDetailBusy(true);
    try {
      const q = `?cycleId=${encodeURIComponent(cycleId)}&plateNo=${encodeURIComponent(row.plateNo)}&branchId=${encodeURIComponent(row.branchId)}`;
      const r = await fetch(api(`/api/experimental/fuel-report/plate-detail${q}`));
      const j = await r.json();
      setDetail({ key, items: j.items || [] });
    } catch { setDetail({ key, items: [] }); } finally { setDetailBusy(false); }
  };

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
              const isOpen = expanded === key;
              return (
                <Fragment key={key}>
                <tr className="border-b last:border-0 hover:bg-natural-50">
                  <td className="px-3 py-2 font-semibold">
                    {/* คลิกทะเบียนเพื่อย่อ/ขยายดูรายละเอียดต่อใบ */}
                    <button onClick={() => toggleExpand(r)} className="flex items-center gap-1 hover:text-brand-red">
                      <span className={`transition-transform ${isOpen ? 'rotate-90' : ''}`}>▶</span>{r.plateNo}
                    </button>
                  </td>
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
                {isOpen && (
                  <tr className="bg-natural-50/60">
                    <td colSpan={branchId ? 6 : 7} className="px-3 py-2">
                      <PlateDetail busy={detailBusy} items={detail?.key === key ? detail.items : null} />
                    </td>
                  </tr>
                )}
                </Fragment>
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
          ต้นทุนน้ำมันคิดจาก DOH ระยะถนนจริง (loop nearest-neighbor) × ราคาดีเซลล่าสุดของสาขา · รวมน้ำมันขึ้นเขา (ถ้าปลายทางอยู่ในลิสต์ ⛰️)
        </p>
      </div>
    </div>
  );
}

// ===== ตั้งค่าสูตร (fuelPolicy — ชุดเดียวทั้งบริษัท) =====
type Policy = {
  nearSpeedKmh: number; farSpeedKmh: number; speedThresholdKm: number; fuelEfficiencyKmPerL: number;
  driverHourlyRate: number; unloadingMinutesPerStore: number; baseBoxes: number; lumpSumBoxThreshold: number; finalRoundingDecimals: number;
};
// ช่องที่ให้แก้ใน UI (baseBoxes/finalRoundingDecimals เชิงเทคนิค ไม่โชว์)
const POLICY_FIELDS: { key: keyof Policy; label: string; unit: string; step?: string }[] = [
  { key: 'nearSpeedKmh', label: 'ความเร็วขับรถใกล้', unit: 'กม./ชม.' },
  { key: 'farSpeedKmh', label: 'ความเร็วขับรถไกล', unit: 'กม./ชม.' },
  { key: 'speedThresholdKm', label: 'จุดตัดใกล้/ไกล (ระยะสูงสุด ≤ นี้ = ใช้ความเร็วใกล้)', unit: 'กม.' },
  { key: 'fuelEfficiencyKmPerL', label: 'อัตราสิ้นเปลืองน้ำมัน', unit: 'กม./ลิตร', step: '0.1' },
  { key: 'driverHourlyRate', label: 'ค่าแรงคนขับ', unit: 'บาท/ชม.', step: '0.01' },
  { key: 'unloadingMinutesPerStore', label: 'เวลาลงสินค้าต่อจุด', unit: 'นาที' },
  { key: 'lumpSumBoxThreshold', label: 'เส้นแบ่งเหมา (กล่อง ≤ นี้ = เหมาทั้งเที่ยว)', unit: 'กล่อง' },
];
function PolicySection() {
  const [policy, setPolicy] = useState<Policy | null>(null);
  const [defaults, setDefaults] = useState<Policy | null>(null);
  const [isDefault, setIsDefault] = useState(true);
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState('');

  const load = async () => {
    try {
      const r = await fetch(api('/api/experimental/fuel-policy'));
      const j = await r.json();
      setPolicy(j.policy); setDefaults(j.defaults); setIsDefault(j.isDefault);
    } catch (e: any) { setMsg('❌ ' + e.message); }
  };
  useEffect(() => { load(); }, []); // eslint-disable-line

  const save = async () => {
    if (!policy) return;
    setBusy(true); setMsg('');
    try {
      const r = await fetch(api('/api/experimental/fuel-policy'), {
        method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ policy }),
      });
      const j = await r.json();
      if (!j.success) setMsg('❌ ' + j.error);
      else { setMsg('✅ บันทึกแล้ว — รายงานคิดใหม่ตามค่าใหม่ (ระยะที่คิดไว้ใช้ได้)'); setIsDefault(false); }
    } catch (e: any) { setMsg('❌ ' + e.message); } finally { setBusy(false); }
  };
  const resetDefault = () => { if (defaults) { setPolicy({ ...defaults }); setMsg('ตั้งเป็นค่าเริ่มต้น (ยังไม่บันทึก — กด "บันทึก" เพื่อยืนยัน)'); } };

  if (!policy) return <div className="rounded-xl bg-white border p-4 text-sm text-natural-muted">กำลังโหลด...</div>;
  return (
    <div className="rounded-xl bg-white border p-4 space-y-4 max-w-2xl">
      <div>
        <h3 className="font-bold text-brand-navy">⚙️ ตั้งค่าสูตรคำนวณ (ใช้ร่วมทุกสาขา)</h3>
        <p className="text-sm text-natural-muted mt-1">
          ค่าเหล่านี้ใช้คิดเบี้ยขับ + ค่าน้ำมันทั้งหน้า "คำนวณ 1 เที่ยว" และ "เทียบต้นทุน" · แก้แล้วมีผลทันที (รายงานคิดใหม่ตามค่าใหม่)
          {isDefault && <span className="ml-1 text-amber-600">· ตอนนี้ใช้ค่าเริ่มต้น (ยังไม่เคยตั้งเอง)</span>}
        </p>
      </div>
      <div className="grid md:grid-cols-2 gap-3 text-sm">
        {POLICY_FIELDS.map((f) => (
          <label key={f.key} className="block">
            {f.label}
            <div className="flex items-center gap-2 mt-1">
              <input type="number" step={f.step || '1'} value={policy[f.key]}
                onChange={(e) => setPolicy({ ...policy, [f.key]: e.target.value === '' ? 0 : Number(e.target.value) })}
                className="w-full border rounded-lg px-2 py-1.5" />
              <span className="text-natural-muted whitespace-nowrap text-xs">{f.unit}</span>
            </div>
          </label>
        ))}
      </div>
      {msg && <div className="text-sm rounded-lg bg-natural-50 px-3 py-2">{msg}</div>}
      <div className="flex gap-2">
        <button onClick={save} disabled={busy} className="bg-brand-red text-white rounded-lg px-4 py-2 text-sm font-semibold disabled:opacity-50">💾 บันทึก</button>
        <button onClick={resetDefault} disabled={busy} className="bg-natural-100 text-brand-navy rounded-lg px-4 py-2 text-sm font-semibold">คืนค่าเริ่มต้น</button>
      </div>
      <p className="text-xs text-natural-muted">
        หมายเหตุ: ค่าแรง 150 บาท/ชม. = 1,200 บาท ÷ 8 ชม. · เบี้ยขับ = (เวลาเดินทาง + เวลาลงของ) × ค่าแรง/ชม. · ค่าน้ำมัน = (ระยะ ÷ อัตราสิ้นเปลือง) × ราคาดีเซล
      </p>
    </div>
  );
}

// ===== master น้ำมันขึ้นเขา (ลิตรเพิ่มต่อปลายทาง) =====
interface MountainRoute { id: string; province: string; district?: string; extraLiters: number; note?: string; createdAt: string; }
function MountainSection() {
  const [routes, setRoutes] = useState<MountainRoute[]>([]);
  const [destOpts, setDestOpts] = useState<{ province: string; districts: string[] }[]>([]); // จังหวัด->อำเภอ ที่มีจริงในระบบ
  const [form, setForm] = useState({ province: '', district: '', extraLiters: '', note: '' });
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState('');

  const load = async () => {
    try { const r = await fetch(api('/api/experimental/mountain-routes')); setRoutes((await r.json()).routes || []); }
    catch (e: any) { setMsg('❌ ' + e.message); }
  };
  // โหลดรายการปลายทางที่มีจริง (dropdown) — กันพิมพ์ผิด/ต้อง match เป๊ะ
  useEffect(() => {
    load();
    fetch(api('/api/experimental/dest-options')).then((r) => r.json()).then((j) => setDestOpts(j.provinces || [])).catch(() => {});
  }, []); // eslint-disable-line
  const districtsOf = (prov: string) => destOpts.find((p) => p.province === prov)?.districts || [];

  const add = async () => {
    setBusy(true); setMsg('');
    try {
      const r = await fetch(api('/api/experimental/mountain-routes'), {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ province: form.province, district: form.district, extraLiters: Number(form.extraLiters), note: form.note }),
      });
      const j = await r.json();
      if (!j.success) setMsg('❌ ' + j.error);
      else { setMsg('✅ เพิ่มแล้ว'); setForm({ province: '', district: '', extraLiters: '', note: '' }); await load(); }
    } catch (e: any) { setMsg('❌ ' + e.message); } finally { setBusy(false); }
  };

  const del = async (id: string, label: string) => {
    if (!window.confirm(`ลบเส้นทางขึ้นเขา "${label}"?`)) return;
    setBusy(true); setMsg('');
    try {
      const r = await fetch(api(`/api/experimental/mountain-routes/${id}`), { method: 'DELETE' });
      const j = await r.json();
      if (!j.success) setMsg('❌ ' + j.error); else { setMsg('✅ ลบแล้ว'); await load(); }
    } catch (e: any) { setMsg('❌ ' + e.message); } finally { setBusy(false); }
  };

  return (
    <div className="space-y-4">
      <div className="rounded-xl bg-white border p-4 space-y-3">
        <h3 className="font-bold text-brand-navy">⛰️ เส้นทางขึ้นเขา (น้ำมันเพิ่ม)</h3>
        <p className="text-sm text-natural-muted">
          ปลายทางที่ต้อง <b>บวกน้ำมันเพิ่ม</b> (ขึ้นดอย/ภูเขา) — ระบุเป็น <b>ลิตรเพิ่มต่อเที่ยว</b> จะถูก <b>คูณราคาดีเซลของสาขา</b> แล้วบวกเข้าต้นทุนน้ำมัน ·
          อำเภอว่าง = ทั้งจังหวัด · ใบที่ผ่านปลายทางในลิสต์นี้จะถูกคิดเพิ่มอัตโนมัติ (คิดใหม่หลังแก้)
        </p>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-2 text-sm">
          <label>จังหวัด*
            <select value={form.province} onChange={(e) => setForm({ ...form, province: e.target.value, district: '' })} className="mt-1 w-full border rounded-lg px-2 py-1.5 bg-white">
              <option value="">— เลือกจังหวัด —</option>
              {destOpts.map((p) => <option key={p.province} value={p.province}>{p.province}</option>)}
            </select>
          </label>
          <label>อำเภอ (ว่าง=ทั้งจังหวัด)
            <select value={form.district} onChange={(e) => setForm({ ...form, district: e.target.value })} disabled={!form.province} className="mt-1 w-full border rounded-lg px-2 py-1.5 bg-white disabled:opacity-60">
              <option value="">ทั้งจังหวัด</option>
              {districtsOf(form.province).map((d) => <option key={d} value={d}>{d}</option>)}
            </select>
          </label>
          <label>ลิตรเพิ่ม*<input type="number" step="0.1" value={form.extraLiters} onChange={(e) => setForm({ ...form, extraLiters: e.target.value })} placeholder="เช่น 5" className="mt-1 w-full border rounded-lg px-2 py-1.5" /></label>
          <label>หมายเหตุ<input value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} placeholder="เช่น ขึ้นดอย" className="mt-1 w-full border rounded-lg px-2 py-1.5" /></label>
        </div>
        <button onClick={add} disabled={busy || !form.province || !form.extraLiters} className="bg-brand-red text-white rounded-lg px-4 py-2 text-sm font-semibold disabled:opacity-50">+ เพิ่มเส้นทาง</button>
        {msg && <div className="text-sm rounded-lg bg-natural-50 px-3 py-2">{msg}</div>}
      </div>

      <div className="rounded-xl bg-white border p-4 overflow-x-auto">
        <table className="w-full text-sm">
          <thead><tr className="bg-natural-100 text-left">
            <th className="px-3 py-2 rounded-l-lg">จังหวัด</th><th className="px-3 py-2">อำเภอ</th>
            <th className="px-3 py-2 text-right">ลิตรเพิ่ม</th><th className="px-3 py-2">หมายเหตุ</th>
            <th className="px-3 py-2 rounded-r-lg text-right">ลบ</th>
          </tr></thead>
          <tbody>
            {routes.map((r) => (
              <tr key={r.id} className="border-b last:border-0">
                <td className="px-3 py-2 font-semibold">{r.province}</td>
                <td className="px-3 py-2">{r.district || <span className="text-natural-muted">ทั้งจังหวัด</span>}</td>
                <td className="px-3 py-2 text-right font-bold text-brand-navy">{r.extraLiters} ลิตร</td>
                <td className="px-3 py-2 text-natural-muted">{r.note || '-'}</td>
                <td className="px-3 py-2 text-right">
                  <button onClick={() => del(r.id, `${r.province}${r.district ? ' / ' + r.district : ''}`)} disabled={busy} className="text-brand-red text-xs font-semibold hover:underline">ลบ</button>
                </td>
              </tr>
            ))}
            {!routes.length && <tr><td colSpan={5} className="px-3 py-6 text-center text-natural-muted">ยังไม่มีเส้นทางขึ้นเขา — เพิ่มด้านบน</td></tr>}
          </tbody>
        </table>
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
