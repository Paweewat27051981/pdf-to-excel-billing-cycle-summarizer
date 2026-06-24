import React, { useState, useEffect, useRef } from 'react';
import {
  UploadCloud, AlertTriangle, FileSpreadsheet, Trash2, Plus, Save,
  RefreshCw, Lock, Unlock, Database, Truck, Tag, Filter, Calculator, Fuel, Receipt, Coins,
  Building2, LogOut,
} from 'lucide-react';
import {
  DatabaseState, BillingCycle, Branch, Vehicle, RateMaster, RateOverride, ReceiverGroup, ReceiverGroupAlias,
  ProductConversionRule, TripDocument, FuelEntry, DeductionEntry, ExtractedTripDocument, MoneyCategory, ManualBoxSender,
} from './types';
import { exportCycleToExcel, exportPerVehicleReport } from './excel-export';
import { summarizeByVehicle, isUnspecifiedName, normPlate } from './calc';
import { confirmDelete, confirmAction, confirmPassword, notify } from './ui';

// โมเดลที่มีจริง (ตรวจจาก ListModels API) — flash=เร็ว/ฟรีกว่า, pro=แม่นกว่า
const GEMINI_MODELS = [
  'gemini-3.5-flash',
  'gemini-2.5-flash',
  'gemini-flash-latest',
  'gemini-2.5-pro',
  'gemini-3-pro-preview',
  'gemini-3.1-pro-preview',
  'gemini-pro-latest',
];

const EMPTY: DatabaseState = {
  settings: { geminiModel: 'gemini-3.5-flash' },
  branches: [],
  cycles: [], vehicles: [], rateMasters: [], rateOverrides: [], rateMasterHistory: [], receiverGroups: [],
  receiverGroupAliases: [], conversionRules: [], manualBoxSenders: [], moneyCategories: [], tripDocuments: [], fuelEntries: [], deductions: [],
};

type BranchAuth = { id: string; name: string; isHQ: boolean };
type Tab = 'calc' | 'rates' | 'rules' | 'vehicles' | 'fuel' | 'dashboard' | 'branches' | 'reports';
type Toast = { type: 'success' | 'error' | 'warning'; message: string };

const THAI_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const money = (n: number) => n.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export default function App() {
  const [db, setDb] = useState<DatabaseState>(EMPTY);
  const [loading, setLoading] = useState(true);
  const [selectedCycleId, setSelectedCycleId] = useState('');
  const [tab, setTab] = useState<Tab>('calc');
  const [aiEnabled, setAiEnabled] = useState(true);
  const [auth, setAuth] = useState<BranchAuth | null>(() => {
    try { return JSON.parse(localStorage.getItem('branchAuth') || 'null'); } catch { return null; }
  });
  // สาขาที่กำลังทำงาน: ผู้ใช้สาขา = สาขาตัวเอง; HQ = เลือกได้
  const [workBranchId, setWorkBranchId] = useState('');

  const showToast = (type: Toast['type'], message: string) => notify(type, message);

  const api = async (url: string, method: string, body?: any) => {
    const res = await fetch(url, {
      method,
      headers: { 'Content-Type': 'application/json' },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || 'เกิดข้อผิดพลาด');
    return res.json();
  };

  // branchId ที่ใช้กรอง/บันทึก
  const effBranchId = auth?.isHQ ? workBranchId : (auth?.id || '');

  const fetchState = async (autoCycle?: string) => {
    setLoading(true);
    try {
      const url = effBranchId ? `/api/state?branchId=${encodeURIComponent(effBranchId)}` : '/api/state';
      const res = await fetch(url);
      const data: DatabaseState = await res.json();
      setDb(data);
      if (autoCycle) setSelectedCycleId(autoCycle);
      else if (!selectedCycleId && data.cycles.length) {
        const open = data.cycles.find((c) => c.status === 'open');
        setSelectedCycleId(open ? open.id : data.cycles[data.cycles.length - 1].id);
      }
    } catch (e: any) {
      showToast('error', `โหลดข้อมูลไม่ได้: ${e.message}`);
    } finally {
      setLoading(false);
    }
  };

  // โหลด config + branches เสมอ (สำหรับหน้า login)
  useEffect(() => {
    fetch('/api/config').then((r) => r.json()).then((c) => setAiEnabled(!!c.aiEnabled)).catch(() => setAiEnabled(false));
  }, []);

  // โหลดข้อมูลตามสาขาที่เลือก (HQ workBranchId='' = ทุกสาขา)
  useEffect(() => {
    if (!auth) { fetch('/api/state').then((r) => r.json()).then(setDb).finally(() => setLoading(false)); return; }
    fetchState();
  }, [auth, workBranchId]);

  const doLogin = async (a: BranchAuth) => {
    localStorage.setItem('branchAuth', JSON.stringify(a));
    setAuth(a);
    setWorkBranchId(a.isHQ ? '' : a.id);
    if (a.isHQ) setTab('dashboard'); // HQ เริ่มที่ภาพรวมทุกสาขา
    setLoading(true);
  };
  const logout = () => { localStorage.removeItem('branchAuth'); setAuth(null); setWorkBranchId(''); setDb(EMPTY); };

  // ยังไม่ login -> หน้าเลือกสาขา + รหัสผ่าน
  if (!auth) return <BranchLogin branches={db.branches} api={api} onLogin={doLogin} />;

  const cycle = db.cycles.find((c) => c.id === selectedCycleId) || null;
  const cycleTrips = db.tripDocuments.filter((t) => t.cycleId === selectedCycleId);
  const activeBranchName = db.branches.find((b) => b.id === effBranchId)?.name || (auth.isHQ ? '— เลือกสาขา —' : auth.name);

  const tabs: [Tab, string, any][] = [
    ['calc', 'คำนวณค่าเที่ยว', Calculator],
    ['fuel', 'ค่าน้ำมัน & รายการหัก', Fuel],
    ['dashboard', 'Dashboard', Database],
    ['reports', 'รายงานต่อทะเบียน', FileSpreadsheet],
    ['rates', 'Master ราคาขนส่ง', Tag],
    ['rules', 'เงื่อนไขตัวหาร', Filter],
    ['vehicles', 'รถ & คนขับ', Truck],
  ];
  if (auth.isHQ) tabs.push(['branches', 'จัดการสาขา', Building2]);

  return (
    <div className="min-h-screen bg-natural-bg text-natural-text font-sans flex flex-col">
      {/* Header */}
      <header className="bg-white border-b border-natural-border px-6 py-3 flex flex-col md:flex-row md:items-center justify-between gap-3 sticky top-0 z-30">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-[#1B365D] text-white rounded-xl flex items-center justify-center"><Truck className="w-5 h-5" /></div>
          <div>
            <h1 className="text-lg font-bold text-[#1B365D]">ระบบค่าเที่ยว + ค่าน้ำมันรถร่วม</h1>
            <p className="text-xs text-natural-muted">PDF/Excel ใบกระจาย → Review → คำนวณรอบ → Export</p>
          </div>
        </div>
        <div className="flex items-center gap-3 flex-wrap">
          {/* สาขา */}
          <div className="flex items-center gap-1.5 bg-emerald-50 border border-emerald-200 rounded-full px-3 py-1.5">
            <Building2 className="w-4 h-4 text-emerald-700" />
            {auth.isHQ ? (
              <select aria-label="เลือกสาขา" value={workBranchId} onChange={(e) => setWorkBranchId(e.target.value)}
                className="bg-transparent text-xs font-bold text-emerald-800 outline-none">
                <option value="">🌐 ทุกสาขา (ภาพรวม)</option>
                {db.branches.filter((b) => !b.isHQ && b.status === 'active').map((b) => (
                  <option key={b.id} value={b.id}>{b.name}</option>
                ))}
              </select>
            ) : (
              <span className="text-xs font-bold text-emerald-800">{activeBranchName}</span>
            )}
            {auth.isHQ && <span className="text-[10px] bg-emerald-600 text-white rounded-full px-1.5 py-0.5">HQ</span>}
          </div>
          <CycleBar cycles={db.cycles} selectedCycleId={selectedCycleId} setSelectedCycleId={setSelectedCycleId}
            onCreated={(id: string) => fetchState(id)} api={api} showToast={showToast} />
          <button onClick={logout} title="ออกจากระบบ"
            className="text-natural-muted hover:text-red-600 flex items-center gap-1 text-xs font-semibold"><LogOut className="w-4 h-4" /></button>
        </div>
      </header>

      {/* Tabs */}
      <nav className="bg-white border-b border-natural-border px-6 flex gap-1 overflow-x-auto">
        {tabs.map(([key, label, Icon]) => (
          <button key={key} onClick={() => setTab(key)}
            className={`px-4 py-3 text-sm font-semibold border-b-2 whitespace-nowrap flex items-center gap-1.5 transition-colors ${
              tab === key ? 'border-[#1B365D] text-[#1B365D]' : 'border-transparent text-natural-muted hover:text-natural-text'}`}>
            <Icon className="w-4 h-4" /> {label}
          </button>
        ))}
      </nav>

      <main className="flex-1 max-w-7xl mx-auto w-full p-6">
        {loading ? (
          <div className="flex flex-col items-center justify-center py-20">
            <RefreshCw className="w-10 h-10 text-[#1B365D] animate-spin mb-3" />
            <p className="text-sm text-natural-dark-muted">กำลังโหลด...</p>
          </div>
        ) : (
          <>
            {tab === 'calc' && <CalcTab db={db} cycle={cycle} cycleTrips={cycleTrips} api={api} aiEnabled={aiEnabled} branchId={effBranchId}
              reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'fuel' && <FuelDeductionTab db={db} cycle={cycle} api={api} branchId={effBranchId}
              reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'dashboard' && <DashboardTab db={db} cycle={cycle} branchId={effBranchId} isHQ={auth.isHQ} />}
            {tab === 'reports' && <ReportsTab db={db} cycle={cycle} branchId={effBranchId} showToast={showToast} />}
            {tab === 'rates' && <RatesTab db={db} api={api} branchId={effBranchId} cycle={cycle} reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'rules' && <RulesTab db={db} api={api} branchId={effBranchId} reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'vehicles' && <VehiclesTab db={db} api={api} branchId={effBranchId} reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'branches' && <BranchesTab db={db} api={api} reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
          </>
        )}
      </main>
    </div>
  );
}

// ===========================================================================
// หน้า Login: เลือกสาขา + ใส่รหัสผ่าน
// ===========================================================================
function BranchLogin({ branches, api, onLogin }: any) {
  const active = (branches as Branch[]).filter((b) => b.status === 'active');
  const [branchId, setBranchId] = useState('');
  const [password, setPassword] = useState('');
  const [busy, setBusy] = useState(false);

  useEffect(() => { if (!branchId && active.length) setBranchId(active[0].id); }, [branches]);

  const submit = async () => {
    if (!branchId) return;
    setBusy(true);
    try {
      const r = await api('/api/branch-login', 'POST', { branchId, password });
      onLogin(r.branch);
    } catch (e: any) { notify('error', e.message); }
    finally { setBusy(false); }
  };

  return (
    <div className="min-h-screen bg-natural-bg flex items-center justify-center p-6">
      <div className="bg-white rounded-2xl border border-natural-border shadow-lg p-8 w-full max-w-sm flex flex-col items-center">
        <div className="w-14 h-14 bg-[#1B365D] text-white rounded-2xl flex items-center justify-center mb-3"><Building2 className="w-7 h-7" /></div>
        <h1 className="text-lg font-bold text-[#1B365D]">ระบบค่าเที่ยว + ค่าน้ำมัน</h1>
        <p className="text-xs text-natural-muted mb-5">เลือกสาขา และใส่รหัสผ่านเพื่อเข้าใช้งาน</p>
        <label className="w-full text-xs font-semibold text-natural-dark-muted mb-1">สาขา</label>
        <select aria-label="สาขา" value={branchId} onChange={(e) => setBranchId(e.target.value)}
          className="w-full border border-natural-border rounded-lg px-3 py-2 text-sm mb-3">
          {active.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}
        </select>
        <label className="w-full text-xs font-semibold text-natural-dark-muted mb-1">รหัสผ่าน</label>
        <input type="password" value={password} onChange={(e) => setPassword(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && submit()} aria-label="รหัสผ่าน"
          className="w-full border border-natural-border rounded-lg px-3 py-2 text-sm mb-5" placeholder="••••" />
        <button onClick={submit} disabled={busy || !branchId}
          className="w-full bg-[#1B365D] disabled:bg-natural-muted text-white rounded-lg py-2.5 text-sm font-semibold flex items-center justify-center gap-2">
          {busy ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Unlock className="w-4 h-4" />} เข้าใช้งาน
        </button>
      </div>
    </div>
  );
}

// ===========================================================================
// Tab: จัดการสาขา (เฉพาะ HQ)
// ===========================================================================
function BranchesTab({ db, api, reload, showToast }: any) {
  const realBranches = (db.branches as Branch[]).filter((b) => !b.isHQ);
  const defaultSource = realBranches[0]?.id || '';
  const blank = { name: '', password: '1234', cloneFrom: defaultSource };
  const [form, setForm] = useState(blank);

  const countMaster = (bid: string) =>
    db.conversionRules.filter((r: ProductConversionRule) => r.branchId === bid).length;

  const add = async () => {
    if (!form.name.trim()) return showToast('warning', 'กรอกชื่อสาขา');
    try {
      const b = await api('/api/branches', 'POST', { name: form.name.trim(), password: form.password || '1234', status: 'active' });
      if (form.cloneFrom) {
        await api('/api/branches/clone', 'POST', { sourceBranchId: form.cloneFrom, targetBranchId: b.id, replace: true });
      }
      showToast('success', `เพิ่มสาขา "${form.name}" แล้ว${form.cloneFrom ? ' (คัดลอกกฎ/กลุ่ม/ประเภทมาให้)' : ''}`);
      setForm(blank);
      reload();
    } catch (e: any) { showToast('error', e.message); }
  };

  const cloneInto = async (target: Branch) => {
    const sources = realBranches.filter((b) => b.id !== target.id);
    if (!sources.length) return showToast('warning', 'ไม่มีสาขาต้นแบบให้คัดลอก');
    // เลือกสาขาต้นแบบที่มีกฎมากสุด (ปกติ = นครสวรรค์)
    const src = sources.slice().sort((a, b) => countMaster(b.id) - countMaster(a.id))[0];
    const ok = await confirmAction({
      title: `คัดลอกกฎจาก "${src.name}" → "${target.name}"?`,
      text: `จะแทนที่ กฎตัวหาร/กลุ่มผู้รับ/ประเภทรายได้-หัก/ผู้ส่งกล่อง ของ "${target.name}" ด้วยของ "${src.name}" (ไม่กระทบราคา/รถ)`,
      confirmText: 'คัดลอก',
    });
    if (!ok) return;
    await api('/api/branches/clone', 'POST', { sourceBranchId: src.id, targetBranchId: target.id, replace: true });
    showToast('success', `คัดลอกจาก "${src.name}" แล้ว`);
    reload();
  };
  const setPwd = async (b: Branch) => {
    const np = prompt(`ตั้งรหัสผ่านใหม่ของสาขา "${b.name}"`, '');
    if (np == null || np === '') return;
    await api(`/api/branches/${b.id}`, 'PUT', { password: np });
    showToast('success', 'เปลี่ยนรหัสผ่านแล้ว');
    reload();
  };
  const del = async (b: Branch) => {
    if (b.isHQ) return showToast('warning', 'ลบสำนักงานใหญ่ไม่ได้');
    if (!(await confirmDelete(`สาขา "${b.name}"`))) return;
    await api(`/api/branches/${b.id}`, 'DELETE');
    reload();
  };

  return (
    <div className="flex flex-col gap-4">
      <div className="bg-white rounded-2xl border border-natural-border p-4">
        <h3 className="font-bold text-[#1B365D] mb-3 flex items-center gap-2"><Building2 className="w-4 h-4" />เพิ่มสาขาใหม่</h3>
        <div className="flex flex-wrap gap-2 items-end">
          <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="ชื่อสาขา เช่น ตาก"
            className="border border-natural-border rounded-lg px-3 py-2 text-sm flex-1 min-w-[140px]" />
          <input value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} placeholder="รหัสผ่าน"
            className="border border-natural-border rounded-lg px-3 py-2 text-sm w-28" />
          <label className="flex items-center gap-1 text-[11px] text-natural-muted">
            คัดลอกกฎจาก:
            <select aria-label="สาขาต้นแบบ" value={form.cloneFrom} onChange={(e) => setForm({ ...form, cloneFrom: e.target.value })}
              className="border border-natural-border rounded-lg px-2 py-2 text-xs">
              <option value="">— ไม่คัดลอก (เริ่มว่าง) —</option>
              {realBranches.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}
            </select>
          </label>
          <button onClick={add} className="bg-[#1B365D] text-white rounded-lg px-4 py-2 text-sm font-semibold flex items-center gap-1"><Plus className="w-4 h-4" />เพิ่ม</button>
        </div>
        <p className="text-[11px] text-natural-muted mt-2">💡 คัดลอกกฎตัวหาร/กลุ่มผู้รับ/ประเภทรายได้-หัก จากสาขาต้นแบบมาเป็นจุดเริ่ม (ราคา/รถ ตั้งใหม่แยกต่อสาขา)</p>
      </div>

      <div className="bg-white rounded-2xl border border-natural-border overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-natural-bg text-natural-dark-muted text-xs">
            <tr><th className="text-left px-4 py-2">สาขา</th><th className="text-left px-4 py-2">รถ</th><th className="text-left px-4 py-2">ราคา</th><th className="text-left px-4 py-2">กฎตัวหาร</th><th className="px-4 py-2"></th></tr>
          </thead>
          <tbody>
            {(db.branches as Branch[]).map((b) => (
              <tr key={b.id} className="border-t border-natural-border">
                <td className="px-4 py-2 font-semibold text-[#1B365D]">{b.name} {b.isHQ && <span className="text-[10px] bg-emerald-600 text-white rounded-full px-1.5 py-0.5 ml-1">HQ</span>}</td>
                <td className="px-4 py-2 text-natural-muted">{db.vehicles.filter((v: Vehicle) => v.branchId === b.id).length}</td>
                <td className="px-4 py-2 text-natural-muted">{db.rateMasters.filter((r: RateMaster) => r.branchId === b.id).length}</td>
                <td className="px-4 py-2 text-natural-muted">{countMaster(b.id)}</td>
                <td className="px-4 py-2 text-right whitespace-nowrap">
                  {!b.isHQ && <button type="button" onClick={() => cloneInto(b)} className="text-xs text-emerald-700 font-semibold mr-3">⬇ คัดลอกกฎ</button>}
                  <button onClick={() => setPwd(b)} className="text-xs text-[#1B365D] font-semibold mr-3">ตั้งรหัสผ่าน</button>
                  {!b.isHQ && <button type="button" title="ลบสาขา" onClick={() => del(b)} className="text-red-500"><Trash2 className="w-4 h-4 inline" /></button>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ===========================================================================
// Cycle bar (เลือกเดือน/รอบ + เปิดรอบใหม่)
// ===========================================================================
function CycleBar({ cycles, selectedCycleId, setSelectedCycleId, onCreated, api, showToast }: any) {
  const [open, setOpen] = useState(false);
  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth() + 1);
  const [half, setHalf] = useState<'first' | 'second'>('first');

  const create = async () => {
    try {
      const c = await api('/api/cycles', 'POST', { year, month, half });
      showToast('success', `เปิดรอบ "${c.name}" สำเร็จ`);
      setOpen(false);
      onCreated(c.id);
    } catch (e: any) { showToast('error', e.message); }
  };

  const cur = cycles.find((c: BillingCycle) => c.id === selectedCycleId);
  const toggleLock = async () => {
    const closing = cur.status === 'open';
    const ok = await confirmAction({
      title: closing ? `ปิดรอบ "${cur.name}"?` : `เปิดรอบ "${cur.name}" อีกครั้ง?`,
      text: closing ? 'เมื่อปิดรอบแล้วจะเพิ่ม/แก้ไขข้อมูลในรอบนี้ไม่ได้จนกว่าจะเปิดใหม่' : 'เปิดเพื่อให้บันทึก/แก้ไขข้อมูลในรอบนี้ได้',
      confirmText: closing ? 'ปิดรอบ' : 'เปิดรอบ',
      danger: closing,
    });
    if (!ok) return;
    try {
      await api(`/api/cycles/${cur.id}`, 'PUT', { status: closing ? 'closed' : 'open' });
      onCreated(cur.id);
      showToast('success', closing ? 'ปิดรอบแล้ว' : 'เปิดรอบแล้ว');
    } catch (e: any) { showToast('error', e.message); }
  };

  return (
    <div className="flex items-center gap-2 flex-wrap relative">
      <select aria-label="เลือกรอบคำนวณ" value={selectedCycleId} onChange={(e) => setSelectedCycleId(e.target.value)}
        className="bg-natural-secondary border border-natural-border text-sm font-semibold rounded-full px-4 py-2">
        <option value="" disabled>— เลือกรอบ —</option>
        {cycles.map((c: BillingCycle) => <option key={c.id} value={c.id}>{c.name} {c.status === 'closed' ? '🔒' : ''}</option>)}
      </select>
      {cur && (
        <button onClick={toggleLock} className="border border-natural-border rounded-full px-3 py-2 text-xs font-semibold flex items-center gap-1">
          {cur.status === 'open' ? <><Lock className="w-3.5 h-3.5" />ปิดรอบ</> : <><Unlock className="w-3.5 h-3.5" />เปิดรอบ</>}
        </button>
      )}
      <button onClick={() => setOpen(!open)} className="bg-[#1B365D] text-white rounded-full px-4 py-2 text-xs font-semibold flex items-center gap-1">
        <Plus className="w-4 h-4" />เปิดรอบใหม่
      </button>
      {open && (
        <div className="absolute top-12 right-0 bg-white border border-natural-border rounded-2xl shadow-lg p-4 z-40 flex flex-col gap-3 w-72">
          <h4 className="font-bold text-sm text-[#1B365D]">เปิดรอบคำนวณใหม่</h4>
          <div className="flex gap-2">
            <select aria-label="เลือกเดือน" value={month} onChange={(e) => setMonth(+e.target.value)} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm flex-1">
              {THAI_MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
            </select>
            <input type="number" aria-label="ปี ค.ศ." value={year} onChange={(e) => setYear(+e.target.value)} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-24" />
          </div>
          <div className="flex gap-2">
            {(['first', 'second'] as const).map((h) => (
              <button key={h} onClick={() => setHalf(h)} className={`flex-1 py-1.5 rounded-lg text-sm font-semibold border ${half === h ? 'bg-[#1B365D] text-white border-[#1B365D]' : 'border-natural-border'}`}>
                {h === 'first' ? 'รอบ 1-15' : 'รอบ 16-31'}
              </button>
            ))}
          </div>
          <button onClick={create} className="bg-[#1B365D] text-white rounded-lg py-2 text-sm font-semibold">สร้างรอบ</button>
        </div>
      )}
    </div>
  );
}

// ===========================================================================
// Tab: คำนวณค่าเที่ยว (upload + review + list)
// ===========================================================================
function CalcTab({ db, cycle, cycleTrips, api, aiEnabled, branchId, reload, showToast }: any) {
  const [extracting, setExtracting] = useState(false);
  const [pending, setPending] = useState<{ extracted: ExtractedTripDocument; fileName: string; preview: TripDocument } | null>(null);
  const [filter, setFilter] = useState<'all' | 'divider' | 'warning'>('all');
  const [importing, setImporting] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);
  const excelRef = useRef<HTMLInputElement>(null);

  if (!cycle) return <EmptyHint text="กรุณาเลือกหรือเปิดรอบคำนวณก่อน" />;
  if (!branchId) return <EmptyHint text="กรุณาเลือกสาขาก่อน (มุมเมนูบน) เพื่อเริ่มทำงาน" />;

  const preview = async (extracted: ExtractedTripDocument, fileName: string) => {
    const p: TripDocument = await api('/api/trips/preview', 'POST', { cycleId: cycle.id, extracted, fileName, branchId });
    setPending({ extracted, fileName, preview: p });
  };

  const onFiles = async (files: FileList) => {
    if (!aiEnabled) {
      showToast('warning', 'ยังไม่ได้ตั้งค่า GEMINI_API_KEY ใน .env.local — กรุณาใช้ปุ่ม "กรอกเอง" เพื่อทดสอบ');
      return;
    }
    for (const file of Array.from(files)) {
      if (file.type !== 'application/pdf') { showToast('error', 'รองรับเฉพาะ PDF'); continue; }
      const b64 = await new Promise<string>((resolve) => {
        const r = new FileReader();
        r.onload = () => resolve((r.result as string).split(',')[1]);
        r.readAsDataURL(file);
      });
      setExtracting(true);
      try {
        const data = await api('/api/extract-pdf', 'POST', { pdfBase64: b64 });
        await preview(data.result, file.name);
        showToast('success', `อ่าน ${file.name} สำเร็จ — ตรวจสอบด้านล่าง`);
      } catch (e: any) { showToast('error', e.message); }
      finally { setExtracting(false); }
    }
  };

  const onExcelFiles = async (files: FileList) => {
    for (const file of Array.from(files)) {
      const name = file.name.toLowerCase();
      if (!name.endsWith('.xls') && !name.endsWith('.xlsx')) { showToast('error', 'รองรับเฉพาะไฟล์ Excel (.xls/.xlsx)'); continue; }
      const b64 = await new Promise<string>((resolve) => {
        const r = new FileReader();
        r.onload = () => resolve((r.result as string).split(',')[1]);
        r.readAsDataURL(file);
      });
      setImporting(true);
      try {
        const data = await api('/api/import-excel', 'POST', { fileBase64: b64 });
        const docs: ExtractedTripDocument[] = data.results || [];
        for (const doc of docs) await preview(doc, file.name);
        showToast('success', `นำเข้า ${file.name} สำเร็จ — พบ ${docs.length} ใบกระจาย ตรวจสอบด้านล่าง`);
      } catch (e: any) { showToast('error', e.message); }
      finally { setImporting(false); }
    }
  };

  const manual = () => {
    const today = new Date().toISOString().slice(0, 10);
    const ext: ExtractedTripDocument = {
      documentNo: `MAN-${Date.now().toString().slice(-5)}`, documentDate: today, plateNo: '',
      provinceRaw: '', districtRaw: '', receipts: [{ receiptNo: '', receiverName: '', senderName: '', items: [{ productName: '', quantity: 0 }] }],
    };
    preview(ext, 'manual.pdf');
  };

  const save = async () => {
    if (!pending) return;
    try {
      await api('/api/trips', 'POST', { cycleId: cycle.id, extracted: pending.extracted, fileName: pending.fileName, branchId });
      showToast('success', 'บันทึกใบกระจายลงฐานข้อมูลแล้ว');
      setPending(null);
      reload();
    } catch (e: any) { showToast('error', e.message); }
  };

  const del = async (id: string) => {
    if (!(await confirmDelete('ใบกระจายนี้'))) return;
    await api(`/api/trips/${id}`, 'DELETE');
    reload();
  };

  const recalc = async () => {
    const ok = await confirmAction({
      title: 'คำนวณใหม่ทั้งรอบ?',
      text: 'ระบบจะคำนวณใบกระจายทุกใบในรอบนี้ใหม่ด้วย Master ราคา/เงื่อนไขปัจจุบัน ยอดที่บันทึกไว้อาจเปลี่ยน',
      confirmText: 'คำนวณใหม่',
    });
    if (!ok) return;
    await api(`/api/cycles/${cycle.id}/recalculate`, 'POST');
    showToast('success', 'คำนวณใหม่ทั้งรอบด้วย Master ปัจจุบันแล้ว');
    reload();
  };

  const exportExcel = async () => {
    if (!cycleTrips.length) { showToast('warning', 'ยังไม่มีข้อมูลในรอบนี้'); return; }
    await exportCycleToExcel(cycle, db.tripDocuments, db.fuelEntries, db.deductions, db.vehicles, db.rateMasters);
    showToast('success', 'Export Excel สำเร็จ');
  };

  const visibleTrips = cycleTrips.filter((t: TripDocument) => {
    if (filter === 'divider') return t.receipts.some((r) => r.hasAdjustment);
    if (filter === 'warning') return t.warnings.length > 0;
    return true;
  });

  const totalTrip = cycleTrips.reduce((s: number, t: TripDocument) => s + t.tripAmount, 0);

  return (
    <div className="flex flex-col gap-5">
      {/* action bar */}
      <div className="bg-white rounded-2xl border border-natural-border p-4 flex flex-wrap items-center justify-between gap-3">
        <div className="text-sm">
          <span className="font-bold text-[#1B365D]">{cycle.name}</span>
          <span className="text-natural-muted ml-2">{cycleTrips.length} ใบกระจาย · ค่าเที่ยวรวม ฿{money(totalTrip)}</span>
        </div>
        <div className="flex gap-2">
          <button onClick={recalc} className="border border-natural-border rounded-full px-3 py-2 text-xs font-semibold flex items-center gap-1"><RefreshCw className="w-3.5 h-3.5" />Recalculate</button>
          <button onClick={exportExcel} className="bg-[#1B365D] text-white rounded-full px-4 py-2 text-xs font-semibold flex items-center gap-1"><FileSpreadsheet className="w-4 h-4" />Export Excel</button>
        </div>
      </div>

      {/* upload */}
      <div onDragOver={(e) => e.preventDefault()}
        onDrop={(e) => {
          e.preventDefault();
          const all = Array.from(e.dataTransfer.files) as File[];
          const xls = all.filter((f: File) => /\.xlsx?$/i.test(f.name));
          const pdfs = all.filter((f: File) => !/\.xlsx?$/i.test(f.name));
          if (xls.length) onExcelFiles(xls as unknown as FileList);
          if (pdfs.length) onFiles(pdfs as unknown as FileList);
        }}
        className="bg-white rounded-2xl border-2 border-dashed border-natural-border p-6 text-center flex flex-col items-center">
        <input ref={fileRef} type="file" aria-label="เลือกไฟล์ PDF ใบกระจาย" accept="application/pdf" multiple className="hidden" onChange={(e) => e.target.files && onFiles(e.target.files)} />
        <input ref={excelRef} type="file" aria-label="นำเข้า Excel ใบกระจาย" accept=".xls,.xlsx" multiple className="hidden" onChange={(e) => e.target.files && onExcelFiles(e.target.files)} />
        {extracting || importing ? (
          <div className="flex flex-col items-center gap-2 py-2"><RefreshCw className="w-8 h-8 text-[#1B365D] animate-spin" /><p className="text-sm font-semibold text-[#1B365D]">{importing ? 'กำลังอ่าน Excel ใบกระจาย...' : 'AI กำลังอ่าน PDF ใบกระจาย...'}</p></div>
        ) : (
          <>
            <UploadCloud className="w-8 h-8 text-[#1B365D] mb-2" />
            <p className="font-semibold text-sm text-[#1B365D]">ลากวางไฟล์ใบกระจาย — Excel (.xls/.xlsx) หรือ PDF (หลายไฟล์ได้)</p>
            <p className="text-[11px] text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-full px-3 py-1 mt-2">
              💡 แนะนำ Excel — ฟรี ไม่ใช้ AI ชื่อสินค้าเป๊ะ 100% ตัวหารจับอัตโนมัติ
            </p>
            {!aiEnabled && (
              <p className="text-[11px] text-amber-700 bg-amber-50 border border-amber-200 rounded-full px-3 py-1 mt-2">
                ⚠️ ยังไม่ได้ตั้งค่า GEMINI_API_KEY — อ่าน PDF จริงไม่ได้ (Excel ยังใช้ได้ปกติ)
              </p>
            )}
            <div className="flex gap-2 mt-3 items-center flex-wrap justify-center">
              <button onClick={() => excelRef.current?.click()}
                className="bg-emerald-600 text-white rounded-full px-4 py-2 text-xs font-semibold flex items-center gap-1"><FileSpreadsheet className="w-4 h-4" />นำเข้า Excel</button>
              <button onClick={() => fileRef.current?.click()} disabled={!aiEnabled}
                className="bg-[#1B365D] disabled:bg-natural-muted disabled:cursor-not-allowed text-white rounded-full px-4 py-2 text-xs font-semibold">เลือกไฟล์ PDF</button>
              <button onClick={manual} className="border border-natural-border rounded-full px-4 py-2 text-xs font-semibold">กรอกเอง</button>
              <label className="flex items-center gap-1 text-[11px] text-natural-muted ml-1">
                โมเดล AI:
                <select aria-label="เลือกโมเดล AI" value={db.settings?.geminiModel || 'gemini-3.5-flash'}
                  onChange={async (e) => {
                    await api('/api/settings', 'PUT', { geminiModel: e.target.value });
                    showToast('success', `เปลี่ยนโมเดลเป็น ${e.target.value}`);
                    reload();
                  }}
                  className="border border-natural-border rounded-lg px-2 py-1 text-[11px] font-semibold text-[#1B365D]">
                  {GEMINI_MODELS.map((m) => <option key={m} value={m}>{m}</option>)}
                </select>
              </label>
            </div>
          </>
        )}
      </div>

      {/* review */}
      {pending && <ReviewBoard pending={pending} setPending={setPending} onPreview={preview} onSave={save} locked={cycle.status === 'closed'} existingTrips={db.tripDocuments} cycles={db.cycles} cycleId={cycle.id} />}

      {/* filter */}
      <div className="flex gap-2">
        {([['all', 'ทั้งหมด'], ['divider', '🟧 เฉพาะมีตัวหาร'], ['warning', '⚠️ ต้องตรวจสอบ']] as const).map(([k, l]) => (
          <button key={k} onClick={() => setFilter(k)} className={`px-3 py-1.5 rounded-full text-xs font-semibold border ${filter === k ? 'bg-[#1B365D] text-white border-[#1B365D]' : 'border-natural-border'}`}>{l}</button>
        ))}
      </div>

      {/* saved trips */}
      {visibleTrips.length === 0 ? <EmptyHint text="ยังไม่มีใบกระจายในรอบนี้" /> :
        visibleTrips.map((t: TripDocument) => <TripCard key={t.id} trip={t} onDelete={() => { void del(t.id); }} />)}
    </div>
  );
}

// Review board — แก้ไข extracted + แสดงผล preview พร้อม badge
function ReviewBoard({ pending, setPending, onPreview, onSave, locked, existingTrips = [], cycles = [], cycleId }: any) {
  const ext: ExtractedTripDocument = pending.extracted;
  const prev: TripDocument = pending.preview;
  const needsBox = prev.receipts.some((r) => r.requiresManualBox && (r.manualBoxQty == null || r.manualBoxQty <= 0));

  // 🔒 กฎเหล็ก: เลขใบกระจายห้ามซ้ำในสาขา (ทุกรอบ)
  const docNo = (ext.documentNo || '').trim();
  const dupTrip = docNo ? (existingTrips as TripDocument[]).find((t) => (t.documentNo || '').trim() === docNo) : null;
  const dupCycleName = dupTrip ? ((cycles as BillingCycle[]).find((c) => c.id === dupTrip.cycleId)?.name || dupTrip.cycleId) : '';
  const isDup = !!dupTrip;

  const update = (patch: Partial<ExtractedTripDocument>) => onPreview({ ...ext, ...patch }, pending.fileName);
  const updReceipt = (ri: number, patch: any) => {
    const receipts = ext.receipts.map((r, i) => i === ri ? { ...r, ...patch } : r);
    onPreview({ ...ext, receipts }, pending.fileName);
  };
  const updItem = (ri: number, ii: number, patch: any) => {
    const receipts = ext.receipts.map((r, i) => i === ri ? { ...r, items: r.items.map((it, j) => j === ii ? { ...it, ...patch } : it) } : r);
    onPreview({ ...ext, receipts }, pending.fileName);
  };

  return (
    <div className="bg-white rounded-2xl border border-natural-border p-5 flex flex-col gap-4">
      <div className="flex items-center justify-between border-b border-natural-border pb-2">
        <h3 className="font-bold text-sm text-[#1B365D]">ขั้นตอนตรวจสอบก่อนยืนยัน — {pending.fileName}</h3>
        <button onClick={() => setPending(null)} className="text-xs text-natural-muted hover:text-rose-600 font-semibold">ยกเลิก</button>
      </div>

      {/* warnings */}
      {(prev.warnings || []).length > 0 && (
        <div className="bg-[#FCE4D6] border border-amber-200 rounded-xl p-3 text-xs text-[#9C0006] flex flex-col gap-1">
          {(prev.warnings || []).map((w: string, i: number) => <div key={i} className="flex gap-1.5"><AlertTriangle className="w-4 h-4 shrink-0" />{w}</div>)}
        </div>
      )}

      {/* 🔒 เตือนเลขใบกระจายซ้ำ */}
      {isDup && (
        <div className="bg-rose-50 border-2 border-rose-400 rounded-xl p-3 text-xs text-rose-800 flex gap-1.5 font-semibold">
          <AlertTriangle className="w-4 h-4 shrink-0" />
          เลขใบกระจาย <b>{docNo}</b> ซ้ำ — มีอยู่แล้ว{dupTrip.cycleId === cycleId ? 'ในรอบนี้' : `ในรอบ "${dupCycleName}"`} · บันทึกไม่ได้ (ถ้าต้องการแก้ ให้ลบใบเดิมก่อน หรือเปลี่ยนเลขให้ถูกต้อง)
        </div>
      )}

      {/* header fields */}
      <div className="grid grid-cols-2 md:grid-cols-3 gap-3 text-xs">
        <Field label="เลขที่ใบกระจาย" value={ext.documentNo} onChange={(v) => update({ documentNo: v })} />
        <Field label="วันที่ออก" type="date" value={ext.documentDate} onChange={(v) => update({ documentDate: v })} />
        <Field label="ทะเบียนรถ" value={ext.plateNo} onChange={(v) => update({ plateNo: v })} />
        <Field label="จังหวัด (ค่าเริ่มต้น)" value={ext.provinceRaw} onChange={(v) => update({ provinceRaw: v })} />
        <Field label="อำเภอ (ค่าเริ่มต้น)" value={ext.districtRaw} onChange={(v) => update({ districtRaw: v })} />
      </div>

      {/* แถบเลือกราคาขนส่ง — จุดสำคัญ ทำให้เด่นชัด */}
      <div className={`rounded-2xl border-2 p-3.5 flex flex-col sm:flex-row sm:items-center gap-3 ${
        prev.rateType === 'piece' ? 'border-[#C65911] bg-[#FFF2CC]' : prev.rateType === 'flat' ? 'border-[#1B365D] bg-[#EAF2F8]' : 'border-rose-400 bg-rose-50'}`}>
        <div className="flex items-center gap-2.5 shrink-0">
          <div className={`rounded-xl p-2 text-white ${prev.rateType === 'piece' ? 'bg-[#C65911]' : prev.rateType === 'flat' ? 'bg-[#1B365D]' : 'bg-rose-500'}`}>
            <Coins className="w-5 h-5" />
          </div>
          <div>
            <div className="font-bold text-sm text-natural-text flex items-center gap-1.5">เลือกราคาขนส่ง <span className="text-rose-600 text-xs">*สำคัญ</span></div>
            <div className="text-[11px] text-natural-muted">ทั้งใบใช้แบบเดียวกัน — เลือกผิดราคาจะเพี้ยน</div>
          </div>
        </div>
        {((prev.breakdown?.peat || 0) > 0 || (prev.breakdown?.collect || 0) > 0) && (prev.breakdown?.normal || 0) === 0 ? (
          <span className="sm:ml-auto font-bold text-emerald-700 flex items-center gap-1.5 text-sm">🟢 คิดตามสินค้าพิเศษ — ไม่ใช้เหมา/ชิ้น</span>
        ) : prev.rateOptions.flat != null || prev.rateOptions.piece != null ? (
          <div className="flex items-center gap-2.5 sm:ml-auto w-full sm:w-auto">
            <select aria-label="เลือกราคาขนส่ง" value={ext.rateChoice || prev.rateType || ''}
              onChange={(e) => update({ rateChoice: e.target.value as any })}
              className="flex-1 sm:flex-none border-2 border-natural-border bg-white rounded-xl px-3 py-2.5 text-base font-bold text-[#1B365D] focus:border-[#1B365D] outline-none cursor-pointer shadow-xs">
              {prev.rateOptions.flat != null && <option value="flat">🔵 เหมา ฿{money(prev.rateOptions.flat)} (สูงสุด)</option>}
              {prev.rateOptions.piece != null && <option value="piece">🟠 ชิ้น รวมทุกจุด ฿{money(prev.rateOptions.piece)}</option>}
            </select>
            <span className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-bold text-white ${prev.rateType === 'piece' ? 'bg-[#C65911]' : 'bg-[#1B365D]'}`}>
              {prev.rateType === 'piece' ? 'คิดแบบ ชิ้น' : 'คิดแบบ เหมา'}
            </span>
          </div>
        ) : (
          <span className="sm:ml-auto font-bold text-[#9C0006] flex items-center gap-1.5"><AlertTriangle className="w-4 h-4" />— ไม่เจอราคาขนส่ง —</span>
        )}
      </div>

      <datalist id="units">
        {['กล่อง', 'หีบ', 'ลัง', 'ชิ้น', 'แพ็ค', 'ถุง', 'โหล'].map((u) => <option key={u} value={u} />)}
      </datalist>

      {/* receipts */}
      {ext.receipts.map((r, ri) => {
        const pr = prev.receipts[ri];
        return (
          <div key={ri} className={`rounded-xl border p-3 ${pr?.requiresManualBox ? 'bg-sky-50 border-l-4 border-l-sky-500 border-natural-border' : pr?.hasAdjustment ? 'bg-[#FFF2CC] border-l-4 border-l-[#C65911] border-natural-border' : 'border-natural-border'}`}>
            <div className="grid grid-cols-2 md:grid-cols-3 gap-2 text-xs mb-2">
              <Field label="เลขใบรับสินค้า" value={r.receiptNo} onChange={(v) => updReceipt(ri, { receiptNo: v })} />
              <Field label="ผู้รับสินค้า" value={r.receiverName} onChange={(v) => updReceipt(ri, { receiverName: v })} />
              <Field label="ผู้ส่งสินค้า" value={r.senderName} onChange={(v) => updReceipt(ri, { senderName: v })} />
              <Field label="จังหวัดปลายทาง" value={r.provinceRaw || ''} onChange={(v) => updReceipt(ri, { provinceRaw: v })} />
              <Field label="อำเภอปลายทาง" value={r.districtRaw || ''} onChange={(v) => updReceipt(ri, { districtRaw: v })} />
              <div className="flex flex-col justify-end text-[11px]">
                <span className="text-natural-muted text-[10px] font-bold uppercase">ราคา/ค่าเที่ยวจุดนี้</span>
                <span className="font-semibold text-[#1B365D]">
                  {prev.rateType === 'piece'
                    ? (pr?.piecePrice != null ? `ชิ้น ฿${money(pr.piecePrice)} × ${pr.billingQty} = ฿${money(pr.receiptAmount)}` : '⚠️ ไม่เจอราคาชิ้น')
                    : (pr?.flatPrice != null ? `เหมา ฿${money(pr.flatPrice)}` : '⚠️ ไม่เจอราคาเหมา')}
                </span>
              </div>
            </div>
            <table className="w-full text-xs">
              <thead><tr className="text-natural-muted text-left"><th className="py-1">รายการสินค้า</th><th className="w-16 text-center">จำนวน</th><th className="w-20 text-center">หน่วย</th><th className="w-10"></th></tr></thead>
              <tbody>
                {r.items.map((it, ii) => {
                  const unspecified = isUnspecifiedName(it.productName);
                  return (
                  <tr key={ii} title={unspecified ? 'ชื่อสินค้ายังไม่ระบุ — นับเข้ายอดตามเอกสาร' : ''}>
                    <td className="py-0.5">
                      <div className="flex items-center gap-1">
                        <input value={it.productName} aria-label="ชื่อสินค้า" onChange={(e) => updItem(ri, ii, { productName: e.target.value })} className={`w-full border-b border-dashed border-natural-border bg-transparent p-1 ${unspecified ? 'text-amber-700' : ''}`} placeholder="ชื่อสินค้า" />
                        {unspecified && <span className="shrink-0 text-[9px] font-bold bg-amber-50 text-amber-700 border border-amber-300 rounded px-1 py-0.5">ระบุชื่อ?</span>}
                      </div>
                    </td>
                    <td><input type="number" aria-label="จำนวนสินค้า" value={it.quantity || ''} onChange={(e) => updItem(ri, ii, { quantity: +e.target.value || 0 })} className="w-full text-center border-b border-dashed border-natural-border bg-transparent p-1 font-bold" /></td>
                    <td><input list="units" aria-label="หน่วยนับ" value={it.unit || ''} onChange={(e) => updItem(ri, ii, { unit: e.target.value })} placeholder="หน่วย" className="w-full text-center border-b border-dashed border-natural-border bg-transparent p-1 text-natural-muted" /></td>
                    <td className="text-center"><button type="button" aria-label="ลบรายการสินค้า" title="ลบรายการสินค้า" onClick={() => updReceipt(ri, { items: r.items.filter((_, j) => j !== ii) })} className="text-natural-muted hover:text-rose-600"><Trash2 className="w-3.5 h-3.5" /></button></td>
                  </tr>
                  );
                })}
              </tbody>
            </table>
            <button onClick={() => updReceipt(ri, { items: [...r.items, { productName: '', quantity: 0 }] })} className="text-[#1B365D] text-xs font-semibold mt-1 flex items-center gap-1"><Plus className="w-3 h-3" />เพิ่มสินค้า</button>
            {pr?.hasAdjustment && (
              <div className="mt-2 text-[11px] text-[#C65911] font-semibold flex flex-wrap gap-2">
                {(pr.adjustments || []).map((a: any, i: number) => <span key={i} className="bg-white border border-[#C65911]/40 rounded-full px-2 py-0.5">🟧÷{a.divisor} {a.productName}: {a.note}</span>)}
              </div>
            )}
            {pr?.requiresManualBox && (
              <div className="mt-2 bg-sky-100 border border-sky-300 rounded-lg p-2 flex flex-wrap items-center gap-2 text-xs">
                <span className="font-bold text-sky-800">📦 ผู้ส่งนี้ส่งเป็นชิ้น (อ่านได้ {pr.totalQty} ชิ้น) — กรอกจำนวนกล่อง:</span>
                <input type="number" aria-label="จำนวนกล่อง (บังคับ)" placeholder="จำนวนกล่อง"
                  value={r.manualBoxQty ?? ''} onChange={(e) => updReceipt(ri, { manualBoxQty: e.target.value === '' ? undefined : +e.target.value })}
                  className={`border rounded-lg px-2 py-1 w-24 font-bold ${(r.manualBoxQty ?? 0) > 0 ? 'border-sky-400 text-sky-800' : 'border-rose-400 bg-rose-50'}`} />
                {!((r.manualBoxQty ?? 0) > 0) && <span className="text-rose-600 font-semibold">* บังคับกรอกก่อนบันทึก</span>}
              </div>
            )}
            <div className="mt-2 text-xs font-semibold flex gap-4">
              <span>จำนวนจริง: <b>{pr?.totalQty}</b>{pr?.requiresManualBox ? ' ชิ้น' : ''}</span>
              <span className="text-[#C00000]">คิดค่าเที่ยว: <b>{pr?.billingQty}</b>{pr?.requiresManualBox ? ' กล่อง' : ''}</span>
            </div>
          </div>
        );
      })}
      <button onClick={() => update({ receipts: [...ext.receipts, { receiptNo: '', receiverName: '', senderName: '', items: [{ productName: '', quantity: 0 }] }] })}
        className="text-[#1B365D] text-xs font-semibold flex items-center gap-1"><Plus className="w-3.5 h-3.5" />เพิ่มใบรับสินค้า</button>

      {/* totals + save */}
      <div className="flex items-center justify-between border-t border-natural-border pt-3">
        <div className="text-sm font-bold text-[#1B365D]">ค่าเที่ยวรวม: ฿{money(prev.tripAmount)} <span className="text-natural-muted font-normal">(คิด {prev.billingQty}/{prev.totalQty} ลัง)</span>
          {prev.breakdown && ((prev.breakdown.collect || 0) > 0 || (prev.breakdown.peat || 0) > 0) && (
            <span className="block text-[11px] text-natural-muted font-normal mt-0.5">
              งานปกติ ฿{money(prev.breakdown.normal)}
              {(prev.breakdown.peat || 0) > 0 && <> + Peat mass ฿{money(prev.breakdown.peat)}</>}
              {(prev.breakdown.collect || 0) > 0 && <> + เก็บคืน ฿{money(prev.breakdown.collect)}</>}
            </span>
          )}
          {needsBox && <span className="block text-rose-600 text-xs font-semibold mt-0.5">⚠️ มีใบรับที่ต้องกรอกจำนวนกล่องก่อนบันทึก</span>}
          {isDup && <span className="block text-rose-600 text-xs font-semibold mt-0.5">🔒 เลขใบกระจายซ้ำ — บันทึกไม่ได้</span>}
        </div>
        <button onClick={onSave} disabled={locked || needsBox || isDup} title={isDup ? 'เลขใบกระจายซ้ำ' : needsBox ? 'กรอกจำนวนกล่องให้ครบก่อน' : ''} className="bg-[#1B365D] disabled:bg-natural-muted disabled:cursor-not-allowed text-white rounded-full px-5 py-2 text-sm font-semibold flex items-center gap-1.5"><Save className="w-4 h-4" />ยืนยันบันทึก</button>
      </div>
    </div>
  );
}

const TripCard: React.FC<{ trip: TripDocument; onDelete: () => void }> = ({ trip, onDelete }) => {
  const hasDiv = trip.receipts.some((r) => r.hasAdjustment);
  return (
    <div className={`bg-white rounded-2xl border p-4 ${trip.warnings.length ? 'border-l-4 border-l-[#9C0006]' : hasDiv ? 'border-l-4 border-l-[#C65911]' : 'border-natural-border'}`}>
      <div className="flex items-center justify-between">
        <div>
          <span className="font-bold text-[#1B365D] text-sm">{trip.documentNo}</span>
          <span className="text-natural-muted text-xs ml-2">{trip.documentDate} · {trip.plateNo} {trip.driverName && `(${trip.driverName})`} · {trip.rateType === 'flat' ? 'เหมา' : 'ชิ้น'}</span>
        </div>
        <div className="flex items-center gap-3">
          <span className="text-sm font-bold text-[#C00000]">฿{money(trip.tripAmount)}</span>
          <button type="button" aria-label="ลบใบกระจาย" title="ลบใบกระจาย" onClick={onDelete} className="text-natural-muted hover:text-rose-600"><Trash2 className="w-4 h-4" /></button>
        </div>
      </div>
      {trip.warnings.length > 0 && <div className="mt-1 text-[11px] text-[#9C0006]">⚠️ {trip.warnings.join(' · ')}</div>}
      <div className="mt-2 overflow-x-auto">
        <table className="w-full text-xs min-w-[560px]">
          <thead><tr className="text-natural-muted text-left border-b border-natural-border"><th className="py-1">ใบรับสินค้า</th><th>ผู้รับ</th><th>ปลายทาง</th><th className="text-center w-16">จำนวนจริง</th><th className="text-center w-16">คิดค่าเที่ยว</th><th className="text-right w-20">ค่าเที่ยวจุด</th></tr></thead>
          <tbody>
            {trip.receipts.map((r) => (
              <tr key={r.id} className={r.hasAdjustment ? 'bg-[#FFF2CC]' : ''}>
                <td className="py-1">{r.hasAdjustment && r.adjustments?.[0] && <span className="text-[#C65911] font-bold">🟧÷{r.adjustments[0].divisor} </span>}{r.receiptNo}</td>
                <td>{r.receiverName}</td>
                <td>{r.districtRaw} {r.provinceRaw}</td>
                <td className="text-center">{r.totalQty}</td>
                <td className="text-center font-bold text-[#C00000]" title={(r.adjustments || []).map((a) => a.note).join(' | ')}>{r.billingQty}</td>
                <td className="text-right">{trip.rateType === 'piece' ? money(r.receiptAmount) : (r.flatPrice != null ? money(r.flatPrice) : '-')}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ===========================================================================
// Tab: ค่าน้ำมัน & รายการหัก
// ===========================================================================
const ALL_BRANCH_HINT = 'อยู่โหมดภาพรวมทุกสาขา — เลือกสาขาที่มุมบนขวาเพื่อจัดการข้อมูล';
function FuelDeductionTab({ db, cycle, api, branchId, reload, showToast }: any) {
  const cats: MoneyCategory[] = db.moneyCategories || [];
  const incomeCats = cats.filter((c) => c.kind === 'income' && c.status === 'active');
  const dedCats = cats.filter((c) => c.kind === 'deduction' && c.status === 'active');

  const fuel = db.fuelEntries.filter((f: FuelEntry) => cycle && f.cycleId === cycle.id);
  const allDed = db.deductions.filter((d: DeductionEntry) => cycle && d.cycleId === cycle.id);
  const ded = allDed.filter((d: DeductionEntry) => d.kind === 'deduction');
  const incomes = allDed.filter((d: DeductionEntry) => d.kind === 'income');

  const [fForm, setFForm] = useState({ plateNo: '', refNo: '', date: cycle?.startDate || '', amount: 0 });
  const [dForm, setDForm] = useState({ plateNo: '', categoryId: '', amount: 0, docNo: '' });
  const [bForm, setBForm] = useState({ plateNo: '', categoryId: '', amount: 0, docNo: '' });

  if (!cycle) return <EmptyHint text="กรุณาเลือกรอบก่อน" />;
  if (!branchId) return <EmptyHint text={ALL_BRANCH_HINT} />;

  const addFuel = async () => {
    if (!fForm.plateNo || !fForm.amount) return showToast('warning', 'กรอกทะเบียนและจำนวนเงิน');
    await api('/api/fuel', 'POST', { ...fForm, cycleId: cycle.id, branchId });
    setFForm({ plateNo: '', refNo: '', date: cycle.startDate, amount: 0 }); reload();
  };
  const addEntry = async (plateNo: string, categoryId: string, amount: number, kind: 'income' | 'deduction', docNo: string, reset: () => void) => {
    const cat = cats.find((c) => c.id === categoryId) || (kind === 'income' ? incomeCats[0] : dedCats[0]);
    if (!plateNo || !amount || !cat) return showToast('warning', 'กรอกทะเบียน/จำนวนเงิน และเลือกประเภท');
    await api('/api/deductions', 'POST', { plateNo, categoryId: cat.id, kind, label: cat.name, amount, docNo: docNo.trim(), cycleId: cycle.id, branchId });
    reset(); reload();
  };

  return (
    <div className="grid md:grid-cols-2 gap-5">
      <Section title="ค่าน้ำมัน (แยกตามทะเบียน)" icon={Fuel}>
        <div className="flex flex-wrap gap-2 mb-3">
          <input list="plates" aria-label="ทะเบียนรถ" placeholder="ทะเบียน" value={fForm.plateNo} onChange={(e) => setFForm({ ...fForm, plateNo: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-28" />
          <input aria-label="เลขใบสั่งเติม" placeholder="เลขใบสั่งเติม" value={fForm.refNo} onChange={(e) => setFForm({ ...fForm, refNo: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-32" />
          <input type="date" aria-label="วันที่เติมน้ำมัน" value={fForm.date} onChange={(e) => setFForm({ ...fForm, date: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm" />
          <input type="number" aria-label="จำนวนเงินค่าน้ำมัน" placeholder="จำนวนเงิน" value={fForm.amount || ''} onChange={(e) => setFForm({ ...fForm, amount: +e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-28" />
          <button onClick={addFuel} className="bg-[#1B365D] text-white rounded-lg px-3 text-sm font-semibold">เพิ่ม</button>
        </div>
        <SimpleTable rows={fuel.map((f: FuelEntry) => [f.plateNo, f.refNo, f.date, money(f.amount)])} cols={['ทะเบียน', 'ใบสั่งเติม', 'วันที่', 'จำนวน']}
          onDelete={async (i: number) => { await api(`/api/fuel/${fuel[i].id}`, 'DELETE'); reload(); }} />
      </Section>

      {/* รายได้เพิ่ม (income) — dropdown ดึงจาก Master ประเภท */}
      <Section title="รายได้เพิ่ม (+)" icon={Plus}>
        <p className="text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-lg px-2 py-1 mb-3">เงินที่บวกรวมกับค่าเที่ยววิ่ง เช่น ค่าอัพเดทบิล</p>
        <div className="flex flex-wrap gap-2 mb-3">
          <input list="plates" aria-label="ทะเบียนรถ" placeholder="ทะเบียน" value={bForm.plateNo} onChange={(e) => setBForm({ ...bForm, plateNo: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-28" />
          <select aria-label="ประเภทรายได้เพิ่ม" value={bForm.categoryId} onChange={(e) => setBForm({ ...bForm, categoryId: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm">
            {incomeCats.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
          <input type="number" aria-label="จำนวนเงินรายได้เพิ่ม" placeholder="จำนวนเงิน" value={bForm.amount || ''} onChange={(e) => setBForm({ ...bForm, amount: +e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-28" />
          <input aria-label="ใบกระจายเลขที่" placeholder="ใบกระจายเลขที่" value={bForm.docNo} onChange={(e) => setBForm({ ...bForm, docNo: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-36" />
          <button onClick={() => addEntry(bForm.plateNo, bForm.categoryId, bForm.amount, 'income', bForm.docNo, () => setBForm({ plateNo: '', categoryId: '', amount: 0, docNo: '' }))} className="bg-emerald-600 text-white rounded-lg px-3 text-sm font-semibold">เพิ่ม</button>
        </div>
        <SimpleTable rows={incomes.map((d: DeductionEntry) => [d.plateNo, d.label, d.docNo || '-', `+${money(d.amount)}`])} cols={['ทะเบียน', 'รายการ', 'ใบกระจาย', 'จำนวนเพิ่ม']}
          onDelete={async (i: number) => { await api(`/api/deductions/${incomes[i].id}`, 'DELETE'); reload(); }} />
      </Section>

      {/* รายการหัก (deduction) */}
      <Section title="รายการหัก (−)" icon={Receipt}>
        <div className="flex flex-wrap gap-2 mb-3">
          <input list="plates" aria-label="ทะเบียนรถ" placeholder="ทะเบียน" value={dForm.plateNo} onChange={(e) => setDForm({ ...dForm, plateNo: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-28" />
          <select aria-label="ประเภทรายการหัก" value={dForm.categoryId} onChange={(e) => setDForm({ ...dForm, categoryId: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm">
            {dedCats.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
          <input type="number" aria-label="จำนวนเงินรายการหัก" placeholder="จำนวนเงิน" value={dForm.amount || ''} onChange={(e) => setDForm({ ...dForm, amount: +e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-28" />
          <input aria-label="ใบกระจายเลขที่" placeholder="ใบกระจายเลขที่" value={dForm.docNo} onChange={(e) => setDForm({ ...dForm, docNo: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-36" />
          <button onClick={() => addEntry(dForm.plateNo, dForm.categoryId, dForm.amount, 'deduction', dForm.docNo, () => setDForm({ plateNo: '', categoryId: '', amount: 0, docNo: '' }))} className="bg-[#1B365D] text-white rounded-lg px-3 text-sm font-semibold">เพิ่ม</button>
        </div>
        <SimpleTable rows={ded.map((d: DeductionEntry) => [d.plateNo, d.label, d.docNo || '-', money(d.amount)])} cols={['ทะเบียน', 'รายการ', 'ใบกระจาย', 'จำนวน']}
          onDelete={async (i: number) => { await api(`/api/deductions/${ded[i].id}`, 'DELETE'); reload(); }} />
      </Section>

      {/* จัดการประเภท — เพิ่มชื่อใน dropdown ได้เองโดยไม่ต้องแก้โค้ด */}
      <CategoryManager cats={cats} api={api} branchId={branchId} reload={reload} showToast={showToast} />

      <datalist id="plates">{db.vehicles.map((v: Vehicle) => <option key={v.id} value={v.plateNo} />)}</datalist>
    </div>
  );
}

// จัดการ Master ประเภทรายได้เพิ่ม / รายการหัก (เติม dropdown)
function CategoryManager({ cats, api, branchId, reload, showToast }: any) {
  const [form, setForm] = useState<{ name: string; kind: 'income' | 'deduction' }>({ name: '', kind: 'deduction' });
  const add = async () => {
    if (!form.name.trim()) return showToast('warning', 'กรอกชื่อประเภท');
    await api('/api/money-categories', 'POST', { name: form.name.trim(), kind: form.kind, status: 'active', builtin: false, branchId });
    setForm({ name: '', kind: 'deduction' }); reload(); showToast('success', 'เพิ่มประเภทแล้ว — ใช้ใน dropdown ได้เลย');
  };
  const del = async (c: MoneyCategory) => {
    if (c.builtin) return showToast('warning', 'ประเภทของระบบ ลบไม่ได้ (ปิดใช้งานแทนได้)');
    await api(`/api/money-categories/${c.id}`, 'DELETE'); reload();
  };
  return (
    <Section title="จัดการประเภท (รายได้เพิ่ม / รายการหัก)" icon={Tag}>
      <p className="text-xs text-natural-muted mb-3">เพิ่มประเภทใหม่ที่นี่ แล้วชื่อจะไปโผล่ใน dropdown ของ "รายได้เพิ่ม" หรือ "รายการหัก" อัตโนมัติ</p>
      <div className="flex flex-wrap gap-2 mb-3">
        <input aria-label="ชื่อประเภท" placeholder="ชื่อประเภท เช่น ค่าปรับ" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-40" />
        <select aria-label="ทิศทาง" value={form.kind} onChange={(e) => setForm({ ...form, kind: e.target.value as any })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm">
          <option value="deduction">รายการหัก (−)</option>
          <option value="income">รายได้เพิ่ม (+)</option>
        </select>
        <button onClick={add} className="bg-[#1B365D] text-white rounded-lg px-3 text-sm font-semibold">เพิ่มประเภท</button>
      </div>
      <SimpleTable cols={['ชื่อประเภท', 'ทิศทาง', 'ระบบ']}
        rows={cats.map((c: MoneyCategory) => [c.name, c.kind === 'income' ? 'รายได้เพิ่ม (+)' : 'หักออก (−)', c.builtin ? '✓' : ''])}
        onDelete={async (i: number) => { await del(cats[i]); }} />
    </Section>
  );
}

// ===========================================================================
// Tab: Dashboard (สรุปรับสุทธิต่อทะเบียน)
// ===========================================================================
function DashboardTab({ db, cycle, branchId, isHQ }: any) {
  if (!cycle) return <EmptyHint text="กรุณาเลือกรอบก่อน" />;

  // โหมด HQ ภาพรวมทุกสาขา (เลือก "ทุกสาขา")
  if (isHQ && !branchId) return <HQDashboard db={db} cycle={cycle} />;

  const sums = summarizeByVehicle(cycle.id, db.tripDocuments, db.fuelEntries, db.deductions, db.vehicles);
  const g = sums.reduce((a, s) => ({ trip: a.trip + s.totalTripAmount, fuel: a.fuel + s.fuelTotal, net: a.net + s.netReceive }), { trip: 0, fuel: 0, net: 0 });

  return (
    <div className="flex flex-col gap-4">
      <div className="grid grid-cols-3 gap-4">
        <Stat label="รายได้ค่าเที่ยวรวม" value={`฿${money(g.trip)}`} />
        <Stat label="ค่าน้ำมันรวม" value={`฿${money(g.fuel)}`} />
        <Stat label="รับสุทธิรวม" value={`฿${money(g.net)}`} highlight />
      </div>
      <div className="bg-white rounded-2xl border border-natural-border overflow-x-auto">
        <table className="w-full text-xs min-w-[800px]">
          <thead className="bg-[#1B365D] text-white"><tr>
            {['ทะเบียน', 'คนขับ', 'รายได้', 'หัก 1%', 'ค่าน้ำมัน', '+ รายได้เพิ่ม', 'รวมรายการหัก', 'รับสุทธิ'].map((h) => <th key={h} className="p-2 font-semibold">{h}</th>)}
          </tr></thead>
          <tbody>
            {sums.map((s, i) => (
              <tr key={s.plateNo} className={i % 2 ? 'bg-[#F9FAFC]' : ''}>
                <td className="p-2 font-semibold">{s.plateNo}</td>
                <td className="p-2">{s.driverName}</td>
                <td className="p-2 text-right">{money(s.totalTripAmount)}</td>
                <td className="p-2 text-right text-rose-700">{money(s.deduction1Percent)}</td>
                <td className="p-2 text-right text-rose-700">{money(s.fuelTotal)}</td>
                <td className="p-2 text-right text-emerald-700">{money(s.incomeAdd)}</td>
                <td className="p-2 text-right text-rose-700">{money(s.deductionTotal)}</td>
                <td className="p-2 text-right font-bold text-[#C00000]">{money(s.netReceive)}</td>
              </tr>
            ))}
            {sums.length === 0 && <tr><td colSpan={8} className="p-6 text-center text-natural-muted">ยังไม่มีข้อมูล</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ===========================================================================
// Dashboard รวมทุกสาขา (HQ) — เทียบยอดระหว่างสาขา
// ===========================================================================
function HQDashboard({ db, cycle }: any) {
  const branches = (db.branches as Branch[]).filter((b) => !b.isHQ && b.status === 'active');
  const rows = branches.map((b) => {
    const trips = db.tripDocuments.filter((t: TripDocument) => t.cycleId === cycle.id && t.branchId === b.id);
    const sums = summarizeByVehicle(
      cycle.id,
      db.tripDocuments.filter((t: TripDocument) => t.branchId === b.id),
      db.fuelEntries.filter((f: FuelEntry) => f.branchId === b.id),
      db.deductions.filter((d: DeductionEntry) => d.branchId === b.id),
      db.vehicles.filter((v: Vehicle) => v.branchId === b.id),
    );
    return {
      branch: b,
      docs: trips.length,
      trucks: sums.length,
      trip: sums.reduce((a, s) => a + s.totalTripAmount, 0),
      fuel: sums.reduce((a, s) => a + s.fuelTotal, 0),
      income: sums.reduce((a, s) => a + s.incomeAdd, 0),
      deduct: sums.reduce((a, s) => a + s.deductionTotal + s.deduction1Percent, 0),
      net: sums.reduce((a, s) => a + s.netReceive, 0),
    };
  });
  const g = rows.reduce((a, r) => ({
    docs: a.docs + r.docs, trip: a.trip + r.trip, fuel: a.fuel + r.fuel, net: a.net + r.net,
  }), { docs: 0, trip: 0, fuel: 0, net: 0 });
  const maxTrip = Math.max(1, ...rows.map((r) => r.trip));

  return (
    <div className="flex flex-col gap-4">
      <div className="bg-emerald-50 border border-emerald-200 rounded-xl px-4 py-2 text-sm font-semibold text-emerald-800 flex items-center gap-2">
        <Building2 className="w-4 h-4" /> ภาพรวมทุกสาขา · {cycle.name}
      </div>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <Stat label="ใบกระจายรวม" value={`${g.docs}`} />
        <Stat label="ค่าเที่ยวรวมทุกสาขา" value={`฿${money(g.trip)}`} />
        <Stat label="ค่าน้ำมันรวม" value={`฿${money(g.fuel)}`} />
        <Stat label="รับสุทธิรวม" value={`฿${money(g.net)}`} highlight />
      </div>

      <div className="bg-white rounded-2xl border border-natural-border overflow-x-auto">
        <table className="w-full text-xs min-w-[760px]">
          <thead className="bg-[#1B365D] text-white"><tr>
            {['สาขา', 'ใบกระจาย', 'รถ', 'ค่าเที่ยว', 'ค่าน้ำมัน', '+รายได้เพิ่ม', 'รวมหัก', 'รับสุทธิ', 'สัดส่วน'].map((h) => <th key={h} className="p-2 font-semibold text-left">{h}</th>)}
          </tr></thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={r.branch.id} className={i % 2 ? 'bg-[#F9FAFC]' : ''}>
                <td className="p-2 font-semibold text-[#1B365D]">{r.branch.name}</td>
                <td className="p-2">{r.docs}</td>
                <td className="p-2">{r.trucks}</td>
                <td className="p-2 text-right">{money(r.trip)}</td>
                <td className="p-2 text-right text-rose-700">{money(r.fuel)}</td>
                <td className="p-2 text-right text-emerald-700">{money(r.income)}</td>
                <td className="p-2 text-right text-rose-700">{money(r.deduct)}</td>
                <td className="p-2 text-right font-bold text-[#C00000]">{money(r.net)}</td>
                <td className="p-2 w-32"><div className="bg-natural-bg rounded-full h-2 overflow-hidden"><div className="bg-[#1B365D] h-2 rounded-full" style={{ width: `${(r.trip / maxTrip) * 100}%` }} /></div></td>
              </tr>
            ))}
            <tr className="border-t-2 border-[#1B365D] font-bold bg-[#FFF2CC]">
              <td className="p-2">รวมทุกสาขา</td>
              <td className="p-2">{g.docs}</td>
              <td className="p-2"></td>
              <td className="p-2 text-right">{money(g.trip)}</td>
              <td className="p-2 text-right">{money(g.fuel)}</td>
              <td className="p-2"></td>
              <td className="p-2"></td>
              <td className="p-2 text-right text-[#C00000]">{money(g.net)}</td>
              <td className="p-2"></td>
            </tr>
          </tbody>
        </table>
      </div>
      <p className="text-[11px] text-natural-muted">💡 เลือกสาขาที่มุมบนเพื่อดู/แก้ไขรายละเอียดของแต่ละสาขา</p>
    </div>
  );
}

// ===========================================================================
// Tab: รายงานต่อทะเบียน (Export Excel)
// ===========================================================================
const fmtD = (s: string) => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ''); return m ? `${+m[3]}/${+m[2]}/${m[1]}` : (s || ''); };

function ReportsTab({ db, cycle, branchId, showToast }: any) {
  if (!branchId) return <EmptyHint text={ALL_BRANCH_HINT} />;
  if (!cycle) return <EmptyHint text="กรุณาเลือกรอบก่อน" />;
  const branchName = (db.branches as Branch[]).find((b) => b.id === branchId)?.name || '';
  const cycleTrips = db.tripDocuments.filter((t: TripDocument) => t.cycleId === cycle.id);
  const sums = summarizeByVehicle(cycle.id, db.tripDocuments, db.fuelEntries, db.deductions, db.vehicles);
  const exp = async () => {
    if (!cycleTrips.length) return showToast('warning', 'ยังไม่มีใบกระจายในรอบนี้');
    try { await exportPerVehicleReport(cycle, branchName, db.tripDocuments, db.fuelEntries, db.deductions, db.vehicles); showToast('success', 'Export สำเร็จ — ดาวน์โหลดแล้ว'); }
    catch (e: any) { showToast('error', e.message); }
  };
  const TH = ({ children, r }: any) => <th className={`px-2 py-1.5 font-semibold ${r ? 'text-right' : 'text-left'}`}>{children}</th>;
  const TD = ({ children, r, b }: any) => <td className={`px-2 py-1 ${r ? 'text-right' : 'text-left'} ${b ? 'font-bold' : ''}`}>{children}</td>;

  return (
    <div className="flex flex-col gap-4">
      {/* action */}
      <div className="bg-white rounded-2xl border border-natural-border p-4 flex flex-wrap items-center justify-between gap-3">
        <div className="text-sm"><span className="font-bold text-[#1B365D]">รายงานต่อทะเบียน</span> <span className="text-natural-muted ml-2">รอบ {cycle.name} · สาขา {branchName} · {sums.length} ทะเบียน</span></div>
        <button onClick={exp} className="bg-emerald-600 text-white rounded-full px-4 py-2 text-xs font-semibold flex items-center gap-1.5"><FileSpreadsheet className="w-4 h-4" />Export Excel (.xlsx)</button>
      </div>

      {sums.length === 0 ? <EmptyHint text="ยังไม่มีข้อมูลในรอบนี้" /> : (<>
        {/* สรุปรวม */}
        <div className="bg-white rounded-2xl border border-natural-border overflow-x-auto">
          <div className="px-4 pt-3 font-bold text-[#1B365D] text-sm">📊 สรุปรวมต่อทะเบียน</div>
          <table className="w-full text-xs min-w-[760px] mt-2">
            <thead className="bg-[#1B365D] text-white"><tr><TH>ทะเบียน</TH><TH>คนขับ</TH><TH r>รายได้</TH><TH r>หัก 1%</TH><TH r>ค่าน้ำมัน</TH><TH r>+รายได้เพิ่ม</TH><TH r>รวมหัก</TH><TH r>รับสุทธิ</TH></tr></thead>
            <tbody>
              {sums.map((s: any, i: number) => (
                <tr key={s.plateNo} className={i % 2 ? 'bg-[#F9FAFC]' : ''}>
                  <TD b>{s.plateNo}</TD><TD>{s.driverName}</TD><TD r>{money(s.totalTripAmount)}</TD><TD r>{money(s.deduction1Percent)}</TD><TD r>{money(s.fuelTotal)}</TD><TD r>{money(s.incomeAdd)}</TD><TD r>{money(s.deductionTotal)}</TD>
                  <td className="px-2 py-1 text-right font-bold text-[#C00000]">{money(s.netReceive)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* รายละเอียดต่อทะเบียน */}
        {sums.map((s: any) => {
          const np = normPlate(s.plateNo);
          const vTrips = cycleTrips.filter((t: TripDocument) => normPlate(t.plateNo) === np);
          const vFuel = db.fuelEntries.filter((f: FuelEntry) => f.cycleId === cycle.id && normPlate(f.plateNo) === np);
          const vIncome = db.deductions.filter((d: DeductionEntry) => d.cycleId === cycle.id && d.kind === 'income' && normPlate(d.plateNo) === np);
          const inDocIncome = vIncome.filter((d: DeductionEntry) => (d.docNo || '').trim()).reduce((a: number, d: DeductionEntry) => a + d.amount, 0);
          const perCycleInc: { label: string; amount: number }[] = Object.values(
            vIncome.filter((d: DeductionEntry) => !(d.docNo || '').trim()).reduce((m: any, d: DeductionEntry) => {
              const k = d.label || 'รายได้เพิ่ม'; (m[k] = m[k] || { label: k, amount: 0 }).amount += d.amount; return m;
            }, {})
          );
          const dedLines = s.lines.filter((l: any) => l.kind === 'deduction');
          return (
            <div key={s.plateNo} className="bg-white rounded-2xl border border-natural-border p-4">
              <div className="font-bold text-[#1B365D] mb-2">🚚 {s.plateNo} <span className="text-natural-muted font-normal text-sm">· {s.driverName}</span></div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs min-w-[720px]">
                  <thead className="bg-[#1B365D] text-white"><tr><TH>วันที่</TH><TH>ปลายทาง</TH><TH>เลขใบกระจาย</TH><TH r>จำนวน</TH><TH>แบบ</TH><TH r>ราคา</TH><TH r>เป็นเงิน</TH><TH r>พิเศษ</TH><TH r>รวม</TH><TH>หมายเหตุ</TH></tr></thead>
                  <tbody>
                    {vTrips.map((t: TripDocument) => {
                      const piecePrice = t.receipts?.find((r) => r.piecePrice != null)?.piecePrice ?? null;
                      const unitRate = t.rateType === 'flat' ? t.rateValue : piecePrice;
                      const extra = vIncome.filter((d: DeductionEntry) => (d.docNo || '').trim() && (d.docNo || '').trim() === (t.documentNo || '').trim()).reduce((a: number, d: DeductionEntry) => a + d.amount, 0);
                      const hasDiv = t.receipts?.some((r) => r.hasAdjustment);
                      return (
                        <tr key={t.id} className={hasDiv ? 'bg-[#FFF2CC]' : ''}>
                          <TD>{fmtD(t.documentDate)}</TD><TD>{t.districtRaw ? 'อ.' + t.districtRaw : ''} {t.provinceRaw ? 'จ.' + t.provinceRaw : ''}</TD><TD>{t.documentNo}</TD>
                          <TD r>{t.billingQty}</TD><TD>{t.rateType === 'piece' ? 'ชิ้น' : t.rateType === 'flat' ? 'เหมา' : '-'}</TD>
                          <TD r>{unitRate != null ? money(unitRate) : '-'}</TD><TD r>{money(t.tripAmount)}</TD><TD r>{extra ? money(extra) : '-'}</TD>
                          <td className="px-2 py-1 text-right font-bold text-[#1B365D]">{money(t.tripAmount + extra)}</td>
                          <TD>{hasDiv ? <span className="text-[#C65911] font-semibold">มีหาร</span> : ''}</TD>
                        </tr>
                      );
                    })}
                    {vTrips.length === 0 && <tr><td colSpan={10} className="px-2 py-3 text-center text-natural-muted">ไม่มีใบกระจาย</td></tr>}
                  </tbody>
                </table>
              </div>
              {/* สรุปย่อ + น้ำมัน */}
              <div className="grid md:grid-cols-2 gap-4 mt-3">
                <div className="text-xs">
                  <div className="font-semibold text-[#1B365D] mb-1">สรุป</div>
                  <div className="flex justify-between"><span>รายได้ค่าเที่ยว</span><b>{money(s.totalTripAmount)}</b></div>
                  {inDocIncome > 0 && <div className="flex justify-between text-emerald-700"><span>+ รายได้เพิ่มในใบ (พิเศษ)</span><span>+{money(inDocIncome)}</span></div>}
                  {perCycleInc.map((l) => <div key={l.label} className="flex justify-between text-emerald-700"><span>+ {l.label}</span><span>+{money(l.amount)}</span></div>)}
                  <div className="flex justify-between text-rose-700"><span>หัก 1%</span><span>-{money(s.deduction1Percent)}</span></div>
                  <div className="flex justify-between text-rose-700"><span>หักค่าน้ำมัน</span><span>-{money(s.fuelTotal)}</span></div>
                  {dedLines.map((ln: any) => <div key={ln.categoryId} className="flex justify-between text-rose-700"><span>หัก {ln.label}</span><span>-{money(ln.amount)}</span></div>)}
                  <div className="flex justify-between border-t border-natural-border mt-1 pt-1 font-bold text-[#C00000]"><span>รวมรับสุทธิ</span><span>{money(s.netReceive)}</span></div>
                </div>
                <div className="text-xs">
                  <div className="font-semibold text-[#1B365D] mb-1">ใบสั่งเติมน้ำมัน</div>
                  {vFuel.length === 0 ? <div className="text-natural-muted">ไม่มี</div> : (
                    <table className="w-full">
                      <thead className="text-natural-muted border-b border-natural-border"><tr><th className="text-left py-0.5">วัน/เดือน/ปี</th><th className="text-left">เลขใบสั่ง</th><th className="text-right">จำนวนเงิน</th></tr></thead>
                      <tbody>
                        {vFuel.map((f: FuelEntry) => <tr key={f.id}><td className="py-0.5">{fmtD(f.date)}</td><td>{f.refNo}</td><td className="text-right">{money(f.amount)}</td></tr>)}
                        <tr className="border-t border-natural-border font-bold"><td colSpan={2} className="text-right pr-2">ผลรวม</td><td className="text-right">{money(vFuel.reduce((a: number, f: FuelEntry) => a + f.amount, 0))}</td></tr>
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            </div>
          );
        })}
      </>)}
    </div>
  );
}

// ===========================================================================
// Tab: Master ราคาขนส่ง
// ===========================================================================
function RatesTab({ db, api, branchId, cycle, reload, showToast }: any) {
  const blank = { destinationName: '', provinceName: '', provinceShort: '', districtName: '', priceType: 'flat', price: 0, pieceThreshold: '', productCategory: 'normal', minQty: '', maxQty: '', receiverKeyword: '', senderKeyword: '', productKeyword: '', rateGroup: '', effectiveFrom: '2020-01-01', effectiveTo: null, status: 'active' };
  const [form, setForm] = useState<any>(blank);
  const [sel, setSel] = useState<Set<string>>(new Set());
  const [mode, setMode] = useState<'base' | 'cycle'>('base');
  const [adv, setAdv] = useState(false);
  const [filterGroup, setFilterGroup] = useState('__all__');
  const [editId, setEditId] = useState<string | null>(null);
  const add = async () => {
    if (!form.provinceName || !form.price) return showToast('warning', 'กรอกจังหวัดและราคา');
    // เก็บคืน/Peat mass คิดเป็นชิ้นเสมอ
    const priceType = form.productCategory === 'normal' ? form.priceType : 'piece';
    const payload = {
      ...form, branchId, priceType,
      pieceThreshold: form.pieceThreshold ? +form.pieceThreshold : null,
      minQty: form.minQty ? +form.minQty : null,
      maxQty: form.maxQty ? +form.maxQty : null,
      receiverKeyword: form.receiverKeyword?.trim() || '',
      senderKeyword: form.senderKeyword?.trim() || '',
      productKeyword: form.productKeyword?.trim() || '',
      rateGroup: form.rateGroup?.trim() || '',
    };
    if (editId) {
      await api(`/api/rate-masters/${editId}`, 'PUT', payload);
      showToast('success', 'บันทึกการแก้ไขแล้ว');
    } else {
      await api('/api/rate-masters', 'POST', payload);
      showToast('success', 'เพิ่มราคาแล้ว');
    }
    setForm(blank); setEditId(null); reload();
  };
  const startEdit = async (r: RateMaster) => {
    if (!(await confirmPassword('แก้ไขราคา'))) return;
    setForm({
      destinationName: r.destinationName || '', provinceName: r.provinceName || '', provinceShort: r.provinceShort || '',
      districtName: r.districtName || '', priceType: r.priceType, price: r.price, pieceThreshold: r.pieceThreshold ?? '',
      productCategory: r.productCategory || 'normal', minQty: r.minQty ?? '', maxQty: r.maxQty ?? '',
      receiverKeyword: r.receiverKeyword || '', senderKeyword: r.senderKeyword || '', productKeyword: r.productKeyword || '',
      rateGroup: r.rateGroup || '', effectiveFrom: r.effectiveFrom, effectiveTo: r.effectiveTo, status: r.status,
    });
    setEditId(r.id); setAdv(true);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };
  const cancelEdit = () => { setForm(blank); setEditId(null); };
  const updateGroup = async (r: RateMaster, g: string) => {
    try { await api(`/api/rate-masters/${r.id}`, 'PUT', { rateGroup: g }); showToast('success', 'เปลี่ยนกลุ่มแล้ว'); reload(); }
    catch (e: any) { showToast('error', e.message); }
  };
  const branchGroups: string[] = ((db.branches as Branch[]).find((b) => b.id === branchId)?.rateGroups || []).map((g) => g.name);
  const catLabel = (c?: string) => c === 'collect_back' ? 'เก็บคืน' : c === 'peat_mass' ? 'Peat mass' : 'งานปกติ';
  if (!branchId) return <EmptyHint text={ALL_BRANCH_HINT} />;

  const allRates: RateMaster[] = db.rateMasters;
  const rates = filterGroup === '__all__' ? allRates : allRates.filter((r) => (r.rateGroup || '') === (filterGroup === '__none__' ? '' : filterGroup));
  const toggle = (id: string) => setSel((s) => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const allChecked = rates.length > 0 && rates.every((r) => sel.has(r.id));
  const toggleAll = () => setSel(allChecked ? new Set() : new Set(rates.map((r) => r.id)));
  const delOne = async (r: RateMaster) => {
    if (!(await confirmDelete(`ราคา ${r.destinationName || r.provinceName}`))) return;
    await api(`/api/rate-masters/${r.id}`, 'DELETE'); reload();
  };
  const bulkDel = async () => {
    if (!sel.size) return;
    if (!(await confirmDelete(`ราคา ${sel.size} รายการที่เลือก`))) return;
    await api('/api/rate-masters/bulk-delete', 'POST', { ids: [...sel] });
    showToast('success', `ลบ ${sel.size} รายการแล้ว`);
    setSel(new Set()); reload();
  };

  // ---- ราคาเฉพาะรอบ ----
  const cycleMode = mode === 'cycle' && cycle;
  const ovFor = (r: RateMaster): RateOverride | null =>
    (db.rateOverrides || []).find((o: RateOverride) => o.cycleId === cycle?.id && o.rateMasterId === r.id) || null;
  const effPrice = (r: RateMaster) => { const o = cycleMode ? ovFor(r) : null; return o ? o.price : r.price; };
  const effTh = (r: RateMaster) => { const o = cycleMode ? ovFor(r) : null; return o ? (o.pieceThreshold ?? null) : (r.pieceThreshold ?? null); };
  // บันทึกช่อง: โหมดหลัก -> แก้ราคาหลัก; โหมดเฉพาะรอบ -> upsert override (เก็บราคา+จุดตัดคู่กัน)
  const saveCell = async (r: RateMaster, field: 'price' | 'pieceThreshold', value: number | null) => {
    try {
      if (cycleMode) {
        const price = field === 'price' ? (value as number) : effPrice(r);
        const pieceThreshold = field === 'pieceThreshold' ? value : effTh(r);
        await api('/api/rate-overrides/upsert', 'POST', { branchId, cycleId: cycle.id, rateMasterId: r.id, price, pieceThreshold });
        showToast('success', `บันทึกราคาเฉพาะรอบ ${cycle.name}`);
      } else {
        await api(`/api/rate-masters/${r.id}`, 'PUT', { [field]: value });
        showToast('success', 'บันทึกราคาหลักแล้ว');
      }
      reload();
    } catch (e: any) { showToast('error', e.message); }
  };
  const removeOverride = async (r: RateMaster) => {
    const o = ovFor(r); if (!o) return;
    await api(`/api/rate-overrides/${o.id}`, 'DELETE');
    showToast('success', 'กลับไปใช้ราคาหลักแล้ว'); reload();
  };

  return (
    <Section title="Master ราคาขนส่ง" icon={Tag}>
      {/* สวิตช์ ราคาหลัก / ราคาเฉพาะรอบ */}
      <div className="flex flex-wrap items-center gap-2 mb-3 text-xs">
        <button onClick={() => setMode('base')} className={`px-3 py-1.5 rounded-full font-semibold border ${mode === 'base' ? 'bg-[#1B365D] text-white border-[#1B365D]' : 'border-natural-border text-natural-muted'}`}>ราคาหลัก (ทุกรอบ)</button>
        <button onClick={() => setMode('cycle')} disabled={!cycle} className={`px-3 py-1.5 rounded-full font-semibold border disabled:opacity-40 ${mode === 'cycle' ? 'bg-amber-500 text-white border-amber-500' : 'border-natural-border text-natural-muted'}`}>
          🏷️ ราคาเฉพาะรอบ{cycle ? `: ${cycle.name}` : ' (เลือกรอบก่อน)'}
        </button>
        {cycleMode && <span className="text-amber-700">แก้ราคา/จุดตัด = ทับเฉพาะรอบนี้ · รอบอื่นใช้ราคาหลัก</span>}
        {branchGroups.length > 0 && (
          <label className="flex items-center gap-1 text-xs text-natural-muted ml-auto">
            กรองกลุ่ม:
            <select aria-label="กรองกลุ่มราคา" value={filterGroup} onChange={(e) => setFilterGroup(e.target.value)} className="border border-natural-border rounded-lg px-2 py-1 text-xs font-semibold text-[#1B365D]">
              <option value="__all__">ทุกกลุ่ม</option>
              {branchGroups.map((g) => <option key={g} value={g}>{g}</option>)}
              <option value="__none__">— ไม่ระบุกลุ่ม (ใช้ร่วม) —</option>
            </select>
          </label>
        )}
      </div>
      <div className="flex flex-wrap gap-2 mb-3 text-sm items-center">
        <input aria-label="ปลายทาง" placeholder="ปลายทาง" value={form.destinationName} onChange={(e) => setForm({ ...form, destinationName: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-32" />
        <input aria-label="จังหวัด" placeholder="จังหวัด" value={form.provinceName} onChange={(e) => setForm({ ...form, provinceName: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-28" />
        <input aria-label="อำเภอ" placeholder="อำเภอ" value={form.districtName} onChange={(e) => setForm({ ...form, districtName: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-24" />
        <select aria-label="ประเภทสินค้า" value={form.productCategory} onChange={(e) => setForm({ ...form, productCategory: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5">
          <option value="normal">งานปกติ</option><option value="collect_back">เก็บคืน</option><option value="peat_mass">Peat mass</option>
        </select>
        <select aria-label="ประเภทราคา" value={form.priceType} onChange={(e) => setForm({ ...form, priceType: e.target.value })} disabled={form.productCategory !== 'normal'} className="border border-natural-border rounded-lg px-2 py-1.5 disabled:opacity-50">
          <option value="flat">ราคาเหมา</option><option value="piece">ราคาชิ้น</option>
        </select>
        <input type="number" aria-label="ราคา" placeholder="ราคา" value={form.price || ''} onChange={(e) => setForm({ ...form, price: +e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-24" />
        <input type="number" aria-label="จุดตัดชิ้น" placeholder="จุดตัดชิ้น" title="≤จุดตัด=เหมา, >จุดตัด=ชิ้น (เว้นว่าง=ไม่ใช้)" value={form.pieceThreshold} onChange={(e) => setForm({ ...form, pieceThreshold: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-24" />
        <button onClick={add} className={`${editId ? 'bg-amber-600' : 'bg-[#1B365D]'} text-white rounded-lg px-3 py-1.5 font-semibold`}>{editId ? '✎ บันทึกแก้ไข' : 'เพิ่ม'}</button>
        {editId && <button onClick={cancelEdit} className="border border-natural-border rounded-lg px-3 py-1.5 text-sm font-semibold">ยกเลิก</button>}
        <button onClick={() => setAdv(!adv)} className="text-xs text-[#1B365D] font-semibold underline">{adv ? 'ซ่อน' : 'เงื่อนไขพิเศษ'}</button>
        {sel.size > 0 && (
          <button onClick={bulkDel} className="bg-red-600 text-white rounded-lg px-3 py-1.5 font-semibold flex items-center gap-1 ml-auto">
            <Trash2 className="w-4 h-4" />ลบที่เลือก ({sel.size})
          </button>
        )}
      </div>
      {adv && (
        <div className="flex flex-wrap gap-2 mb-3 text-sm items-center bg-amber-50 border border-amber-200 rounded-lg p-2">
          <span className="text-[11px] text-amber-800 font-semibold">เงื่อนไขพิเศษ (เว้นว่าง=ไม่ใช้):</span>
          {branchGroups.length > 0 ? (
            <select aria-label="กลุ่มราคา" value={form.rateGroup} onChange={(e) => setForm({ ...form, rateGroup: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5">
              <option value="">ทุกกลุ่ม</option>
              {branchGroups.map((g) => <option key={g} value={g}>{g}</option>)}
            </select>
          ) : (
            <input aria-label="กลุ่มราคา" placeholder="กลุ่มราคา" value={form.rateGroup} onChange={(e) => setForm({ ...form, rateGroup: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-24" />
          )}
          <input aria-label="ผู้รับ" placeholder="เฉพาะผู้รับ" value={form.receiverKeyword} onChange={(e) => setForm({ ...form, receiverKeyword: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-28" />
          <input aria-label="ผู้ส่ง" placeholder="เฉพาะผู้ส่ง" value={form.senderKeyword} onChange={(e) => setForm({ ...form, senderKeyword: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-28" />
          <input aria-label="สินค้า" placeholder="เฉพาะสินค้า" value={form.productKeyword} onChange={(e) => setForm({ ...form, productKeyword: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-28" />
          <input type="number" aria-label="กล่องต่ำสุด" placeholder="กล่อง≥" value={form.minQty} onChange={(e) => setForm({ ...form, minQty: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-20" />
          <input type="number" aria-label="กล่องสูงสุด" placeholder="กล่อง≤" value={form.maxQty} onChange={(e) => setForm({ ...form, maxQty: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-20" />
        </div>
      )}
      <div className="overflow-x-auto">
        <table className="w-full text-xs">
          <thead>
            <tr className="text-natural-muted text-left border-b border-natural-border">
              <th className="w-8 py-1.5 px-1"><input type="checkbox" aria-label="เลือกทั้งหมด" checked={allChecked} onChange={toggleAll} /></th>
              {branchGroups.length > 0 && <th className="py-1.5 px-1">กลุ่ม</th>}
              {['ปลายทาง', 'จังหวัด', 'อำเภอ', 'หมวด', 'ประเภท', 'ราคา', 'จุดตัดชิ้น', 'เริ่มใช้'].map((c) => <th key={c} className="py-1.5 px-1">{c}</th>)}
              <th className="w-8"></th>
            </tr>
          </thead>
          <tbody>
            {rates.map((r) => (
              <tr key={r.id} className={`border-b border-natural-border/60 ${sel.has(r.id) ? 'bg-red-50' : ''}`}>
                <td className="py-1.5 px-1"><input type="checkbox" aria-label={`เลือก ${r.destinationName}`} checked={sel.has(r.id)} onChange={() => toggle(r.id)} /></td>
                {branchGroups.length > 0 && (
                  <td className="py-1.5 px-1">
                    <select aria-label={`กลุ่ม ${r.destinationName}`} value={r.rateGroup || ''} onChange={(e) => updateGroup(r, e.target.value)}
                      className="border border-natural-border rounded px-1 py-0.5 text-xs font-semibold text-[#1B365D]">
                      <option value="">ใช้ร่วม</option>
                      {branchGroups.map((g) => <option key={g} value={g}>{g}</option>)}
                    </select>
                  </td>
                )}
                <td className="py-1.5 px-1 font-semibold text-[#1B365D]">{r.destinationName}</td>
                <td className="py-1.5 px-1">{r.provinceName}</td>
                <td className="py-1.5 px-1">{r.districtName}</td>
                <td className="py-1.5 px-1">{(r.productCategory && r.productCategory !== 'normal') ? <span className="text-amber-700 font-semibold">{catLabel(r.productCategory)}</span> : <span className="text-natural-muted">ปกติ</span>}</td>
                <td className="py-1.5 px-1">{r.priceType === 'flat' ? 'เหมา' : 'ชิ้น'}</td>
                <td className="py-1.5 px-1">
                  <input type="number" key={`p-${r.id}-${cycleMode ? 'c' : 'b'}-${effPrice(r)}`} defaultValue={effPrice(r)} aria-label={`ราคา ${r.destinationName}`}
                    onBlur={(e) => { const v = +e.target.value; if (v !== effPrice(r)) saveCell(r, 'price', v); }}
                    className={`w-20 border rounded px-1 py-0.5 text-xs text-right outline-none ${cycleMode && ovFor(r) ? 'border-amber-400 bg-amber-50' : 'border-natural-border'} focus:border-[#1B365D]`} />
                  {cycleMode && ovFor(r) && <span className="text-[9px] text-amber-700 ml-0.5">เฉพาะรอบ</span>}
                </td>
                <td className="py-1.5 px-1">
                  <input type="number" key={`t-${r.id}-${cycleMode ? 'c' : 'b'}-${effTh(r) ?? ''}`} defaultValue={effTh(r) ?? ''} placeholder="-" aria-label={`จุดตัด ${r.destinationName}`}
                    onBlur={(e) => { const raw = e.target.value.trim(); const v = raw === '' ? null : +raw; if (v !== effTh(r)) saveCell(r, 'pieceThreshold', v); }}
                    className={`w-16 border rounded px-1 py-0.5 text-xs text-right outline-none ${cycleMode && ovFor(r) ? 'border-amber-400 bg-amber-50' : 'border-natural-border'} focus:border-[#1B365D]`} />
                </td>
                <td className="py-1.5 px-1">{r.effectiveFrom}</td>
                <td className="py-1.5 px-1 whitespace-nowrap">
                  {cycleMode && ovFor(r) && <button type="button" title="กลับไปใช้ราคาหลัก" onClick={() => removeOverride(r)} className="text-amber-600 hover:text-amber-800 mr-1"><RefreshCw className="w-3.5 h-3.5 inline" /></button>}
                  <button type="button" title="แก้ไข (ใส่รหัส)" onClick={() => startEdit(r)} className="text-[#1B365D] hover:text-amber-700 font-semibold mr-2">✎</button>
                  <button type="button" title="ลบ" onClick={() => delOne(r)} className="text-red-400 hover:text-red-600"><Trash2 className="w-3.5 h-3.5 inline" /></button>
                </td>
              </tr>
            ))}
            {rates.length === 0 && <tr><td colSpan={branchGroups.length > 0 ? 11 : 10} className="py-6 text-center text-natural-muted">ไม่มีราคา (ลองเปลี่ยนตัวกรองกลุ่ม)</td></tr>}
          </tbody>
        </table>
      </div>
    </Section>
  );
}

// ===========================================================================
// Tab: เงื่อนไขตัวหาร + กลุ่มผู้รับ
// ===========================================================================
function RulesTab({ db, api, branchId, reload, showToast }: any) {
  const blank = { ruleName: '', senderKeyword: 'ซีโน', receiverGroupId: db.receiverGroups[0]?.id || '', productKeyword: '', productSizeKeyword: '', divisor: 3, roundingMethod: 'half_up', applyLevel: 'receipt', status: 'active', effectiveFrom: '2020-01-01', effectiveTo: null };
  const [form, setForm] = useState<any>(blank);
  const [fSender, setFSender] = useState('');
  const [fGroup, setFGroup] = useState('');
  const add = async () => {
    if (!form.productKeyword) return showToast('warning', 'กรอกชื่อสินค้า');
    await api('/api/conversion-rules', 'POST', { ...form, branchId, ruleName: form.ruleName || `${form.productKeyword} หาร ${form.divisor}` });
    setForm(blank); reload(); showToast('success', 'เพิ่มกฎแล้ว');
  };
  const norm = (s: string) => (s || '').toLowerCase().replace(/\s+/g, '');
  const filteredRules: ProductConversionRule[] = db.conversionRules.filter((r: ProductConversionRule) =>
    (!fSender || norm(r.senderKeyword).includes(norm(fSender))) &&
    (!fGroup || r.receiverGroupId === fGroup)
  );
  // ค่าที่เคยใช้ (สำหรับ datalist combobox)
  const uniq = (arr: string[]) => [...new Set(arr.filter(Boolean))];
  const senderOpts = uniq(db.conversionRules.map((r: ProductConversionRule) => r.senderKeyword));
  const productOpts = uniq(db.conversionRules.map((r: ProductConversionRule) => r.productKeyword));
  const sizeOpts = uniq(db.conversionRules.map((r: ProductConversionRule) => r.productSizeKeyword));
  if (!branchId) return <EmptyHint text={ALL_BRANCH_HINT} />;
  return (
    <div className="flex flex-col gap-5">
      <Section title="เงื่อนไขแปลงจำนวนสินค้า (ตัวหาร)" icon={Filter}>
        <p className="text-xs text-natural-muted mb-3">ต้องตรงทุกข้อ: ผู้ส่งเข้าคำว่า "ซีโน" + กลุ่มผู้รับตรง + ชื่อสินค้า + ขนาด → หารตามตัวหาร (คำนวณแยกตามเลขใบรับสินค้า)</p>
        <div className="flex flex-wrap gap-2 mb-3 text-sm">
          <input list="rule-senders" aria-label="ผู้ส่ง (keyword)" placeholder="ผู้ส่ง (พิมพ์ใหม่ได้)" value={form.senderKeyword} onChange={(e) => setForm({ ...form, senderKeyword: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-32" />
          <select aria-label="กลุ่มผู้รับสินค้า" value={form.receiverGroupId} onChange={(e) => setForm({ ...form, receiverGroupId: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5">
            {db.receiverGroups.map((g: ReceiverGroup) => <option key={g.id} value={g.id}>{g.groupName}</option>)}
          </select>
          <input list="rule-products" aria-label="ชื่อสินค้า" placeholder="สินค้า (พิมพ์ใหม่ได้)" value={form.productKeyword} onChange={(e) => setForm({ ...form, productKeyword: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-32" />
          <input list="rule-sizes" aria-label="ขนาดสินค้า" placeholder="ขนาด เช่น 14 กรัม" value={form.productSizeKeyword} onChange={(e) => setForm({ ...form, productSizeKeyword: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-28" />
          <input type="number" aria-label="ตัวหาร" placeholder="ตัวหาร" value={form.divisor} onChange={(e) => setForm({ ...form, divisor: +e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-20" />
          <button onClick={add} className="bg-[#1B365D] text-white rounded-lg px-3 font-semibold">เพิ่ม</button>
          <datalist id="rule-senders">{senderOpts.map((s) => <option key={s} value={s} />)}</datalist>
          <datalist id="rule-products">{productOpts.map((s) => <option key={s} value={s} />)}</datalist>
          <datalist id="rule-sizes">{sizeOpts.map((s) => <option key={s} value={s} />)}</datalist>
        </div>

        {/* ตัวกรอง: ผู้ส่ง / กลุ่มผู้รับ */}
        <div className="flex flex-wrap items-center gap-2 mb-2 text-sm bg-natural-secondary/60 border border-natural-border rounded-lg p-2">
          <span className="text-natural-muted font-semibold flex items-center gap-1"><Filter className="w-3.5 h-3.5" />กรอง:</span>
          <input aria-label="กรองผู้ส่ง" placeholder="ค้นหาผู้ส่ง" value={fSender} onChange={(e) => setFSender(e.target.value)} className="border border-natural-border rounded-lg px-2 py-1 w-32" />
          <select aria-label="กรองกลุ่มผู้รับ" value={fGroup} onChange={(e) => setFGroup(e.target.value)} className="border border-natural-border rounded-lg px-2 py-1">
            <option value="">ทุกกลุ่มผู้รับ</option>
            {db.receiverGroups.map((g: ReceiverGroup) => <option key={g.id} value={g.id}>{g.groupName}</option>)}
          </select>
          {(fSender || fGroup) && (
            <button onClick={() => { setFSender(''); setFGroup(''); }} className="text-xs text-natural-muted hover:text-rose-600 font-semibold underline">ล้างตัวกรอง</button>
          )}
          <span className="text-xs text-natural-muted ml-auto">{filteredRules.length}/{db.conversionRules.length} กฎ</span>
        </div>

        <SimpleTable cols={['ผู้ส่ง', 'กลุ่มผู้รับ', 'สินค้า', 'ขนาด', 'หาร', 'ปัดเศษ']}
          rows={filteredRules.map((r: ProductConversionRule) => [r.senderKeyword, db.receiverGroups.find((g: ReceiverGroup) => g.id === r.receiverGroupId)?.groupName || '-', r.productKeyword, r.productSizeKeyword, `÷${r.divisor}`, r.roundingMethod === 'half_up' ? '.5 ปัดขึ้น' : r.roundingMethod])}
          onDelete={async (i: number) => { await api(`/api/conversion-rules/${filteredRules[i].id}`, 'DELETE'); reload(); }} />
      </Section>
      <GroupManager db={db} api={api} branchId={branchId} reload={reload} showToast={showToast} />
      <ManualBoxSenderManager db={db} api={api} branchId={branchId} reload={reload} showToast={showToast} />
    </div>
  );
}

// จัดการผู้ส่งที่ส่งเป็นชิ้น (ต้องกรอกจำนวนกล่องเอง)
function ManualBoxSenderManager({ db, api, branchId, reload, showToast }: any) {
  const [form, setForm] = useState({ senderKeyword: '', note: '' });
  const add = async () => {
    if (!form.senderKeyword.trim()) return showToast('warning', 'กรอกคำในชื่อผู้ส่ง');
    await api('/api/manual-box-senders', 'POST', { senderKeyword: form.senderKeyword.trim(), note: form.note.trim(), status: 'active', branchId });
    setForm({ senderKeyword: '', note: '' }); reload(); showToast('success', 'เพิ่มแล้ว');
  };
  return (
    <Section title="ผู้ส่งที่ส่งเป็นชิ้น (ต้องกรอกจำนวนกล่องเอง)" icon={Filter}>
      <p className="text-xs text-natural-muted mb-3">เมื่อใบรับสินค้ามีผู้ส่งตรงคำที่ระบุ ระบบจะบังคับให้กรอก "จำนวนกล่อง" ในหน้า Review (ใช้แทนจำนวนชิ้นที่อ่านได้) — บันทึกไม่ได้จนกว่าจะกรอก</p>
      <div className="flex flex-wrap gap-2 mb-3 text-sm">
        <input aria-label="คำในชื่อผู้ส่ง" placeholder="คำในชื่อผู้ส่ง เช่น คอนซูเมอร์" value={form.senderKeyword} onChange={(e) => setForm({ ...form, senderKeyword: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-48" />
        <input aria-label="หมายเหตุ" placeholder="หมายเหตุ (ถ้ามี)" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-48" />
        <button onClick={add} className="bg-[#1B365D] text-white rounded-lg px-3 font-semibold">เพิ่ม</button>
      </div>
      <SimpleTable cols={['คำในชื่อผู้ส่ง', 'หมายเหตุ']}
        rows={db.manualBoxSenders.map((s: ManualBoxSender) => [s.senderKeyword, s.note || ''])}
        onDelete={async (i: number) => { await api(`/api/manual-box-senders/${db.manualBoxSenders[i].id}`, 'DELETE'); reload(); }} />
    </Section>
  );
}

// จัดการกลุ่มผู้รับสินค้า + ชื่อพ้อง (alias) — เพิ่ม "ชื่อผู้รับใหม่" ได้เอง
function GroupManager({ db, api, branchId, reload, showToast }: any) {
  const [newGroup, setNewGroup] = useState('');
  const [aliasInputs, setAliasInputs] = useState<Record<string, string>>({});

  const addGroup = async () => {
    if (!newGroup.trim()) return showToast('warning', 'กรอกชื่อกลุ่ม');
    await api('/api/receiver-groups', 'POST', { groupName: newGroup.trim(), status: 'active', branchId });
    setNewGroup(''); reload(); showToast('success', 'เพิ่มกลุ่มแล้ว — ใช้ใน dropdown ได้เลย');
  };
  const delGroup = async (g: ReceiverGroup) => {
    if (!(await confirmDelete(g.groupName))) return;
    await api(`/api/receiver-groups/${g.id}`, 'DELETE'); reload();
  };
  const addAlias = async (groupId: string) => {
    const name = (aliasInputs[groupId] || '').trim();
    if (!name) return showToast('warning', 'กรอกชื่อพ้อง (ชื่อผู้รับที่ปรากฏใน PDF)');
    await api('/api/receiver-aliases', 'POST', { receiverGroupId: groupId, aliasName: name, status: 'active', branchId });
    setAliasInputs({ ...aliasInputs, [groupId]: '' }); reload();
  };
  const delAlias = async (a: ReceiverGroupAlias) => {
    if (!(await confirmDelete(a.aliasName))) return;
    await api(`/api/receiver-aliases/${a.id}`, 'DELETE'); reload();
  };

  return (
    <Section title="กลุ่มผู้รับสินค้า & ชื่อพ้อง (alias)" icon={Tag}>
      <p className="text-xs text-natural-muted mb-3">"ชื่อพ้อง" คือชื่อผู้รับที่ปรากฏจริงใน PDF (เช่น แม็คโคร, MK, CP AXTRA) — ระบบใช้จับว่าใบรับนั้นอยู่กลุ่มไหน · เพิ่มกลุ่มใหม่ที่นี่แล้วจะไปโผล่ใน dropdown "กลุ่มผู้รับ" อัตโนมัติ</p>

      {/* เพิ่มกลุ่มใหม่ */}
      <div className="flex flex-wrap gap-2 mb-4">
        <input aria-label="ชื่อกลุ่มผู้รับใหม่" placeholder="ชื่อกลุ่มใหม่ เช่น โลตัส" value={newGroup} onChange={(e) => setNewGroup(e.target.value)} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-44" />
        <button onClick={addGroup} className="bg-[#1B365D] text-white rounded-lg px-3 text-sm font-semibold">+ เพิ่มกลุ่ม</button>
      </div>

      <div className="flex flex-col gap-3">
        {db.receiverGroups.map((g: ReceiverGroup) => (
          <div key={g.id} className="border border-natural-border rounded-xl p-3">
            <div className="flex items-center justify-between mb-2">
              <span className="font-semibold text-sm text-[#1B365D]">{g.groupName}</span>
              <button type="button" aria-label="ลบกลุ่ม" title="ลบกลุ่ม" onClick={() => delGroup(g)} className="text-natural-muted hover:text-rose-600"><Trash2 className="w-3.5 h-3.5" /></button>
            </div>
            <div className="flex flex-wrap gap-1.5 items-center mb-2">
              {db.receiverGroupAliases.filter((a: ReceiverGroupAlias) => a.receiverGroupId === g.id).map((a: ReceiverGroupAlias) => (
                <span key={a.id} className="bg-natural-secondary border border-natural-border rounded-full pl-2 pr-1 py-0.5 text-xs flex items-center gap-1">
                  {a.aliasName}
                  <button type="button" aria-label={`ลบ ${a.aliasName}`} title="ลบชื่อพ้อง" onClick={() => delAlias(a)} className="text-natural-muted hover:text-rose-600">×</button>
                </span>
              ))}
            </div>
            <div className="flex gap-2">
              <input aria-label={`เพิ่มชื่อพ้องให้ ${g.groupName}`} placeholder="+ ชื่อผู้รับใหม่ในกลุ่มนี้" value={aliasInputs[g.id] || ''}
                onChange={(e) => setAliasInputs({ ...aliasInputs, [g.id]: e.target.value })}
                onKeyDown={(e) => { if (e.key === 'Enter') addAlias(g.id); }}
                className="border border-natural-border rounded-lg px-2 py-1 text-xs w-52" />
              <button onClick={() => addAlias(g.id)} className="border border-natural-border rounded-lg px-2.5 py-1 text-xs font-semibold">เพิ่มชื่อพ้อง</button>
            </div>
          </div>
        ))}
      </div>
    </Section>
  );
}

// ===========================================================================
// Tab: รถ & คนขับ
// ===========================================================================
function VehiclesTab({ db, api, branchId, reload, showToast }: any) {
  const blankV = { plateNo: '', driverName: '', vehicleType: '6 ล้อ', rateGroup: '', status: 'active' };
  const [form, setForm] = useState<any>(blankV);
  const [sel, setSel] = useState<Set<string>>(new Set());
  const [editId, setEditId] = useState<string | null>(null);
  const branchGroups: string[] = ((db.branches as Branch[]).find((b) => b.id === branchId)?.rateGroups || []).map((g) => g.name);
  const add = async () => {
    if (!form.plateNo) return showToast('warning', 'กรอกทะเบียน');
    if (editId) { await api(`/api/vehicles/${editId}`, 'PUT', form); showToast('success', 'บันทึกการแก้ไขแล้ว'); }
    else { await api('/api/vehicles', 'POST', { ...form, branchId }); showToast('success', 'เพิ่มรถแล้ว'); }
    setForm(blankV); setEditId(null); reload();
  };
  const startEdit = async (v: Vehicle) => {
    if (!(await confirmPassword('แก้ไขรถ'))) return;
    setForm({ plateNo: v.plateNo, driverName: v.driverName, vehicleType: v.vehicleType, rateGroup: v.rateGroup || '', status: v.status });
    setEditId(v.id);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };
  const cancelEdit = () => { setForm(blankV); setEditId(null); };
  const setGroup = async (v: Vehicle, g: string) => { await api(`/api/vehicles/${v.id}`, 'PUT', { rateGroup: g }); reload(); };
  if (!branchId) return <EmptyHint text={ALL_BRANCH_HINT} />;

  const vehicles: Vehicle[] = db.vehicles;
  const toggle = (id: string) => setSel((s) => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const allChecked = vehicles.length > 0 && vehicles.every((v) => sel.has(v.id));
  const toggleAll = () => setSel(allChecked ? new Set() : new Set(vehicles.map((v) => v.id)));
  const delOne = async (v: Vehicle) => {
    if (!(await confirmDelete(`รถ ${v.plateNo}`))) return;
    await api(`/api/vehicles/${v.id}`, 'DELETE'); reload();
  };
  const bulkDel = async () => {
    if (!sel.size) return;
    if (!(await confirmDelete(`รถ ${sel.size} คันที่เลือก`))) return;
    await api('/api/vehicles/bulk-delete', 'POST', { ids: [...sel] });
    showToast('success', `ลบ ${sel.size} รายการแล้ว`);
    setSel(new Set()); reload();
  };

  return (
    <Section title="Master รถร่วม & คนขับ" icon={Truck}>
      <div className="flex flex-wrap gap-2 mb-3 text-sm items-center">
        <input aria-label="ทะเบียนรถ" placeholder="ทะเบียน" value={form.plateNo} onChange={(e) => setForm({ ...form, plateNo: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-28" />
        <input aria-label="ชื่อคนขับ" placeholder="คนขับ" value={form.driverName} onChange={(e) => setForm({ ...form, driverName: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-36" />
        <input aria-label="ประเภทรถ" placeholder="ประเภทรถ" value={form.vehicleType} onChange={(e) => setForm({ ...form, vehicleType: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-24" />
        {branchGroups.length > 0 && (
          <select aria-label="กลุ่มราคา" value={form.rateGroup} onChange={(e) => setForm({ ...form, rateGroup: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5">
            <option value="">กลุ่มราคา?</option>
            {branchGroups.map((g) => <option key={g} value={g}>{g}</option>)}
          </select>
        )}
        <button onClick={add} className={`${editId ? 'bg-amber-600' : 'bg-[#1B365D]'} text-white rounded-lg px-3 py-1.5 font-semibold`}>{editId ? '✎ บันทึกแก้ไข' : 'เพิ่ม'}</button>
        {editId && <button onClick={cancelEdit} className="border border-natural-border rounded-lg px-3 py-1.5 font-semibold">ยกเลิก</button>}
        {sel.size > 0 && (
          <button onClick={bulkDel} className="bg-red-600 text-white rounded-lg px-3 py-1.5 font-semibold flex items-center gap-1 ml-auto">
            <Trash2 className="w-4 h-4" />ลบที่เลือก ({sel.size})
          </button>
        )}
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-xs">
          <thead>
            <tr className="text-natural-muted text-left border-b border-natural-border">
              <th className="w-8 py-1.5 px-1"><input type="checkbox" aria-label="เลือกทั้งหมด" checked={allChecked} onChange={toggleAll} /></th>
              <th className="py-1.5 px-1">ทะเบียน</th><th className="py-1.5 px-1">คนขับ</th><th className="py-1.5 px-1">ประเภท</th>{branchGroups.length > 0 && <th className="py-1.5 px-1">กลุ่มราคา</th>}<th className="w-8"></th>
            </tr>
          </thead>
          <tbody>
            {vehicles.map((v) => (
              <tr key={v.id} className={`border-b border-natural-border/60 ${sel.has(v.id) ? 'bg-red-50' : ''}`}>
                <td className="py-1.5 px-1"><input type="checkbox" aria-label={`เลือก ${v.plateNo}`} checked={sel.has(v.id)} onChange={() => toggle(v.id)} /></td>
                <td className="py-1.5 px-1 font-semibold text-[#1B365D]">{v.plateNo}</td>
                <td className="py-1.5 px-1">{v.driverName}</td>
                <td className="py-1.5 px-1">{v.vehicleType}</td>
                {branchGroups.length > 0 && (
                  <td className="py-1.5 px-1">
                    <select aria-label={`กลุ่ม ${v.plateNo}`} value={v.rateGroup || ''} onChange={(e) => setGroup(v, e.target.value)} className="border border-natural-border rounded px-1 py-0.5 text-xs">
                      <option value="">—</option>
                      {branchGroups.map((g) => <option key={g} value={g}>{g}</option>)}
                    </select>
                  </td>
                )}
                <td className="py-1.5 px-1 whitespace-nowrap">
                  <button type="button" title="แก้ไข (ใส่รหัส)" onClick={() => startEdit(v)} className="text-[#1B365D] hover:text-amber-700 font-semibold mr-2">✎</button>
                  <button type="button" title="ลบ" onClick={() => delOne(v)} className="text-red-400 hover:text-red-600"><Trash2 className="w-3.5 h-3.5 inline" /></button>
                </td>
              </tr>
            ))}
            {vehicles.length === 0 && <tr><td colSpan={6} className="py-6 text-center text-natural-muted">ยังไม่มีรถในสาขานี้</td></tr>}
          </tbody>
        </table>
      </div>
    </Section>
  );
}

// ===========================================================================
// Shared small components
// ===========================================================================
function Field({ label, value, onChange, type = 'text' }: { label: string; value: string; onChange: (v: string) => void; type?: string }) {
  return (
    <div className="flex flex-col gap-0.5">
      <span className="text-natural-muted text-[10px] font-bold uppercase">{label}</span>
      <input type={type} value={value} aria-label={label} onChange={(e) => onChange(e.target.value)} className="border border-natural-border rounded-lg px-2 py-1.5 focus:border-[#1B365D] outline-none" />
    </div>
  );
}
function Section({ title, icon: Icon, children }: any) {
  return (
    <div className="bg-white rounded-2xl border border-natural-border p-5">
      <h3 className="font-bold text-sm text-[#1B365D] flex items-center gap-1.5 mb-3"><Icon className="w-4 h-4" />{title}</h3>
      {children}
    </div>
  );
}
function Stat({ label, value, highlight }: { label: string; value: string; highlight?: boolean }) {
  return (
    <div className={`rounded-2xl border p-4 text-center ${highlight ? 'bg-[#E6EEF8] border-[#1B365D]' : 'bg-white border-natural-border'}`}>
      <div className="text-[10px] uppercase font-bold text-natural-muted">{label}</div>
      <div className={`text-xl font-bold mt-1 ${highlight ? 'text-[#C00000]' : 'text-[#1B365D]'}`}>{value}</div>
    </div>
  );
}
function SimpleTable({ cols, rows, onDelete }: { cols: string[]; rows: any[][]; onDelete?: (i: number) => void }) {
  return (
    <div className="overflow-x-auto">
      <table className="w-full text-xs">
        <thead><tr className="text-natural-muted text-left border-b border-natural-border">{cols.map((c) => <th key={c} className="py-1.5 px-1">{c}</th>)}{onDelete && <th className="w-8"></th>}</tr></thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} className={i % 2 ? 'bg-[#F9FAFC]' : ''}>
              {row.map((cell, j) => <td key={j} className="py-1.5 px-1">{cell}</td>)}
              {onDelete && <td className="text-center"><button type="button" aria-label="ลบรายการ" title="ลบรายการ" onClick={async () => { if (await confirmDelete()) onDelete(i); }} className="text-natural-muted hover:text-rose-600"><Trash2 className="w-3.5 h-3.5" /></button></td>}
            </tr>
          ))}
          {rows.length === 0 && <tr><td colSpan={cols.length + 1} className="py-4 text-center text-natural-muted">ยังไม่มีข้อมูล</td></tr>}
        </tbody>
      </table>
    </div>
  );
}
function EmptyHint({ text }: { text: string }) {
  return <div className="bg-white rounded-2xl border border-natural-border p-10 text-center text-natural-muted text-sm">{text}</div>;
}
