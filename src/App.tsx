import React, { useState, useEffect, useRef, Fragment } from 'react';
import {
  UploadCloud, AlertTriangle, FileSpreadsheet, Trash2, Plus, Save,
  RefreshCw, Lock, Unlock, Database, Truck, Tag, Filter, Calculator, Fuel, Receipt, Coins,
  Building2, LogOut, Search, Calendar, Menu, X, ChevronsLeft, ChevronsRight, TrendingUp, MapPin,
} from 'lucide-react';
import {
  DatabaseState, BillingCycle, Branch, Vehicle, RateMaster, RateOverride, ReceiverGroup, ReceiverGroupAlias,
  ProductConversionRule, TripDocument, TripReceipt, FuelEntry, DeductionEntry, ExtractedTripDocument, MoneyCategory, ManualBoxSender, DestinationOverride,
} from './types';
import { exportCycleToExcel, exportPerVehicleReport, downloadRateTemplate, downloadFuelTemplate, exportBranchSummary, tripSubRows, exportDriverKpi, exportCostAreas } from './excel-export';
import { summarizeByVehicle, isUnspecifiedName, normPlate, normDoc } from './calc';
import { confirmDelete, confirmAction, confirmPassword, notify, alertBox } from './ui';

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
  receiverGroupAliases: [], conversionRules: [], manualBoxSenders: [], destinationOverrides: [], moneyCategories: [], tripDocuments: [], fuelEntries: [], deductions: [],
};

// ---- พื้นที่ให้บริการของสาขา (ตรวจตอนบันทึก) ----
const _normP = (s: string) => (s || '').replace(/\s/g, '');
function parseServiceAreas(text?: string): { prov: string; dists: string[] | null }[] {
  return (text || '').split('\n').map((l) => l.trim()).filter(Boolean).map((line) => {
    const idx = line.search(/[:：]/);
    if (idx < 0) return { prov: line.trim(), dists: null };
    const dists = line.slice(idx + 1).split(/[,，]/).map((s) => s.trim()).filter(Boolean);
    return { prov: line.slice(0, idx).trim(), dists: dists.length ? dists : null };
  });
}
function inServiceArea(areas: { prov: string; dists: string[] | null }[], prov: string, dist: string): boolean {
  if (!areas.length) return true; // ไม่ตั้งพื้นที่ = ไม่ตรวจ
  const np = _normP(prov), nd = _normP(dist);
  for (const a of areas) {
    const ap = _normP(a.prov);
    if (np && ap && (np.includes(ap) || ap.includes(np))) {
      if (!a.dists) return true; // ทั้งจังหวัด
      return a.dists.some((d) => { const ad = _normP(d); return nd && ad && (nd.includes(ad) || ad.includes(nd)); });
    }
  }
  return false;
}

type BranchAuth = { id: string; name: string; isHQ: boolean };
type Tab = 'calc' | 'rates' | 'rules' | 'vehicles' | 'fuel' | 'dashboard' | 'branches' | 'reports' | 'driverkpi' | 'costarea' | 'destfix';
type Toast = { type: 'success' | 'error' | 'warning'; message: string };

const THAI_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const money = (n: number) => n.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
// จำนวน: ตัดทศนิยมที่ 2 ตำแหน่ง (ไม่ปัดขึ้น) กัน float error เช่น 1.9999999998 -> 1.99
const qtyFmt = (n: number) => String(Math.floor(((n ?? 0) + 1e-9) * 100) / 100);

export default function App() {
  const [db, setDb] = useState<DatabaseState>(EMPTY);
  const [loading, setLoading] = useState(true);
  const [booted, setBooted] = useState(false); // โหลดสำเร็จครั้งแรกแล้ว — reload หลังจากนี้ไม่ขึ้นสปินเนอร์เต็มจอ (กันแท็บ remount/รีเซ็ตสถานะ)
  const [selectedCycleId, setSelectedCycleId] = useState('');
  const [tab, setTab] = useState<Tab>('calc');
  const [navOpen, setNavOpen] = useState(false);
  const [navCollapsed, setNavCollapsed] = useState(false);
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
      setBooted(true);
    }
  };

  // โหลด config + branches เสมอ (สำหรับหน้า login)
  useEffect(() => {
    fetch('/api/config').then((r) => r.json()).then((c) => setAiEnabled(!!c.aiEnabled)).catch(() => setAiEnabled(false));
  }, []);

  // โหลดข้อมูลตามสาขาที่เลือก (HQ workBranchId='' = ทุกสาขา)
  useEffect(() => {
    if (!auth) { fetch('/api/state').then((r) => r.json()).then(setDb).finally(() => { setLoading(false); setBooted(true); }); return; }
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
    ['costarea', 'พื้นที่ต้นทุนสูง', MapPin],
    ['rates', 'Master ราคาขนส่ง', Tag],
    ['rules', 'เงื่อนไขตัวหาร', Filter],
    ['destfix', 'แก้ปลายทาง 📍', MapPin],
    ['vehicles', 'รถ & คนขับ', Truck],
  ];
  if (auth.isHQ) tabs.push(['driverkpi', 'วิเคราะห์รายได้ พขร 🔒', TrendingUp]);
  if (auth.isHQ) tabs.push(['branches', 'จัดการสาขา', Building2]);

  const activeTabLabel = tabs.find(([k]) => k === tab)?.[1] || '';

  return (
    <div className="min-h-screen bg-natural-bg text-natural-text font-sans flex">
      {/* Mobile overlay */}
      {navOpen && <div className="fixed inset-0 bg-black/50 z-40 md:hidden" onClick={() => setNavOpen(false)} />}

      {/* ===== Sidebar ===== */}
      <aside className={`fixed md:sticky top-0 z-50 md:z-10 h-screen w-60 ${navCollapsed ? 'md:w-16' : 'md:w-60'} shrink-0 bg-brand-navy text-white flex flex-col transition-all duration-200 ${navOpen ? 'translate-x-0' : '-translate-x-full md:translate-x-0'}`}>
        <div className={`py-4 border-b border-white/10 flex items-center gap-2.5 ${navCollapsed ? 'md:px-3 md:justify-center px-5' : 'px-5'}`}>
          <img src="/iconneo.png" alt="NEOSIAM" className="w-10 h-10 rounded-lg object-cover shadow-md shrink-0" />
          <div className={`leading-tight text-center ${navCollapsed ? 'md:hidden' : ''}`}>
            <div className="brand-wordmark text-xl">NEOSIAM</div>
            <div className="text-[10px] text-brand-gold font-semibold tracking-wide">ส่งด่วน · ส่งไว · แน่นอน</div>
          </div>
          <button type="button" aria-label="ปิดเมนู" onClick={() => setNavOpen(false)} className="ml-auto md:hidden text-white/70 hover:text-white"><X className="w-5 h-5" /></button>
        </div>
        <nav className="flex-1 overflow-y-auto py-3 px-2 flex flex-col gap-0.5">
          {tabs.map(([key, label, Icon]) => (
            <button key={key} onClick={() => { setTab(key); setNavOpen(false); }} title={navCollapsed ? label : ''}
              className={`relative flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-semibold text-left transition ${navCollapsed ? 'md:justify-center md:px-0' : ''} ${
                tab === key ? 'bg-white/10 text-white' : 'text-white/60 hover:text-white hover:bg-white/5'}`}>
              {tab === key && <span className="absolute left-0 top-2 bottom-2 w-1 rounded-full bg-brand-red" />}
              <Icon className={`w-[18px] h-[18px] shrink-0 ${tab === key ? 'text-brand-red' : ''}`} /> <span className={navCollapsed ? 'md:hidden' : ''}>{label}</span>
            </button>
          ))}
        </nav>
        <div className="px-2 py-2 border-t border-white/10 space-y-0.5">
          {/* ปุ่มย่อ/ขยาย (เฉพาะจอใหญ่) */}
          <button type="button" onClick={() => setNavCollapsed((c) => !c)} title={navCollapsed ? 'ขยายเมนู' : 'ย่อเมนู'}
            className={`hidden md:flex w-full items-center gap-3 px-3 py-2 rounded-lg text-xs font-semibold text-white/45 hover:text-white hover:bg-white/5 ${navCollapsed ? 'justify-center px-0' : ''}`}>
            {navCollapsed ? <ChevronsRight className="w-[18px] h-[18px]" /> : <><ChevronsLeft className="w-[18px] h-[18px]" /> ย่อเมนู</>}
          </button>
          <button type="button" onClick={logout} title={navCollapsed ? 'ออกจากระบบ' : ''}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-semibold text-white/60 hover:text-white hover:bg-white/5 ${navCollapsed ? 'md:justify-center md:px-0' : ''}`}>
            <LogOut className="w-[18px] h-[18px] shrink-0" /> <span className={navCollapsed ? 'md:hidden' : ''}>ออกจากระบบ</span>
          </button>
        </div>
      </aside>

      {/* ===== Main column ===== */}
      <div className="flex-1 min-w-0 flex flex-col">
        {/* Topbar */}
        <header className="bg-white border-b border-natural-border px-4 md:px-6 py-3 flex items-center gap-3 sticky top-0 z-30">
          <button type="button" aria-label="เปิดเมนู" onClick={() => setNavOpen(true)} className="md:hidden text-natural-muted"><Menu className="w-6 h-6" /></button>
          <h1 className="text-base md:text-xl font-bold text-brand-navy truncate">{activeTabLabel}</h1>
          <div className="flex items-center gap-2.5 flex-wrap ml-auto">
            <div className="flex items-center gap-1.5 bg-natural-secondary border border-natural-border rounded-full px-3 py-1.5">
              <Building2 className="w-4 h-4 text-brand-navy" />
              {auth.isHQ ? (
                <select aria-label="เลือกสาขา" value={workBranchId} onChange={(e) => setWorkBranchId(e.target.value)}
                  className="bg-transparent text-xs font-bold text-brand-navy outline-none cursor-pointer">
                  <option value="">🌐 ทุกสาขา (ภาพรวม)</option>
                  {db.branches.filter((b) => !b.isHQ && b.status === 'active').map((b) => (
                    <option key={b.id} value={b.id}>{b.name}</option>
                  ))}
                </select>
              ) : (
                <span className="text-xs font-bold text-brand-navy">{activeBranchName}</span>
              )}
              {auth.isHQ && <span className="text-[10px] bg-brand-red text-white rounded-full px-1.5 py-0.5 font-bold">HQ</span>}
            </div>
            <CycleBar cycles={db.cycles} selectedCycleId={selectedCycleId} setSelectedCycleId={setSelectedCycleId}
              onCreated={(id: string) => fetchState(id)} api={api} showToast={showToast} isHQ={auth.isHQ} />
          </div>
        </header>

        <main className="flex-1 max-w-[1440px] w-full p-4 md:p-6">
          {loading && !booted ? (
            <div className="flex flex-col items-center justify-center py-20">
              <RefreshCw className="w-10 h-10 text-brand-red animate-spin mb-3" />
              <p className="text-sm text-natural-dark-muted">กำลังโหลด...</p>
            </div>
          ) : (
            <>
            {tab === 'calc' && <CalcTab db={db} cycle={cycle} cycleTrips={cycleTrips} api={api} aiEnabled={aiEnabled} branchId={effBranchId}
              reload={() => fetchState(selectedCycleId)} gotoCycle={(id: string) => fetchState(id)} showToast={showToast} />}
            {tab === 'fuel' && <FuelDeductionTab db={db} cycle={cycle} api={api} branchId={effBranchId}
              reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'dashboard' && <DashboardTab db={db} cycle={cycle} branchId={effBranchId} isHQ={auth.isHQ} />}
            {tab === 'driverkpi' && <DriverKpiTab db={db} cycle={cycle} />}
            {tab === 'costarea' && <CostAreaTab db={db} cycle={cycle} branchId={effBranchId} showToast={showToast} />}
            {tab === 'reports' && <ReportsTab db={db} cycle={cycle} branchId={effBranchId} showToast={showToast} />}
            {tab === 'rates' && <RatesTab db={db} api={api} branchId={effBranchId} cycle={cycle} reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'rules' && <RulesTab db={db} api={api} branchId={effBranchId} reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'destfix' && <DestFixTab db={db} api={api} branchId={effBranchId} reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'vehicles' && <VehiclesTab db={db} api={api} branchId={effBranchId} reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'branches' && <BranchesTab db={db} api={api} reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
          </>
        )}
        </main>
      </div>
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
    <div className="min-h-screen bg-gradient-to-br from-brand-navy to-[#081627] flex items-center justify-center p-6">
      <div className="bg-white rounded-2xl shadow-2xl p-8 w-full max-w-sm flex flex-col items-center">
        <img src="/iconneo.png" alt="NEOSIAM" className="w-16 h-16 rounded-2xl shadow-md mb-3 object-cover" />
        <div className="brand-wordmark text-3xl text-brand-navy">NEOSIAM</div>
        <div className="text-[11px] text-brand-red font-bold tracking-wide">ส่งด่วน · ส่งไว · แน่นอน</div>
        <p className="text-xs text-natural-muted mt-1 mb-5">ระบบค่าเที่ยว + ค่าน้ำมันรถร่วม</p>
        <label className="w-full text-xs font-semibold text-natural-dark-muted mb-1">สาขา</label>
        <select aria-label="สาขา" value={branchId} onChange={(e) => setBranchId(e.target.value)}
          className="w-full border border-natural-border rounded-lg px-3 py-2 text-sm mb-3 focus:border-brand-navy outline-none">
          {active.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}
        </select>
        <label className="w-full text-xs font-semibold text-natural-dark-muted mb-1">รหัสผ่าน</label>
        <input type="password" value={password} onChange={(e) => setPassword(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && submit()} aria-label="รหัสผ่าน"
          className="w-full border border-natural-border rounded-lg px-3 py-2 text-sm mb-5 focus:border-brand-navy outline-none" placeholder="••••" />
        <button type="button" onClick={submit} disabled={busy || !branchId}
          className="w-full bg-brand-red hover:bg-brand-red-hover disabled:bg-natural-muted text-white rounded-lg py-2.5 text-sm font-bold flex items-center justify-center gap-2 transition-colors">
          {busy ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Unlock className="w-4 h-4" />} เข้าใช้งาน
        </button>
        <p className="text-[10px] text-natural-muted/70 mt-5 font-semibold tracking-wide">NEOSIAM LOGISTICS &amp; TRANSPORT</p>
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
  const [areaEdit, setAreaEdit] = useState<string | null>(null);
  const [areaText, setAreaText] = useState('');
  const openArea = (b: Branch) => { setAreaEdit(b.id); setAreaText(b.serviceAreaText || ''); };
  const autoArea = (b: Branch) => {
    const provMap = new Map<string, Set<string>>();
    for (const r of (db.rateMasters as RateMaster[]).filter((x) => x.branchId === b.id && x.status === 'active')) {
      const prov = (r.provinceName || '').trim(), dist = (r.districtName || '').trim();
      if (!prov) continue;
      if (!provMap.has(prov)) provMap.set(prov, new Set());
      if (dist && !dist.includes('+')) provMap.get(prov)!.add(dist);
    }
    setAreaText([...provMap.entries()].map(([p, ds]) => ds.size ? `${p}: ${[...ds].join(', ')}` : p).join('\n'));
  };
  const saveArea = async (b: Branch) => {
    await api(`/api/branches/${b.id}`, 'PUT', { serviceAreaText: areaText });
    showToast('success', `บันทึกพื้นที่สาขา ${b.name} แล้ว`); setAreaEdit(null); reload();
  };
  // สลับโหมด "เก็บสินค้าคืน = ครึ่งราคาชิ้นปกติ" (เช่น นครสวรรค์)
  const toggleCollectHalf = async (b: Branch) => {
    const next = !b.collectBackHalfPiece;
    await api(`/api/branches/${b.id}`, 'PUT', { collectBackHalfPiece: next });
    showToast('success', `${b.name}: เก็บคืนคิดครึ่งราคาชิ้น — ${next ? 'เปิด ✅' : 'ปิด'}`);
    reload();
  };

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
        <h3 className="font-bold text-brand-navy mb-3 flex items-center gap-2"><Building2 className="w-4 h-4" />เพิ่มสาขาใหม่</h3>
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
          <button onClick={add} className="bg-brand-red text-white rounded-lg px-4 py-2 text-sm font-semibold flex items-center gap-1"><Plus className="w-4 h-4" />เพิ่ม</button>
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
              <Fragment key={b.id}>
              <tr className="border-t border-natural-border">
                <td className="px-4 py-2 font-semibold text-brand-navy">{b.name} {b.isHQ && <span className="text-[10px] bg-emerald-600 text-white rounded-full px-1.5 py-0.5 ml-1">HQ</span>}{!b.isHQ && (b.serviceAreaText || '').trim() && <span className="text-[10px] bg-brand-navy text-white rounded-full px-1.5 py-0.5 ml-1" title={b.serviceAreaText}>📍 ตั้งพื้นที่แล้ว</span>}{!b.isHQ && b.collectBackHalfPiece && <span className="text-[10px] bg-amber-500 text-white rounded-full px-1.5 py-0.5 ml-1" title="เก็บสินค้าคืน คิดครึ่งราคาชิ้นปกติอัตโนมัติ">♻ เก็บคืน½</span>}</td>
                <td className="px-4 py-2 text-natural-muted">{db.vehicles.filter((v: Vehicle) => v.branchId === b.id).length}</td>
                <td className="px-4 py-2 text-natural-muted">{db.rateMasters.filter((r: RateMaster) => r.branchId === b.id).length}</td>
                <td className="px-4 py-2 text-natural-muted">{countMaster(b.id)}</td>
                <td className="px-4 py-2 text-right whitespace-nowrap">
                  {!b.isHQ && <button type="button" onClick={() => (areaEdit === b.id ? setAreaEdit(null) : openArea(b))} className="text-xs text-brand-navy font-semibold mr-3">📍 พื้นที่</button>}
                  {!b.isHQ && <button type="button" title="เปิด/ปิด: เก็บสินค้าคืน คิด=ครึ่งราคาชิ้นปกติของปลายทาง อัตโนมัติ" onClick={() => toggleCollectHalf(b)} className={`text-xs font-semibold mr-3 ${b.collectBackHalfPiece ? 'text-amber-600' : 'text-natural-muted'}`}>♻ เก็บคืน½ {b.collectBackHalfPiece ? 'เปิด' : 'ปิด'}</button>}
                  {!b.isHQ && <button type="button" onClick={() => cloneInto(b)} className="text-xs text-emerald-700 font-semibold mr-3">⬇ คัดลอกกฎ</button>}
                  <button onClick={() => setPwd(b)} className="text-xs text-brand-navy font-semibold mr-3">ตั้งรหัสผ่าน</button>
                  {!b.isHQ && <button type="button" title="ลบสาขา" onClick={() => del(b)} className="text-red-500"><Trash2 className="w-4 h-4 inline" /></button>}
                </td>
              </tr>
              {areaEdit === b.id && (
                <tr className="border-t border-natural-border bg-natural-secondary/40"><td colSpan={5} className="px-4 py-3">
                  <div className="text-xs font-semibold text-brand-navy mb-1">📍 พื้นที่ให้บริการของสาขา {b.name} (ตรวจตอนบันทึกใบกระจาย)</div>
                  <p className="text-[11px] text-natural-muted mb-2">บรรทัดละ 1 จังหวัด · ใส่ทั้งจังหวัด เช่น <b>พิจิตร</b> หรือเจาะอำเภอ เช่น <b>เพชรบูรณ์: ชนแดน, เมือง</b> (นครสวรรค์/พิษณุโลกมีเพชรบูรณ์เหมือนกัน ใส่อำเภอแยก) · เว้นว่าง = ไม่ตรวจ</p>
                  <textarea value={areaText} onChange={(e) => setAreaText(e.target.value)} rows={6} placeholder={'นครสวรรค์\nพิจิตร\nเพชรบูรณ์: ชนแดน, เมือง'} className="w-full border border-natural-border rounded-lg px-2 py-1.5 text-sm font-mono" />
                  <div className="flex gap-2 mt-2">
                    <button type="button" onClick={() => autoArea(b)} className="text-xs border border-emerald-500 text-emerald-700 rounded-lg px-3 py-1.5 font-semibold">⚡ สร้างจากราคาที่มี ({db.rateMasters.filter((r: RateMaster) => r.branchId === b.id).length})</button>
                    <button type="button" onClick={() => saveArea(b)} className="text-xs bg-brand-red text-white rounded-lg px-4 py-1.5 font-bold">บันทึก</button>
                    <button type="button" onClick={() => setAreaEdit(null)} className="text-xs text-natural-muted border border-natural-border rounded-lg px-3 py-1.5">ยกเลิก</button>
                  </div>
                </td></tr>
              )}
              </Fragment>
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
function CycleBar({ cycles, selectedCycleId, setSelectedCycleId, onCreated, api, showToast, isHQ }: any) {
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

  const delCycle = async () => {
    const ok = await confirmDelete(`รอบ "${cur.name}" (รอบที่มีข้อมูลจะลบไม่ได้)`);
    if (!ok) return;
    try {
      await api(`/api/cycles/${cur.id}`, 'DELETE');
      showToast('success', `ลบรอบ "${cur.name}" แล้ว`);
      const other = cycles.find((c: BillingCycle) => c.id !== cur.id);
      setSelectedCycleId(other ? other.id : '');
      onCreated(other ? other.id : '');
    } catch (e: any) { showToast('error', e.message); }
  };

  return (
    <div className="flex items-center gap-2 flex-wrap relative">
      <select aria-label="เลือกรอบคำนวณ" value={selectedCycleId} onChange={(e) => setSelectedCycleId(e.target.value)}
        className="bg-natural-secondary border border-natural-border text-sm font-semibold rounded-full px-4 py-2">
        <option value="" disabled>— เลือกรอบ —</option>
        {cycles.map((c: BillingCycle) => <option key={c.id} value={c.id}>{c.name} {c.status === 'closed' ? '🔒' : ''}</option>)}
      </select>
      {cur && isHQ && (
        <button onClick={toggleLock} className="border border-natural-border rounded-full px-3 py-2 text-xs font-semibold flex items-center gap-1">
          {cur.status === 'open' ? <><Lock className="w-3.5 h-3.5" />ปิดรอบ</> : <><Unlock className="w-3.5 h-3.5" />เปิดรอบ</>}
        </button>
      )}
      {cur && isHQ && (
        <button onClick={delCycle} title="ลบรอบนี้" className="border border-rose-200 text-rose-600 hover:bg-rose-50 rounded-full px-3 py-2 text-xs font-semibold flex items-center gap-1">
          <Trash2 className="w-3.5 h-3.5" />ลบรอบ
        </button>
      )}
      {cur && !isHQ && cur.status === 'closed' && (
        <span className="border border-natural-border rounded-full px-3 py-2 text-xs font-semibold flex items-center gap-1 text-natural-muted bg-natural-secondary"><Lock className="w-3.5 h-3.5" />รอบนี้ถูกปิด (เฉพาะ HQ เปิด/ปิดได้)</span>
      )}
      <button onClick={() => setOpen(!open)} className="bg-brand-red text-white rounded-full px-4 py-2 text-xs font-semibold flex items-center gap-1">
        <Plus className="w-4 h-4" />เปิดรอบใหม่
      </button>
      {open && (
        <div className="absolute top-12 right-0 bg-white border border-natural-border rounded-2xl shadow-lg p-4 z-40 flex flex-col gap-3 w-72">
          <h4 className="font-bold text-sm text-brand-navy">เปิดรอบคำนวณใหม่</h4>
          <div className="flex gap-2">
            <select aria-label="เลือกเดือน" value={month} onChange={(e) => setMonth(+e.target.value)} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm flex-1">
              {THAI_MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
            </select>
            <input type="number" aria-label="ปี ค.ศ." value={year} onChange={(e) => setYear(+e.target.value)} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-24" />
          </div>
          <div className="flex gap-2">
            {(['first', 'second'] as const).map((h) => (
              <button key={h} onClick={() => setHalf(h)} className={`flex-1 py-1.5 rounded-lg text-sm font-semibold border ${half === h ? 'bg-brand-navy text-white border-brand-navy' : 'border-natural-border'}`}>
                {h === 'first' ? 'รอบ 1-15' : 'รอบ 16-31'}
              </button>
            ))}
          </div>
          <button onClick={create} className="bg-brand-red text-white rounded-lg py-2 text-sm font-semibold">สร้างรอบ</button>
        </div>
      )}
    </div>
  );
}

// ===========================================================================
// Tab: คำนวณค่าเที่ยว (upload + review + list)
// ===========================================================================
function CalcTab({ db, cycle, cycleTrips, api, aiEnabled, branchId, reload, gotoCycle, showToast }: any) {
  const [extracting, setExtracting] = useState(false);
  const [pending, setPending] = useState<{ extracted: ExtractedTripDocument; fileName: string; preview: TripDocument } | null>(null);
  const [filter, setFilter] = useState<'all' | 'divider' | 'warning'>('all');
  const [search, setSearch] = useState('');
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
    const list = Array.from(files);
    if (list.length > 1) showToast('warning', 'นำเข้า Excel ได้ทีละ 1 ไฟล์เท่านั้น — ใช้ไฟล์แรก');
    const file = list[0];
    if (!file) return;
    const name = file.name.toLowerCase();
    if (!name.endsWith('.xls') && !name.endsWith('.xlsx')) { showToast('error', 'รองรับเฉพาะไฟล์ Excel (.xls/.xlsx)'); return; }
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
      const saved = await api('/api/trips', 'POST', { extracted: pending.extracted, fileName: pending.fileName, branchId });
      const cy = saved?._cycle;
      showToast('success', saved?._cycleCreated ? `เปิดรอบ "${cy?.name}" อัตโนมัติ + บันทึกแล้ว` : `บันทึกเข้ารอบ "${cy?.name}" แล้ว`);
      setPending(null);
      if (cy?.id && cy.id !== cycle?.id) gotoCycle(cy.id); // สลับไปรอบที่ใบนั้นเข้าจริง
      else reload();
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

  const q = search.trim().toLowerCase();
  const visibleTrips = cycleTrips.filter((t: TripDocument) => {
    if (filter === 'divider' && !t.receipts.some((r) => r.hasAdjustment)) return false;
    if (filter === 'warning' && t.warnings.length === 0) return false;
    if (q) {
      const hay = [
        t.documentNo, t.plateNo, t.driverName, t.provinceRaw, t.districtRaw,
        ...t.receipts.map((r) => r.receiptNo),
        ...t.receipts.map((r) => r.receiverName),
        ...t.receipts.flatMap((r) => (r.items || []).map((it) => it.productName)),
      ].filter(Boolean).join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  const totalTrip = cycleTrips.reduce((s: number, t: TripDocument) => s + t.tripAmount, 0);

  return (
    <div className="flex flex-col gap-5">
      {/* action bar */}
      <div className="bg-white rounded-2xl border border-natural-border p-4 flex flex-wrap items-center justify-between gap-3">
        <div className="text-sm">
          <span className="font-bold text-brand-navy">{cycle.name}</span>
          <span className="text-natural-muted ml-2">{cycleTrips.length} ใบกระจาย · ค่าเที่ยวรวม ฿{money(totalTrip)}</span>
        </div>
        <div className="flex gap-2">
          <button onClick={recalc} className="border border-natural-border rounded-full px-3 py-2 text-xs font-semibold flex items-center gap-1"><RefreshCw className="w-3.5 h-3.5" />Recalculate</button>
          <button onClick={exportExcel} className="bg-brand-red text-white rounded-full px-4 py-2 text-xs font-semibold flex items-center gap-1"><FileSpreadsheet className="w-4 h-4" />Export Excel</button>
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
        <input ref={excelRef} type="file" aria-label="นำเข้า Excel ใบกระจาย" accept=".xls,.xlsx" className="hidden" onChange={(e) => e.target.files && onExcelFiles(e.target.files)} />
        {extracting || importing ? (
          <div className="flex flex-col items-center gap-2 py-2"><RefreshCw className="w-8 h-8 text-brand-navy animate-spin" /><p className="text-sm font-semibold text-brand-navy">{importing ? 'กำลังอ่าน Excel ใบกระจาย...' : 'AI กำลังอ่าน PDF ใบกระจาย...'}</p></div>
        ) : (
          <>
            <UploadCloud className="w-8 h-8 text-brand-navy mb-2" />
            <p className="font-semibold text-sm text-brand-navy">ลากวางไฟล์ใบกระจาย — Excel (.xls/.xlsx) หรือ PDF (หลายไฟล์ได้)</p>
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
                className="bg-brand-red disabled:bg-natural-muted disabled:cursor-not-allowed text-white rounded-full px-4 py-2 text-xs font-semibold">เลือกไฟล์ PDF</button>
              <button onClick={manual} className="border border-natural-border rounded-full px-4 py-2 text-xs font-semibold">กรอกเอง</button>
              <label className="flex items-center gap-1 text-[11px] text-natural-muted ml-1">
                โมเดล AI:
                <select aria-label="เลือกโมเดล AI" value={db.settings?.geminiModel || 'gemini-3.5-flash'}
                  onChange={async (e) => {
                    await api('/api/settings', 'PUT', { geminiModel: e.target.value });
                    showToast('success', `เปลี่ยนโมเดลเป็น ${e.target.value}`);
                    reload();
                  }}
                  className="border border-natural-border rounded-lg px-2 py-1 text-[11px] font-semibold text-brand-navy">
                  {GEMINI_MODELS.map((m) => <option key={m} value={m}>{m}</option>)}
                </select>
              </label>
            </div>
          </>
        )}
      </div>

      {/* review */}
      {pending && <ReviewBoard pending={pending} setPending={setPending} onPreview={preview} onSave={save} locked={cycle.status === 'closed'} existingTrips={db.tripDocuments} cycles={db.cycles} cycleId={cycle.id} serviceAreaText={(db.branches as Branch[]).find((b) => b.id === branchId)?.serviceAreaText || ''} branchLabel={(db.branches as Branch[]).find((b) => b.id === branchId)?.name || ''} />}

      {/* filter + search */}
      <div className="flex flex-wrap items-center gap-2">
        {([['all', 'ทั้งหมด'], ['divider', '🟧 เฉพาะมีตัวหาร'], ['warning', '⚠️ ต้องตรวจสอบ']] as const).map(([k, l]) => (
          <button key={k} onClick={() => setFilter(k)} className={`px-3 py-1.5 rounded-full text-xs font-semibold border ${filter === k ? 'bg-brand-navy text-white border-brand-navy' : 'border-natural-border'}`}>{l}</button>
        ))}
        <div className="relative ml-auto">
          <Search className="w-4 h-4 text-natural-muted absolute left-2.5 top-1/2 -translate-y-1/2" />
          <input
            value={search} onChange={(e) => setSearch(e.target.value)}
            placeholder="ค้นหา: เลขใบกระจาย / ทะเบียน / คนขับ / เลขใบรับ / ผู้รับ / ชื่อสินค้า"
            className="border border-natural-border rounded-full pl-8 pr-8 py-1.5 text-xs w-72 focus:outline-none focus:border-brand-navy"
          />
          {search && <button onClick={() => setSearch('')} title="ล้าง" className="absolute right-2.5 top-1/2 -translate-y-1/2 text-natural-muted hover:text-rose-600 font-bold">×</button>}
        </div>
      </div>
      {q && <div className="text-xs text-natural-muted -mt-2">พบ {visibleTrips.length} ใบ จากทั้งหมด {cycleTrips.length} ใบ</div>}

      {/* saved trips */}
      {visibleTrips.length === 0 ? <EmptyHint text={q ? `ไม่พบใบกระจายที่ตรงกับ "${search}"` : 'ยังไม่มีใบกระจายในรอบนี้'} /> :
        visibleTrips.map((t: TripDocument) => <TripCard key={t.id} trip={t} onDelete={() => { void del(t.id); }} branchName={(db.branches as Branch[]).find((b) => b.id === branchId)?.name || ''} cycleName={cycle?.name || ''} />)}
    </div>
  );
}

// Review board — แก้ไข extracted + แสดงผล preview พร้อม badge
function ReviewBoard({ pending, setPending, onPreview, onSave, existingTrips = [], cycles = [], cycleId, serviceAreaText = '', branchLabel = '' }: any) {
  const ext: ExtractedTripDocument = pending.extracted;
  const prev: TripDocument = pending.preview;
  const needsBox = prev.receipts.some((r) => r.requiresManualBox && (r.manualBoxQty == null || r.manualBoxQty <= 0));

  // 🔒 กฎเหล็ก: เลขใบกระจายห้ามซ้ำในสาขา (ทุกรอบ)
  const docNo = (ext.documentNo || '').trim();
  const dupTrip = docNo ? (existingTrips as TripDocument[]).find((t) => (t.documentNo || '').trim() === docNo) : null;
  const dupCycleName = dupTrip ? ((cycles as BillingCycle[]).find((c) => c.id === dupTrip.cycleId)?.name || dupTrip.cycleId) : '';
  const isDup = !!dupTrip;

  // 📅 รอบที่ใบนี้จะเข้า (คำนวณจากวันที่ในใบ ฝั่ง server)
  const tgtCycle: string = (prev as any)._cycleName || '';
  const cycleClosed: boolean = !!(prev as any)._cycleClosed;
  const cycleNew: boolean = !!(prev as any)._cycleCreated;

  // 🔒 ด่านตรวจก่อนยืนยัน — ต้องแก้ให้ครบก่อนบันทึก
  const noVehicle = (prev.warnings || []).some((w) => /ไม่อยู่ใน Master รายชื่อรถ/.test(w));
  const noPrice = (prev.warnings || []).some((w) => /ไม่เจอ.*ราคา/.test(w));
  // ปลายทางไม่อยู่ในพื้นที่ของสาขา (ตรวจระดับจังหวัด/อำเภอ)
  // ปลายทางที่จับราคาได้ (เหมา/ชิ้น/เก็บคืน/Peat) = สาขาวิ่งจริง -> ถือว่าอยู่ในพื้นที่อัตโนมัติ
  // (ราคาเป็นแหล่งความจริงแหล่งเดียว ไม่ต้อง sync รายการพื้นที่ตามทุกครั้งที่เพิ่มราคา)
  const areas = parseServiceAreas(serviceAreaText);
  const hasRate = (r: TripReceipt) =>
    r.flatPrice != null || r.piecePrice != null || r.collectPrice != null || r.collectFlatPrice != null || r.peatPrice != null;
  const offArea = [...new Set(prev.receipts
    .filter((r) => (r.totalQty || 0) > 0 && !hasRate(r) && !inServiceArea(areas, r.provinceRaw, r.districtRaw))
    .map((r) => `${r.districtRaw ? 'อ.' + r.districtRaw + ' ' : ''}จ.${r.provinceRaw || '?'}`))];
  const blockReasons: string[] = [];
  if (offArea.length) blockReasons.push(`ปลายทางไม่อยู่ในพื้นที่ของสาขา${branchLabel ? ' ' + branchLabel : ''}: ${offArea.join(', ')} — ตรวจปลายทาง/แก้ปลายทาง หรือเพิ่มพื้นที่ในเมนูจัดการสาขา`);
  if (cycleClosed) blockReasons.push('รอบถูกปิด — ให้ HQ เปิดรอบก่อน');
  if (isDup) blockReasons.push(`เลขใบกระจายซ้ำ (มีอยู่แล้วใน ${dupCycleName})`);
  if (needsBox) blockReasons.push('ยังไม่กรอกจำนวนกล่อง (ผู้ส่งส่งเป็นชิ้น)');
  if (noVehicle) blockReasons.push('ไม่มีทะเบียนรถนี้ใน Master — เพิ่มรถที่เมนู "รถ & คนขับ" ก่อน');
  if (noPrice) blockReasons.push('ไม่เจอราคาขนส่งของปลายทาง — เพิ่มราคาใน Master / แก้ปลายทางให้ตรง');
  const blocked = blockReasons.length > 0;
  const trySave = () => {
    if (blocked) { alertBox('🔒 บันทึกไม่ได้ — ต้องแก้ให้ครบก่อน', blockReasons.map((r, i) => `${i + 1}. ${r}`).join('\n'), 'error'); return; }
    onSave();
  };

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
      <div className="flex items-center justify-between border-b border-natural-border pb-2 gap-2">
        <h3 className="font-bold text-sm text-brand-navy truncate">ขั้นตอนตรวจสอบก่อนยืนยัน — {pending.fileName}</h3>
        <div className="flex items-center gap-2 shrink-0">
          <button onClick={() => setPending(null)} className="text-xs text-natural-muted hover:text-rose-600 font-semibold border border-natural-border rounded-full px-3 py-1.5">ยกเลิก</button>
          <button onClick={trySave} title={blockReasons.join(' · ')} className={`${blocked ? 'bg-natural-muted' : 'bg-brand-red'} text-white rounded-full px-4 py-1.5 text-xs font-bold flex items-center gap-1.5`}><Save className="w-3.5 h-3.5" />ยืนยันบันทึก</button>
        </div>
      </div>
      {blocked && (
        <div className="rounded-xl bg-rose-50 border border-rose-300 px-3 py-2 text-xs text-rose-800">
          <div className="font-bold mb-0.5">🔒 บันทึกไม่ได้ — ต้องแก้ให้ครบก่อน:</div>
          <ul className="list-disc ml-5">{blockReasons.map((r, i) => <li key={i}>{r}</li>)}</ul>
        </div>
      )}

      {/* 📅 รอบที่ใบนี้จะเข้า (อัตโนมัติจากวันที่ในใบ) */}
      {tgtCycle && (
        <div className={`rounded-xl p-3 text-xs flex items-center gap-1.5 font-semibold border ${cycleClosed ? 'bg-rose-50 border-rose-400 text-rose-800' : 'bg-[#EAF2F8] border-brand-navy/30 text-brand-navy'}`}>
          <Calendar className="w-4 h-4 shrink-0" />
          ใบนี้จะเข้ารอบ: <b>{tgtCycle}</b> (ตามวันที่ {ext.documentDate})
          {cycleNew && <span className="bg-emerald-600 text-white rounded-full px-2 py-0.5">เปิดรอบใหม่อัตโนมัติ</span>}
          {cycleClosed && <span className="ml-1">⚠️ รอบนี้ถูกปิดอยู่ — บันทึกไม่ได้ (ให้ HQ เปิดรอบก่อน)</span>}
        </div>
      )}

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
        prev.rateType === 'piece' ? 'border-[#C65911] bg-[#FFF2CC]' : prev.rateType === 'flat' ? 'border-brand-navy bg-[#EAF2F8]' : 'border-rose-400 bg-rose-50'}`}>
        <div className="flex items-center gap-2.5 shrink-0">
          <div className={`rounded-xl p-2 text-white ${prev.rateType === 'piece' ? 'bg-[#C65911]' : prev.rateType === 'flat' ? 'bg-brand-navy' : 'bg-rose-500'}`}>
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
              className="flex-1 sm:flex-none border-2 border-natural-border bg-white rounded-xl px-3 py-2.5 text-base font-bold text-brand-navy focus:border-brand-navy outline-none cursor-pointer shadow-xs">
              {prev.rateOptions.flat != null && <option value="flat">🔵 เหมา ฿{money(prev.rateOptions.flat)} (สูงสุด)</option>}
              {prev.rateOptions.piece != null && <option value="piece">🟠 ชิ้น รวมทุกจุด ฿{money(prev.rateOptions.piece)}</option>}
            </select>
            <span className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-bold text-white ${prev.rateType === 'piece' ? 'bg-[#C65911]' : 'bg-brand-navy'}`}>
              {prev.rateType === 'piece' ? 'คิดแบบ ชิ้น' : 'คิดแบบ เหมา'}
            </span>
          </div>
        ) : (
          <span className="sm:ml-auto font-bold text-[#9C0006] flex items-center gap-1.5"><AlertTriangle className="w-4 h-4" />— ไม่เจอราคาขนส่ง —</span>
        )}
      </div>

      {/* 💰 ยอดรวมทั้งใบ (รวมเก็บคืน/Peat) — เด่นชัด */}
      <div className="rounded-xl bg-brand-navy text-white px-4 py-2.5 flex flex-wrap items-center gap-x-4 gap-y-1">
        <span className="text-base font-bold">💰 ยอดรวมทั้งใบ: ฿{money(prev.tripAmount)}</span>
        <span className="text-[11px] opacity-90">
          งานปกติ ฿{money(prev.breakdown?.normal || 0)}
          {(prev.breakdown?.collect || 0) > 0 && <> + 🔄 เก็บคืน ฿{money(prev.breakdown.collect)}</>}
          {(prev.breakdown?.peat || 0) > 0 && <> + 🌱 Peat ฿{money(prev.breakdown.peat)}</>}
          {(prev.breakdown?.addon || 0) > 0 && <> + 📌 ค่าเหมาเพิ่มตายตัว ฿{money(prev.breakdown.addon)}</>}
        </span>
      </div>

      {/* 📝 โน้ตท้ายใบ (เช่น "วิ่งย่อยไม่เกิน 13 จุด") — ใช้จับราคาแบบมีเงื่อนไข */}
      {(prev.docNote || (ext as any).docNote) && (
        <div className="rounded-lg bg-amber-50 border border-amber-200 text-amber-800 px-3 py-1.5 text-xs font-semibold flex items-center gap-1.5">
          📝 โน้ตท้ายใบ: {prev.docNote || (ext as any).docNote}
        </div>
      )}

      <datalist id="units">
        {['กล่อง', 'หีบ', 'ลัง', 'ชิ้น', 'แพ็ค', 'ถุง', 'โหล'].map((u) => <option key={u} value={u} />)}
      </datalist>

      {/* receipts */}
      {ext.receipts.map((r, ri) => {
        const pr = prev.receipts[ri];
        return (
          <div key={ri} className={`rounded-xl border p-3 ${pr?.destCorrected ? 'bg-violet-50 border-l-4 border-l-violet-500 border-natural-border' : pr?.requiresManualBox ? 'bg-sky-50 border-l-4 border-l-sky-500 border-natural-border' : pr?.hasAdjustment ? 'bg-[#FFF2CC] border-l-4 border-l-[#C65911] border-natural-border' : 'border-natural-border'}`}>
            {pr?.destCorrected && (
              <div className="mb-2 text-[11px] bg-violet-100 border border-violet-300 text-violet-800 rounded-lg px-2 py-1 font-semibold">
                📍 แก้ปลายทาง: {pr.origDistrict} จ.{pr.origProvince} → <b>อ.{pr.districtRaw} จ.{pr.provinceRaw}</b> (คิดราคาตามปลายทางจริง · คีย์ "{pr.destFixKeyword}")
              </div>
            )}
            {pr?.destNote && !pr?.destCorrected && (
              <div className="mb-2 text-[11px] bg-amber-50 border border-amber-300 text-amber-800 rounded-lg px-2 py-1 font-semibold">
                ⚠️ พบโน้ต <b>{pr.destNote}</b> — อาจส่งคนละปลายทาง ตรวจสอบ/ตั้งกฎที่เมนู "แก้ปลายทาง 📍" หรือแก้จังหวัด/อำเภอด้านล่างเอง
              </div>
            )}
            <div className="grid grid-cols-2 md:grid-cols-3 gap-2 text-xs mb-2">
              <Field label="เลขใบรับสินค้า" value={r.receiptNo} onChange={(v) => updReceipt(ri, { receiptNo: v })} />
              <Field label="ผู้รับสินค้า" value={r.receiverName} onChange={(v) => updReceipt(ri, { receiverName: v })} />
              <Field label="ผู้ส่งสินค้า" value={r.senderName} onChange={(v) => updReceipt(ri, { senderName: v })} />
              <Field label="จังหวัดปลายทาง" value={r.provinceRaw || ''} onChange={(v) => updReceipt(ri, { provinceRaw: v })} />
              <Field label="อำเภอปลายทาง" value={r.districtRaw || ''} onChange={(v) => updReceipt(ri, { districtRaw: v })} />
              <div className="flex flex-col justify-end text-[11px]">
                <span className="text-natural-muted text-[10px] font-bold uppercase">ราคา/ค่าเที่ยวจุดนี้</span>
                <span className="font-semibold text-brand-navy">
                  {(pr?.collectQty > 0)
                    ? <span className="text-emerald-700">🔄 เก็บคืน ฿{money(pr.collectPrice || 0)} × {qtyFmt(pr.collectQty)} = ฿{money((pr.collectPrice || 0) * pr.collectQty)}{pr.collectFlatPrice != null ? ` · เหมา ฿${money(pr.collectFlatPrice)}` : ''}</span>
                    : (pr?.peatQty > 0)
                    ? <span className="text-emerald-700">🌱 Peat ฿{money(pr.peatPrice || 0)} × {qtyFmt(pr.peatQty)} = ฿{money((pr.peatPrice || 0) * pr.peatQty)}</span>
                    : prev.rateType === 'piece'
                    ? (pr?.piecePrice != null ? `ชิ้น ฿${money(pr.piecePrice)} × ${qtyFmt(pr.billingQty)} = ฿${money(pr.receiptAmount)}` : (pr?.normalQty > 0 ? '⚠️ ไม่เจอราคาชิ้น' : '—'))
                    : (pr?.flatPrice != null ? `เหมา ฿${money(pr.flatPrice)}` : (pr?.normalQty > 0 ? '⚠️ ไม่เจอราคาเหมา' : '—'))}
                </span>
              </div>
            </div>
            <table className="w-full text-xs">
              <thead><tr className="text-natural-muted text-left"><th className="py-1">รายการสินค้า</th><th className="w-16 text-center">จำนวน</th><th className="w-20 text-center">หน่วย</th><th className="w-10"></th></tr></thead>
              <tbody>
                {r.items.map((it, ii) => {
                  const unspecified = isUnspecifiedName(it.productName);
                  const adj = (pr?.adjustments || []).find((a: any) => (a.items || [a.productName]).includes(it.productName));
                  const skip = !adj && (pr?.divisorSkipped || []).find((s: any) => s.productName === it.productName);
                  const single = adj && (adj.items?.length ?? 1) === 1;
                  return (
                  <tr key={ii} title={adj ? `หาร ${adj.divisor}: ${adj.note}` : skip ? `เข้ากฎ÷${skip.divisor} แต่รวม ${skip.groupTotal} < ${skip.divisor} — ไม่หาร` : unspecified ? 'ชื่อสินค้ายังไม่ระบุ — นับเข้ายอดตามเอกสาร' : ''}
                    className={adj ? 'bg-[#FFE0B2] outline outline-1 outline-[#C65911]' : skip ? 'bg-[#FFF2CC] outline outline-1 outline-dashed outline-[#C65911]' : ''}>
                    <td className="py-0.5">
                      <div className="flex items-center gap-1">
                        {adj && <span className="shrink-0 text-[10px] font-extrabold bg-[#C65911] text-white rounded px-1.5 py-0.5">÷{adj.divisor}</span>}
                        {skip && <span className="shrink-0 text-[10px] font-bold bg-white text-[#C65911] border border-[#C65911] rounded px-1.5 py-0.5" title={`รวม ${skip.groupTotal} < ${skip.divisor}`}>÷{skip.divisor}·ไม่ถึง</span>}
                        <input value={it.productName} aria-label="ชื่อสินค้า" onChange={(e) => updItem(ri, ii, { productName: e.target.value })} className={`w-full border-b border-dashed border-natural-border bg-transparent p-1 ${adj ? 'text-[#7a3a00] font-bold' : skip ? 'text-[#9a5a10] font-semibold' : unspecified ? 'text-amber-700' : ''}`} placeholder="ชื่อสินค้า" />
                        {unspecified && <span className="shrink-0 text-[9px] font-bold bg-amber-50 text-amber-700 border border-amber-300 rounded px-1 py-0.5">ระบุชื่อ?</span>}
                      </div>
                    </td>
                    <td><input type="number" aria-label="จำนวนสินค้า" value={it.quantity || ''} onChange={(e) => updItem(ri, ii, { quantity: +e.target.value || 0 })} className={`w-full text-center border-b border-dashed border-natural-border bg-transparent p-1 font-bold ${adj ? 'text-[#7a3a00]' : ''}`} /></td>
                    <td className="text-center">{single
                      ? <span className="text-[11px] font-bold text-[#C65911]">→ {adj.convertedQty}</span>
                      : adj ? <span className="text-[10px] font-semibold text-[#C65911]">(รวมหาร)</span>
                      : <input list="units" aria-label="หน่วยนับ" value={it.unit || ''} onChange={(e) => updItem(ri, ii, { unit: e.target.value })} placeholder="หน่วย" className="w-full text-center border-b border-dashed border-natural-border bg-transparent p-1 text-natural-muted" />}</td>
                    <td className="text-center"><button type="button" aria-label="ลบรายการสินค้า" title="ลบรายการสินค้า" onClick={() => updReceipt(ri, { items: r.items.filter((_, j) => j !== ii) })} className="text-natural-muted hover:text-rose-600"><Trash2 className="w-3.5 h-3.5" /></button></td>
                  </tr>
                  );
                })}
              </tbody>
            </table>
            <button onClick={() => updReceipt(ri, { items: [...r.items, { productName: '', quantity: 0 }] })} className="text-brand-navy text-xs font-semibold mt-1 flex items-center gap-1"><Plus className="w-3 h-3" />เพิ่มสินค้า</button>
            {(pr?.hasAdjustment || pr?.divisorSkipped?.length > 0) && (
              <div className="mt-2 text-[11px] font-semibold flex flex-wrap gap-2">
                {(pr.adjustments || []).map((a: any, i: number) => <span key={i} className="text-[#C65911] bg-white border border-[#C65911]/40 rounded-full px-2 py-0.5">🟧÷{a.divisor} {a.productName}: {a.note}</span>)}
                {(pr.divisorSkipped || []).length > 0 && (() => { const s = pr.divisorSkipped[0]; const tot = s.groupTotal; return <span className="text-[#9a5a10] bg-amber-50 border border-[#C65911]/40 rounded-full px-2 py-0.5">⚠️ เข้ากฎ÷{s.divisor} แต่รวม {tot} {'<'} {s.divisor} — ไม่หาร (เก็บตามจริง)</span>; })()}
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
            <div className="mt-2 text-xs font-semibold flex flex-wrap gap-4">
              <span>จำนวนจริง: <b>{qtyFmt(pr?.totalQty)}</b>{pr?.requiresManualBox ? ' ชิ้น' : ''}</span>
              <span className="text-[#C00000]">คิดค่าเที่ยว: <b>{qtyFmt(pr?.billingQty)}</b>{pr?.requiresManualBox ? ' กล่อง' : ''}</span>
              {pr?.collectQty > 0 && <span className="text-emerald-700">🔄 เก็บคืน: <b>{qtyFmt(pr.collectQty)} × {money(pr.collectPrice || 0)} = ฿{money((pr.collectPrice || 0) * pr.collectQty)}</b>{pr.collectFlatPrice != null ? <> · เหมา <b>฿{money(pr.collectFlatPrice)}</b> (เก็บคืนอย่างเดียวเลือกสูงกว่า)</> : ''}</span>}
              {pr?.peatQty > 0 && <span className="text-emerald-700">🌱 Peat: <b>{qtyFmt(pr.peatQty)} × {money(pr.peatPrice || 0)} = ฿{money((pr.peatPrice || 0) * pr.peatQty)}</b></span>}
            </div>
          </div>
        );
      })}
      <button onClick={() => update({ receipts: [...ext.receipts, { receiptNo: '', receiverName: '', senderName: '', items: [{ productName: '', quantity: 0 }] }] })}
        className="text-brand-navy text-xs font-semibold flex items-center gap-1"><Plus className="w-3.5 h-3.5" />เพิ่มใบรับสินค้า</button>

      {/* totals + save */}
      <div className="flex items-center justify-between border-t border-natural-border pt-3">
        <div className="text-sm font-bold text-brand-navy">ค่าเที่ยวรวม: ฿{money(prev.tripAmount)} <span className="text-natural-muted font-normal">(คิด {qtyFmt(prev.billingQty)}/{qtyFmt(prev.totalQty)} ลัง)</span>
          {prev.breakdown && ((prev.breakdown.collect || 0) > 0 || (prev.breakdown.peat || 0) > 0) && (
            <span className="block text-[11px] text-natural-muted font-normal mt-0.5">
              งานปกติ ฿{money(prev.breakdown.normal)}
              {(prev.breakdown.peat || 0) > 0 && <> + Peat mass ฿{money(prev.breakdown.peat)}</>}
              {(prev.breakdown.collect || 0) > 0 && <> + เก็บคืน ฿{money(prev.breakdown.collect)}</>}
            </span>
          )}
          {blocked && <span className="block text-rose-600 text-xs font-semibold mt-0.5">🔒 บันทึกไม่ได้ — {blockReasons.length} ปัญหา (กดปุ่มเพื่อดูรายละเอียด)</span>}
        </div>
        <button onClick={trySave} title={blockReasons.join(' · ')} className={`${blocked ? 'bg-natural-muted' : 'bg-brand-red'} text-white rounded-full px-5 py-2 text-sm font-semibold flex items-center gap-1.5`}><Save className="w-4 h-4" />ยืนยันบันทึก</button>
      </div>
    </div>
  );
}

// ---- พิมพ์/บันทึก PDF ใบกระจาย รูปแบบ "สรุปการจัดส่ง" ของ NEOSIAM (เป๊ะตามต้นฉบับ) ----
const _esc = (s: any) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' } as any)[c]);
const _fmtPD = (s: string) => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ''); return m ? `${+m[3]}/${+m[2]}/${m[1]}` : (s || ''); };

// สร้าง HTML 1 ใบ (1 หน้า A4) — รูปแบบ "สรุปการจัดส่ง" (ตารางมีเส้นกรอบ, 1 ใบรับ=1 แถว)
function tripDocSection(trip: TripDocument): string {
  const t = new Date(); const pad = (n: number) => String(n).padStart(2, '0');
  const printDate = `${t.getDate()}/${t.getMonth() + 1}/${t.getFullYear()}  ${pad(t.getHours())}:${pad(t.getMinutes())}:${pad(t.getSeconds())}`;
  const groups = new Map<string, { prov: string; dist: string; receipts: TripReceipt[] }>();
  for (const r of trip.receipts || []) {
    const key = (r.provinceRaw || '') + '|' + (r.districtRaw || '');
    if (!groups.has(key)) groups.set(key, { prov: r.provinceRaw || '', dist: r.districtRaw || '', receipts: [] });
    groups.get(key)!.receipts.push(r);
  }
  let no = 0, grandQty = 0, body = '';
  for (const g of groups.values()) {
    body += `<tr><td></td><td><b>จังหวัด : ${_esc(g.prov)}</b></td><td><b>อำเภอ : ${_esc(g.dist)}</b></td><td></td><td></td><td></td><td></td></tr>`;
    for (const r of g.receipts) {
      no++;
      const items = (r.items || []).length ? r.items : [{ productName: '', quantity: r.totalQty, unit: '' } as any];
      const qtyLines = items.map((it: any) => qtyFmt(it.quantity)).join('<br>');
      const unitLines = items.map((it: any) => _esc(it.unit || '')).join('<br>');
      const prodLines = items.map((it: any) => {
        const isDiv = (r.adjustments || []).some((a) => (a.items || [a.productName]).includes(it.productName)) || (r.divisorSkipped || []).some((s) => s.productName === it.productName);
        return `<span style="${isDiv ? 'background:#FFE0B2;padding:0 2px' : ''}">${_esc(it.productName)}</span>`;
      }).join('<br>');
      items.forEach((it: any) => { grandQty += it.quantity || 0; });
      body += `<tr>
        <td style="text-align:center">${no}</td>
        <td>${_esc(r.receiverName)}</td>
        <td>${_esc(r.senderName)}</td>
        <td style="text-align:center">${qtyLines}</td>
        <td style="text-align:center">${unitLines}</td>
        <td>${prodLines}</td>
        <td>${_esc(r.receiptNo)}</td>
      </tr>`;
    }
  }
  return `<div class="doc">
    <div class="hdr"><span class="co">บริษัท นีโอสยาม โลจิสติกส์แอนด์ทรานสปอร์ต จำกัด</span><span class="ver">v 1.00</span></div>
    <div class="title">สรุปการจัดส่ง</div>
    <div class="row"><span class="lbl">ใบกระจายเลขที่</span> &nbsp; ${_esc(trip.documentNo)} &nbsp;&nbsp;&nbsp;&nbsp; <span class="lbl">วันที่ออก</span> &nbsp; ${_fmtPD(trip.documentDate)}</div>
    <div class="row2"><span><span class="lbl">ทะเบียนรถ</span> &nbsp; ${_esc(trip.plateNo)}${trip.driverName ? ' (' + _esc(trip.driverName) + ')' : ''}</span><span><span class="lbl">วันที่พิมพ์</span> &nbsp; ${printDate}</span></div>
    <table class="grid"><thead><tr>
      <th style="width:32px">ลำดับ</th><th style="width:20%">ผู้รับสินค้า</th><th style="width:20%">ผู้ส่งสินค้า</th><th style="width:48px">จำนวน</th><th style="width:42px">หน่วย</th><th style="width:22%">รายการ</th><th>เลขที่ใบรับสินค้า</th>
    </tr></thead><tbody>${body}</tbody></table>
    <div class="footer">
      <div class="signbox"><div class="signline"></div>ผู้ออกเอกสาร</div>
      <div class="totals"><div><span class="lbl">ยอดรวมสินค้า</span> &nbsp; <b>${qtyFmt(grandQty)}</b> &nbsp; ชิ้น</div><div class="lbl" style="margin-top:3px">X - ยังไม่บันทึกส่งเสร็จ</div></div>
      <div class="signbox"><div class="signline"></div>พนักงานขับ</div>
    </div>
    <div class="ft">FM-OP01-05 REV.00</div>
  </div>`;
}

const PRINT_STYLE = `@page{size:A4;margin:14mm}
  *{font-family:Tahoma,"TH Sarabun New","Sarabun",sans-serif;box-sizing:border-box}
  body{color:#000;font-size:13px;margin:0}
  .doc{page-break-after:always}
  .doc:last-child{page-break-after:auto}
  .hdr{overflow:hidden}
  .co{font-weight:bold;font-size:13px;float:left}
  .ver{font-size:11px;float:right}
  .title{text-align:center;font-weight:bold;font-size:20px;margin:6px 0 10px;clear:both}
  .row{font-size:13px;margin:2px 0}
  .row2{display:flex;justify-content:space-between;font-size:13px;margin:2px 0 8px}
  .lbl{font-weight:bold}
  table.grid{width:100%;border-collapse:collapse}
  table.grid th,table.grid td{border:1px solid #000;padding:3px 5px;font-size:11px;vertical-align:top}
  table.grid th{font-weight:bold;text-align:center}
  .footer{display:flex;justify-content:space-between;align-items:flex-end;margin-top:14px}
  .signbox{border:1px solid #000;width:210px;text-align:center;padding:6px;font-size:13px}
  .signline{border-top:1px dotted #000;margin:26px 14px 4px}
  .totals{text-align:center;font-size:13px}
  .totals b{font-size:15px}
  .ft{text-align:right;font-size:11px;margin-top:12px}`;

function openPrint(title: string, inner: string) {
  const html = `<!doctype html><html lang="th"><head><meta charset="utf-8"><title>${_esc(title)}</title><style>${PRINT_STYLE}</style></head><body onload="window.print()">${inner}</body></html>`;
  const w = window.open('', '_blank', 'width=900,height=1000');
  if (!w) { alert('เบราว์เซอร์บล็อกหน้าต่างพิมพ์ — อนุญาต popup แล้วลองใหม่'); return; }
  w.document.write(html); w.document.close();
}

function printTripDocument(trip: TripDocument, _b?: string, _c?: string) { openPrint(trip.documentNo, tripDocSection(trip)); }
function printTripDocuments(trips: TripDocument[], title: string) {
  if (!trips.length) { alert('ไม่มีใบกระจายให้พิมพ์'); return; }
  openPrint(title, trips.map(tripDocSection).join(''));
}

const TripCard: React.FC<{ trip: TripDocument; onDelete: () => void; branchName?: string; cycleName?: string }> = ({ trip, onDelete, branchName = '', cycleName = '' }) => {
  const hasDiv = trip.receipts.some((r) => r.hasAdjustment);
  const [open, setOpen] = useState<Record<string, boolean>>({});
  const totalItems = trip.receipts.reduce((s, r) => s + (r.items || []).length, 0);
  const allOpen = trip.receipts.every((r) => open[r.id]);
  const toggleAll = () => setOpen(allOpen ? {} : Object.fromEntries(trip.receipts.map((r) => [r.id, true])));
  return (
    <div className={`bg-white rounded-2xl border p-4 ${trip.warnings.length ? 'border-l-4 border-l-[#9C0006]' : hasDiv ? 'border-l-4 border-l-[#C65911]' : 'border-natural-border'}`}>
      <div className="flex items-center justify-between">
        <div>
          <span className="font-bold text-brand-navy text-sm">{trip.documentNo}</span>
          <span className="text-natural-muted text-xs ml-2">{trip.documentDate} · {trip.plateNo} {trip.driverName && `(${trip.driverName})`} · {trip.rateType === 'flat' ? 'เหมา' : 'ชิ้น'}</span>
        </div>
        <div className="flex items-center gap-3">
          <button type="button" onClick={toggleAll} className="text-[11px] border border-natural-border rounded-full px-2.5 py-1 text-brand-navy hover:bg-natural-secondary font-semibold">{allOpen ? '▾ ย่อสินค้า' : `▸ ดูสินค้า (${totalItems})`}</button>
          <button type="button" onClick={() => printTripDocument(trip, branchName, cycleName)} title="พิมพ์ / บันทึก PDF (A4)" className="text-[11px] border border-natural-border rounded-full px-2.5 py-1 text-brand-navy hover:bg-natural-secondary font-semibold">🖨️ พิมพ์/PDF</button>
          <span className="text-sm font-bold text-[#C00000]">฿{money(trip.tripAmount)}</span>
          <button type="button" aria-label="ลบใบกระจาย" title="ลบใบกระจาย" onClick={onDelete} className="text-natural-muted hover:text-rose-600"><Trash2 className="w-4 h-4" /></button>
        </div>
      </div>
      {trip.warnings.length > 0 && <div className="mt-1 text-[11px] text-[#9C0006]">⚠️ {trip.warnings.join(' · ')}</div>}
      <div className="mt-2 overflow-x-auto">
        <table className="w-full text-xs min-w-[620px]">
          <thead><tr className="text-natural-muted text-left border-b border-natural-border"><th className="py-1">ใบรับสินค้า</th><th>ผู้รับ</th><th>ปลายทาง</th><th className="text-center w-16">จำนวนจริง</th><th className="text-center w-16">คิดค่าเที่ยว</th><th className="text-right w-20">ค่าเที่ยวจุด</th><th className="text-center w-14">รายการ</th></tr></thead>
          <tbody>
            {trip.receipts.map((r) => {
              const o = !!open[r.id]; const nItems = (r.items || []).length;
              const isCollect = (r.collectQty || 0) > 0;        // เก็บสินค้าคืน
              const isPeat = (r.peatQty || 0) > 0;              // Peat mass
              const collectAmt = isCollect && r.collectPrice != null ? r.collectQty * r.collectPrice : 0;
              const addon = (trip.addonByDest || []).find((a) => (a.prov || '') === (r.provinceRaw || '') && (a.dist || '') === (r.districtRaw || '')); // เหมาบวกเพิ่มตายตัว (ท่าสองยาง +700)
              return (
                <Fragment key={r.id}>
                  <tr onClick={() => setOpen({ ...open, [r.id]: !o })} className={`cursor-pointer ${isCollect ? 'bg-emerald-50' : isPeat ? 'bg-teal-50' : addon ? 'bg-amber-50' : r.hasAdjustment ? 'bg-[#FFF2CC]' : ''} hover:bg-natural-secondary/50`}>
                    <td className="py-1">{isCollect && <span className="text-emerald-700 font-bold">🔄 </span>}{isPeat && <span className="text-teal-700 font-bold">🌱 </span>}{addon && <span className="text-amber-700 font-bold">📌 </span>}{r.hasAdjustment && r.adjustments?.[0] && <span className="text-[#C65911] font-bold">🟧÷{r.adjustments[0].divisor} </span>}{r.receiptNo}</td>
                    <td>{r.receiverName}{isCollect && <span className="text-emerald-700 font-semibold"> · เก็บสินค้าคืน</span>}{isPeat && <span className="text-teal-700 font-semibold"> · Peat</span>}{addon && <span className="text-amber-700 font-semibold"> · เหมาเพิ่มตายตัว</span>}</td>
                    <td>{r.districtRaw} {r.provinceRaw}</td>
                    <td className="text-center">{qtyFmt(r.totalQty)}</td>
                    <td className="text-center font-bold text-[#C00000]" title={(r.adjustments || []).map((a) => a.note).join(' | ')}>{qtyFmt(r.billingQty)}</td>
                    <td className="text-right">{isCollect ? <span className="text-emerald-700 font-semibold whitespace-nowrap">🔄 {qtyFmt(r.collectQty)}×{money(r.collectPrice || 0)}=฿{money(collectAmt)}</span> : (addon && r.receiptAmount === 0 && r.flatPrice == null) ? <span className="text-amber-700 font-semibold whitespace-nowrap">📌 +฿{money(addon.amount)}</span> : (trip.rateType === 'piece' ? money(r.receiptAmount) : (r.flatPrice != null ? money(r.flatPrice) : '-'))}</td>
                    <td className="text-center text-brand-navy font-semibold whitespace-nowrap">{o ? '▾' : '▸'} {nItems}</td>
                  </tr>
                  {o && (
                    <tr className="bg-natural-secondary/30">
                      <td colSpan={7} className="px-3 py-1.5">
                        {nItems === 0 ? <span className="text-natural-muted">— ไม่มีรายการสินค้า —</span> : (
                          <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-x-4 gap-y-0.5">
                            {(r.items || []).map((it, k) => {
                              const adj = (r.adjustments || []).find((a) => (a.items || [a.productName]).includes(it.productName));
                              const skip = !adj && (r.divisorSkipped || []).find((s) => s.productName === it.productName);
                              return (
                                <div key={k} className={`flex justify-between gap-2 items-center rounded px-1 ${adj ? 'bg-[#FFE0B2]' : skip ? 'bg-[#FFF2CC]' : ''}`}>
                                  <span className="truncate flex items-center gap-1 min-w-0">
                                    {adj && <span className="shrink-0 text-[9px] font-extrabold bg-[#C65911] text-white rounded px-1">÷{adj.divisor}</span>}
                                    {skip && <span className="shrink-0 text-[9px] font-bold text-[#C65911] border border-[#C65911] rounded px-1" title={`เข้ากฎ÷${skip.divisor} แต่รวมไม่ถึง`}>÷{skip.divisor}</span>}
                                    <span className={`truncate ${adj || skip ? 'text-[#7a3a00] font-semibold' : ''}`}>{it.productName || '(ไม่ระบุชื่อ)'}</span>
                                  </span>
                                  <span className="font-semibold whitespace-nowrap">{qtyFmt(it.quantity)} {it.unit || ''}</span>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
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
  const fuelFileRef = useRef<HTMLInputElement>(null);
  const [impFuel, setImpFuel] = useState(false);
  const [fPlate, setFPlate] = useState('');

  if (!cycle) return <EmptyHint text="กรุณาเลือกรอบก่อน" />;
  if (!branchId) return <EmptyHint text={ALL_BRANCH_HINT} />;

  const branchName = (db.branches as Branch[]).find((b) => b.id === branchId)?.name || '';
  const addFuel = async () => {
    if (!fForm.plateNo || !fForm.amount) return showToast('warning', 'กรอกทะเบียนและจำนวนเงิน');
    await api('/api/fuel', 'POST', { ...fForm, cycleId: cycle.id, branchId });
    setFForm({ plateNo: '', refNo: '', date: cycle.startDate, amount: 0 }); reload();
  };
  const dlFuelTemplate = () => {
    const vs = (db.vehicles as Vehicle[]).filter((v) => v.status === 'active').map((v) => ({ plateNo: v.plateNo, driverName: v.driverName }));
    downloadFuelTemplate(branchName, vs);
  };
  const onImportFuel = async (files: FileList) => {
    const file = files[0];
    if (!file) return;
    if (!/\.xlsx?$/i.test(file.name)) return showToast('error', 'รองรับเฉพาะไฟล์ Excel (.xls/.xlsx)');
    const b64 = await new Promise<string>((resolve) => { const r = new FileReader(); r.onload = () => resolve((r.result as string).split(',')[1]); r.readAsDataURL(file); });
    setImpFuel(true);
    try {
      const res = await api('/api/import-fuel', 'POST', { branchId, fileBase64: b64 });
      showToast('success', `นำเข้าค่าน้ำมันสำเร็จ — บันทึก ${res.created} รายการ`);
      if (res.summary?.length) alertBox('สรุปการนำเข้าค่าน้ำมัน', res.summary.join('\n'));
      reload();
    } catch (e: any) { showToast('error', e.message); }
    finally { setImpFuel(false); if (fuelFileRef.current) fuelFileRef.current.value = ''; }
  };
  const addEntry = async (plateNo: string, categoryId: string, amount: number, kind: 'income' | 'deduction', docNo: string, reset: () => void) => {
    const cat = cats.find((c) => c.id === categoryId) || (kind === 'income' ? incomeCats[0] : dedCats[0]);
    if (!plateNo || !amount || !cat) return showToast('warning', 'กรอกทะเบียน/จำนวนเงิน และเลือกประเภท');
    // ล้าง docNo: ตัดอักขระแปลกปลอม (ไทยหลง/เว้นวรรค) + ตัวใหญ่ ให้ตรงกับใบกระจาย
    const cleanDoc = docNo.replace(/[^A-Za-z0-9/-]/g, '').toUpperCase().trim();
    await api('/api/deductions', 'POST', { plateNo, categoryId: cat.id, kind, label: cat.name, amount, docNo: cleanDoc, cycleId: cycle.id, branchId });
    reset(); reload();
  };

  // กรองรายทะเบียน (ใช้ร่วมทั้ง 3 ตาราง)
  const allPlates: string[] = [...new Set([...fuel, ...incomes, ...ded].map((x: any) => x.plateNo).filter(Boolean) as string[])].sort();
  const byPlate = (list: any[]): any[] => fPlate ? list.filter((x) => normPlate(x.plateNo) === normPlate(fPlate)) : list;
  const fuelF = byPlate(fuel), incomesF = byPlate(incomes), dedF = byPlate(ded);
  const fuelSum = fuelF.reduce((s: number, f: FuelEntry) => s + f.amount, 0);

  return (
    <div className="flex flex-col gap-4">
      {/* ตัวกรองรายทะเบียน (ใช้ร่วมทุกตาราง) */}
      <div className="bg-white rounded-2xl border border-natural-border p-3 flex flex-wrap items-center gap-2 text-sm">
        <span className="font-semibold text-brand-navy flex items-center gap-1"><Filter className="w-4 h-4" />กรองทะเบียน:</span>
        <select aria-label="กรองทะเบียนรถ" value={fPlate} onChange={(e) => setFPlate(e.target.value)} className="border border-natural-border rounded-full px-3 py-1.5 text-sm font-semibold text-brand-navy">
          <option value="">🚚 ทุกทะเบียน ({allPlates.length})</option>
          {allPlates.map((p) => <option key={p} value={p}>{p}</option>)}
        </select>
        {fPlate && <>
          <span className="text-natural-muted">ค่าน้ำมัน {fuelF.length} รายการ · รวม ฿{money(fuelSum)}</span>
          <button type="button" onClick={() => setFPlate('')} className="text-xs text-natural-muted hover:text-rose-600 font-semibold underline ml-auto">ล้างตัวกรอง</button>
        </>}
      </div>

    <div className="grid md:grid-cols-2 gap-5">
      <Section title="ค่าน้ำมัน (แยกตามทะเบียน)" icon={Fuel}>
        {/* นำเข้า/เทมเพลตค่าน้ำมันจาก Excel */}
        <div className="flex flex-wrap items-center gap-2 mb-3 bg-emerald-50 border border-emerald-200 rounded-xl p-2">
          <span className="text-xs font-semibold text-emerald-800">📥 นำเข้า Excel:</span>
          <button type="button" onClick={dlFuelTemplate} className="bg-white border border-emerald-400 text-emerald-700 rounded-lg px-2.5 py-1 text-xs font-semibold flex items-center gap-1"><FileSpreadsheet className="w-3.5 h-3.5" />เทมเพลต (แยก sheet ต่อทะเบียน)</button>
          <input ref={fuelFileRef} type="file" aria-label="นำเข้าค่าน้ำมัน Excel" accept=".xls,.xlsx" className="hidden" onChange={(e) => e.target.files && onImportFuel(e.target.files)} />
          <button type="button" disabled={impFuel} onClick={() => fuelFileRef.current?.click()} className="bg-emerald-600 disabled:bg-natural-muted text-white rounded-lg px-2.5 py-1 text-xs font-semibold flex items-center gap-1"><UploadCloud className="w-3.5 h-3.5" />{impFuel ? 'กำลังนำเข้า...' : 'นำเข้าค่าน้ำมัน'}</button>
        </div>
        <div className="flex flex-wrap gap-2 mb-3">
          <input list="plates" aria-label="ทะเบียนรถ" placeholder="ทะเบียน" value={fForm.plateNo} onChange={(e) => setFForm({ ...fForm, plateNo: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-28" />
          <input aria-label="เลขใบสั่งเติม" placeholder="เลขใบสั่งเติม" value={fForm.refNo} onChange={(e) => setFForm({ ...fForm, refNo: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-32" />
          <input type="date" aria-label="วันที่เติมน้ำมัน" value={fForm.date} onChange={(e) => setFForm({ ...fForm, date: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm" />
          <input type="number" aria-label="จำนวนเงินค่าน้ำมัน" placeholder="จำนวนเงิน" value={fForm.amount || ''} onChange={(e) => setFForm({ ...fForm, amount: +e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 text-sm w-28" />
          <button onClick={addFuel} className="bg-brand-red text-white rounded-lg px-3 text-sm font-semibold">เพิ่ม</button>
        </div>
        <SimpleTable rows={fuelF.map((f: FuelEntry) => [f.plateNo, f.refNo, f.date, money(f.amount)])} cols={['ทะเบียน', 'ใบสั่งเติม', 'วันที่', 'จำนวน']}
          footer={[`รวม ${fuelF.length} รายการ`, '', '', `฿${money(fuelSum)}`]}
          onDelete={async (i: number) => { await api(`/api/fuel/${fuelF[i].id}`, 'DELETE'); reload(); }} />
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
        <SimpleTable rows={incomesF.map((d: DeductionEntry) => [d.plateNo, d.label, d.docNo || '-', `+${money(d.amount)}`])} cols={['ทะเบียน', 'รายการ', 'ใบกระจาย', 'จำนวนเพิ่ม']}
          onDelete={async (i: number) => { await api(`/api/deductions/${incomesF[i].id}`, 'DELETE'); reload(); }} />
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
          <button onClick={() => addEntry(dForm.plateNo, dForm.categoryId, dForm.amount, 'deduction', dForm.docNo, () => setDForm({ plateNo: '', categoryId: '', amount: 0, docNo: '' }))} className="bg-brand-red text-white rounded-lg px-3 text-sm font-semibold">เพิ่ม</button>
        </div>
        <SimpleTable rows={dedF.map((d: DeductionEntry) => [d.plateNo, d.label, d.docNo || '-', money(d.amount)])} cols={['ทะเบียน', 'รายการ', 'ใบกระจาย', 'จำนวน']}
          onDelete={async (i: number) => { await api(`/api/deductions/${dedF[i].id}`, 'DELETE'); reload(); }} />
      </Section>

      {/* จัดการประเภท — เพิ่มชื่อใน dropdown ได้เองโดยไม่ต้องแก้โค้ด */}
      <CategoryManager cats={cats} api={api} branchId={branchId} reload={reload} showToast={showToast} />

      <datalist id="plates">{db.vehicles.map((v: Vehicle) => <option key={v.id} value={v.plateNo} />)}</datalist>
    </div>
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
        <button onClick={add} className="bg-brand-red text-white rounded-lg px-3 text-sm font-semibold">เพิ่มประเภท</button>
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
          <thead className="bg-brand-navy text-white"><tr>
            {['ทะเบียน', 'คนขับ', 'รายได้', 'หัก 1%', 'ค่าน้ำมัน', '+ รายได้เพิ่ม', 'รวมรายการหัก', 'รับสุทธิ'].map((h) => <th key={h} className={`p-2 font-semibold ${h === 'ทะเบียน' || h === 'คนขับ' ? 'text-left' : 'text-right'}`}>{h}</th>)}
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
      <div className="bg-emerald-50 border border-emerald-200 rounded-xl px-4 py-2 text-sm font-semibold text-emerald-800 flex flex-wrap items-center gap-2">
        <Building2 className="w-4 h-4" /> ภาพรวมทุกสาขา · {cycle.name}
        <button type="button" onClick={() => exportBranchSummary(cycle.name, rows.map((r) => ({ branchName: r.branch.name, docs: r.docs, trucks: r.trucks, trip: r.trip, fuel: r.fuel, income: r.income, deduct: r.deduct, net: r.net })))}
          className="ml-auto bg-brand-red hover:bg-brand-red-hover text-white rounded-full px-3.5 py-1.5 text-xs font-bold flex items-center gap-1.5"><FileSpreadsheet className="w-3.5 h-3.5" />Export Excel</button>
      </div>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <Stat label="ใบกระจายรวม" value={`${g.docs}`} />
        <Stat label="ค่าเที่ยวรวมทุกสาขา" value={`฿${money(g.trip)}`} />
        <Stat label="ค่าน้ำมันรวม" value={`฿${money(g.fuel)}`} />
        <Stat label="รับสุทธิรวม" value={`฿${money(g.net)}`} highlight />
      </div>

      <div className="bg-white rounded-2xl border border-natural-border overflow-x-auto">
        <table className="w-full text-xs min-w-[760px]">
          <thead className="bg-brand-navy text-white"><tr>
            {['สาขา', 'ใบกระจาย', 'รถ', 'ค่าเที่ยว', 'ค่าน้ำมัน', '+รายได้เพิ่ม', 'รวมหัก', 'รับสุทธิ', 'สัดส่วน'].map((h) => <th key={h} className={`p-2 font-semibold ${h === 'สาขา' || h === 'สัดส่วน' ? 'text-left' : 'text-right'}`}>{h}</th>)}
          </tr></thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={r.branch.id} className={i % 2 ? 'bg-[#F9FAFC]' : ''}>
                <td className="p-2 font-semibold text-brand-navy">{r.branch.name}</td>
                <td className="p-2 text-right">{r.docs}</td>
                <td className="p-2 text-right">{r.trucks}</td>
                <td className="p-2 text-right">{money(r.trip)}</td>
                <td className="p-2 text-right text-rose-700">{money(r.fuel)}</td>
                <td className="p-2 text-right text-emerald-700">{money(r.income)}</td>
                <td className="p-2 text-right text-rose-700">{money(r.deduct)}</td>
                <td className="p-2 text-right font-bold text-[#C00000]">{money(r.net)}</td>
                <td className="p-2 w-40">
                  <div className="flex items-center gap-2">
                    <div className="flex-1 bg-natural-bg rounded-full h-2 overflow-hidden"><div className="bg-brand-navy h-2 rounded-full" style={{ width: `${maxTrip > 0 ? (r.trip / maxTrip) * 100 : 0}%` }} /></div>
                    <span className="text-[11px] font-bold text-brand-navy w-9 text-right">{g.trip > 0 ? Math.round((r.trip / g.trip) * 100) : 0}%</span>
                  </div>
                </td>
              </tr>
            ))}
            <tr className="border-t-2 border-brand-navy font-bold bg-[#FFF2CC]">
              <td className="p-2">รวมทุกสาขา</td>
              <td className="p-2 text-right">{g.docs}</td>
              <td className="p-2"></td>
              <td className="p-2 text-right">{money(g.trip)}</td>
              <td className="p-2 text-right">{money(g.fuel)}</td>
              <td className="p-2"></td>
              <td className="p-2"></td>
              <td className="p-2 text-right text-[#C00000]">{money(g.net)}</td>
              <td className="p-2 text-right text-[11px] font-bold text-brand-navy">100%</td>
            </tr>
          </tbody>
        </table>
      </div>
      <p className="text-[11px] text-natural-muted">💡 "สัดส่วน" = ส่วนแบ่ง<b>ค่าเที่ยว</b>ของแต่ละสาขา (% ของค่าเที่ยวรวมทุกสาขา) · แถบยาว = สาขาที่ค่าเที่ยวมากสุด · เลือกสาขาที่มุมบนเพื่อดู/แก้ไขรายละเอียด</p>
    </div>
  );
}

// ===========================================================================
// Tab: วิเคราะห์รายได้ พขร (Driver KPI) — เฉพาะ HQ · มีรหัสเปิดแยก
// ===========================================================================
const KPI_PW = '2468';
const KPI_DEFAULT_ITEMS = [
  { label: 'เงินเดือน', amount: 15000 }, { label: 'ค่าผ่อนรถ', amount: 7000 },
  { label: 'ค่าเช่าบ้าน', amount: 2500 }, { label: 'ค่าดูแลครอบครัว', amount: 5000 },
];
const KPI_DEFAULT = { boxPerDay: 350, daysPerMonth: 25, items: KPI_DEFAULT_ITEMS, pw: KPI_PW };
function DriverKpiTab({ db, cycle }: any) {
  const [unlocked, setUnlocked] = useState(false);
  const [pwInput, setPwInput] = useState('');
  const [pwErr, setPwErr] = useState(false);
  const [cfg, setCfg] = useState<any>(() => {
    try {
      const s = JSON.parse(localStorage.getItem('kpiCfg') || '{}');
      const base: any = { ...KPI_DEFAULT, ...s };
      if (!Array.isArray(base.items)) {
        // migrate รูปแบบเก่า (salary/carPayment/rent/family) -> items[]
        base.items = (s.salary || s.carPayment || s.rent || s.family)
          ? [{ label: 'เงินเดือน', amount: s.salary || 0 }, { label: 'ค่าผ่อนรถ', amount: s.carPayment || 0 }, { label: 'ค่าเช่าบ้าน', amount: s.rent || 0 }, { label: 'ค่าดูแลครอบครัว', amount: s.family || 0 }]
          : KPI_DEFAULT_ITEMS.map((x) => ({ ...x }));
      }
      return base;
    } catch { return { ...KPI_DEFAULT, items: KPI_DEFAULT_ITEMS.map((x) => ({ ...x })) }; }
  });
  useEffect(() => { localStorage.setItem('kpiCfg', JSON.stringify(cfg)); }, [cfg]);
  const [openBranch, setOpenBranch] = useState<Record<string, boolean>>({});
  const [incomeMode, setIncomeMode] = useState<'net' | 'gross'>('net');

  if (!cycle) return <EmptyHint text="กรุณาเลือกรอบก่อน" />;

  if (!unlocked) {
    const tryUnlock = () => { if (pwInput === (cfg.pw || KPI_PW)) { setUnlocked(true); setPwErr(false); } else setPwErr(true); };
    return (
      <div className="max-w-sm mx-auto mt-10 bg-white rounded-2xl border border-natural-border shadow-lg p-8 flex flex-col items-center">
        <div className="w-14 h-14 bg-brand-red text-white rounded-2xl flex items-center justify-center mb-3"><Lock className="w-7 h-7" /></div>
        <h2 className="font-bold text-brand-navy">วิเคราะห์รายได้ พขร</h2>
        <p className="text-xs text-natural-muted mb-4">หน้านี้เป็นความลับ — ใส่รหัสเพื่อเปิด</p>
        <input type="password" value={pwInput} autoFocus onChange={(e) => { setPwInput(e.target.value); setPwErr(false); }}
          onKeyDown={(e) => e.key === 'Enter' && tryUnlock()} placeholder="รหัสผ่าน" aria-label="รหัสผ่าน"
          className={`w-full border rounded-lg px-3 py-2 text-sm mb-2 outline-none ${pwErr ? 'border-rose-400 bg-rose-50' : 'border-natural-border focus:border-brand-navy'}`} />
        {pwErr && <p className="text-xs text-rose-600 mb-2 self-start">รหัสไม่ถูกต้อง</p>}
        <button type="button" onClick={tryUnlock} className="w-full bg-brand-red text-white rounded-lg py-2.5 text-sm font-bold">เปิดดู</button>
      </div>
    );
  }

  const bx = (n: number) => Math.round(n).toLocaleString('th-TH');
  const items: { label: string; amount: number }[] = cfg.items || [];
  const setItem = (i: number, patch: any) => setCfg({ ...cfg, items: items.map((it, j) => (j === i ? { ...it, ...patch } : it)) });
  const addItem = () => setCfg({ ...cfg, items: [...items, { label: 'รายการใหม่', amount: 0 }] });
  const delItem = (i: number) => setCfg({ ...cfg, items: items.filter((_, j) => j !== i) });
  const boxesPerMonth = Math.max(1, cfg.boxPerDay * cfg.daysPerMonth);
  const totalExpense = items.reduce((a, it) => a + (+it.amount || 0), 0);
  const targetPerBox = totalExpense / boxesPerMonth;
  const targetPerCycle = totalExpense / 2;
  const targetBoxCycle = boxesPerMonth / 2;

  const branches = (db.branches as Branch[]).filter((b) => !b.isHQ && b.status === 'active');
  const drivers: any[] = [];
  const tripCountBy = new Map<string, number>();
  for (const b of branches) {
    const bTrips = db.tripDocuments.filter((t: TripDocument) => t.cycleId === cycle.id && t.branchId === b.id);
    if (!bTrips.length) continue;
    tripCountBy.set(b.name, bTrips.length);
    const sums = summarizeByVehicle(cycle.id,
      db.tripDocuments.filter((t: TripDocument) => t.branchId === b.id),
      db.fuelEntries.filter((f: FuelEntry) => f.branchId === b.id),
      db.deductions.filter((d: DeductionEntry) => d.branchId === b.id),
      db.vehicles.filter((v: Vehicle) => v.branchId === b.id));
    for (const s of sums) {
      const boxes = bTrips.filter((t: TripDocument) => normPlate(t.plateNo) === normPlate(s.plateNo)).reduce((a: number, t: TripDocument) => a + (t.totalQty || 0), 0);
      if (boxes <= 0 && s.totalTripAmount <= 0) continue;
      drivers.push({ plate: s.plateNo, driver: s.driverName, branch: b.name, boxes, net: s.netReceive, trip: s.totalTripAmount });
    }
  }
  drivers.sort((a, b) => b.net - a.net);

  const byBranch = new Map<string, { branch: string; boxes: number; trip: number; net: number; nDrivers: number; nHit: number }>();
  for (const d of drivers) {
    const g = byBranch.get(d.branch) || { branch: d.branch, boxes: 0, trip: 0, net: 0, nDrivers: 0, nHit: 0 };
    g.boxes += d.boxes; g.trip += d.trip; g.net += d.net; g.nDrivers++; if (d.boxes >= targetBoxCycle) g.nHit++;
    byBranch.set(d.branch, g);
  }
  const branchRows = [...byBranch.values()].map((g) => {
    const nTrips = tripCountBy.get(g.branch) || 0;
    const income = incomeMode === 'net' ? g.net : g.trip;
    return { ...g, nTrips, income, perBox: g.boxes > 0 ? income / g.boxes : 0, boxPerTrip: nTrips > 0 ? g.boxes / nTrips : 0, boxPerDriver: g.nDrivers > 0 ? g.boxes / g.nDrivers : 0, pctHit: g.nDrivers > 0 ? g.nHit / g.nDrivers * 100 : 0 };
  }).filter((g) => g.boxes > 0).sort((a, b) => a.perBox - b.perBox);

  // จัดกลุ่มคนรถตามสาขา (สำหรับตารางย่อ/ขยาย) เรียงตาม %ถึงเป้า มาก->น้อย
  const driverGroups = (() => {
    const m = new Map<string, { branch: string; list: any[]; boxes: number; net: number; gross: number }>();
    for (const d of drivers) {
      const g = m.get(d.branch) || { branch: d.branch, list: [], boxes: 0, net: 0, gross: 0 };
      g.list.push(d); g.boxes += d.boxes; g.net += d.net; g.gross += d.trip;
      m.set(d.branch, g);
    }
    return [...m.values()].map((g) => {
      const tgt = g.list.length * targetBoxCycle;
      return { ...g, pct: tgt > 0 ? g.boxes / tgt * 100 : 0, nHit: g.list.filter((d) => d.boxes >= targetBoxCycle).length };
    }).sort((a, b) => b.pct - a.pct);
  })();
  const expandAll = () => setOpenBranch(Object.fromEntries(driverGroups.map((g) => [g.branch, true])));
  const collapseAll = () => setOpenBranch({});

  const statusOf = (boxes: number) => {
    const p = boxes / Math.max(1, targetBoxCycle);
    return p >= 1 ? { t: '🟢 ถึงเป้า', c: 'text-emerald-700' } : p >= 0.8 ? { t: '🟡 ใกล้เป้า', c: 'text-amber-600' } : { t: '🔴 ต่ำกว่าเป้า', c: 'text-rose-600' };
  };
  const numIn = (v: number, set: (n: number) => void) => (<input type="number" aria-label="ตั้งค่า" value={v || ''} onChange={(e) => set(+e.target.value || 0)} className="w-24 border border-natural-border rounded px-2 py-1 text-sm text-right" />);

  const doExport = () => exportDriverKpi(cycle.name, { totalExpense, boxesPerMonth, targetPerBox, targetPerCycle, targetBoxCycle, items }, drivers, branchRows);

  // ตาราง KPI สาขา (ใช้ซ้ำ 2 แบบ: เรียงค่ากระจายต่ำสุด / เรียงอัดงานเก่งสุด)
  const renderKpiTable = (title: string, sorted: typeof branchRows, metric: 'cost' | 'pack') => (
    <div className="bg-white rounded-2xl border border-natural-border overflow-x-auto">
      <div className="px-4 pt-3 font-bold text-brand-navy text-sm">{title}</div>
      <table className="w-full text-xs min-w-[920px] mt-2">
        <thead className="bg-brand-navy text-white"><tr>
          {['อันดับ', 'สาขา', 'กล่องรวม', 'เที่ยว', 'กล่อง/เที่ยว', 'กล่อง/คนรถ', incomeMode === 'net' ? 'รับสุทธิรวม' : 'ค่าเที่ยวรวม', 'ค่ากระจาย บาท/กล่อง', '%ถึงเป้า'].map((h, i) => <th key={i} className={`p-2 font-semibold ${i <= 1 ? 'text-left' : 'text-right'} ${(metric === 'pack' && i === 4) || (metric === 'cost' && i === 7) ? 'bg-brand-red/30' : ''}`}>{h}</th>)}
        </tr></thead>
        <tbody>
          {sorted.map((g, i) => (<tr key={g.branch} className={i % 2 ? 'bg-natural-secondary/40' : ''}>
            <td className="p-2 font-bold">{i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : i + 1}</td>
            <td className="p-2 font-semibold text-brand-navy">{g.branch}</td>
            <td className="p-2 text-right">{bx(g.boxes)}</td><td className="p-2 text-right">{bx(g.nTrips)}</td>
            <td className={`p-2 text-right font-semibold ${metric === 'pack' ? `bg-amber-50 ${i === 0 ? 'text-emerald-700 font-bold' : ''}` : ''}`}>{bx(g.boxPerTrip)}</td>
            <td className="p-2 text-right font-semibold">{bx(g.boxPerDriver)}</td>
            <td className="p-2 text-right">{money(g.income)}</td>
            <td className={`p-2 text-right font-bold ${metric === 'cost' ? `bg-amber-50 ${i === 0 ? 'text-emerald-700' : ''}` : ''}`}>฿{money(g.perBox)}</td>
            <td className={`p-2 text-right font-semibold ${g.pctHit >= 100 ? 'text-emerald-700' : g.pctHit >= 50 ? 'text-amber-600' : 'text-rose-600'}`}>{g.pctHit.toFixed(0)}% <span className="text-natural-muted font-normal">({g.nHit}/{g.nDrivers})</span></td>
          </tr>))}
          {sorted.length === 0 && <tr><td colSpan={9} className="p-6 text-center text-natural-muted">ยังไม่มีข้อมูล (เลือก "ทุกสาขา" ที่มุมบนเพื่อเทียบทุกสาขา)</td></tr>}
        </tbody>
      </table>
    </div>
  );
  const costRanked = branchRows; // เรียงค่ากระจายต่ำ->สูงอยู่แล้ว
  const packRanked = [...branchRows].sort((a, b) => b.boxPerTrip - a.boxPerTrip);

  return (
    <div className="flex flex-col gap-4">
      <div className="flex justify-end">
        <button type="button" onClick={doExport} disabled={!drivers.length}
          className="flex items-center gap-2 bg-brand-red hover:bg-brand-red-hover disabled:opacity-40 text-white rounded-lg px-4 py-2 text-sm font-bold">
          <FileSpreadsheet className="w-4 h-4" /> Export Excel
        </button>
      </div>
      <Section title="ตั้งค่าฐานรายได้ (ปรับเองได้ · บันทึกอัตโนมัติ)" icon={TrendingUp}>
        <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3 text-xs mb-4">
          <label className="flex items-center justify-between gap-2">กล่อง/วัน {numIn(cfg.boxPerDay, (n) => setCfg({ ...cfg, boxPerDay: n }))}</label>
          <label className="flex items-center justify-between gap-2">วันทำงาน/เดือน {numIn(cfg.daysPerMonth, (n) => setCfg({ ...cfg, daysPerMonth: n }))}</label>
          <div className="flex items-center justify-between gap-2 font-semibold text-brand-navy">กล่อง/เดือน <span>{bx(boxesPerMonth)}</span></div>
        </div>
        <div className="border border-natural-border rounded-xl p-3 bg-natural-secondary/30">
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs font-semibold text-brand-navy">รายการค่าใช้จ่าย/เดือน (เป้ารายได้คนรถ)</span>
            <button type="button" onClick={addItem} className="flex items-center gap-1 bg-brand-navy text-white rounded-lg px-2.5 py-1 text-xs font-semibold"><Plus className="w-3 h-3" /> เพิ่มรายการ</button>
          </div>
          <div className="flex flex-col gap-2">
            {items.map((it, i) => (
              <div key={i} className="flex items-center gap-2">
                <input type="text" aria-label="ชื่อรายการ" value={it.label} onChange={(e) => setItem(i, { label: e.target.value })}
                  className="flex-1 border border-natural-border rounded px-2 py-1 text-sm" placeholder="ชื่อรายการ เช่น ค่าโทรศัพท์" />
                <input type="number" aria-label="จำนวนเงิน" value={it.amount || ''} onChange={(e) => setItem(i, { amount: +e.target.value || 0 })}
                  className="w-28 border border-natural-border rounded px-2 py-1 text-sm text-right" placeholder="บาท" />
                <button type="button" onClick={() => delItem(i)} className="text-rose-500 hover:text-rose-700 p-1" aria-label="ลบรายการ"><Trash2 className="w-4 h-4" /></button>
              </div>
            ))}
            {items.length === 0 && <p className="text-xs text-natural-muted py-2 text-center">ยังไม่มีรายการ — กด "เพิ่มรายการ"</p>}
          </div>
          <div className="flex items-center justify-end gap-2 mt-2 pt-2 border-t border-natural-border text-sm font-bold text-brand-navy">
            รวมรายจ่าย/เดือน <span className="text-brand-red">฿{money(totalExpense)}</span>
          </div>
        </div>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-4">
          <Stat label="รายจ่ายเป้า/เดือน" value={`฿${money(totalExpense)}`} />
          <Stat label="เป้า บาท/กล่อง" value={`฿${money(targetPerBox)}`} highlight />
          <Stat label="เป้ารายรับ/รอบ" value={`฿${money(targetPerCycle)}`} />
          <Stat label="เป้ากล่อง/รอบ" value={bx(targetBoxCycle)} />
        </div>
      </Section>

      <div className="bg-white rounded-2xl border border-natural-border overflow-x-auto">
        <div className="px-4 pt-3 flex items-center justify-between flex-wrap gap-2">
          <span className="font-bold text-brand-navy text-sm">👤 รายได้ต่อคนรถ (จัดกลุ่มตามสาขา) · {cycle.name}</span>
          <div className="flex items-center gap-2 flex-wrap">
            <div className="flex rounded-lg border border-natural-border overflow-hidden text-xs">
              <button type="button" onClick={() => setIncomeMode('net')} className={`px-2.5 py-1 ${incomeMode === 'net' ? 'bg-brand-navy text-white font-semibold' : 'hover:bg-natural-secondary'}`}>รับสุทธิ</button>
              <button type="button" onClick={() => setIncomeMode('gross')} className={`px-2.5 py-1 border-l border-natural-border ${incomeMode === 'gross' ? 'bg-brand-navy text-white font-semibold' : 'hover:bg-natural-secondary'}`}>ค่าเที่ยว (ก่อนหัก)</button>
            </div>
            <button type="button" onClick={expandAll} className="text-xs border border-natural-border rounded-lg px-2.5 py-1 hover:bg-natural-secondary">ขยายทั้งหมด</button>
            <button type="button" onClick={collapseAll} className="text-xs border border-natural-border rounded-lg px-2.5 py-1 hover:bg-natural-secondary">ย่อทั้งหมด</button>
          </div>
        </div>
        <table className="w-full text-xs min-w-[760px] mt-2">
          <thead className="bg-brand-navy text-white"><tr>
            {['ทะเบียน', 'คนขับ', 'กล่องวิ่ง', '%เป้า', incomeMode === 'net' ? 'รับสุทธิ' : 'ค่าเที่ยว (ก่อนหัก)', 'บาท/กล่อง', 'สถานะ'].map((h, i) => <th key={i} className={`p-2 font-semibold ${i <= 1 ? 'text-left' : 'text-right'}`}>{h}</th>)}
          </tr></thead>
          <tbody>
            {driverGroups.map((g) => {
              const open = !!openBranch[g.branch];
              const gc = g.pct >= 100 ? 'text-emerald-700' : g.pct >= 80 ? 'text-amber-600' : 'text-rose-600';
              const gIncome = incomeMode === 'net' ? g.net : g.gross;
              return (
                <Fragment key={g.branch}>
                  <tr onClick={() => setOpenBranch({ ...openBranch, [g.branch]: !open })} className="cursor-pointer bg-brand-navy/5 hover:bg-brand-navy/10 font-bold text-brand-navy border-t border-natural-border">
                    <td className="p-2" colSpan={2}>{open ? '▾' : '▸'} 🏢 {g.branch} <span className="text-natural-muted font-normal">· {g.list.length} คน</span></td>
                    <td className="p-2 text-right">{bx(g.boxes)}</td>
                    <td className={`p-2 text-right ${gc}`}>{g.pct.toFixed(0)}%</td>
                    <td className="p-2 text-right">{money(gIncome)}</td>
                    <td className="p-2 text-right">{money(g.boxes > 0 ? gIncome / g.boxes : 0)}</td>
                    <td className={`p-2 text-right ${gc}`}>{g.nHit}/{g.list.length} ถึงเป้า</td>
                  </tr>
                  {open && g.list.map((d: any, i: number) => {
                    const st = statusOf(d.boxes); const pct = d.boxes / Math.max(1, targetBoxCycle) * 100;
                    const income = incomeMode === 'net' ? d.net : d.trip; const perBox = d.boxes > 0 ? income / d.boxes : 0;
                    return (<tr key={d.plate + i} className={i % 2 ? 'bg-natural-secondary/40' : ''}>
                      <td className="p-2 pl-6 font-semibold text-brand-navy">{d.plate}</td><td className="p-2">{d.driver}</td>
                      <td className="p-2 text-right">{bx(d.boxes)}</td><td className={`p-2 text-right font-semibold ${st.c}`}>{pct.toFixed(0)}%</td>
                      <td className="p-2 text-right font-bold">{money(income)}</td><td className="p-2 text-right">{money(perBox)}</td>
                      <td className={`p-2 text-right font-semibold ${st.c}`}>{st.t}</td>
                    </tr>);
                  })}
                </Fragment>
              );
            })}
            {driverGroups.length === 0 && <tr><td colSpan={7} className="p-6 text-center text-natural-muted">ยังไม่มีข้อมูลในรอบนี้</td></tr>}
          </tbody>
        </table>
      </div>

      {renderKpiTable('🏆 KPI สาขา — ค่ากระจาย บาท/กล่อง ต่ำสุด (ต้นทุนถูกสุด)', costRanked, 'cost')}
      {renderKpiTable('💪 KPI สาขา — อัดงานเก่งสุด (กล่อง/เที่ยว มากสุด)', packRanked, 'pack')}
      <p className="text-[11px] text-natural-muted">🔒 ความลับ (เฉพาะ HQ + รหัส) · 💡 ตาราง 1 เรียง <b>ค่ากระจาย/กล่อง ต่ำสุด</b> (ต้นทุนถูก) · ตาราง 2 เรียง <b>กล่อง/เที่ยว มากสุด</b> (อัดงานหนาแน่น) — คนละมุม อาจคนละสาขาชนะ · รหัสเริ่มต้น 2468 (เปลี่ยนได้ในโค้ด KPI_PW)</p>
    </div>
  );
}

// ===========================================================================
// Tab: พื้นที่ต้นทุนสูง — ต่อสาขา -> จังหวัด เรียงตาม บาท/กล่อง
// ===========================================================================
function computeCostAreas(db: any, cycle: BillingCycle, branches: Branch[]) {
  return branches.map((b) => {
    const trips = db.tripDocuments.filter((t: TripDocument) => t.cycleId === cycle.id && t.branchId === b.id);
    type Agg = { boxes: number; cost: number; docs: Set<string> };
    const provMap = new Map<string, Agg & { prov: string; dists: Map<string, Agg & { dist: string }> }>();
    const add = (prov: string, dist: string, boxes: number, cost: number, docId: string) => {
      const g = provMap.get(prov) || { prov, boxes: 0, cost: 0, docs: new Set<string>(), dists: new Map() };
      g.boxes += boxes; g.cost += cost; g.docs.add(docId);
      const dg = g.dists.get(dist) || { dist, boxes: 0, cost: 0, docs: new Set<string>() };
      dg.boxes += boxes; dg.cost += cost; dg.docs.add(docId); g.dists.set(dist, dg);
      provMap.set(prov, g);
    };
    // จุดหลักของใบ = อำเภอ/จังหวัดที่กล่อง "งานปกติ" มากสุด (กันที่อยู่ สนญ. ของเก็บคืน/ปลายทางแปลก)
    const pickPrimary = (list: any[]) => {
      const m = new Map<string, number>();
      for (const r of list) { const k = ((r.provinceRaw || '').trim()) + '|' + ((r.districtRaw || '').trim()); m.set(k, (m.get(k) || 0) + (r.totalQty || 0)); }
      let pk = '|', mx = -1;
      for (const [k, q] of m) if (q > mx) { mx = q; pk = k; }
      const [p, d] = pk.split('|');
      return { prov: p.trim() || '(ไม่ระบุจังหวัด)', dist: d.trim() || '(ไม่ระบุอำเภอ)' };
    };
    for (const t of trips) {
      const recs = t.receipts || [];
      const normalRecs = recs.filter((r: any) => (r.collectQty || 0) === 0 && (r.peatQty || 0) === 0);
      const primary = pickPrimary(normalRecs.length ? normalRecs : recs);
      if (t.rateType === 'piece') {
        // ราคาชิ้น: งานปกติคิดต่อใบรับ · เก็บคืน/Peat (ส่งคืน สนญ.) -> ยกไปจุดหลักของใบ
        for (const r of recs) {
          const isReturn = (r.collectQty || 0) > 0 || (r.peatQty || 0) > 0;
          const prov = isReturn ? primary.prov : ((r.provinceRaw || '').trim() || '(ไม่ระบุจังหวัด)');
          const dist = isReturn ? primary.dist : ((r.districtRaw || '').trim() || '(ไม่ระบุอำเภอ)');
          add(prov, dist, r.totalQty || 0, r.receiptAmount || 0, t.id);
        }
      } else {
        // ราคาเหมา: ยกค่าเที่ยว+กล่องทั้งใบไปจุดหลัก
        const boxes = recs.reduce((a: number, r: any) => a + (r.totalQty || 0), 0);
        add(primary.prov, primary.dist, boxes, t.tripAmount, t.id);
      }
    }
    const provs = [...provMap.values()].map((g) => ({
      prov: g.prov, boxes: g.boxes, cost: g.cost, docs: g.docs.size, perBox: g.boxes > 0 ? g.cost / g.boxes : 0,
      dists: [...g.dists.values()].map((d) => ({ dist: d.dist, boxes: d.boxes, cost: d.cost, docs: d.docs.size, perBox: d.boxes > 0 ? d.cost / d.boxes : 0 })).sort((a, b) => b.perBox - a.perBox),
    })).filter((g) => g.boxes > 0 || g.cost > 0).sort((a, b) => b.perBox - a.perBox);
    const total = provs.reduce((a, p) => a + p.cost, 0);
    return { branch: b.name, branchId: b.id, provs, total };
  }).filter((d) => d.provs.length > 0).sort((a, b) => b.total - a.total);
}

function CostAreaTab({ db, cycle, branchId, showToast }: any) {
  const [open, setOpen] = useState<Record<string, boolean>>({});
  const [openProv, setOpenProv] = useState<Record<string, boolean>>({});
  const [thrStr, setThrStr] = useState('');
  if (!cycle) return <EmptyHint text="กรุณาเลือกรอบก่อน" />;
  const bx = (n: number) => Math.round(n).toLocaleString('th-TH');
  const thr = +thrStr || 0;
  const branches = (db.branches as Branch[]).filter((b) => !b.isHQ && b.status === 'active' && (!branchId || b.id === branchId));
  const data = computeCostAreas(db, cycle, branches);
  // กรองเฉพาะจังหวัดที่ บาท/กล่อง >= เกณฑ์
  const fdata = data.map((d) => ({ ...d, fprovs: thr > 0 ? d.provs.filter((p) => p.perBox >= thr) : d.provs })).filter((d) => d.fprovs.length > 0);
  const expandAll = () => setOpen(Object.fromEntries(fdata.map((d) => [d.branchId, true])));
  const collapseAll = () => { setOpen({}); setOpenProv({}); };
  const doExport = async () => {
    if (!fdata.length) return showToast('warning', 'ยังไม่มีข้อมูลในรอบนี้');
    try { await exportCostAreas(cycle.name, fdata.map((d) => ({ branch: d.branch, total: d.total, provs: d.fprovs })), thr); showToast('success', 'Export สำเร็จ — ดาวน์โหลดแล้ว'); }
    catch (e: any) { showToast('error', e.message); }
  };

  return (
    <div className="flex flex-col gap-4">
      <div className="bg-white rounded-2xl border border-natural-border p-4 flex flex-wrap items-center justify-between gap-3">
        <div className="text-sm"><span className="font-bold text-brand-navy">พื้นที่ต้นทุนสูง</span> <span className="text-natural-muted ml-2">รอบ {cycle.name} · เรียงปลายทางตาม บาท/กล่อง มาก→น้อย</span></div>
        <div className="flex items-center gap-2 flex-wrap">
          <div className="flex items-center gap-1 border border-natural-border rounded-lg px-2 py-1 text-xs">
            <span className="text-natural-muted">บาท/กล่อง ≥</span>
            <input type="number" aria-label="เกณฑ์ บาท/กล่อง" value={thrStr} onChange={(e) => setThrStr(e.target.value)} placeholder="ทั้งหมด" className="w-16 text-right outline-none" />
            {thr > 0 && <button type="button" onClick={() => setThrStr('')} className="text-rose-500 font-bold px-1" title="ล้างตัวกรอง">✕</button>}
          </div>
          <button type="button" onClick={expandAll} className="text-xs border border-natural-border rounded-lg px-2.5 py-1.5 hover:bg-natural-secondary">ขยายทั้งหมด</button>
          <button type="button" onClick={collapseAll} className="text-xs border border-natural-border rounded-lg px-2.5 py-1.5 hover:bg-natural-secondary">ย่อทั้งหมด</button>
          <button type="button" onClick={doExport} className="bg-brand-red hover:bg-brand-red-hover text-white rounded-lg px-4 py-2 text-sm font-bold flex items-center gap-1.5"><FileSpreadsheet className="w-4 h-4" />Export Excel</button>
        </div>
      </div>

      {fdata.length === 0 ? <EmptyHint text={thr > 0 ? `ไม่มีจังหวัดที่ บาท/กล่อง ≥ ${thr}` : 'ยังไม่มีข้อมูลในรอบนี้'} /> : fdata.map((d) => {
        const o = !!open[d.branchId];
        return (
          <div key={d.branchId} className="bg-white rounded-2xl border border-natural-border overflow-hidden">
            <button type="button" onClick={() => setOpen({ ...open, [d.branchId]: !o })} className="w-full flex items-center justify-between px-4 py-3 hover:bg-brand-navy/5 text-left">
              <span className="font-bold text-brand-navy">{o ? '▾' : '▸'} 🏢 {d.branch}</span>
              <span className="text-xs text-natural-muted">ต้นทุนรวม <b className="text-brand-navy">฿{money(d.total)}</b> · {d.fprovs.length}{thr > 0 ? `/${d.provs.length}` : ''} จังหวัด</span>
            </button>
            {o && (
              <div className="overflow-x-auto border-t border-natural-border">
                <table className="w-full text-xs min-w-[620px]">
                  <thead className="bg-brand-navy text-white"><tr>
                    {['ปลายทาง', 'กล่อง', 'จำนวนใบ', 'ค่าขนส่งรวม', 'บาท/กล่อง'].map((h, i) => <th key={h} className={`p-2 font-semibold ${i === 0 ? 'text-left' : 'text-right'}`}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {d.fprovs.map((p, i) => {
                      const hi = i === 0, lo = i === d.fprovs.length - 1 && d.fprovs.length > 1;
                      const pc = hi ? 'text-rose-600' : lo ? 'text-emerald-700' : '';
                      const pk = d.branchId + '|' + p.prov, po = !!openProv[pk];
                      return (
                        <Fragment key={p.prov}>
                          <tr onClick={() => setOpenProv({ ...openProv, [pk]: !po })} className={`cursor-pointer hover:bg-brand-navy/5 ${i % 2 ? 'bg-natural-secondary/40' : ''}`}>
                            <td className="p-2 font-semibold text-brand-navy">{po ? '▾' : '▸'} จ.{p.prov} <span className="text-natural-muted font-normal">({p.dists.length} อ.)</span></td>
                            <td className="p-2 text-right">{bx(p.boxes)}</td><td className="p-2 text-right">{p.docs}</td>
                            <td className="p-2 text-right">{money(p.cost)}</td>
                            <td className={`p-2 text-right font-bold ${pc}`}>฿{money(p.perBox)} {hi ? '🔴' : lo ? '🟢' : ''}</td>
                          </tr>
                          {po && p.dists.map((dd, k) => {
                            const dhi = k === 0, dlo = k === p.dists.length - 1 && p.dists.length > 1;
                            return (<tr key={dd.dist} className="bg-natural-secondary/20 text-natural-dark-muted">
                              <td className="p-2 pl-7">↳ อ.{dd.dist}</td>
                              <td className="p-2 text-right">{bx(dd.boxes)}</td><td className="p-2 text-right">{dd.docs}</td>
                              <td className="p-2 text-right">{money(dd.cost)}</td>
                              <td className={`p-2 text-right font-semibold ${dhi ? 'text-rose-600' : dlo ? 'text-emerald-700' : ''}`}>฿{money(dd.perBox)}</td>
                            </tr>);
                          })}
                        </Fragment>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        );
      })}
      <p className="text-[11px] text-natural-muted">💡 "ต้นทุน" = ค่าเที่ยวที่จ่ายส่งพื้นที่นั้น (ราคาชิ้น=ตามใบรับ · ราคาเหมา=เฉลี่ยตามสัดส่วนกล่อง) · กดแถวจังหวัดเพื่อดูรายอำเภอ · 🔴 แพงสุด/กล่อง 🟢 ถูกสุด/กล่อง</p>
    </div>
  );
}

// ===========================================================================
// Tab: แก้ปลายทาง — คีย์เวิร์ด -> จังหวัด/อำเภอจริง (ใบกระจายระบุผิด)
// ===========================================================================
function DestFixTab({ db, api, branchId, reload, showToast }: any) {
  const [form, setForm] = useState({ keyword: '', province: '', district: '', note: '' });
  if (!branchId) return <EmptyHint text={ALL_BRANCH_HINT} />;
  const list = (db.destinationOverrides || []).filter((d: DestinationOverride) => d.branchId === branchId);
  const inputCls = 'border border-natural-border rounded-lg px-2 py-1.5 text-sm';
  const add = async () => {
    if (!form.keyword.trim() || !form.province.trim()) return showToast('warning', 'กรอกคีย์เวิร์ดและจังหวัดจริง');
    await api('/api/destination-overrides', 'POST', { branchId, keyword: form.keyword.trim(), province: form.province.trim(), district: form.district.trim(), note: form.note.trim(), status: 'active' });
    showToast('success', 'เพิ่มแล้ว'); setForm({ keyword: '', province: '', district: '', note: '' }); reload();
  };
  return (
    <div className="flex flex-col gap-4">
      <Section title="แก้ปลายทาง (ใบกระจายระบุจังหวัด/อำเภอผิด)" icon={MapPin}>
        <p className="text-xs text-natural-dark-muted bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 mb-3 leading-relaxed">
          กรณีใบกระจายระบุปลายทางผิด เช่น เขียน <b>อุตรดิตถ์</b> แต่จริงๆ ส่ง <b>พิษณุโลก</b> — ตั้งคีย์เวิร์ด (เช่น <b>เจ๊ไพร</b>) ที่อยู่ใน<b>ชื่อผู้รับ</b> · <b>ชื่อผู้ส่ง</b> (งานเก็บคืน) · หรือ <b>บรรทัดโน้ต *...*</b> ระบบจะคิดราคาตาม<b>ปลายทางจริง</b>ให้อัตโนมัติ + ติดป้าย 📍 ทั้งในระบบและ Excel
        </p>
        <div className="flex flex-wrap gap-2 mb-3">
          <input value={form.keyword} onChange={(e) => setForm({ ...form, keyword: e.target.value })} placeholder="คีย์เวิร์ด เช่น เจ๊ไพร" aria-label="คีย์เวิร์ด" className={inputCls + ' w-44'} />
          <input value={form.province} onChange={(e) => setForm({ ...form, province: e.target.value })} placeholder="จังหวัดจริง เช่น พิษณุโลก" aria-label="จังหวัดจริง" className={inputCls + ' w-40'} />
          <input value={form.district} onChange={(e) => setForm({ ...form, district: e.target.value })} placeholder="อำเภอจริง เช่น เมือง" aria-label="อำเภอจริง" className={inputCls + ' w-36'} />
          <input value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} placeholder="หมายเหตุ (ถ้ามี)" aria-label="หมายเหตุ" className={inputCls + ' w-40'} />
          <button type="button" onClick={add} className="bg-brand-red hover:bg-brand-red-hover text-white rounded-lg px-4 py-1.5 text-sm font-bold">เพิ่ม</button>
        </div>
        <SimpleTable cols={['คีย์เวิร์ด', 'จังหวัดจริง', 'อำเภอจริง', 'หมายเหตุ']}
          rows={list.map((d: DestinationOverride) => [d.keyword, d.province, d.district, d.note || ''])}
          onDelete={async (i: number) => { await api(`/api/destination-overrides/${list[i].id}`, 'DELETE'); reload(); }} />
      </Section>
    </div>
  );
}

// ===========================================================================
// Tab: รายงานต่อทะเบียน (Export Excel)
// ===========================================================================
const fmtD = (s: string) => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ''); return m ? `${+m[3]}/${+m[2]}/${m[1]}` : (s || ''); };

function ReportsTab({ db, cycle, branchId, showToast }: any) {
  const [selPlate, setSelPlate] = useState('');
  if (!branchId) return <EmptyHint text={ALL_BRANCH_HINT} />;
  if (!cycle) return <EmptyHint text="กรุณาเลือกรอบก่อน" />;
  const branchName = (db.branches as Branch[]).find((b) => b.id === branchId)?.name || '';
  const cycleTrips = db.tripDocuments.filter((t: TripDocument) => t.cycleId === cycle.id);
  const sums = summarizeByVehicle(cycle.id, db.tripDocuments, db.fuelEntries, db.deductions, db.vehicles);
  const shownSums = selPlate ? sums.filter((s: any) => s.plateNo === selPlate) : sums;
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
        <div className="text-sm"><span className="font-bold text-brand-navy">รายงานต่อทะเบียน</span> <span className="text-natural-muted ml-2">รอบ {cycle.name} · สาขา {branchName} · {sums.length} ทะเบียน</span></div>
        <div className="flex items-center gap-2 ml-auto">
          <label className="text-xs text-natural-muted font-semibold">เลือกทะเบียน:</label>
          <select aria-label="เลือกทะเบียนรถ" value={selPlate} onChange={(e) => setSelPlate(e.target.value)} className="border border-natural-border rounded-full px-3 py-1.5 text-xs focus:outline-none focus:border-brand-navy">
            <option value="">🚚 ทุกคัน ({sums.length})</option>
            {sums.map((s: any) => <option key={s.plateNo} value={s.plateNo}>{s.plateNo} — {s.driverName}</option>)}
          </select>
          <button type="button" onClick={() => printTripDocuments([...cycleTrips].sort((a: TripDocument, b: TripDocument) => (a.plateNo || '').localeCompare(b.plateNo || '') || (a.documentDate || '').localeCompare(b.documentDate || '') || (a.documentNo || '').localeCompare(b.documentNo || '')), `ใบกระจายทั้งสาขา ${branchName}`)} disabled={!cycleTrips.length} className="bg-brand-navy disabled:opacity-40 text-white rounded-full px-4 py-2 text-xs font-semibold flex items-center gap-1.5">🖨️ พิมพ์ทุกคันทั้งสาขา ({cycleTrips.length})</button>
          <button onClick={exp} className="bg-emerald-600 text-white rounded-full px-4 py-2 text-xs font-semibold flex items-center gap-1.5"><FileSpreadsheet className="w-4 h-4" />Export Excel (.xlsx)</button>
        </div>
      </div>

      {sums.length === 0 ? <EmptyHint text="ยังไม่มีข้อมูลในรอบนี้" /> : (<>
        {/* สรุปรวม */}
        <div className="bg-white rounded-2xl border border-natural-border overflow-x-auto">
          <div className="px-4 pt-3 font-bold text-brand-navy text-sm">📊 สรุปรวมต่อทะเบียน</div>
          <table className="w-full text-xs min-w-[760px] mt-2">
            <thead className="bg-brand-navy text-white"><tr><TH>ทะเบียน</TH><TH>คนขับ</TH><TH r>รายได้</TH><TH r>หัก 1%</TH><TH r>ค่าน้ำมัน</TH><TH r>+รายได้เพิ่ม</TH><TH r>รวมหัก</TH><TH r>รับสุทธิ</TH></tr></thead>
            <tbody>
              {sums.map((s: any, i: number) => (
                <tr key={s.plateNo} onClick={() => setSelPlate(selPlate === s.plateNo ? '' : s.plateNo)} className={`cursor-pointer ${selPlate === s.plateNo ? 'bg-[#FFF2CC] ring-1 ring-brand-navy' : i % 2 ? 'bg-[#F9FAFC]' : ''} hover:bg-[#EAF2F8]`}>
                  <TD b>{s.plateNo}</TD><TD>{s.driverName}</TD><TD r>{money(s.totalTripAmount)}</TD><TD r>{money(s.deduction1Percent)}</TD><TD r>{money(s.fuelTotal)}</TD><TD r>{money(s.incomeAdd)}</TD><TD r>{money(s.deductionTotal)}</TD>
                  <td className="px-2 py-1 text-right font-bold text-[#C00000]">{money(s.netReceive)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* รายละเอียดต่อทะเบียน */}
        {shownSums.map((s: any) => {
          const np = normPlate(s.plateNo);
          const vTrips = cycleTrips.filter((t: TripDocument) => normPlate(t.plateNo) === np).sort((a: TripDocument, b: TripDocument) => (a.documentDate || '').localeCompare(b.documentDate || '') || (a.documentNo || '').localeCompare(b.documentNo || ''));
          const vFuel = db.fuelEntries.filter((f: FuelEntry) => f.cycleId === cycle.id && normPlate(f.plateNo) === np).sort((a: FuelEntry, b: FuelEntry) => (a.date || '').localeCompare(b.date || ''));
          const vIncome = db.deductions.filter((d: DeductionEntry) => d.cycleId === cycle.id && d.kind === 'income' && normPlate(d.plateNo) === np);
          const vDeduct = db.deductions.filter((d: DeductionEntry) => d.cycleId === cycle.id && d.kind === 'deduction' && normPlate(d.plateNo) === np);
          const byLabel = (list: DeductionEntry[]): { label: string; amount: number }[] => Object.values(list.reduce((m: any, d: DeductionEntry) => {
            const k = d.label || 'รายการ'; (m[k] = m[k] || { label: k, amount: 0 }).amount += d.amount; return m;
          }, {}));
          const inDocInc = byLabel(vIncome.filter((d: DeductionEntry) => normDoc(d.docNo || '')));
          const perCycleInc = byLabel(vIncome.filter((d: DeductionEntry) => !normDoc(d.docNo || '')));
          const dedLines = s.lines.filter((l: any) => l.kind === 'deduction');
          return (
            <div key={s.plateNo} className="bg-white rounded-2xl border border-natural-border p-4">
              <div className="flex items-center justify-between mb-2 gap-2 flex-wrap">
                <div className="font-bold text-brand-navy">🚚 {s.plateNo} <span className="text-natural-muted font-normal text-sm">· {s.driverName}</span></div>
                <button type="button" onClick={() => printTripDocuments(vTrips, `ใบกระจาย ${s.plateNo}`)} disabled={!vTrips.length}
                  className="text-xs border border-brand-navy text-brand-navy rounded-full px-3 py-1.5 font-semibold hover:bg-brand-navy hover:text-white disabled:opacity-40 whitespace-nowrap">🖨️ พิมพ์ใบกระจายทั้งหมด ({vTrips.length} ใบ)</button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs min-w-[720px]">
                  <thead className="bg-brand-navy text-white"><tr><TH>วันที่</TH><TH>ปลายทาง</TH><TH>เลขใบกระจาย</TH><TH r>จำนวน</TH><TH>แบบ</TH><TH r>ราคา</TH><TH r>เป็นเงิน</TH><TH r>พิเศษ</TH><TH r>รวม</TH><TH>หมายเหตุ</TH></tr></thead>
                  <tbody>
                    {vTrips.flatMap((t: TripDocument) => {
                      const docInc = vIncome.filter((d: DeductionEntry) => normDoc(d.docNo || '') && normDoc(d.docNo || '') === normDoc(t.documentNo || ''));
                      const docDed = vDeduct.filter((d: DeductionEntry) => normDoc(d.docNo || '') && normDoc(d.docNo || '') === normDoc(t.documentNo || ''));
                      const extra = docInc.reduce((a: number, d: DeductionEntry) => a + d.amount, 0);
                      return tripSubRows(t).map((sub, i) => {
                        const rowExtra = sub.first ? extra : 0;
                        const parts: React.ReactNode[] = [];
                        if (sub.hasDiv) parts.push(<span className="text-[#C65911] font-semibold">มีหาร</span>);
                        if (sub.first) {
                          if ((t.breakdown?.collect || 0) > 0) parts.push(<span className="text-emerald-700 font-semibold">🔄 เก็บคืน ฿{money(t.breakdown.collect)}</span>);
                          if ((t.breakdown?.peat || 0) > 0) parts.push(<span className="text-teal-700 font-semibold">🌱 Peat ฿{money(t.breakdown.peat)}</span>);
                          if ((t.breakdown?.addon || 0) > 0) parts.push(<span className="text-amber-700 font-semibold">📌 เหมาเพิ่ม ฿{money(t.breakdown.addon)}{(t.addonByDest || []).length ? ` (${(t.addonByDest || []).map((a) => a.dist || a.prov).filter(Boolean).join(', ')})` : ''}</span>);
                          docInc.forEach((d: DeductionEntry) => parts.push(<span className="text-emerald-700 font-semibold">➕{d.label} ฿{money(d.amount)}</span>));
                          docDed.forEach((d: DeductionEntry) => parts.push(<span className="text-rose-600 font-semibold">➖{d.label} ฿{money(d.amount)}</span>));
                        }
                        return (
                          <tr key={t.id + '-' + i} className={(t.breakdown?.collect || 0) > 0 && sub.first ? 'bg-emerald-50' : (t.breakdown?.addon || 0) > 0 && sub.first ? 'bg-amber-50' : sub.hasDiv ? 'bg-[#FFF2CC]' : ''}>
                            <TD>{fmtD(sub.date)}</TD><TD>{sub.dest}</TD><TD>{sub.docNo}</TD>
                            <TD r>{qtyFmt(sub.qty)}</TD><TD>{sub.rateType === 'piece' ? 'ชิ้น' : sub.rateType === 'flat' ? 'เหมา' : '-'}</TD>
                            <TD r>{sub.price != null ? money(sub.price) : (sub.rateType === 'piece' ? <span className="text-natural-muted">หลายราคา</span> : '-')}</TD><TD r>{money(sub.amount)}</TD><TD r>{rowExtra ? money(rowExtra) : '-'}</TD>
                            <td className="px-2 py-1 text-right font-bold text-brand-navy">{money(sub.amount + rowExtra)}</td>
                            <TD>{parts.map((p, j) => <span key={j}>{j > 0 && <span className="text-natural-muted"> · </span>}{p}</span>)}</TD>
                          </tr>
                        );
                      });
                    })}
                    {vTrips.length === 0 && <tr><td colSpan={10} className="px-2 py-3 text-center text-natural-muted">ไม่มีใบกระจาย</td></tr>}
                  </tbody>
                </table>
              </div>
              {/* สรุปย่อ + น้ำมัน */}
              <div className="grid md:grid-cols-2 gap-4 mt-3">
                <div className="text-xs">
                  <div className="font-semibold text-brand-navy mb-1">สรุป</div>
                  <div className="flex justify-between"><span>รายได้ค่าเที่ยว</span><b>{money(s.totalTripAmount)}</b></div>
                  {inDocInc.map((l) => <div key={l.label} className="flex justify-between text-emerald-700"><span>+ {l.label} (ในใบ)</span><span>+{money(l.amount)}</span></div>)}
                  {perCycleInc.map((l) => <div key={l.label} className="flex justify-between text-emerald-700"><span>+ {l.label}</span><span>+{money(l.amount)}</span></div>)}
                  <div className="flex justify-between text-rose-700"><span>หัก 1%</span><span>-{money(s.deduction1Percent)}</span></div>
                  <div className="flex justify-between text-rose-700"><span>หักค่าน้ำมัน</span><span>-{money(s.fuelTotal)}</span></div>
                  {dedLines.map((ln: any) => <div key={ln.categoryId} className="flex justify-between text-rose-700"><span>หัก {ln.label}</span><span>-{money(ln.amount)}</span></div>)}
                  <div className="flex justify-between border-t border-natural-border mt-1 pt-1 font-bold text-[#C00000]"><span>รวมรับสุทธิ</span><span>{money(s.netReceive)}</span></div>
                </div>
                <div className="text-xs">
                  <div className="font-semibold text-brand-navy mb-1">ใบสั่งเติมน้ำมัน</div>
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
  const [fProv, setFProv] = useState('');
  const [fDist, setFDist] = useState('');
  const [fCat, setFCat] = useState('');
  const [fType, setFType] = useState('');
  const [editId, setEditId] = useState<string | null>(null);
  const rateFileRef = useRef<HTMLInputElement>(null);
  const [importing, setImporting] = useState(false);
  const [batch, setBatch] = useState(false);
  const [pending, setPending] = useState<Record<string, { price?: number; pieceThreshold?: number | null }>>({});
  const [resetKey, setResetKey] = useState(0);
  const onImportRates = async (files: FileList) => {
    const file = files[0];
    if (!file) return;
    if (!/\.xlsx?$/i.test(file.name)) return showToast('error', 'รองรับเฉพาะไฟล์ Excel (.xls/.xlsx)');
    const existing = (db.rateMasters as RateMaster[]).filter((r) => (r.productCategory || 'normal') === 'normal').length;
    if (existing > 0) {
      const ok = await confirmAction({ title: 'นำเข้าราคาจาก Excel', text: `จะลบราคาเดิมของสาขานี้เฉพาะหมวดที่อยู่ในไฟล์ (งานปกติ + ชีต "พิเศษ" ถ้ามี เช่น เก็บคืน/Peat/บวกเพิ่ม) แล้วแทนที่ด้วยไฟล์นี้`, confirmText: 'ลบของเดิม + นำเข้า', danger: true });
      if (!ok) return;
    }
    const replaceExisting = existing > 0;
    const b64 = await new Promise<string>((resolve) => { const r = new FileReader(); r.onload = () => resolve((r.result as string).split(',')[1]); r.readAsDataURL(file); });
    setImporting(true);
    try {
      const res = await api('/api/import-rates', 'POST', { branchId, fileBase64: b64, replaceExisting });
      showToast('success', `นำเข้าราคาสำเร็จ — สร้าง ${res.created} แถว${res.removed ? ` (ลบเดิม ${res.removed})` : ''}`);
      if (res.summary?.length) alertBox('สรุปการนำเข้าราคา', res.summary.join('\n'));
      reload();
    } catch (e: any) { showToast('error', e.message); }
    finally { setImporting(false); if (rateFileRef.current) rateFileRef.current.value = ''; }
  };
  const add = async () => {
    if (!form.provinceName || !form.price) return showToast('warning', 'กรอกจังหวัดและราคา');
    // Peat mass คิดชิ้นเสมอ · บวกเพิ่มตายตัว = เหมา · เก็บคืน/งานปกติ เลือกเหมา/ชิ้นได้ (เก็บคืนเหมา = ใช้เมื่อเก็บคืนอย่างเดียว)
    const priceType = form.productCategory === 'peat_mass' ? 'piece' : form.productCategory === 'fixed_addon' ? 'flat' : form.priceType;
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
  const catLabel = (c?: string) => c === 'collect_back' ? 'เก็บคืน' : c === 'peat_mass' ? 'Peat mass' : c === 'fixed_addon' ? 'บวกเพิ่มตายตัว' : 'งานปกติ';
  if (!branchId) return <EmptyHint text={ALL_BRANCH_HINT} />;

  const allRates: RateMaster[] = db.rateMasters;
  const nm = (s: string) => (s || '').toLowerCase().replace(/\s+/g, '');
  const rates = allRates.filter((r) => {
    if (filterGroup !== '__all__' && (r.rateGroup || '') !== (filterGroup === '__none__' ? '' : filterGroup)) return false;
    if (fProv && !nm(r.provinceName).includes(nm(fProv))) return false;
    if (fDist && !nm(r.districtName).includes(nm(fDist)) && !nm(r.destinationName).includes(nm(fDist))) return false;
    if (fCat && (r.productCategory || 'normal') !== fCat) return false;
    if (fType && r.priceType !== fType) return false;
    return true;
  });
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

  // ---- โหมดแก้หลายช่อง (batch) ----
  const pendingCount = Object.keys(pending).length;
  const markPending = (r: RateMaster, field: 'price' | 'pieceThreshold', value: number | null) =>
    setPending((p) => ({ ...p, [r.id]: { ...p[r.id], [field]: value } }));
  const clearBatch = () => { setPending({}); setResetKey((k) => k + 1); };
  const saveAll = async () => {
    const ids = Object.keys(pending);
    if (!ids.length) return;
    try {
      for (const id of ids) {
        const r = (db.rateMasters as RateMaster[]).find((x) => x.id === id);
        if (!r) continue;
        const ch = pending[id];
        if (cycleMode) {
          const price = ch.price ?? effPrice(r);
          const pieceThreshold = ch.pieceThreshold !== undefined ? ch.pieceThreshold : effTh(r);
          await api('/api/rate-overrides/upsert', 'POST', { branchId, cycleId: cycle.id, rateMasterId: id, price, pieceThreshold });
        } else {
          const body: any = {};
          if (ch.price !== undefined) body.price = ch.price;
          if (ch.pieceThreshold !== undefined) body.pieceThreshold = ch.pieceThreshold;
          await api(`/api/rate-masters/${id}`, 'PUT', body);
        }
      }
      showToast('success', `บันทึก ${ids.length} ช่อง${cycleMode ? ` (เฉพาะรอบ ${cycle.name})` : ''} แล้ว`);
      setPending({}); reload();
    } catch (e: any) { showToast('error', e.message); }
  };

  return (
    <Section title="Master ราคาขนส่ง" icon={Tag}>
      {/* นำเข้า/เทมเพลตราคาจาก Excel */}
      <div className="flex flex-wrap items-center gap-2 mb-3 bg-emerald-50 border border-emerald-200 rounded-xl p-2.5">
        <span className="text-xs font-semibold text-emerald-800">📥 นำเข้าราคาทั้งสาขาจาก Excel:</span>
        <button type="button" onClick={() => downloadRateTemplate()} className="bg-white border border-emerald-400 text-emerald-700 rounded-lg px-3 py-1.5 text-xs font-semibold flex items-center gap-1"><FileSpreadsheet className="w-3.5 h-3.5" />ดาวน์โหลดเทมเพลต</button>
        <input ref={rateFileRef} type="file" aria-label="นำเข้าราคา Excel" accept=".xls,.xlsx" className="hidden" onChange={(e) => e.target.files && onImportRates(e.target.files)} />
        <button type="button" disabled={importing || !branchId} onClick={() => rateFileRef.current?.click()} className="bg-emerald-600 disabled:bg-natural-muted text-white rounded-lg px-3 py-1.5 text-xs font-semibold flex items-center gap-1"><UploadCloud className="w-3.5 h-3.5" />{importing ? 'กำลังนำเข้า...' : 'นำเข้าราคา (.xlsx)'}</button>
        <span className="text-[11px] text-emerald-700/80">เหมาต่ออำเภอ + ชิ้นต่อจังหวัด/อำเภอ · ระบบเทียบ max ให้อัตโนมัติ</span>
      </div>

      {/* สวิตช์ ราคาหลัก / ราคาเฉพาะรอบ */}
      <div className="flex flex-wrap items-center gap-2 mb-3 text-xs">
        <button onClick={() => setMode('base')} className={`px-3 py-1.5 rounded-full font-semibold border ${mode === 'base' ? 'bg-brand-navy text-white border-brand-navy' : 'border-natural-border text-natural-muted'}`}>ราคาหลัก (ทุกรอบ)</button>
        <button onClick={() => setMode('cycle')} disabled={!cycle} className={`px-3 py-1.5 rounded-full font-semibold border disabled:opacity-40 ${mode === 'cycle' ? 'bg-amber-500 text-white border-amber-500' : 'border-natural-border text-natural-muted'}`}>
          🏷️ ราคาเฉพาะรอบ{cycle ? `: ${cycle.name}` : ' (เลือกรอบก่อน)'}
        </button>
        {cycleMode && <span className="text-amber-700">แก้ราคา/จุดตัด = ทับเฉพาะรอบนี้ · รอบอื่นใช้ราคาหลัก</span>}
        <button type="button" onClick={() => { if (batch && pendingCount) clearBatch(); setBatch(!batch); }}
          className={`px-3 py-1.5 rounded-full font-semibold border ${batch ? 'bg-violet-600 text-white border-violet-600' : 'border-natural-border text-natural-muted'}`}>
          ✏️ แก้หลายช่อง{batch ? ' (เปิด)' : ''}
        </button>
        {batch && (
          <span className="flex items-center gap-2">
            <button type="button" onClick={saveAll} disabled={!pendingCount}
              className="bg-emerald-600 disabled:opacity-40 text-white rounded-full px-4 py-1.5 font-bold flex items-center gap-1"><Save className="w-3.5 h-3.5" />บันทึกทั้งหมด ({pendingCount})</button>
            {pendingCount > 0 && <button type="button" onClick={clearBatch} className="text-rose-600 underline">ยกเลิก</button>}
          </span>
        )}
        {branchGroups.length > 0 && (
          <label className="flex items-center gap-1 text-xs text-natural-muted ml-auto">
            กรองกลุ่ม:
            <select aria-label="กรองกลุ่มราคา" value={filterGroup} onChange={(e) => setFilterGroup(e.target.value)} className="border border-natural-border rounded-lg px-2 py-1 text-xs font-semibold text-brand-navy">
              <option value="__all__">ทุกกลุ่ม</option>
              {branchGroups.map((g) => <option key={g} value={g}>{g}</option>)}
              <option value="__none__">— ไม่ระบุกลุ่ม (ใช้ร่วม) —</option>
            </select>
          </label>
        )}
      </div>

      {/* ตัวกรองตารางราคา */}
      <div className="flex flex-wrap items-center gap-2 mb-3 text-xs bg-natural-secondary/60 border border-natural-border rounded-lg p-2">
        <span className="text-natural-muted font-semibold flex items-center gap-1"><Filter className="w-3.5 h-3.5" />กรอง:</span>
        <input aria-label="กรองจังหวัด" placeholder="จังหวัด" value={fProv} onChange={(e) => setFProv(e.target.value)} className="border border-natural-border rounded-lg px-2 py-1 w-28" />
        <input aria-label="กรองอำเภอ/ปลายทาง" placeholder="อำเภอ / ปลายทาง" value={fDist} onChange={(e) => setFDist(e.target.value)} className="border border-natural-border rounded-lg px-2 py-1 w-36" />
        <select aria-label="กรองหมวด" value={fCat} onChange={(e) => setFCat(e.target.value)} className="border border-natural-border rounded-lg px-2 py-1">
          <option value="">ทุกหมวด</option>
          <option value="normal">งานปกติ</option>
          <option value="collect_back">เก็บคืน</option>
          <option value="peat_mass">Peat mass</option>
          <option value="fixed_addon">บวกเพิ่มตายตัว</option>
        </select>
        <select aria-label="กรองประเภท" value={fType} onChange={(e) => setFType(e.target.value)} className="border border-natural-border rounded-lg px-2 py-1">
          <option value="">ทุกประเภท</option>
          <option value="flat">เหมา</option>
          <option value="piece">ชิ้น</option>
        </select>
        {(fProv || fDist || fCat || fType) && (
          <button type="button" onClick={() => { setFProv(''); setFDist(''); setFCat(''); setFType(''); }} className="text-natural-muted hover:text-rose-600 font-semibold underline">ล้างตัวกรอง</button>
        )}
        <span className="text-natural-muted ml-auto">{rates.length}/{allRates.length} รายการ</span>
      </div>

      <div className="flex flex-wrap gap-2 mb-3 text-sm items-center">
        <input aria-label="ปลายทาง" placeholder="ปลายทาง" value={form.destinationName} onChange={(e) => setForm({ ...form, destinationName: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-32" />
        <input aria-label="จังหวัด" placeholder="จังหวัด" value={form.provinceName} onChange={(e) => setForm({ ...form, provinceName: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-28" />
        <input aria-label="อำเภอ" placeholder="อำเภอ" value={form.districtName} onChange={(e) => setForm({ ...form, districtName: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-24" />
        <select aria-label="ประเภทสินค้า" value={form.productCategory} onChange={(e) => setForm({ ...form, productCategory: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5">
          <option value="normal">งานปกติ</option><option value="collect_back">เก็บคืน</option><option value="peat_mass">Peat mass</option><option value="fixed_addon">บวกเพิ่มตายตัว (+700)</option>
        </select>
        <select aria-label="ประเภทราคา" value={form.priceType} onChange={(e) => setForm({ ...form, priceType: e.target.value })} disabled={form.productCategory === 'peat_mass' || form.productCategory === 'fixed_addon'} className="border border-natural-border rounded-lg px-2 py-1.5 disabled:opacity-50">
          <option value="flat">ราคาเหมา</option><option value="piece">ราคาชิ้น</option>
        </select>
        <input type="number" aria-label="ราคา" placeholder="ราคา" value={form.price || ''} onChange={(e) => setForm({ ...form, price: +e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-24" />
        <input type="number" aria-label="จุดตัดชิ้น" placeholder="จุดตัดชิ้น" title="≤จุดตัด=เหมา, >จุดตัด=ชิ้น (เว้นว่าง=ไม่ใช้)" value={form.pieceThreshold} onChange={(e) => setForm({ ...form, pieceThreshold: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-24" />
        <button onClick={add} className={`${editId ? 'bg-amber-600' : 'bg-brand-navy'} text-white rounded-lg px-3 py-1.5 font-semibold`}>{editId ? '✎ บันทึกแก้ไข' : 'เพิ่ม'}</button>
        {editId && <button onClick={cancelEdit} className="border border-natural-border rounded-lg px-3 py-1.5 text-sm font-semibold">ยกเลิก</button>}
        <button onClick={() => setAdv(!adv)} className="text-xs text-brand-navy font-semibold underline">{adv ? 'ซ่อน' : 'เงื่อนไขพิเศษ'}</button>
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
                      className="border border-natural-border rounded px-1 py-0.5 text-xs font-semibold text-brand-navy">
                      <option value="">ใช้ร่วม</option>
                      {branchGroups.map((g) => <option key={g} value={g}>{g}</option>)}
                    </select>
                  </td>
                )}
                <td className="py-1.5 px-1 font-semibold text-brand-navy">{r.destinationName}</td>
                <td className="py-1.5 px-1">{r.provinceName}</td>
                <td className="py-1.5 px-1">{r.districtName}</td>
                <td className="py-1.5 px-1">{(r.productCategory && r.productCategory !== 'normal') ? <span className="text-amber-700 font-semibold">{catLabel(r.productCategory)}</span> : <span className="text-natural-muted">ปกติ</span>}</td>
                <td className="py-1.5 px-1">{r.priceType === 'flat' ? 'เหมา' : 'ชิ้น'}</td>
                <td className="py-1.5 px-1">
                  <input type="number" key={`p-${r.id}-${cycleMode ? 'c' : 'b'}-${effPrice(r)}-${resetKey}`} defaultValue={effPrice(r)} aria-label={`ราคา ${r.destinationName}`}
                    onBlur={(e) => { const v = +e.target.value; if (v !== effPrice(r)) { batch ? markPending(r, 'price', v) : saveCell(r, 'price', v); } }}
                    className={`w-20 border rounded px-1 py-0.5 text-xs text-right outline-none ${pending[r.id]?.price !== undefined ? 'border-violet-500 bg-violet-50 ring-1 ring-violet-300' : cycleMode && ovFor(r) ? 'border-amber-400 bg-amber-50' : 'border-natural-border'} focus:border-brand-navy`} />
                  {pending[r.id]?.price !== undefined ? <span className="text-[9px] text-violet-700 ml-0.5 font-bold">รอบันทึก</span> : cycleMode && ovFor(r) && <span className="text-[9px] text-amber-700 ml-0.5">เฉพาะรอบ</span>}
                </td>
                <td className="py-1.5 px-1">
                  <input type="number" key={`t-${r.id}-${cycleMode ? 'c' : 'b'}-${effTh(r) ?? ''}-${resetKey}`} defaultValue={effTh(r) ?? ''} placeholder="-" aria-label={`จุดตัด ${r.destinationName}`}
                    onBlur={(e) => { const raw = e.target.value.trim(); const v = raw === '' ? null : +raw; if (v !== effTh(r)) { batch ? markPending(r, 'pieceThreshold', v) : saveCell(r, 'pieceThreshold', v); } }}
                    className={`w-16 border rounded px-1 py-0.5 text-xs text-right outline-none ${pending[r.id]?.pieceThreshold !== undefined ? 'border-violet-500 bg-violet-50 ring-1 ring-violet-300' : cycleMode && ovFor(r) ? 'border-amber-400 bg-amber-50' : 'border-natural-border'} focus:border-brand-navy`} />
                </td>
                <td className="py-1.5 px-1">{r.effectiveFrom}</td>
                <td className="py-1.5 px-1 whitespace-nowrap">
                  {cycleMode && ovFor(r) && <button type="button" title="กลับไปใช้ราคาหลัก" onClick={() => removeOverride(r)} className="text-amber-600 hover:text-amber-800 mr-1"><RefreshCw className="w-3.5 h-3.5 inline" /></button>}
                  <button type="button" title="แก้ไข (ใส่รหัส)" onClick={() => startEdit(r)} className="text-brand-navy hover:text-amber-700 font-semibold mr-2">✎</button>
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
  const blank = { ruleName: '', senderKeyword: 'ซีโน', receiverKeyword: '', provinceKeyword: '', receiverGroupId: db.receiverGroups[0]?.id || '', productKeyword: '', productSizeKeyword: '', divisor: 3, roundingMethod: 'half_up', applyLevel: 'receipt', status: 'active', effectiveFrom: '2020-01-01', effectiveTo: null };
  const [form, setForm] = useState<any>(blank);
  const [editId, setEditId] = useState<string | null>(null);
  const [fSender, setFSender] = useState('');
  const [fGroup, setFGroup] = useState('');
  const [newRecv, setNewRecv] = useState('');
  const [newSend, setNewSend] = useState('');
  const add = async () => {
    if (!form.productKeyword) return showToast('warning', 'กรอกชื่อสินค้า');
    if (editId) {
      await api(`/api/conversion-rules/${editId}`, 'PUT', { ...form, ruleName: form.ruleName || `${form.productKeyword} หาร ${form.divisor}` });
      showToast('success', 'บันทึกการแก้ไขแล้ว');
    } else {
      await api('/api/conversion-rules', 'POST', { ...form, branchId, ruleName: form.ruleName || `${form.productKeyword} หาร ${form.divisor}` });
      showToast('success', 'เพิ่มกฎแล้ว');
    }
    setForm(blank); setEditId(null); reload();
  };
  const startEdit = async (r: ProductConversionRule) => {
    if (!(await confirmPassword('แก้ไขกฎตัวหาร'))) return;
    setForm({ ruleName: r.ruleName, senderKeyword: r.senderKeyword, receiverKeyword: r.receiverKeyword || '', provinceKeyword: r.provinceKeyword || '', receiverGroupId: r.receiverGroupId, productKeyword: r.productKeyword, productSizeKeyword: r.productSizeKeyword, divisor: r.divisor, roundingMethod: r.roundingMethod, applyLevel: r.applyLevel, status: r.status, effectiveFrom: r.effectiveFrom, effectiveTo: r.effectiveTo });
    setEditId(r.id);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };
  const cancelEdit = () => { setForm(blank); setEditId(null); };
  const norm = (s: string) => (s || '').toLowerCase().replace(/\s+/g, '');
  const filteredRules: ProductConversionRule[] = db.conversionRules.filter((r: ProductConversionRule) =>
    (!fSender || norm(r.senderKeyword).includes(norm(fSender))) &&
    (!fGroup || r.receiverGroupId === fGroup)
  );
  // ค่าที่เคยใช้ (สำหรับ datalist combobox)
  const uniq = (arr: string[]) => [...new Set(arr.filter(Boolean))];
  const senderOpts = uniq(db.conversionRules.map((r: ProductConversionRule) => r.senderKeyword));
  const receiverOpts = uniq(db.conversionRules.map((r: ProductConversionRule) => r.receiverKeyword || ''));
  const productOpts = uniq(db.conversionRules.map((r: ProductConversionRule) => r.productKeyword));
  const sizeOpts = uniq(db.conversionRules.map((r: ProductConversionRule) => r.productSizeKeyword));
  // รายชื่อ "ที่หาร 3" รวมจากทุกกฎ — เพิ่ม/ลบที่นี่ sync เข้าทุกกฎ
  const splitKw = (s: string) => (s || '').split('|').map((x) => x.trim()).filter(Boolean);
  const allReceivers: string[] = [...new Set(db.conversionRules.flatMap((r: ProductConversionRule) => splitKw(r.receiverKeyword || '')))] as string[];
  const allSenders: string[] = [...new Set(db.conversionRules.flatMap((r: ProductConversionRule) => splitKw(r.senderKeyword || '')))] as string[];
  const syncName = async (field: 'receiverKeyword' | 'senderKeyword', name: string, op: 'add' | 'remove') => {
    const nm = name.trim();
    if (!nm) return;
    if (!db.conversionRules.length) return showToast('warning', 'ยังไม่มีกฎตัวหาร — เพิ่มกฎอย่างน้อย 1 ข้อก่อน');
    for (const r of db.conversionRules as ProductConversionRule[]) {
      const cur = splitKw((r as any)[field] || '');
      const has = cur.includes(nm);
      if (op === 'add' && !has) await api(`/api/conversion-rules/${r.id}`, 'PUT', { [field]: [...cur, nm].join('|') });
      if (op === 'remove' && has) await api(`/api/conversion-rules/${r.id}`, 'PUT', { [field]: cur.filter((x) => x !== nm).join('|') });
    }
    reload();
  };
  if (!branchId) return <EmptyHint text={ALL_BRANCH_HINT} />;
  return (
    <div className="flex flex-col gap-5">
      {/* 📋 รายชื่อที่หาร 3 (ใช้ร่วมทุกกฎ) */}
      <Section title="รายชื่อผู้รับ / ผู้ส่ง ที่หาร 3 (จดจำไว้ใช้ทุกกฎ)" icon={Filter}>
        <p className="text-xs text-natural-muted mb-3">เพิ่ม/ลบชื่อที่นี่ที่เดียว ระบบจะบันทึกเข้า <b>ทุกกฎตัวหาร</b> อัตโนมัติ ({db.conversionRules.length} กฎ) — ทำให้หาร 3 แม่นยำ ไม่ต้องไปแก้ทีละกฎ</p>
        <div className="grid md:grid-cols-2 gap-4">
          {/* ผู้รับ */}
          <div className="bg-natural-bg rounded-xl p-3">
            <div className="font-bold text-sm text-brand-navy mb-2">👤 ผู้รับที่หาร 3 ({allReceivers.length})</div>
            <div className="flex flex-wrap gap-1.5 mb-2 min-h-[28px]">
              {allReceivers.length === 0 && <span className="text-xs text-natural-muted">ยังไม่มี — ว่าง = หารทุกผู้รับ</span>}
              {allReceivers.map((n) => (
                <span key={n} className="inline-flex items-center gap-1 bg-white border border-brand-navy/30 rounded-full pl-2.5 pr-1 py-0.5 text-xs">
                  {n}
                  <button type="button" title="ลบ" onClick={() => syncName('receiverKeyword', n, 'remove')} className="text-natural-muted hover:text-rose-600 font-bold w-4">×</button>
                </span>
              ))}
            </div>
            <div className="flex gap-2">
              <input value={newRecv} onChange={(e) => setNewRecv(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { syncName('receiverKeyword', newRecv, 'add'); setNewRecv(''); } }} placeholder="พิมพ์ชื่อผู้รับ เช่น รำพึงรีสอร์ท" className="flex-1 border border-natural-border rounded-lg px-2 py-1.5 text-sm" />
              <button onClick={() => { syncName('receiverKeyword', newRecv, 'add'); setNewRecv(''); }} className="bg-brand-red text-white rounded-lg px-3 text-sm font-semibold">เพิ่ม</button>
            </div>
          </div>
          {/* ผู้ส่ง */}
          <div className="bg-natural-bg rounded-xl p-3">
            <div className="font-bold text-sm text-brand-navy mb-2">🚚 ผู้ส่งที่หาร 3 ({allSenders.length})</div>
            <div className="flex flex-wrap gap-1.5 mb-2 min-h-[28px]">
              {allSenders.length === 0 && <span className="text-xs text-natural-muted">ยังไม่มี</span>}
              {allSenders.map((n) => (
                <span key={n} className="inline-flex items-center gap-1 bg-white border border-brand-navy/30 rounded-full pl-2.5 pr-1 py-0.5 text-xs">
                  {n}
                  <button type="button" title="ลบ" onClick={() => syncName('senderKeyword', n, 'remove')} className="text-natural-muted hover:text-rose-600 font-bold w-4">×</button>
                </span>
              ))}
            </div>
            <div className="flex gap-2">
              <input value={newSend} onChange={(e) => setNewSend(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { syncName('senderKeyword', newSend, 'add'); setNewSend(''); } }} placeholder="พิมพ์ชื่อผู้ส่ง เช่น ซีโน" className="flex-1 border border-natural-border rounded-lg px-2 py-1.5 text-sm" />
              <button onClick={() => { syncName('senderKeyword', newSend, 'add'); setNewSend(''); }} className="bg-brand-red text-white rounded-lg px-3 text-sm font-semibold">เพิ่ม</button>
            </div>
          </div>
        </div>
        <p className="text-[11px] text-natural-muted mt-2">💡 ชื่อเฉพาะราย ใส่ชื่อให้ครบ (เช่น "รำพึงรีสอร์ท" ไม่ใช่ "รำพึง") · ชื่อแบรนด์ครอบทุกสาขาใส่คำแบรนด์ได้ (เช่น "ไทยลอตเต้")</p>
      </Section>

      <Section title="เงื่อนไขแปลงจำนวนสินค้า (ตัวหาร)" icon={Filter}>
        <p className="text-xs text-natural-muted mb-3">ต้องตรงทุกข้อ: ผู้ส่งเข้าคำ + (ผู้รับเข้าคำ ถ้าระบุ) + กลุ่มผู้รับตรง + ชื่อสินค้า + ขนาด → หารตามตัวหาร (คำนวณแยกตามเลขใบรับสินค้า)<br /><span className="text-brand-navy">💡 ช่อง "สินค้า" และ "ผู้รับ" ใส่หลายคำคั่นด้วย <b>|</b> ได้ เช่น <b>ยูบี|ยูปี</b> · ช่องผู้รับเว้นว่าง = ทุกผู้รับ</span></p>
        <div className="flex flex-wrap gap-2 mb-3 text-sm">
          <input list="rule-senders" aria-label="ผู้ส่ง (keyword)" placeholder="ผู้ส่ง (พิมพ์ใหม่ได้)" value={form.senderKeyword} onChange={(e) => setForm({ ...form, senderKeyword: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-32" />
          <input list="rule-receivers" aria-label="ผู้รับ (keyword)" placeholder="ผู้รับ (ว่าง=ทุกคน)" value={form.receiverKeyword} onChange={(e) => setForm({ ...form, receiverKeyword: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-36" />
          <select aria-label="กลุ่มผู้รับสินค้า" value={form.receiverGroupId} onChange={(e) => setForm({ ...form, receiverGroupId: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5">
            {db.receiverGroups.map((g: ReceiverGroup) => <option key={g.id} value={g.id}>{g.groupName}</option>)}
          </select>
          <input list="rule-products" aria-label="ชื่อสินค้า" placeholder="สินค้า (พิมพ์ใหม่ได้)" value={form.productKeyword} onChange={(e) => setForm({ ...form, productKeyword: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-32" />
          <input list="rule-sizes" aria-label="ขนาดสินค้า" placeholder="ขนาด เช่น 14 กรัม" value={form.productSizeKeyword} onChange={(e) => setForm({ ...form, productSizeKeyword: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-28" />
          <input aria-label="จังหวัด (keyword)" placeholder="จังหวัด (ว่าง=ทุกจังหวัด)" title="ใส่ชื่อจังหวัดปลายทาง เช่น นครปฐม เพื่อหารเฉพาะจังหวัดนั้น (ว่าง=ทุกจังหวัด, หลายคำคั่น |)" value={form.provinceKeyword} onChange={(e) => setForm({ ...form, provinceKeyword: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-36" />
          <input type="number" aria-label="ตัวหาร" placeholder="ตัวหาร" value={form.divisor} onChange={(e) => setForm({ ...form, divisor: +e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-20" />
          <label className="flex items-center gap-1 text-xs text-natural-muted">เริ่มใช้<input type="date" aria-label="วันเริ่มใช้กฎ" value={form.effectiveFrom} onChange={(e) => setForm({ ...form, effectiveFrom: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5" /></label>
          <button onClick={add} className={`${editId ? 'bg-amber-600' : 'bg-brand-navy'} text-white rounded-lg px-3 font-semibold`}>{editId ? '✎ บันทึกแก้ไข' : 'เพิ่ม'}</button>
          {editId && <button onClick={cancelEdit} className="border border-natural-border rounded-lg px-3 font-semibold">ยกเลิก</button>}
          <datalist id="rule-senders">{senderOpts.map((s) => <option key={s} value={s} />)}</datalist>
          <datalist id="rule-receivers">{receiverOpts.map((s) => <option key={s} value={s} />)}</datalist>
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

        <SimpleTable cols={['ผู้ส่ง', 'ผู้รับ', 'จังหวัด', 'กลุ่มผู้รับ', 'สินค้า', 'ขนาด', 'หาร', 'ปัดเศษ', 'เริ่มใช้']}
          rows={filteredRules.map((r: ProductConversionRule) => [r.senderKeyword, r.receiverKeyword || 'ทุกผู้รับ', r.provinceKeyword || 'ทุกจังหวัด', db.receiverGroups.find((g: ReceiverGroup) => g.id === r.receiverGroupId)?.groupName || '-', r.productKeyword, r.productSizeKeyword, `÷${r.divisor}`, r.roundingMethod === 'half_up' ? '.5 ปัดขึ้น' : r.roundingMethod, r.effectiveFrom || '-'])}
          onEdit={(i: number) => startEdit(filteredRules[i])}
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
        <button onClick={add} className="bg-brand-red text-white rounded-lg px-3 font-semibold">เพิ่ม</button>
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
        <button onClick={addGroup} className="bg-brand-red text-white rounded-lg px-3 text-sm font-semibold">+ เพิ่มกลุ่ม</button>
      </div>

      <div className="flex flex-col gap-3">
        {db.receiverGroups.map((g: ReceiverGroup) => (
          <div key={g.id} className="border border-natural-border rounded-xl p-3">
            <div className="flex items-center justify-between mb-2">
              <span className="font-semibold text-sm text-brand-navy">{g.groupName}</span>
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
    if (editId) { await api(`/api/vehicles/${editId}`, 'PUT', form); alertBox('บันทึกการแก้ไขแล้ว ✅', `รถ ${form.plateNo}${form.driverName ? ' · ' + form.driverName : ''}`, 'success'); }
    else { await api('/api/vehicles', 'POST', { ...form, branchId }); alertBox('เพิ่มรถแล้ว ✅', `รถ ${form.plateNo}${form.driverName ? ' · ' + form.driverName : ''}`, 'success'); }
    setForm(blankV); setEditId(null); reload();
  };
  const startEdit = async (v: Vehicle) => {
    if (!(await confirmPassword('แก้ไขรถ'))) return;
    setForm({ plateNo: v.plateNo, driverName: v.driverName, vehicleType: v.vehicleType, rateGroup: v.rateGroup || '', status: v.status });
    setEditId(v.id);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };
  const cancelEdit = () => { setForm(blankV); setEditId(null); };
  const setGroup = async (v: Vehicle, g: string) => { await api(`/api/vehicles/${v.id}`, 'PUT', { rateGroup: g }); showToast('success', `${v.plateNo}: ${g || 'ไม่ระบุกลุ่ม'}`); reload(); };
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
        <button onClick={add} className={`${editId ? 'bg-amber-600' : 'bg-brand-navy'} text-white rounded-lg px-3 py-1.5 font-semibold`}>{editId ? '✎ บันทึกแก้ไข' : 'เพิ่ม'}</button>
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
                <td className="py-1.5 px-1 font-semibold text-brand-navy">{v.plateNo}</td>
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
                  <button type="button" title="แก้ไข (ใส่รหัส)" onClick={() => startEdit(v)} className="text-brand-navy hover:text-amber-700 font-semibold mr-2">✎</button>
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
      <input type={type} value={value} aria-label={label} onChange={(e) => onChange(e.target.value)} className="border border-natural-border rounded-lg px-2 py-1.5 focus:border-brand-navy outline-none" />
    </div>
  );
}
function Section({ title, icon: Icon, children }: any) {
  return (
    <div className="bg-white rounded-2xl border border-natural-border p-5">
      <h3 className="font-bold text-sm text-brand-navy flex items-center gap-1.5 mb-3"><Icon className="w-4 h-4" />{title}</h3>
      {children}
    </div>
  );
}
function Stat({ label, value, highlight }: { label: string; value: string; highlight?: boolean }) {
  return (
    <div className={`relative rounded-2xl border p-4 overflow-hidden shadow-xs ${highlight ? 'bg-white border-brand-red/30' : 'bg-white border-natural-border'}`}>
      <span className={`absolute left-0 top-0 bottom-0 w-1 ${highlight ? 'bg-brand-red' : 'bg-brand-navy'}`} />
      <div className="text-[10px] uppercase font-bold tracking-wide text-natural-muted pl-1">{label}</div>
      <div className={`text-2xl font-extrabold mt-1 pl-1 ${highlight ? 'text-brand-red' : 'text-brand-navy'}`}>{value}</div>
    </div>
  );
}
function SimpleTable({ cols, rows, onDelete, onEdit, footer }: { cols: string[]; rows: any[][]; onDelete?: (i: number) => void; onEdit?: (i: number) => void; footer?: any[] }) {
  const hasAction = !!(onDelete || onEdit);
  return (
    <div className="overflow-x-auto rounded-xl border border-natural-border">
      <table className="w-full text-xs">
        <thead><tr className="bg-brand-navy text-white text-left">{cols.map((c) => <th key={c} className="py-2 px-2 font-semibold">{c}</th>)}{hasAction && <th className="w-16"></th>}</tr></thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} className={`border-t border-natural-border ${i % 2 ? 'bg-natural-secondary/40' : 'bg-white'}`}>
              {row.map((cell, j) => <td key={j} className="py-1.5 px-1">{cell}</td>)}
              {hasAction && <td className="text-center whitespace-nowrap">
                {onEdit && <button type="button" aria-label="แก้ไขรายการ" title="แก้ไข (ใส่รหัส)" onClick={() => onEdit(i)} className="text-brand-navy hover:text-amber-700 font-semibold mr-2">✎</button>}
                {onDelete && <button type="button" aria-label="ลบรายการ" title="ลบรายการ" onClick={async () => { if (await confirmDelete()) onDelete(i); }} className="text-natural-muted hover:text-rose-600"><Trash2 className="w-3.5 h-3.5 inline" /></button>}
              </td>}
            </tr>
          ))}
          {rows.length === 0 && <tr><td colSpan={cols.length + 1} className="py-4 text-center text-natural-muted">ยังไม่มีข้อมูล</td></tr>}
        </tbody>
        {footer && rows.length > 0 && (
          <tfoot><tr className="border-t-2 border-brand-navy bg-brand-navy/5 font-bold text-brand-navy">
            {footer.map((cell, j) => <td key={j} className="py-2 px-2">{cell}</td>)}
            {hasAction && <td></td>}
          </tr></tfoot>
        )}
      </table>
    </div>
  );
}
function EmptyHint({ text }: { text: string }) {
  return <div className="bg-white rounded-2xl border border-natural-border p-10 text-center text-natural-muted text-sm">{text}</div>;
}
