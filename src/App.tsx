import React, { useState, useEffect, useRef } from 'react';
import {
  UploadCloud, AlertTriangle, FileSpreadsheet, Trash2, Plus, Save,
  RefreshCw, Lock, Unlock, Database, Truck, Tag, Filter, Calculator, Fuel, Receipt, Coins,
} from 'lucide-react';
import {
  DatabaseState, BillingCycle, Vehicle, RateMaster, ReceiverGroup, ReceiverGroupAlias,
  ProductConversionRule, TripDocument, FuelEntry, DeductionEntry, ExtractedTripDocument, MoneyCategory, ManualBoxSender,
} from './types';
import { exportCycleToExcel } from './excel-export';
import { summarizeByVehicle, isUnspecifiedName } from './calc';
import { confirmDelete, confirmAction, notify } from './ui';

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
  cycles: [], vehicles: [], rateMasters: [], rateMasterHistory: [], receiverGroups: [],
  receiverGroupAliases: [], conversionRules: [], manualBoxSenders: [], moneyCategories: [], tripDocuments: [], fuelEntries: [], deductions: [],
};

type Tab = 'calc' | 'rates' | 'rules' | 'vehicles' | 'fuel' | 'dashboard';
type Toast = { type: 'success' | 'error' | 'warning'; message: string };

const THAI_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const money = (n: number) => n.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export default function App() {
  const [db, setDb] = useState<DatabaseState>(EMPTY);
  const [loading, setLoading] = useState(true);
  const [selectedCycleId, setSelectedCycleId] = useState('');
  const [tab, setTab] = useState<Tab>('calc');
  const [aiEnabled, setAiEnabled] = useState(true);

  const showToast = (type: Toast['type'], message: string) => notify(type, message);

  const fetchState = async (autoCycle?: string) => {
    setLoading(true);
    try {
      const res = await fetch('/api/state');
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
  useEffect(() => {
    fetchState();
    fetch('/api/config').then((r) => r.json()).then((c) => setAiEnabled(!!c.aiEnabled)).catch(() => setAiEnabled(false));
  }, []);

  const cycle = db.cycles.find((c) => c.id === selectedCycleId) || null;
  const cycleTrips = db.tripDocuments.filter((t) => t.cycleId === selectedCycleId);

  const api = async (url: string, method: string, body?: any) => {
    const res = await fetch(url, {
      method,
      headers: { 'Content-Type': 'application/json' },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || 'เกิดข้อผิดพลาด');
    return res.json();
  };

  return (
    <div className="min-h-screen bg-natural-bg text-natural-text font-sans flex flex-col">
      {/* Header */}
      <header className="bg-white border-b border-natural-border px-6 py-3 flex flex-col md:flex-row md:items-center justify-between gap-3 sticky top-0 z-30">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-[#1B365D] text-white rounded-xl flex items-center justify-center"><Truck className="w-5 h-5" /></div>
          <div>
            <h1 className="text-lg font-bold text-[#1B365D]">ระบบค่าเที่ยว + ค่าน้ำมันรถร่วม</h1>
            <p className="text-xs text-natural-muted">PDF ใบกระจาย → Review → คำนวณรอบ 1-15 / 16-31 → Export Excel</p>
          </div>
        </div>
        <CycleBar cycles={db.cycles} selectedCycleId={selectedCycleId} setSelectedCycleId={setSelectedCycleId}
          onCreated={(id: string) => fetchState(id)} api={api} showToast={showToast} />
      </header>

      {/* Tabs */}
      <nav className="bg-white border-b border-natural-border px-6 flex gap-1 overflow-x-auto">
        {([
          ['calc', 'คำนวณค่าเที่ยว', Calculator],
          ['fuel', 'ค่าน้ำมัน & รายการหัก', Fuel],
          ['dashboard', 'Dashboard', Database],
          ['rates', 'Master ราคาขนส่ง', Tag],
          ['rules', 'เงื่อนไขตัวหาร', Filter],
          ['vehicles', 'รถ & คนขับ', Truck],
        ] as [Tab, string, any][]).map(([key, label, Icon]) => (
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
            {tab === 'calc' && <CalcTab db={db} cycle={cycle} cycleTrips={cycleTrips} api={api} aiEnabled={aiEnabled}
              reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'fuel' && <FuelDeductionTab db={db} cycle={cycle} api={api}
              reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'dashboard' && <DashboardTab db={db} cycle={cycle} />}
            {tab === 'rates' && <RatesTab db={db} api={api} reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'rules' && <RulesTab db={db} api={api} reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
            {tab === 'vehicles' && <VehiclesTab db={db} api={api} reload={() => fetchState(selectedCycleId)} showToast={showToast} />}
          </>
        )}
      </main>
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
function CalcTab({ db, cycle, cycleTrips, api, aiEnabled, reload, showToast }: any) {
  const [extracting, setExtracting] = useState(false);
  const [pending, setPending] = useState<{ extracted: ExtractedTripDocument; fileName: string; preview: TripDocument } | null>(null);
  const [filter, setFilter] = useState<'all' | 'divider' | 'warning'>('all');
  const fileRef = useRef<HTMLInputElement>(null);

  if (!cycle) return <EmptyHint text="กรุณาเลือกหรือเปิดรอบคำนวณก่อน" />;

  const preview = async (extracted: ExtractedTripDocument, fileName: string) => {
    const p: TripDocument = await api('/api/trips/preview', 'POST', { cycleId: cycle.id, extracted, fileName });
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
      await api('/api/trips', 'POST', { cycleId: cycle.id, extracted: pending.extracted, fileName: pending.fileName });
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
      <div onDragOver={(e) => e.preventDefault()} onDrop={(e) => { e.preventDefault(); onFiles(e.dataTransfer.files); }}
        className="bg-white rounded-2xl border-2 border-dashed border-natural-border p-6 text-center flex flex-col items-center">
        <input ref={fileRef} type="file" aria-label="เลือกไฟล์ PDF ใบกระจาย" accept="application/pdf" multiple className="hidden" onChange={(e) => e.target.files && onFiles(e.target.files)} />
        {extracting ? (
          <div className="flex flex-col items-center gap-2 py-2"><RefreshCw className="w-8 h-8 text-[#1B365D] animate-spin" /><p className="text-sm font-semibold text-[#1B365D]">AI กำลังอ่าน PDF ใบกระจาย...</p></div>
        ) : (
          <>
            <UploadCloud className="w-8 h-8 text-[#1B365D] mb-2" />
            <p className="font-semibold text-sm text-[#1B365D]">ลากวางไฟล์ PDF ใบกระจาย (หลายไฟล์ได้)</p>
            {!aiEnabled && (
              <p className="text-[11px] text-amber-700 bg-amber-50 border border-amber-200 rounded-full px-3 py-1 mt-2">
                ⚠️ ยังไม่ได้ตั้งค่า GEMINI_API_KEY — อ่าน PDF จริงไม่ได้ ใช้ "กรอกเอง" เพื่อทดสอบ
              </p>
            )}
            <div className="flex gap-2 mt-3 items-center flex-wrap justify-center">
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
      {pending && <ReviewBoard pending={pending} setPending={setPending} onPreview={preview} onSave={save} locked={cycle.status === 'closed'} />}

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
function ReviewBoard({ pending, setPending, onPreview, onSave, locked }: any) {
  const ext: ExtractedTripDocument = pending.extracted;
  const prev: TripDocument = pending.preview;
  const needsBox = prev.receipts.some((r) => r.requiresManualBox && (r.manualBoxQty == null || r.manualBoxQty <= 0));

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
      {prev.warnings.length > 0 && (
        <div className="bg-[#FCE4D6] border border-amber-200 rounded-xl p-3 text-xs text-[#9C0006] flex flex-col gap-1">
          {prev.warnings.map((w, i) => <div key={i} className="flex gap-1.5"><AlertTriangle className="w-4 h-4 shrink-0" />{w}</div>)}
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
        {prev.rateOptions.flat != null || prev.rateOptions.piece != null ? (
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
                {pr.adjustments.map((a, i) => <span key={i} className="bg-white border border-[#C65911]/40 rounded-full px-2 py-0.5">🟧÷{a.divisor} {a.productName}: {a.note}</span>)}
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
          {needsBox && <span className="block text-rose-600 text-xs font-semibold mt-0.5">⚠️ มีใบรับที่ต้องกรอกจำนวนกล่องก่อนบันทึก</span>}
        </div>
        <button onClick={onSave} disabled={locked || needsBox} title={needsBox ? 'กรอกจำนวนกล่องให้ครบก่อน' : ''} className="bg-[#1B365D] disabled:bg-natural-muted disabled:cursor-not-allowed text-white rounded-full px-5 py-2 text-sm font-semibold flex items-center gap-1.5"><Save className="w-4 h-4" />ยืนยันบันทึก</button>
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
                <td className="py-1">{r.hasAdjustment && <span className="text-[#C65911] font-bold">🟧÷{r.adjustments[0].divisor} </span>}{r.receiptNo}</td>
                <td>{r.receiverName}</td>
                <td>{r.districtRaw} {r.provinceRaw}</td>
                <td className="text-center">{r.totalQty}</td>
                <td className="text-center font-bold text-[#C00000]" title={r.adjustments.map((a) => a.note).join(' | ')}>{r.billingQty}</td>
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
function FuelDeductionTab({ db, cycle, api, reload, showToast }: any) {
  const cats: MoneyCategory[] = db.moneyCategories || [];
  const incomeCats = cats.filter((c) => c.kind === 'income' && c.status === 'active');
  const dedCats = cats.filter((c) => c.kind === 'deduction' && c.status === 'active');

  const fuel = db.fuelEntries.filter((f: FuelEntry) => cycle && f.cycleId === cycle.id);
  const allDed = db.deductions.filter((d: DeductionEntry) => cycle && d.cycleId === cycle.id);
  const ded = allDed.filter((d: DeductionEntry) => d.kind === 'deduction');
  const incomes = allDed.filter((d: DeductionEntry) => d.kind === 'income');

  const [fForm, setFForm] = useState({ plateNo: '', refNo: '', date: cycle?.startDate || '', amount: 0 });
  const [dForm, setDForm] = useState({ plateNo: '', categoryId: '', amount: 0 });
  const [bForm, setBForm] = useState({ plateNo: '', categoryId: '', amount: 0 });

  if (!cycle) return <EmptyHint text="กรุณาเลือกรอบก่อน" />;

  const addFuel = async () => {
    if (!fForm.plateNo || !fForm.amount) return showToast('warning', 'กรอกทะเบียนและจำนวนเงิน');
    await api('/api/fuel', 'POST', { ...fForm, cycleId: cycle.id });
    setFForm({ plateNo: '', refNo: '', date: cycle.startDate, amount: 0 }); reload();
  };
  const addEntry = async (plateNo: string, categoryId: string, amount: number, kind: 'income' | 'deduction', reset: () => void) => {
    const cat = cats.find((c) => c.id === categoryId) || (kind === 'income' ? incomeCats[0] : dedCats[0]);
    if (!plateNo || !amount || !cat) return showToast('warning', 'กรอกทะเบียน/จำนวนเงิน และเลือกประเภท');
    await api('/api/deductions', 'POST', { plateNo, categoryId: cat.id, kind, label: cat.name, amount, cycleId: cycle.id });
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
          <button onClick={() => addEntry(bForm.plateNo, bForm.categoryId, bForm.amount, 'income', () => setBForm({ plateNo: '', categoryId: '', amount: 0 }))} className="bg-emerald-600 text-white rounded-lg px-3 text-sm font-semibold">เพิ่ม</button>
        </div>
        <SimpleTable rows={incomes.map((d: DeductionEntry) => [d.plateNo, d.label, `+${money(d.amount)}`])} cols={['ทะเบียน', 'รายการ', 'จำนวนเพิ่ม']}
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
          <button onClick={() => addEntry(dForm.plateNo, dForm.categoryId, dForm.amount, 'deduction', () => setDForm({ plateNo: '', categoryId: '', amount: 0 }))} className="bg-[#1B365D] text-white rounded-lg px-3 text-sm font-semibold">เพิ่ม</button>
        </div>
        <SimpleTable rows={ded.map((d: DeductionEntry) => [d.plateNo, d.label, money(d.amount)])} cols={['ทะเบียน', 'รายการ', 'จำนวน']}
          onDelete={async (i: number) => { await api(`/api/deductions/${ded[i].id}`, 'DELETE'); reload(); }} />
      </Section>

      {/* จัดการประเภท — เพิ่มชื่อใน dropdown ได้เองโดยไม่ต้องแก้โค้ด */}
      <CategoryManager cats={cats} api={api} reload={reload} showToast={showToast} />

      <datalist id="plates">{db.vehicles.map((v: Vehicle) => <option key={v.id} value={v.plateNo} />)}</datalist>
    </div>
  );
}

// จัดการ Master ประเภทรายได้เพิ่ม / รายการหัก (เติม dropdown)
function CategoryManager({ cats, api, reload, showToast }: any) {
  const [form, setForm] = useState<{ name: string; kind: 'income' | 'deduction' }>({ name: '', kind: 'deduction' });
  const add = async () => {
    if (!form.name.trim()) return showToast('warning', 'กรอกชื่อประเภท');
    await api('/api/money-categories', 'POST', { name: form.name.trim(), kind: form.kind, status: 'active', builtin: false });
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
function DashboardTab({ db, cycle }: any) {
  if (!cycle) return <EmptyHint text="กรุณาเลือกรอบก่อน" />;
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
// Tab: Master ราคาขนส่ง
// ===========================================================================
function RatesTab({ db, api, reload, showToast }: any) {
  const blank = { destinationName: '', provinceName: '', provinceShort: '', districtName: '', priceType: 'flat', price: 0, effectiveFrom: '2020-01-01', effectiveTo: null, status: 'active' };
  const [form, setForm] = useState<any>(blank);
  const add = async () => {
    if (!form.provinceName || !form.price) return showToast('warning', 'กรอกจังหวัดและราคา');
    await api('/api/rate-masters', 'POST', form); setForm(blank); reload(); showToast('success', 'เพิ่มราคาแล้ว');
  };
  return (
    <Section title="Master ราคาขนส่ง" icon={Tag}>
      <div className="flex flex-wrap gap-2 mb-3 text-sm">
        <input aria-label="ปลายทาง" placeholder="ปลายทาง" value={form.destinationName} onChange={(e) => setForm({ ...form, destinationName: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-32" />
        <input aria-label="จังหวัด" placeholder="จังหวัด" value={form.provinceName} onChange={(e) => setForm({ ...form, provinceName: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-28" />
        <input aria-label="อำเภอ" placeholder="อำเภอ" value={form.districtName} onChange={(e) => setForm({ ...form, districtName: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-24" />
        <select aria-label="ประเภทราคา" value={form.priceType} onChange={(e) => setForm({ ...form, priceType: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5">
          <option value="flat">ราคาเหมา</option><option value="piece">ราคาชิ้น</option>
        </select>
        <input type="number" aria-label="ราคา" placeholder="ราคา" value={form.price || ''} onChange={(e) => setForm({ ...form, price: +e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-24" />
        <button onClick={add} className="bg-[#1B365D] text-white rounded-lg px-3 font-semibold">เพิ่ม</button>
      </div>
      <SimpleTable cols={['ปลายทาง', 'จังหวัด', 'อำเภอ', 'ประเภท', 'ราคา', 'เริ่มใช้']}
        rows={db.rateMasters.map((r: RateMaster) => [r.destinationName, r.provinceName, r.districtName, r.priceType === 'flat' ? 'เหมา' : 'ชิ้น', money(r.price), r.effectiveFrom])}
        onDelete={async (i: number) => { await api(`/api/rate-masters/${db.rateMasters[i].id}`, 'DELETE'); reload(); }} />
    </Section>
  );
}

// ===========================================================================
// Tab: เงื่อนไขตัวหาร + กลุ่มผู้รับ
// ===========================================================================
function RulesTab({ db, api, reload, showToast }: any) {
  const blank = { ruleName: '', senderKeyword: 'ซีโน', receiverGroupId: db.receiverGroups[0]?.id || '', productKeyword: '', productSizeKeyword: '', divisor: 3, roundingMethod: 'half_up', applyLevel: 'receipt', status: 'active', effectiveFrom: '2020-01-01', effectiveTo: null };
  const [form, setForm] = useState<any>(blank);
  const [fSender, setFSender] = useState('');
  const [fGroup, setFGroup] = useState('');
  const add = async () => {
    if (!form.productKeyword) return showToast('warning', 'กรอกชื่อสินค้า');
    await api('/api/conversion-rules', 'POST', { ...form, ruleName: form.ruleName || `${form.productKeyword} หาร ${form.divisor}` });
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
      <GroupManager db={db} api={api} reload={reload} showToast={showToast} />
      <ManualBoxSenderManager db={db} api={api} reload={reload} showToast={showToast} />
    </div>
  );
}

// จัดการผู้ส่งที่ส่งเป็นชิ้น (ต้องกรอกจำนวนกล่องเอง)
function ManualBoxSenderManager({ db, api, reload, showToast }: any) {
  const [form, setForm] = useState({ senderKeyword: '', note: '' });
  const add = async () => {
    if (!form.senderKeyword.trim()) return showToast('warning', 'กรอกคำในชื่อผู้ส่ง');
    await api('/api/manual-box-senders', 'POST', { senderKeyword: form.senderKeyword.trim(), note: form.note.trim(), status: 'active' });
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
function GroupManager({ db, api, reload, showToast }: any) {
  const [newGroup, setNewGroup] = useState('');
  const [aliasInputs, setAliasInputs] = useState<Record<string, string>>({});

  const addGroup = async () => {
    if (!newGroup.trim()) return showToast('warning', 'กรอกชื่อกลุ่ม');
    await api('/api/receiver-groups', 'POST', { groupName: newGroup.trim(), status: 'active' });
    setNewGroup(''); reload(); showToast('success', 'เพิ่มกลุ่มแล้ว — ใช้ใน dropdown ได้เลย');
  };
  const delGroup = async (g: ReceiverGroup) => {
    if (!(await confirmDelete(g.groupName))) return;
    await api(`/api/receiver-groups/${g.id}`, 'DELETE'); reload();
  };
  const addAlias = async (groupId: string) => {
    const name = (aliasInputs[groupId] || '').trim();
    if (!name) return showToast('warning', 'กรอกชื่อพ้อง (ชื่อผู้รับที่ปรากฏใน PDF)');
    await api('/api/receiver-aliases', 'POST', { receiverGroupId: groupId, aliasName: name, status: 'active' });
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
function VehiclesTab({ db, api, reload, showToast }: any) {
  const [form, setForm] = useState({ plateNo: '', driverName: '', vehicleType: '6 ล้อ', status: 'active' });
  const add = async () => {
    if (!form.plateNo) return showToast('warning', 'กรอกทะเบียน');
    await api('/api/vehicles', 'POST', form); setForm({ plateNo: '', driverName: '', vehicleType: '6 ล้อ', status: 'active' }); reload();
  };
  return (
    <Section title="Master รถร่วม & คนขับ" icon={Truck}>
      <div className="flex flex-wrap gap-2 mb-3 text-sm">
        <input aria-label="ทะเบียนรถ" placeholder="ทะเบียน" value={form.plateNo} onChange={(e) => setForm({ ...form, plateNo: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-28" />
        <input aria-label="ชื่อคนขับ" placeholder="คนขับ" value={form.driverName} onChange={(e) => setForm({ ...form, driverName: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-36" />
        <input aria-label="ประเภทรถ" placeholder="ประเภทรถ" value={form.vehicleType} onChange={(e) => setForm({ ...form, vehicleType: e.target.value })} className="border border-natural-border rounded-lg px-2 py-1.5 w-24" />
        <button onClick={add} className="bg-[#1B365D] text-white rounded-lg px-3 font-semibold">เพิ่ม</button>
      </div>
      <SimpleTable cols={['ทะเบียน', 'คนขับ', 'ประเภท']}
        rows={db.vehicles.map((v: Vehicle) => [v.plateNo, v.driverName, v.vehicleType])}
        onDelete={async (i: number) => { await api(`/api/vehicles/${db.vehicles[i].id}`, 'DELETE'); reload(); }} />
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
