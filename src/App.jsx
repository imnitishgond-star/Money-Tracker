import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import {
  Landmark, PlusCircle, Trash2, CheckCircle2, AlertTriangle, Calendar,
  TrendingUp, TrendingDown, Wallet, ClipboardList, ChevronLeft, ChevronRight,
  X, PiggyBank, Clock, Download, Users, ArrowUpRight, ArrowDownRight,
  BookOpen, PieChart, ChevronDown, ChevronUp, ShoppingBag,
  Sparkles, Upload, Lightbulb, Target, Wallet2, Activity,
  Calculator, Lock, Unlock, Camera, Repeat, Moon, Sun, MessageCircle,
  FileDown, Trophy, Edit2, Bell, Image as ImageIcon, Settings, Filter,
} from 'lucide-react';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell,
} from 'recharts';
import * as XLSX from 'xlsx';

/* ---------------------------------------------------------------------- */
/* Helpers                                                                */
/* ---------------------------------------------------------------------- */

function ymToDate(ym) { const [y, m] = ym.split('-').map(Number); return new Date(y, m - 1, 1); }
function dateToYM(d) { return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; }
function addMonths(ym, n) { const d = ymToDate(ym); d.setMonth(d.getMonth() + n); return dateToYM(d); }
function monthDiff(fromYM, toYM) {
  const a = ymToDate(fromYM), b = ymToDate(toYM);
  return (b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth());
}
function fmtMonthLabel(ym) {
  const d = ymToDate(ym);
  return d.toLocaleString('en-IN', { month: 'long', year: 'numeric' });
}
function currentYM() { return dateToYM(new Date()); }
function todayISO() { return new Date().toISOString().slice(0, 10); }
const CURRENCIES = {
  INR: { symbol: '₹', locale: 'en-IN' },
  USD: { symbol: '$', locale: 'en-US' },
  EUR: { symbol: '€', locale: 'de-DE' },
  GBP: { symbol: '£', locale: 'en-GB' },
  AED: { symbol: 'AED ', locale: 'en-AE' },
  SGD: { symbol: 'S$', locale: 'en-SG' },
};
let CURRENT_CURRENCY = CURRENCIES.INR;
let CURRENT_CURRENCY_CODE = 'INR';
function setCurrencyCode(code) { if (CURRENCIES[code]) { CURRENT_CURRENCY = CURRENCIES[code]; CURRENT_CURRENCY_CODE = code; } }
function fmtINR(n) { return CURRENT_CURRENCY.symbol + Math.round(n || 0).toLocaleString(CURRENT_CURRENCY.locale); }

/* ---- Amortization helpers (approximate — for interest split & prepayment estimates) ---- */
function monthlyRateFromAnnual(annualPct) { return (Number(annualPct) || 0) / 12 / 100; }
function remainingMonthsForBalance(balance, emi, monthlyRate) {
  if (balance <= 0) return 0;
  if (monthlyRate <= 0) return Math.ceil(balance / emi);
  const ratio = (balance * monthlyRate) / emi;
  if (ratio >= 1) return Infinity; // EMI too small to ever cover interest
  return Math.ceil(Math.log(1 / (1 - ratio)) / Math.log(1 + monthlyRate));
}
function fmtDate(iso) {
  if (!iso) return '—';
  const [y, m, d] = iso.split('-');
  if (!y || !m || !d) return iso;
  return `${d}-${m}-${y.slice(2)}`;
}
function uid() { return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`; }

/* ---- EMI ---- */
const LOAN_CATEGORIES = ['Home Loan', 'Car Loan', 'Personal Loan', 'Education Loan', 'Credit Card', 'Other'];
const LOAN_CATEGORY_COLORS = {
  'Home Loan': { bg: 'bg-[#EDE7F6]', text: 'text-[#7C6BA6]', bar: '#9E8BC4' },
  'Car Loan': { bg: 'bg-[#E3EEF5]', text: 'text-[#5B85A3]', bar: '#7CA6C2' },
  'Personal Loan': { bg: 'bg-[#F3E7EE]', text: 'text-[#B4778E]', bar: '#C99CB0' },
  'Education Loan': { bg: 'bg-[#E3F1EE]', text: 'text-[#4F9187]', bar: '#6FAEA4' },
  'Credit Card': { bg: 'bg-[#F3E9DD]', text: 'text-[#A67C52]', bar: '#C0996F' },
  'Other': { bg: 'bg-[#F1ECE9]', text: 'text-[#8A7D78]', bar: '#ACA09B' },
};
const LOAN_STATUS_COLORS = {
  paid: { bg: 'bg-[#EAF3EA]', text: 'text-[#5C8A5E]', ring: 'ring-[#BFDDC0]', dot: 'bg-[#8FBC94]' },
  partial: { bg: 'bg-[#FBF0DE]', text: 'text-[#B9862E]', ring: 'ring-[#EDD6A8]', dot: 'bg-[#E3B166]' },
  overdue: { bg: 'bg-[#FBEAE7]', text: 'text-[#C4685B]', ring: 'ring-[#F0C6BE]', dot: 'bg-[#E2897C]' },
  pending: { bg: 'bg-[#F1ECE9]', text: 'text-[#A79A95]', ring: 'ring-[#E4D9D3]', dot: 'bg-[#CFC3BD]' },
};

/* ---- Budget ---- */
const EXPENSE_CATEGORIES = [
  'Food', 'Transportation', 'Travel', 'Shopping', 'Bills', 'Entertainment',
  'Medical', 'Rent', 'Utilities', 'Education', 'Personal', 'Groceries',
  'EMI/Loans', 'Savings', 'Other',
];
const EXPENSE_CATEGORY_BARS = {
  'Food': '#E3B166', 'Transportation': '#7CA6C2', 'Travel': '#6FAEA4', 'Shopping': '#C99CB0',
  'Bills': '#C0996F', 'Entertainment': '#E2897C', 'Medical': '#D68C8C', 'Rent': '#9E8BC4',
  'Utilities': '#B6935F', 'Education': '#7C6BA6', 'Personal': '#8FAF8A', 'Groceries': '#8FAF8A',
  'EMI/Loans': '#B6A7D6', 'Savings': '#5C8A5E', 'Other': '#ACA09B',
};
const PAYMENT_MODES = ['Cash', 'UPI', 'Card', 'Credit Card', 'Debit Card', 'Bank Transfer', 'Net Banking', 'Wallet', 'Other'];

/* ---- Auto-categorization from description keywords ---- */
const CATEGORY_KEYWORDS = {
  'Groceries': ['grocery', 'groceries', 'vegetable', 'sabzi', 'kirana', 'supermarket', 'bigbasket', 'blinkit', 'zepto'],
  'Food': ['restaurant', 'food', 'lunch', 'dinner', 'breakfast', 'swiggy', 'zomato', 'cafe', 'coffee', 'pizza', 'burger'],
  'Transportation': ['uber', 'ola', 'taxi', 'cab', 'petrol', 'diesel', 'fuel', 'metro', 'bus', 'train', 'auto', 'parking'],
  'Travel': ['flight', 'hotel', 'trip', 'vacation', 'booking', 'irctc', 'airbnb', 'travel'],
  'Rent': ['rent', 'landlord', 'lease'],
  'Bills': ['electricity', 'water bill', 'wifi', 'internet', 'mobile bill', 'recharge', 'gas bill', 'dth', 'bill'],
  'Utilities': ['electricity', 'water', 'gas', 'utility'],
  'Shopping': ['amazon', 'flipkart', 'myntra', 'shopping', 'clothes', 'shoes', 'mall'],
  'Entertainment': ['movie', 'netflix', 'prime', 'spotify', 'concert', 'game', 'entertainment'],
  'Medical': ['doctor', 'medicine', 'pharmacy', 'hospital', 'clinic', 'health', 'medical'],
  'Personal': ['salon', 'haircut', 'spa', 'personal', 'gym'],
  'Education': ['school', 'college', 'course', 'book', 'tuition', 'fees', 'education'],
  'EMI/Loans': ['emi', 'loan installment'],
  'Savings': ['sip', 'mutual fund', 'investment', 'savings', 'fd', 'rd'],
};
function guessCategory(text) {
  const t = (text || '').toLowerCase();
  if (!t.trim()) return null;
  for (const [cat, words] of Object.entries(CATEGORY_KEYWORDS)) {
    if (words.some((w) => t.includes(w))) return cat;
  }
  return null;
}

/* ---- Voice input parsing: "500 groceries weekly shopping" → amount + category + description ---- */
function parseVoiceExpense(transcript) {
  const t = transcript.trim();
  const amountMatch = t.match(/(\d+(\.\d+)?)/);
  const amount = amountMatch ? parseFloat(amountMatch[1]) : null;
  const rest = t.replace(amountMatch ? amountMatch[0] : '', '').trim();
  const category = guessCategory(rest) || guessCategory(t);
  return { amount, description: rest, category };
}

/* ---- CSV import (no header row): date,amount,category,paymentMode,description ---- */
function parseCSVLine(line) {
  const out = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
      else if (ch === '"') { inQuotes = false; }
      else { cur += ch; }
    } else {
      if (ch === '"') inQuotes = true;
      else if (ch === ',') { out.push(cur); cur = ''; }
      else cur += ch;
    }
  }
  out.push(cur);
  return out.map((s) => s.trim());
}
function normalizeCSVDate(raw) {
  if (!raw) return null;
  const s = raw.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const match = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (match) {
    let [, first, second, y] = match;
    if (y.length === 2) y = `20${y}`;
    const f = parseInt(first, 10), sec = parseInt(second, 10);
    // Primary format is DD-MM-YYYY — use it whenever the second number can be a valid month.
    if (sec >= 1 && sec <= 12 && f >= 1 && f <= 31) {
      return `${y}-${String(sec).padStart(2, '0')}-${String(f).padStart(2, '0')}`;
    }
    // Otherwise it must be MM-DD-YYYY (second number too big to be a month).
    if (f >= 1 && f <= 12 && sec >= 1 && sec <= 31) {
      return `${y}-${String(f).padStart(2, '0')}-${String(sec).padStart(2, '0')}`;
    }
    return null;
  }
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) return dateToYM(parsed) + '-' + String(parsed.getDate()).padStart(2, '0');
  return null;
}
function matchClosest(value, list, fallback) {
  if (!value) return fallback;
  const v = value.trim().toLowerCase();
  const exact = list.find((c) => c.toLowerCase() === v);
  if (exact) return exact;
  const partial = list.find((c) => c.toLowerCase().includes(v) || v.includes(c.toLowerCase()));
  return partial || fallback;
}
function parseExpensesCSV(text) {
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const rows = [];
  let skipped = 0;
  for (const line of lines) {
    const cols = parseCSVLine(line);
    const [rawDate, rawAmount, rawCategory, rawPaymentMode, rawPaidBy, rawDescription] = cols;
    const date = normalizeCSVDate(rawDate);
    const amount = parseFloat(String(rawAmount || '').replace(/[₹,]/g, ''));
    if (!date || isNaN(amount) || amount <= 0) { skipped++; continue; }
    rows.push({
      id: uid(),
      date,
      amount,
      category: matchClosest(rawCategory, EXPENSE_CATEGORIES, 'Other'),
      paymentMode: matchClosest(rawPaymentMode, PAYMENT_MODES, 'Cash'),
      paidBy: (rawPaidBy || '').trim(),
      description: (rawDescription || '').trim(),
    });
  }
  return { rows, skipped };
}

/* ---------------------------------------------------------------------- */
/* Shared building blocks                                                 */
/* ---------------------------------------------------------------------- */

function AnimatedNumber({ value, formatter, duration = 700 }) {
  const [display, setDisplay] = useState(value);
  const prevRef = useRef(value);
  useEffect(() => {
    const from = prevRef.current;
    const to = value;
    if (from === to) return;
    let start = null;
    let raf;
    const step = (ts) => {
      if (start === null) start = ts;
      const progress = Math.min((ts - start) / duration, 1);
      const eased = 1 - Math.pow(1 - progress, 3); // ease-out cubic
      setDisplay(from + (to - from) * eased);
      if (progress < 1) raf = requestAnimationFrame(step);
      else prevRef.current = to;
    };
    raf = requestAnimationFrame(step);
    return () => cancelAnimationFrame(raf);
  }, [value, duration]);
  return <>{formatter ? formatter(display) : Math.round(display)}</>;
}

function KPI({ F, label, value, numeric, sub, icon: Icon, accent }) {
  return (
    <div className="rounded-xl border border-[#EFE1DA] bg-white p-4 hover:border-[#E4D5CE] transition-colors relative overflow-hidden min-w-0">
      <div className="flex items-center justify-between gap-2">
        <p className="text-[10.5px] text-[#A79A95] uppercase tracking-[0.1em] font-medium">{label}</p>
        <div className={`w-6 h-6 rounded-md flex items-center justify-center bg-current/10 shrink-0 ${accent}`}>
          <Icon size={13} className={accent} />
        </div>
      </div>
      <p style={{ fontFamily: F.mono, letterSpacing: '-0.01em' }} className={`text-xl font-bold mt-2 ${accent}`}>
        {typeof numeric === 'number' ? <AnimatedNumber value={numeric} formatter={fmtINR} /> : value}
      </p>
      {sub && <p className="text-[11px] text-[#A79A95] mt-1">{sub}</p>}
    </div>
  );
}

function ProgressBar({ pct, color }) {
  const clamped = Math.max(0, Math.min(pct, 100));
  return (
    <div className="h-1.5 rounded-full bg-[#F1ECE9] overflow-hidden">
      <div className="h-full rounded-full transition-all" style={{ width: `${clamped}%`, backgroundColor: color }} />
    </div>
  );
}

function MonthNav({ F, selectedMonth, setSelectedMonth }) {
  const isRealCurrentMonth = selectedMonth === currentYM();
  return (
    <div className="flex items-center justify-center gap-3 bg-white rounded-xl border border-[#EFE1DA] px-3 py-2">
      <button onClick={() => setSelectedMonth(addMonths(selectedMonth, -1))} className="text-[#A79A95] hover:text-[#3D3436] p-1">
        <ChevronLeft size={18} />
      </button>
      <span style={{ fontFamily: F.display }} className="text-sm font-semibold text-[#3D3436] min-w-[128px] text-center select-none">
        {fmtMonthLabel(selectedMonth)}
      </span>
      <button onClick={() => setSelectedMonth(addMonths(selectedMonth, 1))} className="text-[#A79A95] hover:text-[#3D3436] p-1">
        <ChevronRight size={18} />
      </button>
      {!isRealCurrentMonth && (
        <button onClick={() => setSelectedMonth(currentYM())} className="text-[11px] px-2.5 py-1 rounded-lg border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-colors whitespace-nowrap">
          This month
        </button>
      )}
    </div>
  );
}

function ModalShell({ F, title, onClose, children }) {
  return (
    <div
      className="fixed inset-0 z-50 isolate bg-black/70 backdrop-blur-sm flex items-end sm:items-center justify-center p-0 sm:p-4"
      onClick={onClose}
    >
      <div
        style={{ fontFamily: F.body, backgroundColor: '#FDFAF7' }}
        className="relative w-full sm:max-w-md rounded-t-2xl sm:rounded-2xl p-5 max-h-[90vh] overflow-y-auto shadow-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between mb-4">
          <h3 style={{ fontFamily: F.display }} className="text-base font-bold text-[#3D3436]">{title}</h3>
          <button onClick={onClose} className="text-[#A79A95] hover:text-[#3D3436]"><X size={18} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}

function FieldLabel({ children }) { return <label className="text-[11px] text-[#A79A95] font-medium">{children}</label>; }
const inputCls = "mt-1 w-full text-sm px-3 py-2 rounded-lg border border-[#E4D5CE] focus:outline-none focus:border-[#C4899C] bg-white";

/* ======================================================================
   EMI TAB
   ====================================================================== */

function AddLoanModal({ F, onClose, onSave }) {
  const [form, setForm] = useState({
    name: '', category: 'Home Loan', lender: '', totalAmount: '', emiAmount: '', interestRate: '',
    startMonth: currentYM(), tenureMonths: '', dueDay: '5',
  });
  const [err, setErr] = useState('');
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  const submit = () => {
    if (!form.name.trim()) { setErr('Loan/EMI ka naam daalo.'); return; }
    const total = parseFloat(form.totalAmount);
    const emi = parseFloat(form.emiAmount);
    const tenure = parseInt(form.tenureMonths, 10);
    const dueDay = parseInt(form.dueDay, 10);
    const interestRate = form.interestRate === '' ? 0 : parseFloat(form.interestRate);
    if (isNaN(total) || total <= 0) { setErr('Total loan amount sahi daalo.'); return; }
    if (isNaN(emi) || emi <= 0) { setErr('Monthly EMI amount sahi daalo.'); return; }
    if (isNaN(tenure) || tenure <= 0) { setErr('Tenure (months) sahi daalo.'); return; }
    if (isNaN(dueDay) || dueDay < 1 || dueDay > 31) { setErr('Due day 1-31 ke beech honi chahiye.'); return; }
    if (isNaN(interestRate) || interestRate < 0) { setErr('Interest rate sahi daalo (ya khaali chodo).'); return; }
    onSave({
      id: uid(), name: form.name.trim(), category: form.category, lender: form.lender.trim(),
      totalAmount: total, emiAmount: emi, interestRate, startMonth: form.startMonth, tenureMonths: tenure, dueDay,
    });
  };

  return (
    <ModalShell F={F} title="Naya Loan / EMI Add Karo" onClose={onClose}>
      <div className="space-y-3">
        <div><FieldLabel>Loan ka naam</FieldLabel>
          <input value={form.name} onChange={set('name')} placeholder="e.g. Home Loan - SBI" className={inputCls} />
        </div>
        <div><FieldLabel>Category</FieldLabel>
          <select value={form.category} onChange={set('category')} className={inputCls}>
            {LOAN_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div><FieldLabel>Lender / Bank (optional)</FieldLabel>
          <input value={form.lender} onChange={set('lender')} placeholder="e.g. HDFC Bank" className={inputCls} />
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div><FieldLabel>Total loan amount (₹)</FieldLabel>
            <input type="number" value={form.totalAmount} onChange={set('totalAmount')} placeholder="500000" className={inputCls} />
          </div>
          <div><FieldLabel>Monthly EMI (₹)</FieldLabel>
            <input type="number" value={form.emiAmount} onChange={set('emiAmount')} placeholder="12000" className={inputCls} />
          </div>
        </div>
        <div className="rounded-lg border border-[#EFE1DA] bg-white p-3 space-y-3">
          <div>
            <FieldLabel>Start month</FieldLabel>
            <input type="month" value={form.startMonth} onChange={set('startMonth')} className={inputCls} />
          </div>
          <div>
            <FieldLabel>Tenure (months)</FieldLabel>
            <input type="number" value={form.tenureMonths} onChange={set('tenureMonths')} placeholder="36" className={inputCls} />
          </div>
          {form.tenureMonths && !isNaN(parseInt(form.tenureMonths, 10)) && parseInt(form.tenureMonths, 10) > 0 && (
            <p className="text-[10.5px] text-[#A79A95]">
              Loan {fmtMonthLabel(form.startMonth)} se shuru hoke {fmtMonthLabel(addMonths(form.startMonth, parseInt(form.tenureMonths, 10) - 1))} tak active rahega ({form.tenureMonths} installments).
            </p>
          )}
        </div>
        <div><FieldLabel>Har mahine due date (din)</FieldLabel>
          <input type="number" min="1" max="31" value={form.dueDay} onChange={set('dueDay')} className={inputCls} />
        </div>
        <div><FieldLabel>Interest rate (% p.a., optional)</FieldLabel>
          <input type="number" step="0.1" value={form.interestRate} onChange={set('interestRate')} placeholder="e.g. 8.5" className={inputCls} />
          <p className="text-[10.5px] text-[#A79A95] mt-1">Ye dene se interest/principal split aur prepayment calculator use kar paoge.</p>
        </div>
        {err && <p className="text-xs text-[#C4685B]">{err}</p>}
        <button onClick={submit} style={{ background: 'linear-gradient(to right, #C4677A, #8F7CB8)' }} className="w-full text-white text-sm font-semibold py-2.5 rounded-lg hover:opacity-90 transition-all shadow-md mt-1">
          Loan Add Karo
        </button>
      </div>
    </ModalShell>
  );
}

function StatusBadge({ status }) {
  const c = LOAN_STATUS_COLORS[status];
  const label = { paid: 'Paid', partial: 'Partially paid', overdue: 'Overdue', pending: 'Pending' }[status];
  return (
    <span className={`inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full text-[11px] font-semibold ${c.bg} ${c.text} ring-1 ${c.ring}`}>
      <span className={`w-1.5 h-1.5 rounded-full ${c.dot} ${status === 'overdue' ? 'animate-pulse' : ''}`} />
      {label}
    </span>
  );
}

function PrepaymentModal({ F, loan, derived, onClose }) {
  const [extra, setExtra] = useState('');
  const monthlyRate = monthlyRateFromAnnual(loan.interestRate);
  const extraNum = parseFloat(extra) || 0;

  const currentRemaining = remainingMonthsForBalance(derived.outstanding, loan.emiAmount, monthlyRate);
  const newBalance = Math.max(derived.outstanding - extraNum, 0);
  const newRemaining = remainingMonthsForBalance(newBalance, loan.emiAmount, monthlyRate);
  const monthsSaved = isFinite(currentRemaining) && isFinite(newRemaining) ? currentRemaining - newRemaining : 0;

  const oldTotalPayable = isFinite(currentRemaining) ? currentRemaining * loan.emiAmount : null;
  const newTotalPayable = isFinite(newRemaining) ? newRemaining * loan.emiAmount + extraNum : null;
  const interestSaved = oldTotalPayable !== null && newTotalPayable !== null ? Math.max(oldTotalPayable - newTotalPayable, 0) : null;

  return (
    <ModalShell F={F} title="Prepayment Calculator" onClose={onClose}>
      <div className="space-y-3">
        <p className="text-xs text-[#A79A95]">{loan.name} — current outstanding: <span className="font-semibold text-[#3D3436]">{fmtINR(derived.outstanding)}</span></p>
        {!loan.interestRate && (
          <p className="text-[11px] text-[#B9862E] bg-[#FBF0DE] rounded-lg px-2.5 py-2">Is loan me interest rate set nahi hai, isliye sirf tenure reduction dikhega (linear estimate), interest savings nahi.</p>
        )}
        <div><FieldLabel>Extra one-time payment (₹)</FieldLabel>
          <input type="number" value={extra} onChange={(e) => setExtra(e.target.value)} placeholder="e.g. 50000" className={inputCls} autoFocus />
        </div>
        {extraNum > 0 && (
          <div style={{ fontFamily: F.mono }} className="rounded-lg bg-[#EAF3EA] p-3 space-y-1.5">
            <div className="flex justify-between text-xs"><span className="text-[#5C8A5E] font-sans">New tenure remaining</span><span className="font-semibold text-[#3D3436]">{isFinite(newRemaining) ? `${newRemaining} months` : '—'}</span></div>
            <div className="flex justify-between text-xs"><span className="text-[#5C8A5E] font-sans">Months saved</span><span className="font-semibold text-[#5C8A5E]">{isFinite(monthsSaved) ? `${monthsSaved} months` : '—'}</span></div>
            {loan.interestRate > 0 && interestSaved !== null && (
              <div className="flex justify-between text-xs"><span className="text-[#5C8A5E] font-sans">Est. interest saved</span><span className="font-semibold text-[#5C8A5E]">{fmtINR(interestSaved)}</span></div>
            )}
          </div>
        )}
        <p className="text-[10.5px] text-[#A79A95]">Estimate hai — assume karta hai EMI amount same rahega aur extra payment turant principal se kat jayega.</p>
      </div>
    </ModalShell>
  );
}

function LoanCard({ F, loan, derived, onPay, onDelete }) {
  const [amount, setAmount] = useState('');
  const [showPartial, setShowPartial] = useState(false);
  const [showPrepay, setShowPrepay] = useState(false);
  const cat = LOAN_CATEGORY_COLORS[loan.category] || LOAN_CATEGORY_COLORS.Other;
  const recordCustom = () => {
    const v = parseFloat(amount);
    if (isNaN(v) || v <= 0) return;
    onPay(loan.id, v);
    setAmount('');
    setShowPartial(false);
  };
  const monthlyRate = monthlyRateFromAnnual(loan.interestRate);
  const balanceBeforeThisMonth = Math.max(loan.totalAmount - (derived.totalPaidAllTime - derived.paidThisMonth), 0);
  const interestPortion = loan.interestRate > 0 ? Math.min(balanceBeforeThisMonth * monthlyRate, derived.due) : 0;
  const principalPortion = loan.interestRate > 0 ? Math.max(derived.due - interestPortion, 0) : 0;
  return (
    <div className="rounded-xl border border-[#EFE1DA] bg-white p-4 space-y-3">
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <div className="flex items-center gap-2 flex-wrap">
            <h3 style={{ fontFamily: F.display }} className="text-sm font-bold text-[#3D3436]">{loan.name}</h3>
            <span className={`text-[10px] px-2 py-0.5 rounded-full font-semibold ${cat.bg} ${cat.text}`}>{loan.category}</span>
          </div>
          {loan.lender && <p className="text-xs text-[#A79A95] mt-0.5">{loan.lender}</p>}
        </div>
        <div className="flex items-center gap-2 shrink-0">
          {loan.interestRate > 0 && (
            <button onClick={() => setShowPrepay(true)} className="text-[#A79A95] hover:text-[#7C6BA6]" title="Prepayment calculator"><Calculator size={15} /></button>
          )}
          <button onClick={() => onDelete(loan.id)} className="text-[#C7B9B3] hover:text-[#C4685B]" title="Delete loan"><Trash2 size={15} /></button>
        </div>
      </div>

      {derived.isActive ? (
        <>
          <div className="flex items-center justify-between">
            <StatusBadge status={derived.status} />
            <span style={{ fontFamily: F.mono }} className="text-xs text-[#A79A95]">Installment {derived.installmentNo}/{loan.tenureMonths}</span>
          </div>
          <div style={{ fontFamily: F.mono }} className="grid grid-cols-3 gap-2 text-xs">
            <div><p className="text-[#A79A95]">Due</p><p className="font-semibold text-[#3D3436]">{fmtINR(derived.due)}</p></div>
            <div><p className="text-[#A79A95]">Paid</p><p className="font-semibold text-[#5C8A5E]">{fmtINR(derived.paidThisMonth)}</p></div>
            <div><p className="text-[#A79A95]">Remaining</p><p className="font-semibold text-[#C4685B]">{fmtINR(derived.remaining)}</p></div>
          </div>
          {loan.interestRate > 0 && derived.due > 0 && (
            <div style={{ fontFamily: F.mono }} className="flex items-center justify-between text-[10.5px] text-[#A79A95] bg-[#FAF5F0] rounded-lg px-2.5 py-1.5">
              <span>Interest: <span className="text-[#C4685B] font-semibold">{fmtINR(interestPortion)}</span></span>
              <span>Principal: <span className="text-[#5C8A5E] font-semibold">{fmtINR(principalPortion)}</span></span>
            </div>
          )}
          {derived.remaining > 0 && (
            !showPartial ? (
              <div className="flex items-center gap-3">
                <button onClick={() => onPay(loan.id, derived.remaining)} className="text-xs font-semibold px-3 py-1.5 rounded-lg bg-[#EAF3EA] text-[#5C8A5E] hover:bg-[#DCEBDC] transition-colors flex items-center gap-1">
                  <CheckCircle2 size={13} /> Mark Paid ({fmtINR(derived.remaining)})
                </button>
                <button onClick={() => setShowPartial(true)} className="text-[11px] text-[#A79A95] hover:text-[#7A6B67] underline underline-offset-2">
                  Partial payment?
                </button>
              </div>
            ) : (
              <div className="flex items-center gap-2">
                <input type="number" autoFocus value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="Amount" style={{ fontFamily: F.mono }} className="w-24 text-xs px-2 py-1.5 rounded-lg border border-[#E4D5CE] focus:outline-none focus:border-[#C4899C]" />
                <button onClick={recordCustom} className="text-xs font-semibold px-2.5 py-1.5 rounded-lg bg-[#EAF3EA] text-[#5C8A5E] hover:bg-[#DCEBDC] transition-colors">Add</button>
                <button onClick={() => { setShowPartial(false); setAmount(''); }} className="text-[11px] text-[#A79A95] hover:text-[#7A6B67]">Cancel</button>
              </div>
            )
          )}
        </>
      ) : (
        <p className="text-xs text-[#A79A95]">
          {derived.installmentNo <= 0 ? `Shuru hoga ${fmtMonthLabel(loan.startMonth)} se` : 'Ye loan is month active nahi hai (tenure complete).'}
        </p>
      )}

      <div>
        <div className="flex items-center justify-between mb-1">
          <span className="text-[10.5px] text-[#A79A95] uppercase tracking-[0.08em] font-medium">Overall progress</span>
          <span style={{ fontFamily: F.mono }} className="text-[11px] text-[#A79A95]">{derived.progressPct.toFixed(0)}%</span>
        </div>
        <ProgressBar pct={derived.progressPct} color={cat.bar} />
        <div style={{ fontFamily: F.mono }} className="flex items-center justify-between mt-1.5 text-[11px] text-[#A79A95]">
          <span>Paid till date: {fmtINR(derived.totalPaidAllTime)}</span>
          <span>Outstanding: {fmtINR(derived.outstanding)}</span>
        </div>
      </div>

      {showPrepay && <PrepaymentModal F={F} loan={loan} derived={derived} onClose={() => setShowPrepay(false)} />}
    </div>
  );
}

function EMITab({ F, loans, payments, selectedMonth, addLoan, deleteLoan, recordPayment, deletePayment, showAddLoan, setShowAddLoan, loadExampleLoan }) {
  const today = new Date();
  const isRealCurrentMonth = selectedMonth === currentYM();
  const todayDay = today.getDate();

  const derivedLoans = useMemo(() => {
    return loans.map((loan) => {
      const monthsElapsed = monthDiff(loan.startMonth, selectedMonth);
      const installmentNo = monthsElapsed + 1;
      const isActive = monthsElapsed >= 0 && monthsElapsed < loan.tenureMonths;
      const loanPayments = payments.filter((p) => p.loanId === loan.id);
      const totalPaidAllTime = loanPayments.reduce((s, p) => s + p.amount, 0);
      const outstanding = Math.max(loan.totalAmount - totalPaidAllTime, 0);
      const progressPct = loan.totalAmount > 0 ? Math.min((totalPaidAllTime / loan.totalAmount) * 100, 100) : 0;
      const paidThisMonth = loanPayments.filter((p) => p.month === selectedMonth).reduce((s, p) => s + p.amount, 0);
      const due = isActive ? loan.emiAmount : 0;
      const remaining = Math.max(due - paidThisMonth, 0);
      let status = 'pending';
      if (isActive) {
        if (paidThisMonth >= due) status = 'paid';
        else if (paidThisMonth > 0) status = 'partial';
        else {
          const pastDue = (isRealCurrentMonth && loan.dueDay < todayDay) ||
            (!isRealCurrentMonth && ymToDate(selectedMonth) < ymToDate(currentYM()));
          status = pastDue ? 'overdue' : 'pending';
        }
      }
      return { loan, isActive, installmentNo, due, paidThisMonth, remaining, status, totalPaidAllTime, outstanding, progressPct };
    });
  }, [loans, payments, selectedMonth, isRealCurrentMonth, todayDay]);

  const activeThisMonth = derivedLoans.filter((d) => d.isActive);
  const totalDue = activeThisMonth.reduce((s, d) => s + d.due, 0);
  const totalPaid = activeThisMonth.reduce((s, d) => s + d.paidThisMonth, 0);
  const totalRemaining = activeThisMonth.reduce((s, d) => s + d.remaining, 0);
  const totalOutstandingAll = derivedLoans.reduce((s, d) => s + d.outstanding, 0);
  const overdueLoans = isRealCurrentMonth ? activeThisMonth.filter((d) => d.status === 'overdue') : [];
  const upcomingLoans = isRealCurrentMonth ? activeThisMonth.filter((d) => d.status === 'pending' && d.loan.dueDay >= todayDay && d.loan.dueDay - todayDay <= 7) : [];

  const monthPayments = useMemo(() => {
    return payments.filter((p) => p.month === selectedMonth).map((p) => ({ ...p, loan: loans.find((l) => l.id === p.loanId) }))
      .filter((p) => p.loan).sort((a, b) => (a.date < b.date ? 1 : -1));
  }, [payments, loans, selectedMonth]);

  const chartData = activeThisMonth.map((d) => ({
    name: d.loan.name.length > 12 ? d.loan.name.slice(0, 12) + '…' : d.loan.name,
    Due: d.due, Paid: d.paidThisMonth, fill: (LOAN_CATEGORY_COLORS[d.loan.category] || LOAN_CATEGORY_COLORS.Other).bar,
  }));

  return (
    <div className="space-y-5">
      <div className="flex justify-end gap-2">
        <button onClick={loadExampleLoan} className="flex items-center gap-1.5 text-xs px-3.5 py-2 rounded-lg bg-white border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-all cursor-pointer font-semibold">
          <Sparkles size={14} /> Try Example
        </button>
        <button onClick={() => setShowAddLoan(true)} style={{ background: 'linear-gradient(to right, #C4677A, #8F7CB8)', color: '#FFFFFF' }} className="flex items-center gap-1.5 text-xs px-3.5 py-2 rounded-lg hover:opacity-90 transition-all cursor-pointer shadow-lg font-semibold ring-1 ring-white/30">
          <PlusCircle size={14} /> Add Loan
        </button>
      </div>

      {overdueLoans.length > 0 && (
        <div className="rounded-xl border border-[#F0C6BE] bg-[#FBEAE7] p-3.5 flex items-start gap-2.5">
          <AlertTriangle size={16} className="text-[#C4685B] shrink-0 mt-0.5" />
          <div className="min-w-0">
            <p className="text-sm font-semibold text-[#C4685B]">{overdueLoans.length} EMI{overdueLoans.length > 1 ? 's' : ''} overdue hai{overdueLoans.length > 1 ? 'n' : ''}</p>
            <p className="text-xs text-[#C4685B]/80 mt-0.5">{overdueLoans.map((d) => d.loan.name).join(', ')}</p>
          </div>
        </div>
      )}
      {upcomingLoans.length > 0 && (
        <div className="rounded-xl border border-[#EDD6A8] bg-[#FBF0DE] p-3.5 flex items-start gap-2.5">
          <Clock size={16} className="text-[#B9862E] shrink-0 mt-0.5" />
          <div className="min-w-0">
            <p className="text-sm font-semibold text-[#B9862E]">Agle 7 dino me due</p>
            <p className="text-xs text-[#B9862E]/80 mt-0.5">{upcomingLoans.map((d) => `${d.loan.name} (${d.loan.dueDay} tareek)`).join(', ')}</p>
          </div>
        </div>
      )}

      <div className="grid grid-cols-1 xs:grid-cols-2 md:grid-cols-4 gap-3 items-stretch">
        <KPI F={F} label="Total EMI Due" value={fmtINR(totalDue)} numeric={totalDue} icon={Landmark} accent="text-[#B6A7D6]" sub={`${activeThisMonth.length} active loan${activeThisMonth.length === 1 ? '' : 's'}`} />
        <KPI F={F} label="Total Paid" value={fmtINR(totalPaid)} numeric={totalPaid} icon={TrendingUp} accent="text-[#5C8A5E]" />
        <KPI F={F} label="Total Remaining" value={fmtINR(totalRemaining)} numeric={totalRemaining} icon={TrendingDown} accent="text-[#C4685B]" />
        <KPI F={F} label="Outstanding (all loans)" value={fmtINR(totalOutstandingAll)} numeric={totalOutstandingAll} icon={PiggyBank} accent="text-[#8FAF8A]" />
      </div>

      {chartData.length > 0 && (
        <div className="rounded-xl border border-[#EFE1DA] bg-white p-4">
          <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold mb-3 text-[#3D3436]">Due vs Paid — {fmtMonthLabel(selectedMonth)}</h3>
          <div style={{ width: '100%', height: 220 }}>
            <ResponsiveContainer>
              <BarChart data={chartData} margin={{ top: 4, right: 8, left: -12, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#F1ECE9" vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 10, fill: '#A79A95' }} axisLine={{ stroke: '#EFE1DA' }} tickLine={false} />
                <YAxis tick={{ fontSize: 10, fill: '#A79A95' }} axisLine={false} tickLine={false} tickFormatter={(v) => `₹${v / 1000}k`} />
                <Tooltip formatter={(v) => fmtINR(v)} contentStyle={{ fontSize: 12, borderRadius: 8, border: '1px solid #EFE1DA', fontFamily: F.body }} />
                <Bar dataKey="Due" fill="#E4D5CE" radius={[4, 4, 0, 0]} />
                <Bar dataKey="Paid" radius={[4, 4, 0, 0]}>{chartData.map((d, i) => <Cell key={i} fill={d.fill} />)}</Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      <div className="space-y-3">
        <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold text-[#3D3436]">Loans</h3>
        {loans.length === 0 ? (
          <div className="rounded-xl border border-dashed border-[#E4D5CE] bg-white p-8 text-center">
            <Wallet size={24} className="mx-auto text-[#C7B9B3] mb-2" />
            <p className="text-sm text-[#A79A95]">Abhi koi loan/EMI add nahi hui. "Add Loan" par click karke shuru karo.</p>
          </div>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            {derivedLoans.map((d) => <LoanCard key={d.loan.id} F={F} loan={d.loan} derived={d} onPay={recordPayment} onDelete={deleteLoan} />)}
          </div>
        )}
      </div>

      {monthPayments.length > 0 && (
        <div className="rounded-xl border border-[#EFE1DA] bg-white overflow-hidden">
          <div className="px-4 py-3 border-b border-[#F1ECE9] flex items-center gap-2">
            <ClipboardList size={14} className="text-[#A79A95]" />
            <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold text-[#3D3436]">Payments — {fmtMonthLabel(selectedMonth)}</h3>
          </div>
          <table className="w-full text-xs">
            <thead><tr className="text-left text-[#A79A95] border-b border-[#F1ECE9]">
              <th className="px-4 py-2 font-medium">Date</th><th className="px-4 py-2 font-medium">Loan</th>
              <th className="px-4 py-2 font-medium text-right">Amount</th><th className="px-4 py-2"></th>
            </tr></thead>
            <tbody>
              {monthPayments.map((p) => (
                <tr key={p.id} className="border-b border-[#F8F3EF] last:border-0">
                  <td style={{ fontFamily: F.mono }} className="px-4 py-2 text-[#5C5250] whitespace-nowrap">{fmtDate(p.date)}</td>
                  <td className="px-4 py-2 text-[#3D3436]">{p.loan.name}</td>
                  <td style={{ fontFamily: F.mono }} className="px-4 py-2 text-right font-medium text-[#5C8A5E]">{fmtINR(p.amount)}</td>
                  <td className="px-4 py-2 text-right"><button onClick={() => deletePayment(p.id)} className="text-[#C7B9B3] hover:text-[#C4685B]"><X size={13} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {showAddLoan && <AddLoanModal F={F} onClose={() => setShowAddLoan(false)} onSave={addLoan} />}
    </div>
  );
}

/* ======================================================================
   BUDGET TAB
   ====================================================================== */

function downscaleImage(file, maxWidth, quality) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, maxWidth / img.width);
        const canvas = document.createElement('canvas');
        canvas.width = Math.round(img.width * scale);
        canvas.height = Math.round(img.height * scale);
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL('image/jpeg', quality));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function ExpenseModal({ F, onClose, onSave, initial }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(initial ? {
    category: initial.category, amount: String(initial.amount), date: initial.date,
    paymentMode: initial.paymentMode || 'Cash', paidBy: initial.paidBy || '', description: initial.description || '',
    receiptImage: initial.receiptImage || null,
  } : { category: 'Groceries', amount: '', date: todayISO(), paymentMode: 'Cash', paidBy: '', description: '', receiptImage: null });
  const [makeRecurring, setMakeRecurring] = useState(false);
  const [err, setErr] = useState('');
  const [categoryTouched, setCategoryTouched] = useState(isEdit);
  const [listening, setListening] = useState(false);
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  const setDescription = (e) => {
    const val = e.target.value;
    setForm((f) => {
      const guessed = !categoryTouched ? guessCategory(val) : null;
      return { ...f, description: val, category: guessed || f.category };
    });
  };
  const setCategory = (e) => { setCategoryTouched(true); setForm((f) => ({ ...f, category: e.target.value })); };

  const startVoice = () => {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) { alert('Ye browser voice input support nahi karta.'); return; }
    const rec = new SR();
    rec.lang = 'en-IN';
    rec.onresult = (e) => {
      const transcript = e.results[0][0].transcript;
      const parsed = parseVoiceExpense(transcript);
      setForm((f) => ({
        ...f,
        amount: parsed.amount !== null ? String(parsed.amount) : f.amount,
        description: parsed.description || f.description,
        category: parsed.category || f.category,
      }));
      if (parsed.category) setCategoryTouched(true);
      setListening(false);
    };
    rec.onerror = () => setListening(false);
    rec.onend = () => setListening(false);
    setListening(true);
    rec.start();
  };

  const handlePhoto = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const dataUrl = await downscaleImage(file, 480, 0.6);
      setForm((f) => ({ ...f, receiptImage: dataUrl }));
    } catch (err2) { /* non-fatal */ }
  };

  const submit = () => {
    const amt = parseFloat(form.amount);
    if (isNaN(amt) || amt <= 0) { setErr('Amount sahi daalo.'); return; }
    if (!form.date) { setErr('Date daalo.'); return; }
    onSave({
      id: initial ? initial.id : uid(), category: form.category, amount: amt, date: form.date,
      paymentMode: form.paymentMode, paidBy: form.paidBy.trim(), description: form.description.trim(), receiptImage: form.receiptImage,
    }, makeRecurring);
  };

  return (
    <ModalShell F={F} title={isEdit ? 'Expense Edit Karo' : 'Naya Expense Add Karo'} onClose={onClose}>
      <div className="space-y-3">
        {!isEdit && (
          <button
            onClick={startVoice}
            className={`w-full flex items-center justify-center gap-2 text-xs font-semibold py-2.5 rounded-lg border transition-colors ${listening ? 'bg-[#FBEAE7] border-[#F0C6BE] text-[#C4685B] animate-pulse' : 'border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5]'}`}
          >
            <Sparkles size={14} /> {listening ? 'Sun raha hoon…' : 'Bolke Add Karo (Voice)'}
          </button>
        )}
        <div><FieldLabel>Category {!categoryTouched && form.description && <span className="text-[#7C6BA6] normal-case font-normal">(auto-detected)</span>}</FieldLabel>
          <select value={form.category} onChange={setCategory} className={inputCls}>
            {EXPENSE_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div><FieldLabel>Amount (₹)</FieldLabel><input type="number" value={form.amount} onChange={set('amount')} placeholder="500" className={inputCls} /></div>
          <div><FieldLabel>Date</FieldLabel><input type="date" value={form.date} onChange={set('date')} className={inputCls} /></div>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div><FieldLabel>Payment Mode</FieldLabel>
            <select value={form.paymentMode} onChange={set('paymentMode')} className={inputCls}>
              {PAYMENT_MODES.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
          </div>
          <div><FieldLabel>Paid By (optional)</FieldLabel>
            <input value={form.paidBy} onChange={set('paidBy')} placeholder="e.g. Self, Ramesh" className={inputCls} />
          </div>
        </div>
        <div><FieldLabel>Description (optional)</FieldLabel><input value={form.description} onChange={setDescription} placeholder="e.g. Weekly groceries" className={inputCls} /></div>
        <div>
          <FieldLabel>Receipt photo (optional)</FieldLabel>
          {form.receiptImage ? (
            <div className="mt-1 flex items-center gap-2">
              <img src={form.receiptImage} alt="Receipt" className="w-14 h-14 object-cover rounded-lg border border-[#E4D5CE]" />
              <button onClick={() => setForm((f) => ({ ...f, receiptImage: null }))} className="text-[11px] text-[#C4685B] underline underline-offset-2">Remove</button>
            </div>
          ) : (
            <label className="mt-1 flex items-center justify-center gap-1.5 border border-dashed border-[#E4D5CE] rounded-lg py-2.5 cursor-pointer hover:bg-[#FBF3EF] transition-colors text-xs text-[#7A6B67]">
              <Camera size={14} /> Photo lo ya chuno
              <input type="file" accept="image/*" className="hidden" onChange={handlePhoto} />
            </label>
          )}
        </div>
        {!isEdit && (
          <label className="flex items-center gap-2 text-xs text-[#7A6B67] cursor-pointer">
            <input type="checkbox" checked={makeRecurring} onChange={(e) => setMakeRecurring(e.target.checked)} className="rounded border-[#E4D5CE]" />
            <Repeat size={13} className="text-[#A79A95]" /> Ye har mahine repeat hota hai (recurring)
          </label>
        )}
        {err && <p className="text-xs text-[#C4685B]">{err}</p>}
        <button onClick={submit} style={{ background: 'linear-gradient(to right, #C4677A, #8F7CB8)' }} className="w-full text-white text-sm font-semibold py-2.5 rounded-lg hover:opacity-90 transition-all shadow-md mt-1">{isEdit ? 'Save Changes' : 'Expense Add Karo'}</button>
      </div>
    </ModalShell>
  );
}

function ImportCSVModal({ F, onClose, onImport }) {
  const [fileName, setFileName] = useState('');
  const [preview, setPreview] = useState(null);
  const [err, setErr] = useState('');

  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setErr('');
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const { rows, skipped } = parseExpensesCSV(String(ev.target.result || ''));
        if (rows.length === 0) { setErr('Koi valid row nahi mili. Format check karo: date,amount,category,payment mode,paid by,description'); setPreview(null); return; }
        setPreview({ rows, skipped });
      } catch (e2) { setErr('CSV parse nahi ho payi.'); }
    };
    reader.readAsText(file);
  };

  const confirmImport = () => {
    if (!preview || preview.rows.length === 0) return;
    onImport(preview.rows);
  };

  return (
    <ModalShell F={F} title="Import from CSV" onClose={onClose}>
      <div className="space-y-3">
        <p className="text-xs text-[#A79A95]">
          CSV bina header ke — har row is order me: <span style={{ fontFamily: F.mono }} className="text-[#7A6B67]">date, amount, category, payment mode, paid by, description</span>
        </p>
        <label className="flex flex-col items-center justify-center gap-1.5 border border-dashed border-[#E4D5CE] rounded-xl py-6 cursor-pointer hover:bg-[#FBF3EF] transition-colors">
          <Upload size={18} className="text-[#A79A95]" />
          <span className="text-xs text-[#7A6B67] font-medium">{fileName || 'CSV file chuno'}</span>
          <input type="file" accept=".csv,text/csv" className="hidden" onChange={handleFile} />
        </label>
        {err && <p className="text-xs text-[#C4685B]">{err}</p>}
        {preview && (
          <div className="rounded-lg border border-[#EFE1DA] bg-[#FAF5F0] p-3 space-y-1.5 max-h-40 overflow-y-auto">
            <p className="text-xs font-semibold text-[#5C8A5E]">{preview.rows.length} rows ready to import{preview.skipped > 0 ? `, ${preview.skipped} skipped (invalid)` : ''}</p>
            {preview.rows.slice(0, 5).map((r) => (
              <p key={r.id} style={{ fontFamily: F.mono }} className="text-[10.5px] text-[#7A6B67]">
                {r.date} · {fmtINR(r.amount)} · {r.category} · {r.paymentMode}{r.paidBy ? ` · ${r.paidBy}` : ''}
              </p>
            ))}
            {preview.rows.length > 5 && <p className="text-[10.5px] text-[#A79A95]">…and {preview.rows.length - 5} more</p>}
          </div>
        )}
        <button
          onClick={confirmImport}
          disabled={!preview || preview.rows.length === 0}
          style={{ background: 'linear-gradient(to right, #C4677A, #8F7CB8)' }}
          className="w-full text-white text-sm font-semibold py-2.5 rounded-lg hover:opacity-90 transition-all shadow-md mt-1 disabled:opacity-40 disabled:cursor-not-allowed"
        >
          Import {preview ? `(${preview.rows.length})` : ''}
        </button>
      </div>
    </ModalShell>
  );
}

function RecurringManagerModal({ F, templates, onAdd, onDelete, onClose }) {
  const [form, setForm] = useState({ category: 'Rent', amount: '', paymentMode: 'Cash', paidBy: '', description: '', dayOfMonth: '1' });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const submit = () => {
    const amt = parseFloat(form.amount);
    const day = parseInt(form.dayOfMonth, 10);
    if (isNaN(amt) || amt <= 0 || isNaN(day) || day < 1 || day > 28) return;
    onAdd({ id: uid(), category: form.category, amount: amt, paymentMode: form.paymentMode, paidBy: form.paidBy.trim(), description: form.description.trim(), dayOfMonth: day });
    setForm({ category: 'Rent', amount: '', paymentMode: 'Cash', paidBy: '', description: '', dayOfMonth: '1' });
  };
  return (
    <ModalShell F={F} title="Recurring Expenses" onClose={onClose}>
      <div className="space-y-3">
        <p className="text-[11px] text-[#A79A95]">Ye har mahine automatically expense list me add ho jayenge (jab tum us mahine app khologe).</p>
        {templates.length > 0 && (
          <div className="rounded-lg border border-[#EFE1DA] divide-y divide-[#F1ECE9]">
            {templates.map((t) => (
              <div key={t.id} className="flex items-center justify-between px-3 py-2 text-xs">
                <div>
                  <p className="font-medium text-[#3D3436]">{t.category} — {fmtINR(t.amount)}</p>
                  <p className="text-[#A79A95]">Har mahine din {t.dayOfMonth} ko{t.paidBy ? ` · ${t.paidBy}` : ''}{t.description ? ` · ${t.description}` : ''}</p>
                </div>
                <button onClick={() => onDelete(t.id)} className="text-[#C7B9B3] hover:text-[#C4685B]"><X size={13} /></button>
              </div>
            ))}
          </div>
        )}
        <div className="rounded-lg border border-[#EFE1DA] bg-[#FAF5F0] p-3 space-y-2.5">
          <p className="text-xs font-semibold text-[#3D3436]">Naya recurring expense</p>
          <div className="grid grid-cols-2 gap-2">
            <select value={form.category} onChange={set('category')} className={inputCls}>
              {EXPENSE_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
            <input type="number" value={form.amount} onChange={set('amount')} placeholder="Amount" className={inputCls} />
          </div>
          <div className="grid grid-cols-2 gap-2">
            <select value={form.paymentMode} onChange={set('paymentMode')} className={inputCls}>
              {PAYMENT_MODES.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
            <input value={form.paidBy} onChange={set('paidBy')} placeholder="Paid by (optional)" className={inputCls} />
          </div>
          <input type="number" min="1" max="28" value={form.dayOfMonth} onChange={set('dayOfMonth')} placeholder="Day (1-28)" className={inputCls} />
          <input value={form.description} onChange={set('description')} placeholder="Description (optional)" className={inputCls} />
          <button onClick={submit} className="w-full text-xs font-semibold py-2 rounded-lg bg-[#EAF3EA] text-[#5C8A5E] hover:bg-[#DCEBDC] transition-colors">+ Add Recurring</button>
        </div>
      </div>
    </ModalShell>
  );
}

function BudgetTab({ F, budgets, setBudgetForMonth, expenses, addExpense, updateExpense, importExpenses, deleteExpense, selectedMonth, setSelectedMonth, emiPaidThisMonth, recurringTemplates, addRecurringTemplate, deleteRecurringTemplate }) {
  const [showAddExpense, setShowAddExpense] = useState(false);
  const [showImportCSV, setShowImportCSV] = useState(false);
  const [showRecurring, setShowRecurring] = useState(false);
  const [editingExpense, setEditingExpense] = useState(null);
  const [receiptView, setReceiptView] = useState(null);
  const [editing, setEditing] = useState(false);
  const monthBudget = budgets[selectedMonth] || {};
  const [draft, setDraft] = useState(monthBudget);

  useEffect(() => { setDraft(budgets[selectedMonth] || {}); setEditing(false); }, [selectedMonth, budgets]);

  const monthExpenses = useMemo(() => expenses.filter((e) => e.date.slice(0, 7) === selectedMonth), [expenses, selectedMonth]);

  // Anomaly detection: flag expenses well above the typical amount for their category
  const categoryAverages = useMemo(() => {
    const sums = {}, counts = {};
    expenses.forEach((e) => { sums[e.category] = (sums[e.category] || 0) + e.amount; counts[e.category] = (counts[e.category] || 0) + 1; });
    const avg = {};
    Object.keys(sums).forEach((c) => { avg[c] = sums[c] / counts[c]; });
    return avg;
  }, [expenses]);
  const isAnomaly = (e) => {
    const avg = categoryAverages[e.category];
    return avg && e.amount > avg * 2.2 && e.amount > 500;
  };

  const spentByCategory = useMemo(() => {
    const map = {};
    monthExpenses.forEach((e) => { map[e.category] = (map[e.category] || 0) + e.amount; });
    if (emiPaidThisMonth > 0) map['EMI/Loans'] = (map['EMI/Loans'] || 0) + emiPaidThisMonth;
    return map;
  }, [monthExpenses, emiPaidThisMonth]);

  const totalBudget = EXPENSE_CATEGORIES.reduce((s, c) => s + (Number(monthBudget[c]) || 0), 0);
  const totalSpent = Object.values(spentByCategory).reduce((s, v) => s + v, 0);
  const totalRemaining = totalBudget - totalSpent;

  const chartData = EXPENSE_CATEGORIES.filter((c) => (monthBudget[c] || spentByCategory[c])).map((c) => ({
    name: c.length > 10 ? c.slice(0, 10) + '…' : c, Budget: Number(monthBudget[c]) || 0, Spent: spentByCategory[c] || 0, fill: EXPENSE_CATEGORY_BARS[c],
  }));

  const copyPrevMonth = () => {
    const prev = budgets[addMonths(selectedMonth, -1)] || {};
    setDraft(prev);
  };
  const saveBudgets = () => { setBudgetForMonth(selectedMonth, draft); setEditing(false); };

  return (
    <div className="space-y-5">
      <div className="flex justify-end gap-2 flex-wrap">
        <button onClick={() => setShowRecurring(true)} className="flex items-center gap-1.5 text-xs px-3.5 py-2 rounded-lg bg-white border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-all cursor-pointer font-semibold">
          <Repeat size={14} /> Recurring{recurringTemplates.length > 0 ? ` (${recurringTemplates.length})` : ''}
        </button>
        <button onClick={() => setShowImportCSV(true)} className="flex items-center gap-1.5 text-xs px-3.5 py-2 rounded-lg bg-white border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-all cursor-pointer font-semibold">
          <Upload size={14} /> Import CSV
        </button>
        <button onClick={() => setShowAddExpense(true)} style={{ background: 'linear-gradient(to right, #C4677A, #8F7CB8)', color: '#FFFFFF' }} className="flex items-center gap-1.5 text-xs px-3.5 py-2 rounded-lg hover:opacity-90 transition-all cursor-pointer shadow-lg font-semibold ring-1 ring-white/30">
          <PlusCircle size={14} /> Add Expense
        </button>
      </div>

      <div className="grid grid-cols-1 xs:grid-cols-2 md:grid-cols-3 gap-3 items-stretch">
        <KPI F={F} label="Total Budget" value={fmtINR(totalBudget)} numeric={totalBudget} icon={PiggyBank} accent="text-[#B6A7D6]" />
        <KPI F={F} label="Total Spent" value={fmtINR(totalSpent)} numeric={totalSpent} icon={ShoppingBag} accent="text-[#C4685B]" />
        <KPI F={F} label={totalRemaining >= 0 ? 'Remaining' : 'Over Budget'} value={fmtINR(Math.abs(totalRemaining))} icon={totalRemaining >= 0 ? TrendingUp : TrendingDown} accent={totalRemaining >= 0 ? 'text-[#5C8A5E]' : 'text-[#C4685B]'} />
      </div>

      {chartData.length > 0 && (
        <div className="rounded-xl border border-[#EFE1DA] bg-white p-4">
          <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold mb-3 text-[#3D3436]">Budget vs Spent — {fmtMonthLabel(selectedMonth)}</h3>
          <div style={{ width: '100%', height: 220 }}>
            <ResponsiveContainer>
              <BarChart data={chartData} margin={{ top: 4, right: 8, left: -12, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#F1ECE9" vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 10, fill: '#A79A95' }} axisLine={{ stroke: '#EFE1DA' }} tickLine={false} />
                <YAxis tick={{ fontSize: 10, fill: '#A79A95' }} axisLine={false} tickLine={false} tickFormatter={(v) => `₹${v / 1000}k`} />
                <Tooltip formatter={(v) => fmtINR(v)} contentStyle={{ fontSize: 12, borderRadius: 8, border: '1px solid #EFE1DA', fontFamily: F.body }} />
                <Bar dataKey="Budget" fill="#E4D5CE" radius={[4, 4, 0, 0]} />
                <Bar dataKey="Spent" radius={[4, 4, 0, 0]}>{chartData.map((d, i) => <Cell key={i} fill={d.fill} />)}</Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      <div className="rounded-xl border border-[#EFE1DA] bg-white overflow-hidden">
        <div className="px-4 py-3 border-b border-[#F1ECE9] flex items-center justify-between flex-wrap gap-2">
          <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold text-[#3D3436]">Category-wise Budget</h3>
          {!editing ? (
            <button onClick={() => setEditing(true)} className="text-[11px] px-2.5 py-1.5 rounded-lg border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-colors">Edit budgets</button>
          ) : (
            <div className="flex items-center gap-2">
              <button onClick={copyPrevMonth} className="text-[11px] px-2.5 py-1.5 rounded-lg border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-colors">Copy last month</button>
              <button onClick={saveBudgets} className="text-[11px] px-2.5 py-1.5 rounded-lg bg-[#EAF3EA] text-[#5C8A5E] font-semibold hover:bg-[#DCEBDC] transition-colors">Save</button>
            </div>
          )}
        </div>
        <div className="divide-y divide-[#F8F3EF]">
          {EXPENSE_CATEGORIES.map((c) => {
            const b = editing ? (Number(draft[c]) || 0) : (Number(monthBudget[c]) || 0);
            const spent = spentByCategory[c] || 0;
            const pct = b > 0 ? Math.min((spent / b) * 100, 100) : (spent > 0 ? 100 : 0);
            const over = b > 0 && spent > b;
            return (
              <div key={c} className="px-4 py-3">
                <div className="flex items-center justify-between gap-2 mb-1.5">
                  <span className="text-xs font-medium text-[#3D3436]">{c}</span>
                  {editing ? (
                    <input type="number" value={draft[c] ?? ''} onChange={(e) => setDraft((d) => ({ ...d, [c]: e.target.value }))} placeholder="0" style={{ fontFamily: F.mono }} className="w-24 text-xs px-2 py-1 rounded-lg border border-[#E4D5CE] focus:outline-none focus:border-[#C4899C] text-right" />
                  ) : (
                    <span style={{ fontFamily: F.mono }} className={`text-xs font-semibold ${over ? 'text-[#C4685B]' : 'text-[#3D3436]'}`}>
                      {fmtINR(spent)} <span className="text-[#A79A95] font-normal">/ {fmtINR(b)}</span>
                    </span>
                  )}
                </div>
                {!editing && <ProgressBar pct={pct} color={over ? '#E2897C' : EXPENSE_CATEGORY_BARS[c]} />}
              </div>
            );
          })}
        </div>
      </div>

      {monthExpenses.length > 0 && (
        <div className="rounded-xl border border-[#EFE1DA] bg-white overflow-hidden">
          <div className="px-4 py-3 border-b border-[#F1ECE9] flex items-center gap-2">
            <ClipboardList size={14} className="text-[#A79A95]" />
            <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold text-[#3D3436]">Expenses — {fmtMonthLabel(selectedMonth)}</h3>
          </div>
          <table className="w-full text-xs">
            <thead><tr className="text-left text-[#A79A95] border-b border-[#F1ECE9]">
              <th className="px-4 py-2 font-medium">Date</th><th className="px-4 py-2 font-medium">Category</th>
              <th className="px-4 py-2 font-medium">Mode / By</th>
              <th className="px-4 py-2 font-medium">Description</th><th className="px-4 py-2 font-medium text-right">Amount</th><th className="px-4 py-2"></th>
            </tr></thead>
            <tbody>
              {monthExpenses.slice().sort((a, b) => (a.date < b.date ? 1 : -1)).map((e) => (
                <tr key={e.id} className="border-b border-[#F8F3EF] last:border-0">
                  <td style={{ fontFamily: F.mono }} className="px-4 py-2 text-[#5C5250] whitespace-nowrap">{fmtDate(e.date)}</td>
                  <td className="px-4 py-2 text-[#3D3436]">{e.category}</td>
                  <td className="px-4 py-2 text-[#A79A95] whitespace-nowrap">
                    {e.paymentMode || '—'}{e.paidBy ? <span className="block text-[10px] text-[#C7B9B3]">by {e.paidBy}</span> : null}
                  </td>
                  <td className="px-4 py-2 text-[#A79A95] max-w-[140px] truncate">
                    <div className="flex items-center gap-1.5">
                      {e.receiptImage && (
                        <button onClick={() => setReceiptView(e.receiptImage)} className="shrink-0"><img src={e.receiptImage} alt="Receipt" className="w-6 h-6 object-cover rounded border border-[#E4D5CE]" /></button>
                      )}
                      <span className="truncate">{e.description || '—'}</span>
                    </div>
                  </td>
                  <td style={{ fontFamily: F.mono }} className="px-4 py-2 text-right font-medium text-[#C4685B]">
                    <span className="inline-flex items-center gap-1 justify-end">
                      {isAnomaly(e) && <AlertTriangle size={11} className="text-[#B9862E]" title="Category average se kaafi zyada" />}
                      {fmtINR(e.amount)}
                    </span>
                  </td>
                  <td className="px-4 py-2 text-right whitespace-nowrap">
                    <button onClick={() => setEditingExpense(e)} className="text-[#C7B9B3] hover:text-[#7C6BA6] mr-2"><Edit2 size={13} /></button>
                    <button onClick={() => deleteExpense(e.id)} className="text-[#C7B9B3] hover:text-[#C4685B]"><X size={13} /></button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {showAddExpense && (
        <ExpenseModal
          F={F} onClose={() => setShowAddExpense(false)}
          onSave={(exp, recurring) => {
            addExpense(exp);
            if (recurring) addRecurringTemplate({ id: uid(), category: exp.category, amount: exp.amount, paymentMode: exp.paymentMode, paidBy: exp.paidBy, description: exp.description, dayOfMonth: parseInt(exp.date.slice(8, 10), 10) });
            setShowAddExpense(false);
          }}
        />
      )}
      {editingExpense && (
        <ExpenseModal
          F={F} initial={editingExpense} onClose={() => setEditingExpense(null)}
          onSave={(exp) => { updateExpense(exp); setEditingExpense(null); }}
        />
      )}
      {showImportCSV && (
        <ImportCSVModal
          F={F} onClose={() => setShowImportCSV(false)}
          onImport={(rows) => {
            importExpenses(rows);
            setShowImportCSV(false);
            // jump to the month with the most imported rows, since the table only shows the selected month
            const counts = {};
            rows.forEach((r) => { const ym = r.date.slice(0, 7); counts[ym] = (counts[ym] || 0) + 1; });
            const months = Object.entries(counts).sort((a, b) => b[1] - a[1]);
            if (months.length > 0) {
              setSelectedMonth(months[0][0]);
              const monthList = months.map(([ym, c]) => `${fmtMonthLabel(ym)} (${c})`).join(', ');
              alert(`${rows.length} expenses import ho gaye — ${monthList}. Dikha raha hoon ${fmtMonthLabel(months[0][0])}.`);
            }
          }}
        />
      )}
      {showRecurring && (
        <RecurringManagerModal F={F} templates={recurringTemplates} onAdd={addRecurringTemplate} onDelete={deleteRecurringTemplate} onClose={() => setShowRecurring(false)} />
      )}
      {receiptView && (
        <div className="fixed inset-0 z-50 bg-black/80 flex items-center justify-center p-4" onClick={() => setReceiptView(null)}>
          <img src={receiptView} alt="Receipt" className="max-w-full max-h-full rounded-lg" />
        </div>
      )}
    </div>
  );
}

/* ======================================================================
   KHATA BOOK TAB
   ====================================================================== */

function AddContactModal({ F, onClose, onSave }) {
  const [name, setName] = useState('');
  const [err, setErr] = useState('');
  const submit = () => {
    if (!name.trim()) { setErr('Please enter a name.'); return; }
    onSave({ id: uid(), name: name.trim() });
  };
  return (
    <ModalShell F={F} title="Add New Contact" onClose={onClose}>
      <div className="space-y-3">
        <div><FieldLabel>Name</FieldLabel><input value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. John" className={inputCls} autoFocus /></div>
        {err && <p className="text-xs text-[#C4685B]">{err}</p>}
        <button onClick={submit} style={{ background: 'linear-gradient(to right, #C4677A, #8F7CB8)' }} className="w-full text-white text-sm font-semibold py-2.5 rounded-lg hover:opacity-90 transition-all shadow-md mt-1">Add Contact</button>
      </div>
    </ModalShell>
  );
}

function ContactCard({ F, contact, transactions, onAddTx, onDeleteTx, onDeleteContact, onSetDueDate }) {
  const [open, setOpen] = useState(false);
  const [type, setType] = useState('gave');
  const [amount, setAmount] = useState('');
  const [note, setNote] = useState('');

  const gave = transactions.filter((t) => t.type === 'gave').reduce((s, t) => s + t.amount, 0);
  const got = transactions.filter((t) => t.type === 'got').reduce((s, t) => s + t.amount, 0);
  const balance = gave - got; // positive: they owe you. negative: you owe them.
  const isOverdue = contact.dueDate && balance !== 0 && contact.dueDate < todayISO();

  const submit = () => {
    const v = parseFloat(amount);
    if (isNaN(v) || v <= 0) return;
    onAddTx(contact.id, { id: uid(), type, amount: v, date: todayISO(), note: note.trim() });
    setAmount(''); setNote('');
  };

  const settleUp = () => {
    if (balance === 0) return;
    if (balance > 0) onAddTx(contact.id, { id: uid(), type: 'got', amount: balance, date: todayISO(), note: 'Settled up' });
    else onAddTx(contact.id, { id: uid(), type: 'gave', amount: Math.abs(balance), date: todayISO(), note: 'Settled up' });
  };

  const whatsappRemind = () => {
    let msg;
    if (balance > 0) msg = `Hi ${contact.name}, ek reminder — aapko mujhe ${fmtINR(balance)} dena hai. Jab convenient ho, settle kar dena. Dhanyavaad!`;
    else msg = `Hi ${contact.name}, ek reminder — mujhe aapko ${fmtINR(Math.abs(balance))} dena hai, jaldi settle kar dunga.`;
    window.open(`https://wa.me/?text=${encodeURIComponent(msg)}`, '_blank');
  };

  return (
    <div className="rounded-xl border border-[#EFE1DA] bg-white overflow-hidden">
      <div className="p-4 flex items-center justify-between gap-2 cursor-pointer" onClick={() => setOpen((o) => !o)}>
        <div className="min-w-0 flex items-center gap-3">
          <div className="w-9 h-9 rounded-full bg-[#F5EBE5] flex items-center justify-center shrink-0 font-semibold text-[#8F7CB8]" style={{ fontFamily: F.display }}>
            {contact.name.trim()[0]?.toUpperCase() || '?'}
          </div>
          <div className="min-w-0">
            <p style={{ fontFamily: F.display }} className="text-sm font-bold text-[#3D3436] truncate">{contact.name}</p>
            {balance === 0 ? (
              <p className="text-xs text-[#A79A95]">Settled</p>
            ) : balance > 0 ? (
              <p style={{ fontFamily: F.mono }} className="text-xs font-semibold text-[#5C8A5E]">You'll receive: {fmtINR(balance)}</p>
            ) : (
              <p style={{ fontFamily: F.mono }} className="text-xs font-semibold text-[#C4685B]">You owe: {fmtINR(Math.abs(balance))}</p>
            )}
            {isOverdue && <p className="text-[10.5px] text-[#C4685B] font-semibold flex items-center gap-1 mt-0.5"><AlertTriangle size={10} /> Overdue since {fmtDate(contact.dueDate)}</p>}
          </div>
        </div>
        <div className="flex items-center gap-2 shrink-0">
          <button onClick={(e) => { e.stopPropagation(); onDeleteContact(contact.id); }} className="text-[#C7B9B3] hover:text-[#C4685B]"><Trash2 size={15} /></button>
          {open ? <ChevronUp size={16} className="text-[#A79A95]" /> : <ChevronDown size={16} className="text-[#A79A95]" />}
        </div>
      </div>

      {open && (
        <div className="border-t border-[#F1ECE9] p-4 space-y-3">
          <div className="flex items-center gap-2 flex-wrap">
            <div className="flex rounded-lg border border-[#E4D5CE] overflow-hidden">
              <button onClick={() => setType('gave')} className={`text-xs font-semibold px-3 py-1.5 flex items-center gap-1 ${type === 'gave' ? 'bg-[#FBEAE7] text-[#C4685B]' : 'bg-white text-[#A79A95]'}`}>
                <ArrowUpRight size={12} /> You Gave
              </button>
              <button onClick={() => setType('got')} className={`text-xs font-semibold px-3 py-1.5 flex items-center gap-1 border-l border-[#E4D5CE] ${type === 'got' ? 'bg-[#EAF3EA] text-[#5C8A5E]' : 'bg-white text-[#A79A95]'}`}>
                <ArrowDownRight size={12} /> You Got
              </button>
            </div>
            <input type="number" value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="Amount" style={{ fontFamily: F.mono }} className="w-24 text-xs px-2 py-1.5 rounded-lg border border-[#E4D5CE] focus:outline-none focus:border-[#C4899C]" />
            <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="Note (optional)" className="flex-1 min-w-[100px] text-xs px-2 py-1.5 rounded-lg border border-[#E4D5CE] focus:outline-none focus:border-[#C4899C]" />
            <button onClick={submit} className="text-xs font-semibold px-3 py-1.5 rounded-lg border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-colors">Add</button>
            {balance !== 0 && (
              <button onClick={settleUp} className="text-xs font-semibold px-3 py-1.5 rounded-lg bg-[#EAF3EA] text-[#5C8A5E] hover:bg-[#DCEBDC] transition-colors">Settle Up</button>
            )}
            {balance !== 0 && (
              <button onClick={whatsappRemind} className="text-xs font-semibold px-3 py-1.5 rounded-lg bg-[#DFF3E4] text-[#3C8D4A] hover:bg-[#CDEBD4] transition-colors flex items-center gap-1">
                <MessageCircle size={13} /> Remind on WhatsApp
              </button>
            )}
          </div>

          <div className="flex items-center gap-2 text-xs">
            <Bell size={12} className="text-[#A79A95]" />
            <span className="text-[#A79A95]">Due date:</span>
            <input type="date" value={contact.dueDate || ''} onChange={(e) => onSetDueDate(contact.id, e.target.value)} style={{ fontFamily: F.mono }} className="text-xs px-2 py-1 rounded-lg border border-[#E4D5CE] focus:outline-none focus:border-[#C4899C]" />
          </div>

          {transactions.length > 0 && (
            <div className="space-y-1.5 max-h-48 overflow-y-auto">
              {transactions.slice().sort((a, b) => (a.date < b.date ? 1 : -1)).map((t) => (
                <div key={t.id} className="flex items-center justify-between text-xs bg-[#FAF5F0] rounded-lg px-3 py-2">
                  <div className="flex items-center gap-2 min-w-0">
                    {t.type === 'gave' ? <ArrowUpRight size={13} className="text-[#C4685B] shrink-0" /> : <ArrowDownRight size={13} className="text-[#5C8A5E] shrink-0" />}
                    <div className="min-w-0">
                      <span style={{ fontFamily: F.mono }} className={`font-semibold ${t.type === 'gave' ? 'text-[#C4685B]' : 'text-[#5C8A5E]'}`}>{fmtINR(t.amount)}</span>
                      {t.note && <span className="text-[#A79A95] ml-2 truncate">{t.note}</span>}
                    </div>
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    <span style={{ fontFamily: F.mono }} className="text-[#A79A95]">{fmtDate(t.date)}</span>
                    <button onClick={() => onDeleteTx(t.id)} className="text-[#C7B9B3] hover:text-[#C4685B]"><X size={12} /></button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function KhataTab({ F, contacts, transactions, addContact, deleteContact, addTx, deleteTx, setContactDueDate }) {
  const [showAddContact, setShowAddContact] = useState(false);

  const balances = useMemo(() => {
    const map = {};
    contacts.forEach((c) => {
      const tx = transactions.filter((t) => t.contactId === c.id);
      const gave = tx.filter((t) => t.type === 'gave').reduce((s, t) => s + t.amount, 0);
      const got = tx.filter((t) => t.type === 'got').reduce((s, t) => s + t.amount, 0);
      map[c.id] = gave - got;
    });
    return map;
  }, [contacts, transactions]);

  const toReceive = Object.values(balances).filter((b) => b > 0).reduce((s, b) => s + b, 0);
  const toPay = Object.values(balances).filter((b) => b < 0).reduce((s, b) => s + Math.abs(b), 0);

  const sortedContacts = contacts.slice().sort((a, b) => Math.abs(balances[b.id] || 0) - Math.abs(balances[a.id] || 0));

  return (
    <div className="space-y-5">
      <div className="flex justify-end">
        <button onClick={() => setShowAddContact(true)} style={{ background: 'linear-gradient(to right, #C4677A, #8F7CB8)', color: '#FFFFFF' }} className="flex items-center gap-1.5 text-xs px-3.5 py-2 rounded-lg hover:opacity-90 transition-all cursor-pointer shadow-lg font-semibold ring-1 ring-white/30">
          <PlusCircle size={14} /> Add Contact
        </button>
      </div>

      <div className="grid grid-cols-1 xs:grid-cols-2 md:grid-cols-3 gap-3 items-stretch">
        <KPI F={F} label="To Receive" value={fmtINR(toReceive)} numeric={toReceive} icon={ArrowDownRight} accent="text-[#5C8A5E]" sub="Total receivable" />
        <KPI F={F} label="To Pay" value={fmtINR(toPay)} numeric={toPay} icon={ArrowUpRight} accent="text-[#C4685B]" sub="Total payable" />
        <KPI F={F} label="Net Position" value={fmtINR(toReceive - toPay)} icon={Users} accent={toReceive - toPay >= 0 ? 'text-[#5C8A5E]' : 'text-[#C4685B]'} sub={`${contacts.length} contact${contacts.length === 1 ? '' : 's'}`} />
      </div>

      <div className="space-y-3">
        {contacts.length === 0 ? (
          <div className="rounded-xl border border-dashed border-[#E4D5CE] bg-white p-8 text-center">
            <BookOpen size={24} className="mx-auto text-[#C7B9B3] mb-2" />
            <p className="text-sm text-[#A79A95]">No contacts yet. Start with "Add Contact" — friends, family, vendors, anyone.</p>
          </div>
        ) : (
          sortedContacts.map((c) => (
            <ContactCard
              key={c.id} F={F} contact={c}
              transactions={transactions.filter((t) => t.contactId === c.id)}
              onAddTx={addTx} onDeleteTx={deleteTx} onDeleteContact={deleteContact} onSetDueDate={setContactDueDate}
            />
          ))
        )}
      </div>

      {showAddContact && <AddContactModal F={F} onClose={() => setShowAddContact(false)} onSave={(c) => { addContact(c); setShowAddContact(false); }} />}
    </div>
  );
}

/* ======================================================================
   PASSBOOK TAB — advanced summary, trends & saving tips
   ====================================================================== */

function InsightCard({ icon: Icon, tone, children }) {
  const tones = {
    good: { bg: 'bg-[#EAF3EA]', text: 'text-[#5C8A5E]', ring: 'ring-[#BFDDC0]' },
    warn: { bg: 'bg-[#FBF0DE]', text: 'text-[#B9862E]', ring: 'ring-[#EDD6A8]' },
    bad: { bg: 'bg-[#FBEAE7]', text: 'text-[#C4685B]', ring: 'ring-[#F0C6BE]' },
    info: { bg: 'bg-[#EDE7F6]', text: 'text-[#7C6BA6]', ring: 'ring-[#D6CCEA]' },
  };
  const c = tones[tone] || tones.info;
  return (
    <div className={`rounded-xl p-3.5 flex items-start gap-2.5 ring-1 ${c.bg} ${c.ring}`}>
      <div className={`w-7 h-7 rounded-lg flex items-center justify-center shrink-0 bg-white/70 ${c.text}`}>
        <Icon size={14} />
      </div>
      <p className={`text-xs leading-relaxed ${c.text}`}>{children}</p>
    </div>
  );
}

function GoalModal({ F, onClose, onSave }) {
  const [form, setForm] = useState({ name: '', targetAmount: '', targetDate: '', savedAmount: '' });
  const [err, setErr] = useState('');
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const submit = () => {
    const target = parseFloat(form.targetAmount);
    const saved = form.savedAmount === '' ? 0 : parseFloat(form.savedAmount);
    if (!form.name.trim()) { setErr('Goal ka naam daalo.'); return; }
    if (isNaN(target) || target <= 0) { setErr('Target amount sahi daalo.'); return; }
    onSave({ id: uid(), name: form.name.trim(), targetAmount: target, targetDate: form.targetDate, savedAmount: isNaN(saved) ? 0 : saved });
  };
  return (
    <ModalShell F={F} title="Naya Saving Goal" onClose={onClose}>
      <div className="space-y-3">
        <div><FieldLabel>Goal ka naam</FieldLabel><input value={form.name} onChange={set('name')} placeholder="e.g. Emergency Fund" className={inputCls} autoFocus /></div>
        <div className="grid grid-cols-2 gap-3">
          <div><FieldLabel>Target amount (₹)</FieldLabel><input type="number" value={form.targetAmount} onChange={set('targetAmount')} placeholder="100000" className={inputCls} /></div>
          <div><FieldLabel>Target date (optional)</FieldLabel><input type="date" value={form.targetDate} onChange={set('targetDate')} className={inputCls} /></div>
        </div>
        <div><FieldLabel>Already saved (optional)</FieldLabel><input type="number" value={form.savedAmount} onChange={set('savedAmount')} placeholder="0" className={inputCls} /></div>
        {err && <p className="text-xs text-[#C4685B]">{err}</p>}
        <button onClick={submit} style={{ background: 'linear-gradient(to right, #C4677A, #8F7CB8)' }} className="w-full text-white text-sm font-semibold py-2.5 rounded-lg hover:opacity-90 transition-all shadow-md mt-1">Goal Add Karo</button>
      </div>
    </ModalShell>
  );
}

function GoalCard({ F, goal, onAddSaved, onDelete }) {
  const [amt, setAmt] = useState('');
  const pct = Math.min((goal.savedAmount / goal.targetAmount) * 100, 100);
  const add = () => {
    const v = parseFloat(amt);
    if (isNaN(v) || v <= 0) return;
    onAddSaved(goal.id, v);
    setAmt('');
  };
  return (
    <div className="rounded-lg border border-[#EFE1DA] bg-white p-3 space-y-2">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-1.5"><Trophy size={13} className="text-[#B9862E]" /><span className="text-xs font-semibold text-[#3D3436]">{goal.name}</span></div>
        <button onClick={() => onDelete(goal.id)} className="text-[#C7B9B3] hover:text-[#C4685B]"><X size={12} /></button>
      </div>
      <ProgressBar pct={pct} color="#B9862E" />
      <div style={{ fontFamily: F.mono }} className="flex items-center justify-between text-[10.5px] text-[#A79A95]">
        <span>{fmtINR(goal.savedAmount)} / {fmtINR(goal.targetAmount)} ({pct.toFixed(0)}%)</span>
        {goal.targetDate && <span>by {fmtDate(goal.targetDate)}</span>}
      </div>
      <div className="flex items-center gap-2">
        <input type="number" value={amt} onChange={(e) => setAmt(e.target.value)} placeholder="Add saved amount" style={{ fontFamily: F.mono }} className="flex-1 text-xs px-2 py-1.5 rounded-lg border border-[#E4D5CE] focus:outline-none focus:border-[#C4899C]" />
        <button onClick={add} className="text-xs font-semibold px-2.5 py-1.5 rounded-lg bg-[#EAF3EA] text-[#5C8A5E] hover:bg-[#DCEBDC] transition-colors">Add</button>
      </div>
    </div>
  );
}

function WhatIfModal({ F, expenses, nowYM, thisMonthOutflow, incomeNum, hasIncome, onClose }) {
  const [cutPct, setCutPct] = useState(15);
  const FIXED_CATS = ['Rent', 'EMI/Loans', 'Utilities & Bills'];
  const discretionary = expenses.filter((e) => e.date.slice(0, 7) === nowYM && !FIXED_CATS.includes(e.category)).reduce((s, e) => s + e.amount, 0);
  const saved = discretionary * (cutPct / 100);
  const newOutflow = Math.max(thisMonthOutflow - saved, 0);
  const newSavingsRate = hasIncome ? ((incomeNum - newOutflow) / incomeNum) * 100 : null;
  return (
    <ModalShell F={F} title="What-If Simulator" onClose={onClose}>
      <div className="space-y-4">
        <p className="text-xs text-[#A79A95]">Discretionary spending (Rent/EMI/Utilities chhodke) is month me: <span className="font-semibold text-[#3D3436]">{fmtINR(discretionary)}</span></p>
        <div>
          <div className="flex items-center justify-between mb-1">
            <FieldLabel>Agar main {cutPct}% kam kharch karu…</FieldLabel>
          </div>
          <input type="range" min="0" max="50" value={cutPct} onChange={(e) => setCutPct(parseInt(e.target.value, 10))} className="w-full accent-[#8F7CB8]" />
        </div>
        <div style={{ fontFamily: F.mono }} className="rounded-lg bg-[#EAF3EA] p-3 space-y-1.5 text-xs">
          <div className="flex justify-between"><span className="text-[#5C8A5E] font-sans">Monthly savings</span><span className="font-semibold text-[#5C8A5E]">{fmtINR(saved)}</span></div>
          <div className="flex justify-between"><span className="text-[#5C8A5E] font-sans">New monthly outflow</span><span className="font-semibold text-[#3D3436]">{fmtINR(newOutflow)}</span></div>
          {hasIncome && <div className="flex justify-between"><span className="text-[#5C8A5E] font-sans">New savings rate</span><span className="font-semibold text-[#5C8A5E]">{newSavingsRate.toFixed(0)}%</span></div>}
          <div className="flex justify-between"><span className="text-[#5C8A5E] font-sans">Over 1 year, that's</span><span className="font-semibold text-[#5C8A5E]">{fmtINR(saved * 12)}</span></div>
        </div>
      </div>
    </ModalShell>
  );
}

function PassbookTab({ F, loans, payments, budgets, expenses, contacts, khataTx, income, setIncome, goals, addGoal, deleteGoal, addGoalSaved }) {
  const [incomeDraft, setIncomeDraft] = useState(income || '');
  useEffect(() => { setIncomeDraft(income || ''); }, [income]);
  const [showGoal, setShowGoal] = useState(false);
  const [showWhatIf, setShowWhatIf] = useState(false);
  const [trendRange, setTrendRange] = useState('6m'); // '6m' | 'year'

  const nowYM = currentYM();
  const last6 = useMemo(() => {
    const arr = [];
    for (let i = 5; i >= 0; i--) arr.push(addMonths(nowYM, -i));
    return arr;
  }, [nowYM]);

  const emiPaidByMonth = useMemo(() => {
    const map = {};
    payments.forEach((p) => { map[p.month] = (map[p.month] || 0) + p.amount; });
    return map;
  }, [payments]);

  const expenseByMonth = useMemo(() => {
    const map = {};
    expenses.forEach((e) => { const ym = e.date.slice(0, 7); map[ym] = (map[ym] || 0) + e.amount; });
    return map;
  }, [expenses]);

  const yearlyTotals = useMemo(() => {
    const map = {};
    expenses.forEach((e) => { const y = e.date.slice(0, 4); map[y] = (map[y] || { expenses: 0, emi: 0 }); map[y].expenses += e.amount; });
    payments.forEach((p) => { const y = p.month.slice(0, 4); map[y] = (map[y] || { expenses: 0, emi: 0 }); map[y].emi += p.amount; });
    return Object.entries(map).sort((a, b) => b[0].localeCompare(a[0])).map(([year, v]) => ({ year, ...v, total: v.expenses + v.emi }));
  }, [expenses, payments]);

  const trendData = last6.map((ym) => ({
    name: fmtMonthLabel(ym).split(' ')[0].slice(0, 3),
    Expenses: expenseByMonth[ym] || 0,
    EMI: emiPaidByMonth[ym] || 0,
  }));

  const thisMonthOutflow = (expenseByMonth[nowYM] || 0) + (emiPaidByMonth[nowYM] || 0);
  const prevYM = addMonths(nowYM, -1);
  const prevMonthOutflow = (expenseByMonth[prevYM] || 0) + (emiPaidByMonth[prevYM] || 0);
  const trendPct = prevMonthOutflow > 0 ? ((thisMonthOutflow - prevMonthOutflow) / prevMonthOutflow) * 100 : null;

  const monthsWithData = last6.filter((ym) => (expenseByMonth[ym] || 0) + (emiPaidByMonth[ym] || 0) > 0);
  const avgMonthlyOutflow = monthsWithData.length > 0
    ? monthsWithData.reduce((s, ym) => s + (expenseByMonth[ym] || 0) + (emiPaidByMonth[ym] || 0), 0) / monthsWithData.length
    : 0;

  const categoryTotals = useMemo(() => {
    const map = {};
    expenses.forEach((e) => { map[e.category] = (map[e.category] || 0) + e.amount; });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [expenses]);
  const topCategory = categoryTotals[0];

  const categoryChartData = categoryTotals.slice(0, 6).map(([cat, amt]) => ({
    name: cat.length > 10 ? cat.slice(0, 10) + '…' : cat, Amount: amt, fill: EXPENSE_CATEGORY_BARS[cat] || '#ACA09B',
  }));

  const monthBudget = budgets[nowYM] || {};
  const overBudgetCats = EXPENSE_CATEGORIES.filter((c) => {
    const b = Number(monthBudget[c]) || 0;
    const spent = expenses.filter((e) => e.date.slice(0, 7) === nowYM && e.category === c).reduce((s, e) => s + e.amount, 0);
    return b > 0 && spent > b;
  });

  const totalOutstandingLoans = useMemo(() => {
    return loans.reduce((s, l) => {
      const paid = payments.filter((p) => p.loanId === l.id).reduce((x, p) => x + p.amount, 0);
      return s + Math.max(l.totalAmount - paid, 0);
    }, 0);
  }, [loans, payments]);
  const totalEMIDueThisMonth = useMemo(() => {
    return loans.reduce((s, l) => {
      const monthsElapsed = monthDiff(l.startMonth, nowYM);
      const isActive = monthsElapsed >= 0 && monthsElapsed < l.tenureMonths;
      return isActive ? s + l.emiAmount : s;
    }, 0);
  }, [loans, nowYM]);

  const khataBalances = useMemo(() => {
    const map = {};
    contacts.forEach((c) => {
      const tx = khataTx.filter((t) => t.contactId === c.id);
      const gave = tx.filter((t) => t.type === 'gave').reduce((s, t) => s + t.amount, 0);
      const got = tx.filter((t) => t.type === 'got').reduce((s, t) => s + t.amount, 0);
      map[c.id] = gave - got;
    });
    return map;
  }, [contacts, khataTx]);
  const khataToReceive = Object.values(khataBalances).filter((b) => b > 0).reduce((s, b) => s + b, 0);
  const khataToPay = Object.values(khataBalances).filter((b) => b < 0).reduce((s, b) => s + Math.abs(b), 0);

  const incomeNum = parseFloat(incomeDraft);
  const hasIncome = !isNaN(incomeNum) && incomeNum > 0;
  const savingsRate = hasIncome ? ((incomeNum - thisMonthOutflow) / incomeNum) * 100 : null;

  // --- Spending forecast (linear projection of current month based on days elapsed) ---
  const today = new Date();
  const daysElapsed = today.getDate();
  const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const forecastMonthEnd = daysElapsed > 0 ? (thisMonthOutflow / daysElapsed) * daysInMonth : thisMonthOutflow;

  // --- Budget compliance streak (consecutive past months within budget) ---
  const budgetStreak = useMemo(() => {
    let streak = 0;
    for (let i = 1; i <= 24; i++) {
      const ym = addMonths(nowYM, -i);
      const b = budgets[ym];
      if (!b) break;
      const totalB = EXPENSE_CATEGORIES.reduce((s, c) => s + (Number(b[c]) || 0), 0);
      if (totalB <= 0) break;
      const spent = expenses.filter((e) => e.date.slice(0, 7) === ym).reduce((s, e) => s + e.amount, 0);
      if (spent <= totalB) streak++; else break;
    }
    return streak;
  }, [nowYM, budgets, expenses]);

  // --- EMI on-time streak (consecutive past months where all due EMI was paid) ---
  const emiStreak = useMemo(() => {
    let streak = 0;
    for (let i = 1; i <= 24; i++) {
      const ym = addMonths(nowYM, -i);
      const activeLoans = loans.filter((l) => { const m = monthDiff(l.startMonth, ym); return m >= 0 && m < l.tenureMonths; });
      if (activeLoans.length === 0) break;
      const due = activeLoans.reduce((s, l) => s + l.emiAmount, 0);
      const paid = payments.filter((p) => p.month === ym).reduce((s, p) => s + p.amount, 0);
      if (paid >= due) streak++; else break;
    }
    return streak;
  }, [nowYM, loans, payments]);

  // --- Weekly digest (last 7 days vs previous 7 days) ---
  const weeklyDigest = useMemo(() => {
    const now = new Date();
    const cutoff1 = new Date(now); cutoff1.setDate(now.getDate() - 7);
    const cutoff2 = new Date(now); cutoff2.setDate(now.getDate() - 14);
    const iso = (d) => d.toISOString().slice(0, 10);
    const thisWeek = expenses.filter((e) => e.date >= iso(cutoff1) && e.date <= todayISO());
    const lastWeek = expenses.filter((e) => e.date >= iso(cutoff2) && e.date < iso(cutoff1));
    const thisWeekTotal = thisWeek.reduce((s, e) => s + e.amount, 0);
    const lastWeekTotal = lastWeek.reduce((s, e) => s + e.amount, 0);
    const catMap = {};
    thisWeek.forEach((e) => { catMap[e.category] = (catMap[e.category] || 0) + e.amount; });
    const topCat = Object.entries(catMap).sort((a, b) => b[1] - a[1])[0];
    return { thisWeekTotal, lastWeekTotal, topCat, count: thisWeek.length };
  }, [expenses]);

  // --- Spend heatmap (current month, day-wise intensity) ---
  const heatmapDays = useMemo(() => {
    const dayMap = {};
    expenses.filter((e) => e.date.slice(0, 7) === nowYM).forEach((e) => { dayMap[e.date] = (dayMap[e.date] || 0) + e.amount; });
    const maxDay = Math.max(1, ...Object.values(dayMap));
    const days = [];
    for (let d = 1; d <= daysInMonth; d++) {
      const dateStr = `${nowYM}-${String(d).padStart(2, '0')}`;
      days.push({ day: d, amount: dayMap[dateStr] || 0, intensity: (dayMap[dateStr] || 0) / maxDay });
    }
    return days;
  }, [expenses, nowYM, daysInMonth]);

  // --- Debt trajectory: past (actual) + future (projected) outstanding balance across all loans ---
  const debtTrajectory = useMemo(() => {
    const months = [];
    for (let i = 5; i >= 1; i--) months.push({ ym: addMonths(nowYM, -i), kind: 'past' });
    months.push({ ym: nowYM, kind: 'now' });
    for (let i = 1; i <= 5; i++) months.push({ ym: addMonths(nowYM, i), kind: 'future' });
    return months.map(({ ym, kind }) => {
      let balance;
      if (kind !== 'future') {
        balance = loans.reduce((s, l) => {
          const paidTill = payments.filter((p) => p.loanId === l.id && p.month <= ym).reduce((x, p) => x + p.amount, 0);
          return s + Math.max(l.totalAmount - paidTill, 0);
        }, 0);
      } else {
        const monthsAhead = monthDiff(nowYM, ym);
        balance = loans.reduce((s, l) => {
          const paidTillNow = payments.filter((p) => p.loanId === l.id && p.month <= nowYM).reduce((x, p) => x + p.amount, 0);
          const balNow = Math.max(l.totalAmount - paidTillNow, 0);
          const elapsedFromStart = monthDiff(l.startMonth, ym);
          const stillInTenure = elapsedFromStart < l.tenureMonths;
          const projected = stillInTenure ? Math.max(balNow - l.emiAmount * monthsAhead, 0) : 0;
          return s + projected;
        }, 0);
      }
      return { name: fmtMonthLabel(ym).split(' ')[0].slice(0, 3), Balance: Math.round(balance), kind };
    });
  }, [nowYM, loans, payments]);

  // --- Achievement badges ---
  const badges = useMemo(() => {
    const list = [];
    if (goals.some((g) => g.savedAmount >= g.targetAmount)) list.push({ label: 'Goal Achiever', icon: Trophy });
    if (budgetStreak >= 3) list.push({ label: `${budgetStreak}-Month Budget Streak`, icon: Sparkles });
    if (emiStreak >= 3) list.push({ label: `${emiStreak}-Month EMI Streak`, icon: CheckCircle2 });
    if (loans.some((l) => {
      const paid = payments.filter((p) => p.loanId === l.id).reduce((s, p) => s + p.amount, 0);
      return paid >= l.totalAmount;
    })) list.push({ label: 'Loan Fully Paid', icon: Landmark });
    if (hasIncome && savingsRate !== null && savingsRate >= 20) list.push({ label: '20%+ Saver', icon: PiggyBank });
    const monthsTracked = new Set(expenses.map((e) => e.date.slice(0, 7))).size;
    if (monthsTracked >= 3) list.push({ label: 'Consistent Tracker', icon: Activity });
    return list;
  }, [goals, budgetStreak, emiStreak, loans, payments, hasIncome, savingsRate, expenses]);

  const netWorthLite = khataToReceive - khataToPay - totalOutstandingLoans;

  // Smart, data-driven saving tips — ranked by relevance
  const tips = useMemo(() => {
    const list = [];
    if (hasIncome && savingsRate !== null) {
      if (savingsRate < 0) {
        list.push({ tone: 'bad', icon: AlertTriangle, text: `You're spending ${fmtINR(Math.abs(incomeNum - thisMonthOutflow))} more than your income this month. Trim non-essential categories first before the month closes.` });
      } else if (savingsRate < 20) {
        list.push({ tone: 'warn', icon: Target, text: `You're saving only ${savingsRate.toFixed(0)}% of income this month. Aim for the 50/30/20 rule — 20% saved automatically works best when moved right after payday.` });
      } else {
        list.push({ tone: 'good', icon: Sparkles, text: `Nice — you're saving ${savingsRate.toFixed(0)}% of income this month. Consider directing the surplus into a recurring SIP or high-interest savings account.` });
      }
    } else {
      list.push({ tone: 'info', icon: Lightbulb, text: 'Add your monthly income above to unlock personalised savings-rate insights and a real 50/30/20 breakdown.' });
    }

    if (topCategory) {
      const [cat, amt] = topCategory;
      list.push({ tone: 'info', icon: PieChart, text: `${cat} is your biggest spend overall at ${fmtINR(amt)}. A 10% cut here (${fmtINR(amt * 0.1)}) would meaningfully move your monthly savings.` });
    }

    if (overBudgetCats.length > 0) {
      list.push({ tone: 'bad', icon: AlertTriangle, text: `You're over budget this month in ${overBudgetCats.join(', ')}. Review these first — they're the fastest lever to fix your numbers.` });
    }

    if (trendPct !== null) {
      if (trendPct > 15) {
        list.push({ tone: 'warn', icon: TrendingUp, text: `Spending is up ${trendPct.toFixed(0)}% vs last month. Check for one-off purchases vs a genuine trend before it becomes a habit.` });
      } else if (trendPct < -10) {
        list.push({ tone: 'good', icon: TrendingDown, text: `Spending is down ${Math.abs(trendPct).toFixed(0)}% vs last month — great discipline, keep the streak going.` });
      }
    }

    if (hasIncome && totalEMIDueThisMonth > 0 && totalEMIDueThisMonth / incomeNum > 0.4) {
      list.push({ tone: 'bad', icon: Landmark, text: `EMIs eat up ${((totalEMIDueThisMonth / incomeNum) * 100).toFixed(0)}% of your income. Lenders flag anything above 40% as risky — avoid new debt until this comes down.` });
    }

    if (khataToPay > 0) {
      list.push({ tone: 'warn', icon: ArrowUpRight, text: `You owe ${fmtINR(khataToPay)} across your contacts. Settling this frees up cash flow and avoids awkward reminders.` });
    }
    if (khataToReceive > 0) {
      list.push({ tone: 'good', icon: ArrowDownRight, text: `${fmtINR(khataToReceive)} is owed to you. A gentle nudge to collect it is effectively free money sitting idle.` });
    }

    if (avgMonthlyOutflow > 0) {
      list.push({ tone: 'info', icon: PiggyBank, text: `Your average monthly outflow is ${fmtINR(avgMonthlyOutflow)}. A healthy emergency fund would be ${fmtINR(avgMonthlyOutflow * 3)}–${fmtINR(avgMonthlyOutflow * 6)} (3–6 months of expenses) kept in a liquid, separate account.` });
    }

    return list.slice(0, 6);
  }, [hasIncome, savingsRate, incomeNum, thisMonthOutflow, topCategory, overBudgetCats, trendPct, totalEMIDueThisMonth, khataToPay, khataToReceive, avgMonthlyOutflow]);

  return (
    <div className="space-y-5">
      <div className="flex justify-end gap-2 print:hidden">
        <button onClick={() => setShowWhatIf(true)} className="flex items-center gap-1.5 text-xs px-3.5 py-2 rounded-lg bg-white border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-all cursor-pointer font-semibold">
          <Target size={14} /> What-If Simulator
        </button>
        <button onClick={() => window.print()} className="flex items-center gap-1.5 text-xs px-3.5 py-2 rounded-lg bg-white border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-all cursor-pointer font-semibold">
          <FileDown size={14} /> Download PDF Report
        </button>
      </div>

      {badges.length > 0 && (
        <div className="flex flex-wrap gap-2">
          {badges.map((b, i) => (
            <span key={i} className="inline-flex items-center gap-1.5 text-[11px] font-semibold px-3 py-1.5 rounded-full bg-gradient-to-r from-[#FBF3EF] to-[#F3EEFA] text-[#7C6BA6] ring-1 ring-[#E4D5CE]">
              <b.icon size={12} /> {b.label}
            </span>
          ))}
        </div>
      )}

      <div className="rounded-xl border border-[#EFE1DA] bg-white p-4 flex items-center gap-3 flex-wrap">
        <div className="flex items-center gap-2 text-[#7C6BA6]">
          <Wallet2 size={16} />
          <span className="text-xs font-semibold text-[#3D3436]">Monthly income</span>
          <span className="text-[10.5px] text-[#A79A95]">(optional — powers the tips below)</span>
        </div>
        <input
          type="number" value={incomeDraft} placeholder="e.g. 60000"
          onChange={(e) => setIncomeDraft(e.target.value)}
          onBlur={() => setIncome(incomeDraft)}
          style={{ fontFamily: F.mono }}
          className="ml-auto w-32 text-xs px-2.5 py-1.5 rounded-lg border border-[#E4D5CE] focus:outline-none focus:border-[#C4899C] text-right"
        />
      </div>

      <div className="grid grid-cols-1 xs:grid-cols-2 md:grid-cols-5 gap-3 items-stretch">
        <KPI F={F} label="This Month Outflow" value={fmtINR(thisMonthOutflow)} numeric={thisMonthOutflow} icon={Activity} accent="text-[#C4685B]"
          sub={trendPct !== null ? `${trendPct >= 0 ? '+' : ''}${trendPct.toFixed(0)}% vs last month` : undefined} />
        <KPI F={F} label="Forecast (month end)" value={fmtINR(forecastMonthEnd)} numeric={forecastMonthEnd} icon={Sparkles} accent="text-[#7C6BA6]" sub={`Based on day ${daysElapsed}/${daysInMonth}`} />
        <KPI F={F} label="Outstanding Loans" value={fmtINR(totalOutstandingLoans)} numeric={totalOutstandingLoans} icon={Landmark} accent="text-[#B6A7D6]" />
        <KPI F={F} label="Net Khata Position" value={fmtINR(khataToReceive - khataToPay)} icon={Users} accent={khataToReceive - khataToPay >= 0 ? 'text-[#5C8A5E]' : 'text-[#C4685B]'} />
        <KPI F={F} label="Savings Rate" value={hasIncome ? `${savingsRate.toFixed(0)}%` : '—'} icon={PiggyBank} accent={hasIncome && savingsRate >= 20 ? 'text-[#5C8A5E]' : hasIncome ? 'text-[#B9862E]' : 'text-[#A79A95]'}
          sub={hasIncome ? undefined : 'Add income above'} />
      </div>

      <div className="rounded-xl border border-[#EFE1DA] bg-white p-4">
        <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold mb-3 text-[#3D3436] flex items-center gap-1.5"><Clock size={14} className="text-[#A79A95]" /> Weekly Digest</h3>
        <div style={{ fontFamily: F.mono }} className="grid grid-cols-2 sm:grid-cols-3 gap-3 text-xs">
          <div><p className="text-[#A79A95] font-sans">Last 7 days</p><p className="font-semibold text-[#3D3436]">{fmtINR(weeklyDigest.thisWeekTotal)}</p></div>
          <div><p className="text-[#A79A95] font-sans">Previous 7 days</p><p className="font-semibold text-[#A79A95]">{fmtINR(weeklyDigest.lastWeekTotal)}</p></div>
          <div><p className="text-[#A79A95] font-sans">Top category</p><p className="font-semibold text-[#3D3436]">{weeklyDigest.topCat ? weeklyDigest.topCat[0] : '—'}</p></div>
        </div>
      </div>

      <div className="rounded-xl border border-[#EFE1DA] bg-white p-4">
        <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold mb-3 text-[#3D3436]">Spend Heatmap — {fmtMonthLabel(nowYM)}</h3>
        <div className="grid grid-cols-7 gap-1.5">
          {heatmapDays.map((d) => (
            <div
              key={d.day}
              title={`Day ${d.day}: ${fmtINR(d.amount)}`}
              className="aspect-square rounded flex items-center justify-center text-[9px] font-medium"
              style={{
                backgroundColor: d.amount > 0 ? `rgba(196, 103, 122, ${0.15 + d.intensity * 0.75})` : '#F5EFEA',
                color: d.intensity > 0.5 ? '#FFFFFF' : '#A79A95',
              }}
            >
              {d.day}
            </div>
          ))}
        </div>
      </div>

      {debtTrajectory.some((d) => d.Balance > 0) && (
        <div className="rounded-xl border border-[#EFE1DA] bg-white p-4">
          <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold mb-3 text-[#3D3436]">Debt Payoff Trajectory <span className="text-[#A79A95] font-normal text-xs">(past actual + projected)</span></h3>
          <div style={{ width: '100%', height: 200 }}>
            <ResponsiveContainer>
              <BarChart data={debtTrajectory} margin={{ top: 4, right: 8, left: -12, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#F1ECE9" vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 10, fill: '#A79A95' }} axisLine={{ stroke: '#EFE1DA' }} tickLine={false} />
                <YAxis tick={{ fontSize: 10, fill: '#A79A95' }} axisLine={false} tickLine={false} tickFormatter={(v) => `₹${v / 1000}k`} />
                <Tooltip formatter={(v) => fmtINR(v)} contentStyle={{ fontSize: 12, borderRadius: 8, border: '1px solid #EFE1DA', fontFamily: F.body }} />
                <Bar dataKey="Balance" radius={[4, 4, 0, 0]}>
                  {debtTrajectory.map((d, i) => <Cell key={i} fill={d.kind === 'future' ? '#D6CCEA' : '#9E8BC4'} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <p className="text-[10.5px] text-[#A79A95] mt-1">Halka purple = projected future (assuming EMI schedule as-is continues).</p>
        </div>
      )}

      {trendData.some((d) => d.Expenses || d.EMI) && (
        <div className="rounded-xl border border-[#EFE1DA] bg-white p-4">
          <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold mb-3 text-[#3D3436]">6-Month Outflow Trend</h3>
          <div style={{ width: '100%', height: 220 }}>
            <ResponsiveContainer>
              <BarChart data={trendData} margin={{ top: 4, right: 8, left: -12, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#F1ECE9" vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 10, fill: '#A79A95' }} axisLine={{ stroke: '#EFE1DA' }} tickLine={false} />
                <YAxis tick={{ fontSize: 10, fill: '#A79A95' }} axisLine={false} tickLine={false} tickFormatter={(v) => `₹${v / 1000}k`} />
                <Tooltip formatter={(v) => fmtINR(v)} contentStyle={{ fontSize: 12, borderRadius: 8, border: '1px solid #EFE1DA', fontFamily: F.body }} />
                <Bar dataKey="Expenses" stackId="a" fill="#C99CB0" radius={[0, 0, 0, 0]} />
                <Bar dataKey="EMI" stackId="a" fill="#9E8BC4" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      {categoryChartData.length > 0 && (
        <div className="rounded-xl border border-[#EFE1DA] bg-white p-4">
          <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold mb-3 text-[#3D3436]">Top Spend Categories (all time)</h3>
          <div style={{ width: '100%', height: 200 }}>
            <ResponsiveContainer>
              <BarChart data={categoryChartData} layout="vertical" margin={{ top: 4, right: 16, left: 8, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#F1ECE9" horizontal={false} />
                <XAxis type="number" tick={{ fontSize: 10, fill: '#A79A95' }} axisLine={false} tickLine={false} tickFormatter={(v) => `₹${v / 1000}k`} />
                <YAxis type="category" dataKey="name" tick={{ fontSize: 10, fill: '#A79A95' }} axisLine={false} tickLine={false} width={70} />
                <Tooltip formatter={(v) => fmtINR(v)} contentStyle={{ fontSize: 12, borderRadius: 8, border: '1px solid #EFE1DA', fontFamily: F.body }} />
                <Bar dataKey="Amount" radius={[0, 4, 4, 0]}>{categoryChartData.map((d, i) => <Cell key={i} fill={d.fill} />)}</Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      {yearlyTotals.length > 0 && (
        <div className="rounded-xl border border-[#EFE1DA] bg-white p-4">
          <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold mb-3 text-[#3D3436]">Year-wise Summary</h3>
          <div className="divide-y divide-[#F8F3EF]">
            {yearlyTotals.map((y) => (
              <div key={y.year} className="flex items-center justify-between py-2 text-xs" style={{ fontFamily: F.mono }}>
                <span className="font-semibold text-[#3D3436]">{y.year}</span>
                <span className="text-[#A79A95] font-sans">Expenses {fmtINR(y.expenses)} + EMI {fmtINR(y.emi)}</span>
                <span className="font-semibold text-[#C4685B]">{fmtINR(y.total)}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="rounded-xl border border-[#EFE1DA] bg-white p-4">
        <div className="flex items-center justify-between mb-3">
          <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold text-[#3D3436] flex items-center gap-1.5"><Trophy size={14} className="text-[#B9862E]" /> Saving Goals</h3>
          <button onClick={() => setShowGoal(true)} className="text-[11px] px-2.5 py-1.5 rounded-lg border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-colors flex items-center gap-1"><PlusCircle size={12} /> Add Goal</button>
        </div>
        {goals.length === 0 ? (
          <p className="text-xs text-[#A79A95]">Koi goal set nahi hai abhi. "Add Goal" se ek target set karo — jaise Emergency Fund ya Vacation.</p>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            {goals.map((g) => <GoalCard key={g.id} F={F} goal={g} onAddSaved={addGoalSaved} onDelete={deleteGoal} />)}
          </div>
        )}
        {showGoal && <GoalModal F={F} onClose={() => setShowGoal(false)} onSave={(g) => { addGoal(g); setShowGoal(false); }} />}
      </div>

      <div className="rounded-xl border border-[#EFE1DA] bg-white p-4">
        <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold mb-3 text-[#3D3436]">Compressed Summary</h3>
        <div style={{ fontFamily: F.mono }} className="grid grid-cols-2 sm:grid-cols-3 gap-x-3 gap-y-2.5 text-[11px]">
          <div><p className="text-[#A79A95] font-sans">EMI due this month</p><p className="font-semibold text-[#3D3436]">{fmtINR(totalEMIDueThisMonth)}</p></div>
          <div><p className="text-[#A79A95] font-sans">Loans outstanding</p><p className="font-semibold text-[#C4685B]">{fmtINR(totalOutstandingLoans)}</p></div>
          <div><p className="text-[#A79A95] font-sans">Budgeted this month</p><p className="font-semibold text-[#3D3436]">{fmtINR(EXPENSE_CATEGORIES.reduce((s, c) => s + (Number(monthBudget[c]) || 0), 0))}</p></div>
          <div><p className="text-[#A79A95] font-sans">Spent this month</p><p className="font-semibold text-[#C4685B]">{fmtINR(expenseByMonth[nowYM] || 0)}</p></div>
          <div><p className="text-[#A79A95] font-sans">Khata receivable</p><p className="font-semibold text-[#5C8A5E]">{fmtINR(khataToReceive)}</p></div>
          <div><p className="text-[#A79A95] font-sans">Khata payable</p><p className="font-semibold text-[#C4685B]">{fmtINR(khataToPay)}</p></div>
          <div className="col-span-2 sm:col-span-3 pt-2 mt-1 border-t border-[#F1ECE9] flex items-center justify-between">
            <span className="text-[#A79A95] font-sans">Net position (khata − loans)</span>
            <span className={`font-bold text-sm ${netWorthLite >= 0 ? 'text-[#5C8A5E]' : 'text-[#C4685B]'}`}>{fmtINR(netWorthLite)}</span>
          </div>
        </div>
      </div>

      <div className="rounded-xl p-4" style={{ background: 'linear-gradient(135deg, #FBF3EF, #F3EEFA)' }}>
        <div className="flex items-center gap-2 mb-3">
          <Sparkles size={16} className="text-[#8F7CB8]" />
          <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold text-[#3D3436]">Smart Saving Tips</h3>
        </div>
        <div className="space-y-2">
          {tips.map((t, i) => <InsightCard key={i} icon={t.icon} tone={t.tone}>{t.text}</InsightCard>)}
        </div>
      </div>

      {showWhatIf && (
        <WhatIfModal F={F} expenses={expenses} nowYM={nowYM} thisMonthOutflow={thisMonthOutflow} incomeNum={incomeNum} hasIncome={hasIncome} onClose={() => setShowWhatIf(false)} />
      )}
    </div>
  );
}

/* ======================================================================
   SETTINGS & PIN LOCK
   ====================================================================== */

function SettingsModal({ F, theme, toggleTheme, currencyCode, changeCurrency, pinLock, savePinLock, exportJSON, importJSON, onClose }) {
  const [newPin, setNewPin] = useState('');
  const fileRef = useRef(null);

  const enablePin = () => {
    if (newPin.length < 4) { alert('PIN kam se kam 4 digit ka rakho.'); return; }
    savePinLock({ enabled: true, pin: newPin });
    setNewPin('');
  };
  const disablePin = () => savePinLock({ enabled: false, pin: '' });

  return (
    <ModalShell F={F} title="Settings" onClose={onClose}>
      <div className="space-y-5">
        <div>
          <FieldLabel>Appearance</FieldLabel>
          <button onClick={toggleTheme} className="w-full flex items-center justify-between text-xs px-3 py-2.5 rounded-lg border border-[#E4D5CE] hover:bg-[#F5EBE5] transition-colors">
            <span className="flex items-center gap-2">{theme === 'dark' ? <Moon size={14} /> : <Sun size={14} />} {theme === 'dark' ? 'Dark mode' : 'Light mode'}</span>
            <span className="text-[#A79A95]">Switch to {theme === 'dark' ? 'Light' : 'Dark'}</span>
          </button>
          <p className="text-[10.5px] text-[#A79A95] mt-1">Dark mode ek quick colour-invert filter use karta hai — poora hand-tuned dark palette nahi, lekin low-light me use karne layak hai.</p>
        </div>

        <div>
          <FieldLabel>Currency</FieldLabel>
          <select value={currencyCode} onChange={(e) => changeCurrency(e.target.value)} className={inputCls}>
            {Object.keys(CURRENCIES).map((code) => <option key={code} value={code}>{code} ({CURRENCIES[code].symbol.trim()})</option>)}
          </select>
        </div>

        <div>
          <FieldLabel>Data backup</FieldLabel>
          <div className="flex gap-2">
            <button onClick={exportJSON} className="flex-1 flex items-center justify-center gap-1.5 text-xs px-3 py-2 rounded-lg border border-[#E4D5CE] hover:bg-[#F5EBE5] transition-colors"><FileDown size={13} /> Export JSON</button>
            <button onClick={() => fileRef.current?.click()} className="flex-1 flex items-center justify-center gap-1.5 text-xs px-3 py-2 rounded-lg border border-[#E4D5CE] hover:bg-[#F5EBE5] transition-colors"><Upload size={13} /> Restore JSON</button>
            <input ref={fileRef} type="file" accept="application/json" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) importJSON(f); }} />
          </div>
        </div>

        <div>
          <FieldLabel>PIN lock</FieldLabel>
          {pinLock.enabled ? (
            <button onClick={disablePin} className="w-full flex items-center justify-center gap-1.5 text-xs px-3 py-2 rounded-lg bg-[#FBEAE7] text-[#C4685B] hover:bg-[#F5DAD4] transition-colors"><Unlock size={13} /> Remove PIN lock</button>
          ) : (
            <div className="flex gap-2">
              <input type="password" inputMode="numeric" value={newPin} onChange={(e) => setNewPin(e.target.value.replace(/\D/g, '').slice(0, 8))} placeholder="Set a PIN (4+ digits)" style={{ fontFamily: F.mono }} className="flex-1 text-xs px-2.5 py-2 rounded-lg border border-[#E4D5CE] focus:outline-none focus:border-[#C4899C]" />
              <button onClick={enablePin} className="text-xs font-semibold px-3 py-2 rounded-lg bg-[#EAF3EA] text-[#5C8A5E] hover:bg-[#DCEBDC] transition-colors flex items-center gap-1"><Lock size={13} /> Set</button>
            </div>
          )}
          <p className="text-[10.5px] text-[#A79A95] mt-1">Ye sirf ek app-level PIN screen hai (jaise ek deterrent) — device-level biometric security nahi hai, kyunki ye ek browser-based app hai.</p>
        </div>
      </div>
    </ModalShell>
  );
}

function PinLockScreen({ F, pin, onUnlock }) {
  const [entered, setEntered] = useState('');
  const [err, setErr] = useState(false);
  const submit = () => {
    if (entered === pin) onUnlock();
    else { setErr(true); setEntered(''); }
  };
  return (
    <div style={{ fontFamily: F.body, backgroundColor: '#FAF5F0' }} className="min-h-screen flex items-center justify-center p-4">
      <div className="w-full max-w-xs text-center">
        <Lock size={28} className="mx-auto text-[#8F7CB8] mb-3" />
        <h2 style={{ fontFamily: F.display }} className="text-base font-bold text-[#3D3436] mb-1">Money Tracker Locked</h2>
        <p className="text-xs text-[#A79A95] mb-4">Apna PIN daalo</p>
        <input
          type="password" inputMode="numeric" autoFocus value={entered}
          onChange={(e) => { setEntered(e.target.value.replace(/\D/g, '').slice(0, 8)); setErr(false); }}
          onKeyDown={(e) => e.key === 'Enter' && submit()}
          style={{ fontFamily: F.mono }}
          className={`w-full text-center text-lg tracking-[0.3em] px-3 py-3 rounded-lg border ${err ? 'border-[#C4685B]' : 'border-[#E4D5CE]'} focus:outline-none focus:border-[#C4899C] mb-2`}
        />
        {err && <p className="text-xs text-[#C4685B] mb-2">Galat PIN, dobara try karo.</p>}
        <button onClick={submit} style={{ background: 'linear-gradient(to right, #C4677A, #8F7CB8)' }} className="w-full text-white text-sm font-semibold py-2.5 rounded-lg hover:opacity-90 transition-all shadow-md">Unlock</button>
      </div>
    </div>
  );
}

/* ======================================================================
   EXPLORER — filter by category + date, see who spent what & when
   ====================================================================== */

function ExplorerTab({ F, expenses }) {
  const [category, setCategory] = useState('All');
  const [paidByFilter, setPaidByFilter] = useState('All');
  const [paidByManual, setPaidByManual] = useState('');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  const [savedNote, setSavedNote] = useState(false);

  useEffect(() => {
    (async () => {
      if (!window.storage) return;
      try {
        const r = await window.storage.get('explorer-filter', false);
        if (r && r.value) {
          const f = JSON.parse(r.value);
          setCategory(f.category || 'All');
          setPaidByFilter(f.paidByFilter || 'All');
          setPaidByManual(f.paidByManual || '');
          setDateFrom(f.dateFrom || '');
          setDateTo(f.dateTo || '');
        }
      } catch (e) { /* no saved filter yet */ }
    })();
  }, []);

  const saveFilter = async () => {
    if (!window.storage) return;
    try {
      await window.storage.set('explorer-filter', JSON.stringify({ category, paidByFilter, paidByManual, dateFrom, dateTo }), false);
      setSavedNote(true);
      setTimeout(() => setSavedNote(false), 1800);
    } catch (e) { /* non-fatal */ }
  };

  const paidByOptions = useMemo(() => {
    const set = new Set(expenses.map((e) => e.paidBy?.trim()).filter(Boolean));
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [expenses]);

  const filtered = useMemo(() => {
    return expenses.filter((e) => {
      if (category !== 'All' && e.category !== category) return false;
      if (paidByFilter === 'Custom') {
        if (paidByManual.trim() && !(e.paidBy || '').toLowerCase().includes(paidByManual.trim().toLowerCase())) return false;
      } else if (paidByFilter !== 'All' && (e.paidBy?.trim() || 'Unknown') !== paidByFilter) return false;
      if (dateFrom && e.date < dateFrom) return false;
      if (dateTo && e.date > dateTo) return false;
      return true;
    }).sort((a, b) => (a.date < b.date ? 1 : -1));
  }, [expenses, category, paidByFilter, paidByManual, dateFrom, dateTo]);

  const total = filtered.reduce((s, e) => s + e.amount, 0);

  const byPerson = useMemo(() => {
    const map = {};
    filtered.forEach((e) => {
      const who = e.paidBy?.trim() || 'Unknown';
      map[who] = (map[who] || 0) + e.amount;
    });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [filtered]);

  const byDay = useMemo(() => {
    const map = {};
    filtered.forEach((e) => { map[e.date] = (map[e.date] || 0) + e.amount; });
    return Object.entries(map).sort((a, b) => (a[0] < b[0] ? 1 : -1));
  }, [filtered]);

  const clearFilters = () => { setCategory('All'); setPaidByFilter('All'); setPaidByManual(''); setDateFrom(''); setDateTo(''); };

  return (
    <div className="space-y-5">
      <div className="rounded-xl border border-[#EFE1DA] bg-white p-4 space-y-3">
        <div className="flex items-center gap-2">
          <Filter size={15} className="text-[#7C6BA6]" />
          <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold text-[#3D3436]">Filters</h3>
          <div className="ml-auto flex items-center gap-3">
            {savedNote && <span className="text-[11px] text-[#5C8A5E] font-medium">Saved ✓</span>}
            <button onClick={saveFilter} className="text-[11px] text-[#7C6BA6] hover:text-[#5A4A8A] font-semibold underline underline-offset-2">Save</button>
            {(category !== 'All' || paidByFilter !== 'All' || dateFrom || dateTo) && (
              <button onClick={clearFilters} className="text-[11px] text-[#A79A95] hover:text-[#C4685B] underline underline-offset-2">Clear</button>
            )}
          </div>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div><FieldLabel>Category</FieldLabel>
            <select value={category} onChange={(e) => setCategory(e.target.value)} className={inputCls}>
              <option value="All">All Categories</option>
              {EXPENSE_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div><FieldLabel>Paid By</FieldLabel>
            <select value={paidByFilter} onChange={(e) => setPaidByFilter(e.target.value)} className={inputCls}>
              <option value="All">Everyone</option>
              {paidByOptions.map((p) => <option key={p} value={p}>{p}</option>)}
              <option value="Custom">Type manually…</option>
            </select>
          </div>
        </div>
        {paidByFilter === 'Custom' && (
          <div>
            <FieldLabel>Enter name</FieldLabel>
            <input
              value={paidByManual} onChange={(e) => setPaidByManual(e.target.value)}
              placeholder="e.g. Nitish" autoFocus className={inputCls}
            />
          </div>
        )}
        <div>
          <FieldLabel>From date</FieldLabel>
          <input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} className={inputCls} />
        </div>
        <div>
          <FieldLabel>To date</FieldLabel>
          <input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)} className={inputCls} />
        </div>
      </div>

      <div className="rounded-xl border border-[#EFE1DA] bg-white p-4">
        <div className="flex items-center justify-between">
          <div>
            <p className="text-[10.5px] text-[#A79A95] uppercase tracking-[0.1em] font-medium">Total ({filtered.length} transaction{filtered.length === 1 ? '' : 's'})</p>
            <p style={{ fontFamily: F.mono }} className="text-xl font-bold text-[#C4685B] mt-1">{fmtINR(total)}</p>
          </div>
          <div className="flex items-center gap-1.5 flex-wrap justify-end">
            {category !== 'All' && (
              <span className="text-[10px] px-2.5 py-1 rounded-full font-semibold" style={{ backgroundColor: (EXPENSE_CATEGORY_BARS[category] || '#ACA09B') + '22', color: EXPENSE_CATEGORY_BARS[category] || '#ACA09B' }}>{category}</span>
            )}
            {paidByFilter !== 'All' && (
              <span className="text-[10px] px-2.5 py-1 rounded-full font-semibold bg-[#EDE7F6] text-[#7C6BA6]">
                {paidByFilter === 'Custom' ? (paidByManual.trim() || 'Type manually…') : paidByFilter}
              </span>
            )}
          </div>
        </div>
      </div>

      {byPerson.length > 0 && (
        <div className="rounded-xl border border-[#EFE1DA] bg-white p-4">
          <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold mb-3 text-[#3D3436]">By Person</h3>
          <div className="space-y-2">
            {byPerson.map(([who, amt]) => (
              <div key={who} className="flex items-center justify-between text-xs">
                <span className="text-[#3D3436] font-medium">{who}</span>
                <div className="flex items-center gap-2 flex-1 mx-3">
                  <div className="h-1.5 flex-1 bg-[#F1ECE9] rounded-full overflow-hidden">
                    <div className="h-full rounded-full" style={{ width: `${(amt / total) * 100}%`, backgroundColor: '#8F7CB8' }} />
                  </div>
                </div>
                <span style={{ fontFamily: F.mono }} className="font-semibold text-[#3D3436]">{fmtINR(amt)}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {byDay.length > 0 && (
        <div className="rounded-xl border border-[#EFE1DA] bg-white p-4">
          <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold mb-3 text-[#3D3436]">By Day</h3>
          <div className="divide-y divide-[#F8F3EF]">
            {byDay.map(([date, amt]) => (
              <div key={date} className="flex items-center justify-between py-1.5 text-xs" style={{ fontFamily: F.mono }}>
                <span className="text-[#5C5250]">{fmtDate(date)}</span>
                <span className="font-semibold text-[#C4685B]">{fmtINR(amt)}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="rounded-xl border border-[#EFE1DA] bg-white overflow-hidden">
        <h3 style={{ fontFamily: F.display }} className="text-sm font-semibold p-4 pb-0 text-[#3D3436]">Transactions</h3>
        {filtered.length === 0 ? (
          <p className="text-xs text-[#A79A95] p-4">Is filter ke liye koi expense nahi mila.</p>
        ) : (
          <div className="overflow-x-auto mt-2">
            <table className="w-full text-xs">
              <thead><tr className="text-left text-[#A79A95] border-b border-[#F1ECE9]">
                <th className="px-4 py-2 font-medium">Date</th><th className="px-4 py-2 font-medium">Category</th>
                <th className="px-4 py-2 font-medium">Paid By</th><th className="px-4 py-2 font-medium">Mode</th>
                <th className="px-4 py-2 font-medium">Description</th><th className="px-4 py-2 font-medium text-right">Amount</th>
              </tr></thead>
              <tbody>
                {filtered.map((e) => (
                  <tr key={e.id} className="border-b border-[#F8F3EF] last:border-0">
                    <td style={{ fontFamily: F.mono }} className="px-4 py-2 text-[#5C5250] whitespace-nowrap">{fmtDate(e.date)}</td>
                    <td className="px-4 py-2 text-[#3D3436]">{e.category}</td>
                    <td className="px-4 py-2 text-[#3D3436] whitespace-nowrap">{e.paidBy || '—'}</td>
                    <td className="px-4 py-2 text-[#A79A95] whitespace-nowrap">{e.paymentMode || '—'}</td>
                    <td className="px-4 py-2 text-[#A79A95] max-w-[160px] truncate">{e.description || '—'}</td>
                    <td style={{ fontFamily: F.mono }} className="px-4 py-2 text-right font-medium text-[#C4685B]">{fmtINR(e.amount)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

/* ======================================================================
   CONFETTI & COMMAND PALETTE
   ====================================================================== */

function ConfettiBurst({ fire }) {
  const canvasRef = useRef(null);
  useEffect(() => {
    if (!fire) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;
    const ctx = canvas.getContext('2d');
    const colors = ['#C4677A', '#8F7CB8', '#5C8A5E', '#E3B166', '#7CA6C2'];
    const particles = Array.from({ length: 120 }, () => ({
      x: canvas.width / 2, y: canvas.height / 3,
      vx: (Math.random() - 0.5) * 12, vy: Math.random() * -12 - 4,
      size: Math.random() * 6 + 4, color: colors[Math.floor(Math.random() * colors.length)],
      rotation: Math.random() * 360, rotSpeed: (Math.random() - 0.5) * 12,
    }));
    let frame = 0, raf;
    const animate = () => {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      particles.forEach((p) => {
        p.vy += 0.3; p.x += p.vx; p.y += p.vy; p.rotation += p.rotSpeed;
        ctx.save(); ctx.translate(p.x, p.y); ctx.rotate((p.rotation * Math.PI) / 180);
        ctx.fillStyle = p.color; ctx.fillRect(-p.size / 2, -p.size / 2, p.size, p.size);
        ctx.restore();
      });
      frame++;
      if (frame < 90) raf = requestAnimationFrame(animate);
      else ctx.clearRect(0, 0, canvas.width, canvas.height);
    };
    animate();
    return () => cancelAnimationFrame(raf);
  }, [fire]);
  if (!fire) return null;
  return <canvas ref={canvasRef} className="fixed inset-0 pointer-events-none z-[200]" />;
}

function CommandPalette({ F, onClose, actions }) {
  const [query, setQuery] = useState('');
  const filtered = actions.filter((a) => a.label.toLowerCase().includes(query.toLowerCase()));
  return (
    <div className="fixed inset-0 z-[100] bg-black/60 backdrop-blur-sm flex items-start justify-center pt-24 px-4" onClick={onClose}>
      <div style={{ fontFamily: F.body, backgroundColor: '#FDFAF7' }} className="w-full max-w-md rounded-2xl shadow-2xl overflow-hidden" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center gap-2 px-4 py-3 border-b border-[#F1ECE9]">
          <Sparkles size={16} className="text-[#8F7CB8]" />
          <input
            autoFocus value={query} onChange={(e) => setQuery(e.target.value)}
            placeholder="Type a command… (e.g. Add expense)"
            className="flex-1 text-sm outline-none bg-transparent text-[#3D3436] placeholder:text-[#C7B9B3]"
          />
          <kbd className="text-[10px] text-[#A79A95] border border-[#E4D5CE] rounded px-1.5 py-0.5">Esc</kbd>
        </div>
        <div className="max-h-72 overflow-y-auto py-1.5">
          {filtered.length === 0 && <p className="text-xs text-[#A79A95] px-4 py-3">Koi command nahi mila.</p>}
          {filtered.map((a, i) => (
            <button
              key={i} onClick={() => { a.run(); onClose(); }}
              className="w-full flex items-center gap-2.5 text-left text-xs px-4 py-2.5 hover:bg-[#FBF3EF] transition-colors text-[#3D3436]"
            >
              <a.icon size={14} className="text-[#A79A95]" /> {a.label}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ======================================================================
   MAIN APP
   ====================================================================== */

export default function MoneyTracker() {
  const F = { display: "'Space Grotesk', sans-serif", body: "'Inter', sans-serif", mono: "'JetBrains Mono', monospace" };
  const [activeTab, setActiveTab] = useState('emi');
  const [selectedMonth, setSelectedMonth] = useState(currentYM());
  const [loaded, setLoaded] = useState(false);
  const [storageOK, setStorageOK] = useState(true);

  const [loans, setLoans] = useState([]);
  const [payments, setPayments] = useState([]);
  const [showAddLoan, setShowAddLoan] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showPalette, setShowPalette] = useState(false);
  const [showQuickAdd, setShowQuickAdd] = useState(false);
  const [celebrate, setCelebrate] = useState(false);
  const touchStartX = useRef(null);
  const touchStartY = useRef(null);

  const [budgets, setBudgets] = useState({});
  const [expenses, setExpenses] = useState([]);

  const [contacts, setContacts] = useState([]);
  const [khataTx, setKhataTx] = useState([]);

  const [income, setIncome] = useState('');
  const [recurringTemplates, setRecurringTemplates] = useState([]);
  const [goals, setGoals] = useState([]);
  const [theme, setTheme] = useState('light');
  const [currencyCode, setCurrencyCodeState] = useState('INR');
  const [pinLock, setPinLock] = useState({ enabled: false, pin: '' });
  const [unlocked, setUnlocked] = useState(false);

  useEffect(() => {
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = 'https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500;600;700&display=swap';
    document.head.appendChild(link);
  }, []);

  useEffect(() => {
    const handler = (e) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        setShowPalette((v) => !v);
      }
      if (e.key === 'Escape') { setShowPalette(false); setShowQuickAdd(false); }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, []);

  useEffect(() => {
    (async () => {
      if (!window.storage) { setStorageOK(false); setLoaded(true); return; }
      const keys = [
        ['emi-loans', setLoans], ['emi-payments', setPayments],
        ['budgets', setBudgets], ['expenses', setExpenses],
        ['khata-contacts', setContacts], ['khata-transactions', setKhataTx],
        ['recurring-templates', setRecurringTemplates], ['savings-goals', setGoals],
      ];
      for (const [key, setter] of keys) {
        try {
          const r = await window.storage.get(key, false);
          if (r && r.value) setter(JSON.parse(r.value));
        } catch (e) { /* first run, no data yet */ }
      }
      try {
        const r = await window.storage.get('monthly-income', false);
        if (r && r.value) setIncome(r.value);
      } catch (e) { /* first run, no data yet */ }
      try {
        const r = await window.storage.get('app-theme', false);
        if (r && r.value) setTheme(r.value);
      } catch (e) { /* first run, no data yet */ }
      try {
        const r = await window.storage.get('app-currency', false);
        if (r && r.value) { setCurrencyCodeState(r.value); setCurrencyCode(r.value); }
      } catch (e) { /* first run, no data yet */ }
      try {
        const r = await window.storage.get('pin-lock', false);
        if (r && r.value) setPinLock(JSON.parse(r.value));
      } catch (e) { /* first run, no data yet */ }
      setLoaded(true);
    })();
  }, []);

  const persist = useCallback((key, setter) => async (next) => {
    setter(next);
    if (!window.storage) return;
    try { await window.storage.set(key, JSON.stringify(next), false); } catch (e) { /* non-fatal */ }
  }, []);
  const persistLoans = persist('emi-loans', setLoans);
  const persistPayments = persist('emi-payments', setPayments);
  const persistBudgets = persist('budgets', setBudgets);
  const persistExpenses = persist('expenses', setExpenses);
  const persistContacts = persist('khata-contacts', setContacts);
  const persistKhataTx = persist('khata-transactions', setKhataTx);
  const persistRecurring = persist('recurring-templates', setRecurringTemplates);
  const persistGoals = persist('savings-goals', setGoals);

  // EMI actions
  const addLoan = (loan) => { persistLoans([...loans, loan]); setShowAddLoan(false); };
  const deleteLoan = (id) => {
    if (!confirm('Ye loan aur uske saare payment records delete karne hain?')) return;
    persistLoans(loans.filter((l) => l.id !== id));
    persistPayments(payments.filter((p) => p.loanId !== id));
  };
  const recordPayment = (loanId, amount) => {
    const loan = loans.find((l) => l.id === loanId);
    const paidBefore = payments.filter((p) => p.loanId === loanId).reduce((s, p) => s + p.amount, 0);
    persistPayments([...payments, { id: uid(), loanId, month: selectedMonth, amount, date: todayISO() }]);
    if (loan && paidBefore < loan.totalAmount && paidBefore + amount >= loan.totalAmount) {
      setCelebrate(true);
      setTimeout(() => setCelebrate(false), 2500);
    }
  };
  const deletePayment = (id) => persistPayments(payments.filter((p) => p.id !== id));

  const loadExampleLoan = () => {
    const loanId = uid();
    const start = addMonths(currentYM(), -3);
    const exampleLoan = {
      id: loanId, name: 'Home Loan - SBI (Example)', category: 'Home Loan', lender: 'SBI',
      totalAmount: 500000, emiAmount: 12000, startMonth: start, tenureMonths: 36, dueDay: 5,
    };
    const examplePayments = [
      { id: uid(), loanId, month: addMonths(currentYM(), -3), amount: 12000, date: `${addMonths(currentYM(), -3)}-05` }, // fully paid → "Paid"
      { id: uid(), loanId, month: addMonths(currentYM(), -2), amount: 12000, date: `${addMonths(currentYM(), -2)}-05` }, // fully paid → "Paid"
      { id: uid(), loanId, month: addMonths(currentYM(), -1), amount: 12000, date: `${addMonths(currentYM(), -1)}-05` }, // fully paid → "Paid"
      { id: uid(), loanId, month: currentYM(), amount: 6000, date: `${currentYM()}-05` }, // half paid this month → "Partial"
    ];
    persistLoans([...loans, exampleLoan]);
    persistPayments([...payments, ...examplePayments]);
    setSelectedMonth(currentYM());
  };

  const emiPaidThisMonth = useMemo(() => payments.filter((p) => p.month === selectedMonth).reduce((s, p) => s + p.amount, 0), [payments, selectedMonth]);

  // Budget actions
  const setBudgetForMonth = (month, catMap) => persistBudgets({ ...budgets, [month]: catMap });
  const addExpense = (exp) => persistExpenses([...expenses, exp]);
  const updateExpense = (exp) => persistExpenses(expenses.map((e) => (e.id === exp.id ? exp : e)));
  const importExpenses = (rows) => persistExpenses([...expenses, ...rows]);
  const deleteExpense = (id) => persistExpenses(expenses.filter((e) => e.id !== id));

  // Recurring expense templates
  const addRecurringTemplate = (t) => persistRecurring([...recurringTemplates, t]);
  const deleteRecurringTemplate = (id) => persistRecurring(recurringTemplates.filter((t) => t.id !== id));

  // Auto-apply recurring templates for the current month (once per template per month)
  useEffect(() => {
    if (!loaded || recurringTemplates.length === 0) return;
    const ym = currentYM();
    const alreadyApplied = new Set(
      expenses.filter((e) => e.date.slice(0, 7) === ym && e.fromRecurringId).map((e) => e.fromRecurringId)
    );
    const toAdd = recurringTemplates
      .filter((t) => !alreadyApplied.has(t.id))
      .map((t) => ({
        id: uid(), category: t.category, amount: t.amount, paymentMode: t.paymentMode, paidBy: t.paidBy || '',
        description: t.description, date: `${ym}-${String(t.dayOfMonth).padStart(2, '0')}`, fromRecurringId: t.id,
      }));
    if (toAdd.length > 0) persistExpenses([...expenses, ...toAdd]);
    // eslint-disable-next-line
  }, [loaded, recurringTemplates.length]);

  // Income (used for Passbook savings insights)
  const updateIncome = async (val) => {
    setIncome(val);
    if (!window.storage) return;
    try { await window.storage.set('monthly-income', val, false); } catch (e) { /* non-fatal */ }
  };

  // Khata actions
  const addContact = (c) => persistContacts([...contacts, c]);
  const deleteContact = (id) => {
    if (!confirm('Delete this contact and all their transactions?')) return;
    persistContacts(contacts.filter((c) => c.id !== id));
    persistKhataTx(khataTx.filter((t) => t.contactId !== id));
  };
  const addKhataTx = (contactId, tx) => persistKhataTx([...khataTx, { ...tx, contactId }]);
  const deleteKhataTx = (id) => persistKhataTx(khataTx.filter((t) => t.id !== id));
  const setContactDueDate = (id, date) => persistContacts(contacts.map((c) => (c.id === id ? { ...c, dueDate: date } : c)));

  // Goals
  const addGoal = (g) => persistGoals([...goals, g]);
  const deleteGoal = (id) => persistGoals(goals.filter((g) => g.id !== id));
  const addGoalSaved = (id, amt) => {
    const goal = goals.find((g) => g.id === id);
    if (goal && goal.savedAmount < goal.targetAmount && goal.savedAmount + amt >= goal.targetAmount) {
      setCelebrate(true);
      setTimeout(() => setCelebrate(false), 2500);
    }
    persistGoals(goals.map((g) => (g.id === id ? { ...g, savedAmount: g.savedAmount + amt } : g)));
  };

  // Theme
  const toggleTheme = async () => {
    const next = theme === 'light' ? 'dark' : 'light';
    setTheme(next);
    if (!window.storage) return;
    try { await window.storage.set('app-theme', next, false); } catch (e) { /* non-fatal */ }
  };

  // Currency
  const changeCurrency = async (code) => {
    setCurrencyCode(code);
    setCurrencyCodeState(code);
    if (!window.storage) return;
    try { await window.storage.set('app-currency', code, false); } catch (e) { /* non-fatal */ }
  };

  // PIN lock
  const savePinLock = async (next) => {
    setPinLock(next);
    if (next.enabled) setUnlocked(false);
    if (!window.storage) return;
    try { await window.storage.set('pin-lock', JSON.stringify(next), false); } catch (e) { /* non-fatal */ }
  };

  // JSON backup / restore
  const exportJSON = () => {
    const data = { loans, payments, budgets, expenses, contacts, khataTx, recurringTemplates, goals, income, exportedAt: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `money-tracker-backup-${todayISO()}.json`; a.click();
    URL.revokeObjectURL(url);
  };
  const importJSON = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(String(e.target.result || '{}'));
        if (!confirm('Ye current data ko backup file ke data se replace kar dega. Continue?')) return;
        if (data.loans) persistLoans(data.loans);
        if (data.payments) persistPayments(data.payments);
        if (data.budgets) persistBudgets(data.budgets);
        if (data.expenses) persistExpenses(data.expenses);
        if (data.contacts) persistContacts(data.contacts);
        if (data.khataTx) persistKhataTx(data.khataTx);
        if (data.recurringTemplates) persistRecurring(data.recurringTemplates);
        if (data.goals) persistGoals(data.goals);
        if (data.income !== undefined) updateIncome(data.income);
        alert('Backup restore ho gaya!');
      } catch (err) { alert('Ye file valid backup nahi hai.'); }
    };
    reader.readAsText(file);
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(loans.map((l) => ({
      Name: l.name, Category: l.category, Lender: l.lender, 'Total Amount': l.totalAmount, 'Monthly EMI': l.emiAmount,
      'Start Month': l.startMonth, 'Tenure (months)': l.tenureMonths, 'Due Day': l.dueDay,
    }))), 'Loans');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(payments.map((p) => {
      const l = loans.find((x) => x.id === p.loanId);
      return { Loan: l ? l.name : p.loanId, Month: p.month, Amount: p.amount, Date: p.date };
    })), 'EMI Payments');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(expenses.map((e) => ({
      Date: e.date, Category: e.category, Amount: e.amount, 'Payment Mode': e.paymentMode || '', 'Paid By': e.paidBy || '', Description: e.description,
    }))), 'Expenses');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(khataTx.map((t) => {
      const c = contacts.find((x) => x.id === t.contactId);
      return { Contact: c ? c.name : t.contactId, Type: t.type === 'gave' ? 'You Gave' : 'You Got', Amount: t.amount, Date: t.date, Note: t.note };
    })), 'Khata Book');
    XLSX.writeFile(wb, `money-tracker-${selectedMonth}.xlsx`);
  };

  const resetAllData = () => {
    const first = confirm('⚠ Ye SAARA data permanently erase kar dega — loans, EMI payments, budgets, expenses, Khata Book, income, recurring templates aur goals. Ye undo nahi ho sakta. Continue karna hai?');
    if (!first) return;
    const second = confirm('Pakka? Ek baar erase hone ke baad koi bhi purana data wapas nahi milega.');
    if (!second) return;
    persistLoans([]);
    persistPayments([]);
    persistBudgets({});
    persistExpenses([]);
    persistContacts([]);
    persistKhataTx([]);
    persistRecurring([]);
    persistGoals([]);
    updateIncome('');
    setSelectedMonth(currentYM());
  };

  if (!loaded) {
    return (
      <div style={{ fontFamily: F.body, backgroundColor: '#FAF5F0' }} className="min-h-screen flex items-center justify-center">
        <p className="text-sm text-[#A79A95]">Loading…</p>
      </div>
    );
  }

  if (pinLock.enabled && !unlocked) {
    return <PinLockScreen F={F} pin={pinLock.pin} onUnlock={() => setUnlocked(true)} />;
  }

  const TABS = [
    { id: 'emi', label: 'EMI Tracker', icon: Landmark },
    { id: 'budget', label: 'Budget vs Expense', icon: PieChart },
    { id: 'khata', label: 'Khata Book', icon: BookOpen },
    { id: 'passbook', label: 'Passbook', icon: Sparkles },
    { id: 'explorer', label: 'Explorer', icon: Filter },
  ];

  const handleTouchStart = (e) => {
    const el = e.target.closest && e.target.closest('input, select, textarea, button, a, canvas, [role="slider"], .fixed, table');
    if (el) { touchStartX.current = null; touchStartY.current = null; return; }
    touchStartX.current = e.touches[0].clientX;
    touchStartY.current = e.touches[0].clientY;
  };
  const handleTouchEnd = (e) => {
    if (touchStartX.current === null || touchStartY.current === null) return;
    const dx = e.changedTouches[0].clientX - touchStartX.current;
    const dy = e.changedTouches[0].clientY - touchStartY.current;
    touchStartX.current = null;
    touchStartY.current = null;
    // require a confident, mostly-horizontal swipe so vertical scrolling never triggers a tab change
    if (Math.abs(dx) < 80 || Math.abs(dx) < Math.abs(dy) * 1.8) return;
    const idx = TABS.findIndex((t) => t.id === activeTab);
    if (dx < 0 && idx < TABS.length - 1) setActiveTab(TABS[idx + 1].id); // swipe left → next tab
    if (dx > 0 && idx > 0) setActiveTab(TABS[idx - 1].id); // swipe right → previous tab
  };

  const paletteActions = [
    { label: 'Add Expense', icon: PlusCircle, run: () => { setActiveTab('budget'); setShowQuickAdd(true); } },
    { label: 'Add Loan', icon: Landmark, run: () => { setActiveTab('emi'); setShowAddLoan(true); } },
    { label: 'Go to EMI Tracker', icon: Landmark, run: () => setActiveTab('emi') },
    { label: 'Go to Budget vs Expense', icon: PieChart, run: () => setActiveTab('budget') },
    { label: 'Go to Khata Book', icon: BookOpen, run: () => setActiveTab('khata') },
    { label: 'Go to Passbook', icon: Sparkles, run: () => setActiveTab('passbook') },
    { label: 'Go to Explorer', icon: Filter, run: () => setActiveTab('explorer') },
    { label: 'Export to Excel', icon: Download, run: exportExcel },
    { label: 'Download PDF Report', icon: FileDown, run: () => window.print() },
    { label: 'Open Settings', icon: Settings, run: () => setShowSettings(true) },
    { label: 'Erase All Data', icon: Trash2, run: resetAllData },
  ];

  return (
    <div style={theme === 'dark' ? { filter: 'invert(1) hue-rotate(180deg)' } : undefined}>
    <div style={{ fontFamily: F.body, backgroundColor: '#FAF5F0' }} className="min-h-screen pb-10" onTouchStart={handleTouchStart} onTouchEnd={handleTouchEnd}>
      <div className="max-w-3xl mx-auto px-4 pt-6 space-y-5">

        <div className="flex items-start justify-between gap-3 flex-wrap">
          <div>
            <h1 style={{ fontFamily: F.display, letterSpacing: '-0.02em' }} className="text-lg sm:text-xl font-bold text-[#3D3436]">
              MONEY<span className="text-[#C4899C]"> TRACKER</span>
            </h1>
            <p className="text-xs text-[#A79A95] mt-0.5">EMI · Budget · Khata Book</p>
          </div>
          <div className="flex items-center gap-2 print:hidden">
            <button onClick={() => setShowPalette(true)} className="hidden sm:flex items-center gap-1.5 text-xs px-3 py-2 rounded-lg border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-colors">
              <Sparkles size={14} /> <kbd className="text-[10px] border border-[#E4D5CE] rounded px-1">⌘K</kbd>
            </button>
            <button onClick={() => setShowSettings(true)} className="flex items-center gap-1.5 text-xs px-3 py-2 rounded-lg border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-colors">
              <Settings size={14} />
            </button>
            <button onClick={exportJSON} className="flex items-center gap-1.5 text-xs px-3 py-2 rounded-lg border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-colors">
              <FileDown size={14} /> Save All Data
            </button>
            <button onClick={resetAllData} className="flex items-center gap-1.5 text-xs px-3 py-2 rounded-lg border border-[#F0C6BE] text-[#C4685B] hover:bg-[#FBEAE7] transition-colors">
              <Trash2 size={14} /> Erase All
            </button>
            <button onClick={exportExcel} className="flex items-center gap-1.5 text-xs px-3 py-2 rounded-lg border border-[#E4D5CE] text-[#7A6B67] hover:bg-[#F5EBE5] transition-colors">
              <Download size={14} /> Export
            </button>
          </div>
        </div>

        {!storageOK && (
          <p className="text-[11px] text-[#B9862E] bg-[#FBF0DE] px-3 py-2 rounded-lg">⚠ Storage available nahi hai — data sirf is session tak rahega.</p>
        )}

        <style>{`.tab-scroll::-webkit-scrollbar { display: none; }`}</style>
        <div className="tab-scroll flex gap-2 overflow-x-auto pb-1" style={{ scrollbarWidth: 'none', msOverflowStyle: 'none' }}>
          {TABS.map((t) => {
            const Icon = t.icon;
            const active = activeTab === t.id;
            return (
              <button
                key={t.id} type="button" onClick={() => setActiveTab(t.id)}
                className={`flex items-center gap-1.5 whitespace-nowrap shrink-0 px-3.5 py-1.5 rounded-full text-sm font-semibold transition-all border ${active ? 'border-transparent shadow-md' : 'bg-white border-[#EFE1DA] text-[#6B5F5B] hover:text-[#3D3436] hover:bg-[#FBF3EF]'}`}
                style={active ? { background: 'linear-gradient(to right, #C4677A, #8F7CB8)', color: '#FFFFFF' } : { color: '#6B5F5B' }}
              >
                <Icon size={14} /> {t.label}
              </button>
            );
          })}
        </div>

        {activeTab !== 'khata' && activeTab !== 'passbook' && activeTab !== 'explorer' && <MonthNav F={F} selectedMonth={selectedMonth} setSelectedMonth={setSelectedMonth} />}

        {activeTab === 'emi' && (
          <EMITab
            F={F} loans={loans} payments={payments} selectedMonth={selectedMonth}
            addLoan={addLoan} deleteLoan={deleteLoan} recordPayment={recordPayment} deletePayment={deletePayment}
            showAddLoan={showAddLoan} setShowAddLoan={setShowAddLoan} loadExampleLoan={loadExampleLoan}
          />
        )}
        {activeTab === 'budget' && (
          <BudgetTab
            F={F} budgets={budgets} setBudgetForMonth={setBudgetForMonth} expenses={expenses}
            addExpense={addExpense} updateExpense={updateExpense} importExpenses={importExpenses} deleteExpense={deleteExpense}
            selectedMonth={selectedMonth} setSelectedMonth={setSelectedMonth} emiPaidThisMonth={emiPaidThisMonth}
            recurringTemplates={recurringTemplates} addRecurringTemplate={addRecurringTemplate} deleteRecurringTemplate={deleteRecurringTemplate}
          />
        )}
        {activeTab === 'khata' && (
          <KhataTab F={F} contacts={contacts} transactions={khataTx} addContact={addContact} deleteContact={deleteContact} addTx={addKhataTx} deleteTx={deleteKhataTx} setContactDueDate={setContactDueDate} />
        )}
        {activeTab === 'passbook' && (
          <PassbookTab
            F={F} loans={loans} payments={payments} budgets={budgets} expenses={expenses}
            contacts={contacts} khataTx={khataTx} income={income} setIncome={updateIncome}
            goals={goals} addGoal={addGoal} deleteGoal={deleteGoal} addGoalSaved={addGoalSaved}
          />
        )}
        {activeTab === 'explorer' && <ExplorerTab F={F} expenses={expenses} />}
      </div>
    </div>

    <button
      onClick={() => setShowQuickAdd(true)}
      className="print:hidden fixed bottom-6 right-6 z-40 w-13 h-13 rounded-full flex items-center justify-center text-white shadow-2xl hover:scale-105 active:scale-95 transition-transform"
      style={{ background: 'linear-gradient(135deg, #C4677A, #8F7CB8)', width: 52, height: 52 }}
      title="Quick Add Expense"
    >
      <PlusCircle size={22} />
    </button>

    {showSettings && (
      <SettingsModal
        F={F} theme={theme} toggleTheme={toggleTheme} currencyCode={currencyCode} changeCurrency={changeCurrency}
        pinLock={pinLock} savePinLock={savePinLock} exportJSON={exportJSON} importJSON={importJSON}
        onClose={() => setShowSettings(false)}
      />
    )}
    {showQuickAdd && (
      <ExpenseModal F={F} onClose={() => setShowQuickAdd(false)} onSave={(exp, recurring) => { addExpense(exp); if (recurring) addRecurringTemplate({ id: uid(), category: exp.category, amount: exp.amount, paymentMode: exp.paymentMode, paidBy: exp.paidBy, description: exp.description, dayOfMonth: parseInt(exp.date.slice(8, 10), 10) }); setShowQuickAdd(false); }} />
    )}
    {showPalette && <CommandPalette F={F} actions={paletteActions} onClose={() => setShowPalette(false)} />}
    <ConfettiBurst fire={celebrate} />
    </div>
  );
}
