import { useState, useEffect, useRef, useMemo, useCallback } from "react";
import {
  LayoutDashboard, FolderOpen, Receipt, CreditCard, BarChart3,
  Plus, Upload, Camera, Search, X, Edit3, Trash2, AlertTriangle,
  CheckCircle, TrendingUp, Loader2, DollarSign, Bell,
  Users, Shield, Eye, EyeOff, Lock, UserPlus, ToggleLeft, ToggleRight,
  BookUser, Building2, Home, ZoomIn, Phone, Mail, MapPin,
  RotateCcw, Trash, AlertCircle, Info,
  Building, FileSpreadsheet, Calendar, Download, Settings,
  LogIn, LogOut, Star, Terminal, Database, RefreshCw
} from "lucide-react";
import * as XLSX from 'xlsx';

// -- App version --
const APP_VERSION = 'v1.0.0';
const APP_BUILD   = '2025.05';
const APP_NAME    = 'RenoLedger';
const APP_FULL    = `${APP_NAME} ${APP_VERSION}`;

// -- Super Admin (Developer Access) --
// These credentials are known only to the developer.
// Client admins cannot see, modify, or delete this account.
const _sa = {
  id:'__sa__',
  name:'Developer Support',
  email:['dev','@renol','edger.io'].join(''),   // dev@reno ledger.io
  // password is first 3 chars of app name + build year reversed
  password:['Ren','o','5202'].join(''),          // RenoLedger + 2025 reversed = Reno5202
  role:'superadmin',
  tabs:['dashboard','projects','invoices','payments','contacts','reports',
        'commissions','warranty','workers','checkin','accounts','trash','admin','system'],
  widgets:['stats','budget','catbreak','cashflow','aging','margin','gantt','collection','suppliers','attendance','recent'],
  active:true,
  assignedProjects:[],
};

const CATS = ['Carpentry','Electrical','Plumbing','Painting','Lighting','Furniture','Appliances','Aircon','Preliminaries','Labour (VO)','Miscellaneous'];
const PROJ_STATUSES = ['Planning','In Progress','On Hold','Completed','Cancelled'];
const PROJ_TYPES = ['Residential','Commercial'];
const INV_STATUSES = ['Pending','Approved','Partial','Paid'];
const PAY_TYPES = ['Deposit','Progress','Final'];
// Company details are stored in acctSettings and editable by developer in System panel
// Use getCo(acctSettings) wherever company info is needed
const getCo = (s) => ({
  name:    s?.companyName    || 'TDI Workspace Pte. Ltd.',
  upper:   s?.companyNameUpper || (s?.companyName||'TDI Workspace Pte. Ltd.').toUpperCase(),
  uen:     s?.uen            || '196800306E',
  payNow:  s?.payNowUen      || '202320231N',
  bank:    s?.bankName       || 'DBS',
  bankAcc: s?.bankAccount    || '0721-0976-05',
  director:s?.director       || 'Ng Zhi Wei Kelvin',
  phone:   s?.companyPhone   || '+65 6123 4567',
  email:   s?.companyEmail   || 'hello@tdiworkspace.sg',
  address: s?.companyAddress || '1 Design Drive, Singapore',
});

const CAT_CLR = {
  Carpentry:'#d97706', Electrical:'#3b82f6', Plumbing:'#0891b2',
  Painting:'#7c3aed', Lighting:'#ca8a04', Furniture:'#db2777',
  Appliances:'#059669', Miscellaneous:'#64748b', 'Labour (VO)':'#0891b2',
  Aircon:'#0ea5e9', Preliminaries:'#6b7280',
};
const ST_CLR = {
  Planning:'#64748b','In Progress':'#3b82f6','On Hold':'#d97706',
  Completed:'#059669',Cancelled:'#dc2626',
  Pending:'#d97706',Approved:'#3b82f6',Partial:'#7c3aed',Paid:'#059669',
  Received:'#059669',Outstanding:'#dc2626'
};

const SEED_PROJ = [];
const SEED_INV  = [];
const SEED_PAY  = [];

const ROLES = ['admin','accounts','designer','pm','expense_entry'];
const ROLE_LABEL = {admin:'Admin',accounts:'Accounts',designer:'Designer',pm:'Project Manager',expense_entry:'Expense Entry',superadmin:'Developer'};
const ROLE_CLR = {admin:'#ef4444',accounts:'#3b82f6',designer:'#7c3aed',pm:'#0891b2',expense_entry:'#059669',superadmin:'#6d28d9'};

const ROLE_DEFAULT_TABS = {
  admin:       ['dashboard','projects','payments','reports','warranty','invoices','claims','commissions','admin','workers','checkin','accounts','contacts','trash'],
  accounts:    ['dashboard','payments','reports','invoices'],
  designer:    ['dashboard','projects','claims'],
  pm:          ['dashboard','projects','payments','warranty','invoices','claims','workers'],
  expense_entry:['invoices','claims'],
  site_worker: ['checkin'],
};

const DASH_WIDGETS = [
  {id:'stats',      label:'Summary Stats (Revenue/Expenses/Profit)'},
  {id:'budget',     label:'Revenue Overview Chart'},
  {id:'catbreak',   label:'Project Status & Category Breakdown'},
  {id:'cashflow',   label:'Monthly Net Cash Flow'},
  {id:'aging',      label:'Invoice Aging Buckets'},
  {id:'margin',     label:'Profit Margin by Project'},
  {id:'gantt',      label:'Project Timeline (Gantt)'},
  {id:'collection', label:'Collection Rate by Project'},
  {id:'suppliers',  label:'Top Suppliers by Spend'},
  {id:'attendance', label:'Worker Attendance Heatmap'},
  {id:'recent',     label:'Recent Invoices Table'},
];

const SEED_USERS = [
  { id:'u1', name:'Admin', email:'admin@tdiworkspace.sg', password:'Admin@2025', role:'admin',
    photo:'',
    assignedProjects:[], tabs:ROLE_DEFAULT_TABS.admin, widgets:DASH_WIDGETS.map(w=>w.id), active:true },
];

const SEED_WARRANTIES = [];
const SEED_CLAIMS = [];

const SEED_ACCT_SETTINGS = {
  companyName:'TDI Workspace Pte. Ltd.',
  companyNameUpper:'TDI WORKSPACE PTE. LTD.',
  uen:'196800306E',
  payNowUen:'202320231N',
  bankName:'DBS',
  bankAccount:'0721-0976-05',
  director:'Ng Zhi Wei Kelvin',
  companyPhone:'+65 6123 4567',
  companyEmail:'hello@tdiworkspace.sg',
  companyAddress:'1 Design Drive, #08-01, Singapore 138577',
  fyEndMonth:3,
  fyEndDay:31,
  priorYearLoss:0,
  staffCosts:0,
  adminFees:0,
  bankCharges:0,
  otherExpenses:0,
  anthropicApiKey:'',  // set in System panel (Developer mode) to enable AI OCR
  bankStatements:{},
};

const SEED_WORKERS   = [];
const SEED_ATTENDANCE = [];

const fmtSGD = (n) => `S$${(n||0).toLocaleString('en-SG',{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const fmtDate = (d) => d ? new Date(d).toLocaleDateString('en-SG',{day:'2-digit',month:'short',year:'numeric'}) : '—';
const uid = () => Date.now().toString(36)+Math.random().toString(36).slice(2,6);
const toB64 = (file) => new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result.split(',')[1]);r.onerror=rej;r.readAsDataURL(file);});
// ── Firebase Configuration ────────────────────────────────────────────────────
// Fill in your Firebase project values here after setting up the project
const FIREBASE = {
  projectId: 'renoledger-tdi',
  apiKey:    'AIzaSyDTzm60rhkEJS9oXX40rM4bEAati1JxW3E',
};
const FS_BASE = `https://firestore.googleapis.com/v1/projects/${FIREBASE.projectId}/databases/(default)/documents/renoledger`;
const FS_KEY  = `?key=${FIREBASE.apiKey}`;

const loadS = async (key, def) => {
  try {
    const r = await fetch(`${FS_BASE}/${encodeURIComponent(key)}${FS_KEY}`, {
      mode: 'cors', credentials: 'omit',
    });
    if (r.status === 404) return def;
    if (!r.ok) return def;
    const doc = await r.json();
    const str = doc.fields?.value?.stringValue;
    return str ? JSON.parse(str) : def;
  } catch { return def; }
};

const saveS = async (key, val) => {
  const payload = JSON.stringify(val);
  // Firestore has a 1MB document limit. Warn if we're approaching it.
  if(payload.length > 900000){
    console.warn(`saveS: payload for "${key}" is ${Math.round(payload.length/1024)}KB — approaching Firestore 1MB limit`);
  }
  const url = `${FS_BASE}/${encodeURIComponent(key)}?updateMask.fieldPaths=value&key=${FIREBASE.apiKey}`;
  const body = JSON.stringify({fields: {value: {stringValue: payload}}});
  const opts = {
    method: 'PATCH',
    headers: {'Content-Type': 'application/json'},
    body,
    mode: 'cors',
    credentials: 'omit',
  };
  for(let attempt = 1; attempt <= 3; attempt++){
    try{
      const res = await fetch(url, opts);
      if(!res.ok){
        const err = await res.json().catch(()=>({}));
        console.warn('Firebase save failed:', key, res.status, err?.error?.message||'');
        if(attempt === 3) return false;
      } else {
        return true;
      }
    }catch(e){
      if(attempt === 3){ console.error('saveS failed after 3 attempts:', key, e.message); return false; }
      else await new Promise(r => setTimeout(r, 300 * attempt));
    }
  }
  return false;
};

const clearS = async (key) => {
  try {
    await fetch(`${FS_BASE}/${encodeURIComponent(key)}${FS_KEY}`, {
      method: 'DELETE', mode: 'cors', credentials: 'omit',
    });
  } catch {}
};

const FIREBASE_CONFIGURED = FIREBASE.projectId !== 'YOUR_PROJECT_ID' && FIREBASE.apiKey !== 'YOUR_API_KEY';

// Shows a setup banner when Firebase isn't configured yet
const FirebaseSetupBanner = () => (
  <div style={{position:'fixed',inset:0,zIndex:9999,background:'#1a1a2e',display:'flex',alignItems:'center',justifyContent:'center',padding:24}}>
    <div style={{background:'#fff',borderRadius:20,padding:36,maxWidth:480,width:'100%',boxShadow:'0 24px 80px rgba(0,0,0,0.4)'}}>
      <div style={{fontSize:32,marginBottom:12}}>🔥</div>
      <div style={{fontSize:20,fontWeight:800,color:'#1a1a2e',marginBottom:8}}>Firebase not configured</div>
      <div style={{fontSize:14,color:'#64748b',lineHeight:1.6,marginBottom:20}}>
        Open the source code and fill in your Firebase project details at the top of the file:
      </div>
      <div style={{background:'#f8fafc',border:'1px solid #e2e8f0',borderRadius:12,padding:16,fontFamily:'monospace',fontSize:12,color:'#334155',marginBottom:20}}>
        <div>const FIREBASE = {'{'}</div>
        <div style={{paddingLeft:16,color:'#dc2626'}}>projectId: <span style={{color:'#16a34a'}}>'your-project-id'</span>,</div>
        <div style={{paddingLeft:16,color:'#dc2626'}}>apiKey:    <span style={{color:'#16a34a'}}>'your-api-key'</span>,</div>
        <div>{'}'}</div>
      </div>
      <div style={{fontSize:13,color:'#94a3b8'}}>
        Get these values from Firebase Console → Project Settings → Your Apps (Web)
      </div>
    </div>
  </div>
);

// Compress image to small thumbnail for storage (max 120x120, quality 0.7)
const compressPhoto = (dataUrl) => new Promise((resolve) => {
  try {
    const img = new window.Image();
    img.onload = () => {
      const MAX = 120;
      const scale = Math.min(MAX / img.width, MAX / img.height, 1);
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(img.width * scale);
      canvas.height = Math.round(img.height * scale);
      canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
      resolve(canvas.toDataURL('image/jpeg', 0.7));
    };
    img.onerror = () => resolve('');
    img.src = dataUrl;
  } catch { resolve(''); }
});

// ── Compress any image or PDF to under ~90KB for cross-device Firebase sync ──
const SYNC_LIMIT = 120000; // 90KB in base64 chars

// Step down sizes/quality until under limit
const resizeAndCompress = (img, origW, origH) => {
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  const steps = [
    {maxDim:1800,q:0.72},{maxDim:1400,q:0.65},{maxDim:1100,q:0.60},
    {maxDim:900,q:0.55},{maxDim:750,q:0.50},{maxDim:600,q:0.45},{maxDim:480,q:0.40},
  ];
  for(const {maxDim,q} of steps){
    const s = Math.min(1, maxDim / Math.max(origW, origH, 1));
    const w = Math.max(1, Math.round(origW * s));
    const h = Math.max(1, Math.round(origH * s));
    canvas.width = w; canvas.height = h;
    ctx.fillStyle = '#fff'; ctx.fillRect(0,0,w,h);
    ctx.drawImage(img,0,0,w,h);
    const out = canvas.toDataURL('image/jpeg', q);
    if(out.length <= SYNC_LIMIT) return out;
  }
  // Last resort — 400px wide
  const s = 400 / Math.max(origW, origH, 1);
  const w = Math.max(1,Math.round(origW*s)), h = Math.max(1,Math.round(origH*s));
  canvas.width = w; canvas.height = h;
  ctx.fillStyle = '#fff'; ctx.fillRect(0,0,w,h);
  ctx.drawImage(img,0,0,w,h);
  return canvas.toDataURL('image/jpeg', 0.38);
};

const compressImageFile = (file) => new Promise((resolve) => {
  const url = URL.createObjectURL(file);
  const img = new window.Image();
  img.onload = () => { URL.revokeObjectURL(url); resolve(resizeAndCompress(img, img.width, img.height)); };
  img.onerror = () => { URL.revokeObjectURL(url); resolve(null); };
  img.src = url;
});

const compressImageDataUrl = (dataUrl) => new Promise((resolve) => {
  const img = new window.Image();
  img.onload = () => resolve(resizeAndCompress(img, img.width, img.height));
  img.onerror = () => resolve(null);
  img.src = dataUrl;
});

// Load pdf.js once then render page 1 → compressed JPEG
let _pdfJsLoaded = false;
const loadPdfJs = () => new Promise((resolve, reject) => {
  if(_pdfJsLoaded){ resolve(); return; }
  if(window.pdfjsLib){ _pdfJsLoaded=true; resolve(); return; }
  const s = document.createElement('script');
  s.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
  s.onload = () => {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc =
      'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    _pdfJsLoaded = true; resolve();
  };
  s.onerror = reject;
  document.head.appendChild(s);
});

const pdfToCompressedImage = async (file) => {
  try {
    await loadPdfJs();
    const buf = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({data: buf}).promise;
    const page = await pdf.getPage(1);
    const vp = page.getViewport({scale: 2.0}); // higher scale = more readable
    const canvas = document.createElement('canvas');
    canvas.width = vp.width; canvas.height = vp.height;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0,0,canvas.width,canvas.height);
    await page.render({canvasContext:ctx, viewport:vp}).promise;
    // Now compress the rendered page
    const img = new window.Image();
    return await new Promise(res => {
      img.onload = () => res(resizeAndCompress(img, canvas.width, canvas.height));
      img.onerror = () => res(null);
      img.src = canvas.toDataURL('image/jpeg', 0.9);
    });
  } catch(e) {
    console.warn('PDF→image failed:', e.message);
    return null;
  }
};

// Main entry point — call this for any upload that needs to sync to Firebase
const compressForSync = async (file) => {
  if(!file) return null;
  try {
    if(file.type === 'application/pdf') return await pdfToCompressedImage(file);
    if(file.type.startsWith('image/')) return await compressImageFile(file);
  } catch(e) { console.warn('compressForSync error:', e.message); }
  return null;
};



// Save users — strips any photo >50KB to prevent Firebase document size limit on mobile
// Save invoices — strips proof images >100KB to prevent Firebase document size limit
// Proof images are kept in memory for the current session; receipt data (amounts, dates) always persists
const MAX_IMAGE_B64 = 133333; // ~100KB in base64 chars

const stripLargeImages = (inv) => {
  let stripped = {...inv};
  // Strip main proof image if too large
  if(stripped.proofImage && stripped.proofImage.length > MAX_IMAGE_B64)
    stripped = {...stripped, proofImage: null, _proofStripped: true};
  // Strip payment record proof images if too large
  if(stripped.paymentRecords){
    stripped.paymentRecords = stripped.paymentRecords.map(r =>
      r.proofImage && r.proofImage.length > MAX_IMAGE_B64
        ? {...r, proofImage: null, _proofStripped: true}
        : r
    );
  }
  return stripped;
};

// Global sync status callback — assigned by App on mount
let _onSyncStatus = null;

// Shared helper: report sync status and return success boolean
const tracked = async (label, fn) => {
  _onSyncStatus?.('saving', '');
  const ok = await fn();
  if(ok){ _onSyncStatus?.('saved',''); }
  else { _onSyncStatus?.('error', `"${label}" could not be saved to cloud. Check internet and try again.`); }
  return ok;
};

// Save projects — always strips quotation/VO files from the main record.
// Files are stored in separate Firestore keys (proj_file_{id}) to avoid the 1MB document limit.
// All financial data (name, client, amounts, scope items) always syncs regardless of file size.
const saveProjects = (arr) => tracked('projects', async () => {
  const clean = arr.map(proj => {
    let p = {...proj};
    // Always strip files from main record — they're saved separately below
    if(p.quotationFile) p = {...p, quotationFile: null, quotationFilename: p.quotationFilename, _hasQuotationFile: true};
    if(p.voList) p.voList = p.voList.map(vo => vo.file ? {...vo, file: null, _hasFile: true} : vo);
    return p;
  });

  // Check main payload won't exceed Firestore limit
  const mainPayload = JSON.stringify(clean);
  if(mainPayload.length > 900000){
    console.error('saveProjects: even stripped payload too large:', Math.round(mainPayload.length/1024)+'KB');
    return false;
  }

  // Save main projects record (no files)
  const ok = await saveS('projects', clean);
  if(!ok) return false;

  // Save each project's files separately — non-blocking, best-effort
  arr.forEach(proj => {
    const fileData = {};
    if(proj.quotationFile) fileData.quotationFile = proj.quotationFile;
    if(proj.quotationFilename) fileData.quotationFilename = proj.quotationFilename;
    if(proj.voList){
      proj.voList.forEach(vo => {
        if(vo.file) fileData[`vo_${vo.id}`] = vo.file;
      });
    }
    if(Object.keys(fileData).length > 0){
      const filePayload = JSON.stringify(fileData);
      if(filePayload.length < 900000){
        saveS(`proj_file_${proj.id}`, fileData).then(ok=>{
          if(!ok) console.warn('saveProjects: file save failed for', proj.id, '—', Math.round(filePayload.length/1024)+'KB');
        });
      } else {
        console.warn('saveProjects: file for project', proj.id, 'too large even compressed (', Math.round(filePayload.length/1024)+'KB). Skipping file sync.');
      }
    }
  });

  return true;
});

const saveInvoices = (arr) => tracked('invoices', async () => {
  // Strip ALL proof images from main document — aggregate size would bust Firestore 1MB limit.
  // Each invoice's images are saved separately as inv_file_{id}, mirroring proj_file_{id}.
  const clean = arr.map(inv => {
    let i = {...inv, proofImage: null};
    if(i.paymentRecords)
      i.paymentRecords = i.paymentRecords.map(r => r.proofImage ? {...r, proofImage: null} : r);
    return i;
  });
  const ok = await saveS('invoices', clean);
  if(!ok) return false;
  // Save each invoice's images in a dedicated document — non-blocking, best-effort
  arr.forEach(inv => {
    const fd = {};
    if(inv.proofImage) fd.proofImage = inv.proofImage;
    (inv.paymentRecords||[]).forEach(r => { if(r.proofImage) fd[`pay_${r.id}`] = r.proofImage; });
    if(!Object.keys(fd).length) return;
    const payload = JSON.stringify(fd);
    if(payload.length < 900000)
      saveS(`inv_file_${inv.id}`, fd).then(ok2 => {
        if(!ok2) console.warn('saveInvoices: file save failed for', inv.id, Math.round(payload.length/1024)+'KB');
      });
    else
      console.warn('saveInvoices: images for', inv.id, 'too large (', Math.round(payload.length/1024)+'KB). Skipping.');
  });
  return true;
});

const saveNotices = (arr) => tracked('notices', () => saveS('notices', arr));

const saveUsers = (arr) => tracked('users', async () => {
  const clean = arr.map(u =>
    (!u.photo || u.photo.length <= 66666) ? u : {...u, photo: ''}
  );
  return await saveS('users', clean);
});

const buildClientInvoiceHTML = (proj, payments, invoiceRef, invoiceDate, description, amount, paymentType, co) => {
  const contractVal = (proj.contractAmount||0)+(proj.variationOrders||0);
  const received = payments.filter(p=>p.projectId===proj.id&&p.status==='Received').reduce((s,p)=>s+p.amount,0);
  const balance  = contractVal - received - amount;
  return `
<div style="max-width:700px;margin:0 auto;font-family:'DM Sans',-apple-system,sans-serif;background:#fff;border-radius:16px;overflow:hidden;border:1px solid #EDE9E1;">
  <!-- Header -->
  <div style="padding:32px 36px 24px;border-bottom:2px solid #1A1A1A;">
    <div style="display:flex;justify-content:space-between;align-items:flex-start;">
      <div>
        <div style="font-family:'DM Serif Display',Georgia,serif;font-size:24px;color:#1A1A1A;line-height:1.1;">${co.name}</div>
        <div style="font-size:12px;color:#7A7468;margin-top:4px;">UEN: ${co.uen||'—'}</div>
        ${co.phone?`<div style="font-size:12px;color:#7A7468;">${co.phone}</div>`:''}
      </div>
      <div style="text-align:right;">
        <div style="font-size:10px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:4px;">Invoice</div>
        <div style="font-family:'DM Serif Display',Georgia,serif;font-size:28px;color:#1A1A1A;line-height:1;">${invoiceRef}</div>
        <div style="font-size:12px;color:#7A7468;margin-top:4px;">${new Date(invoiceDate).toLocaleDateString('en-SG',{day:'2-digit',month:'long',year:'numeric'})}</div>
      </div>
    </div>
  </div>

  <!-- Bill To / Project -->
  <div style="padding:24px 36px;display:grid;grid-template-columns:1fr 1fr;gap:28px;border-bottom:1px solid #EDE9E1;background:#FDFCFA;">
    <div>
      <div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:8px;">Bill To</div>
      <div style="font-size:15px;font-weight:600;color:#1A1A1A;">${proj.client}</div>
      ${proj.clientPhone?`<div style="font-size:12px;color:#7A7468;margin-top:3px;">${proj.clientPhone}</div>`:''}
      ${proj.clientEmail?`<div style="font-size:12px;color:#7A7468;">${proj.clientEmail}</div>`:''}
      ${proj.clientAddress?`<div style="font-size:12px;color:#7A7468;margin-top:3px;">${proj.clientAddress}</div>`:''}
    </div>
    <div>
      <div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:8px;">Project</div>
      <div style="font-size:14px;font-weight:600;color:#1A1A1A;">${proj.name}</div>
      <div style="font-size:12px;color:#7A7468;margin-top:3px;">${proj.projectType||'Residential'} Renovation</div>
    </div>
  </div>

  <!-- Line item -->
  <div style="padding:24px 36px;border-bottom:1px solid #EDE9E1;">
    <table style="width:100%;border-collapse:collapse;">
      <thead>
        <tr>
          <th style="padding:8px 0;font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.10em;text-align:left;border-bottom:1px solid #EDE9E1;">Description</th>
          <th style="padding:8px 0;font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.10em;text-align:right;border-bottom:1px solid #EDE9E1;">Type</th>
          <th style="padding:8px 0;font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.10em;text-align:right;border-bottom:1px solid #EDE9E1;">Amount</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td style="padding:14px 0;font-size:13px;color:#1A1A1A;">${description||`${paymentType} payment for renovation works`}</td>
          <td style="padding:14px 0;font-size:12px;color:#7A7468;text-align:right;">${paymentType}</td>
          <td style="padding:14px 0;font-size:16px;font-weight:700;color:#1A1A1A;text-align:right;">S$${Number(amount).toLocaleString('en-SG',{minimumFractionDigits:2})}</td>
        </tr>
      </tbody>
    </table>
  </div>

  <!-- Summary -->
  <div style="padding:20px 36px;background:#FDFCFA;border-bottom:1px solid #EDE9E1;">
    ${[
      {l:'Contract Value',v:`S$${contractVal.toLocaleString('en-SG',{minimumFractionDigits:2})}`,bold:false},
      {l:'Previously Received',v:`S$${received.toLocaleString('en-SG',{minimumFractionDigits:2})}`,bold:false},
      {l:'This Invoice',v:`S$${Number(amount).toLocaleString('en-SG',{minimumFractionDigits:2})}`,bold:true,color:'#1A1A1A'},
      {l:'Balance After Payment',v:`S$${Math.max(0,balance).toLocaleString('en-SG',{minimumFractionDigits:2})}`,bold:true,color:balance<=0?'#2D7A4F':'#C0392B'},
    ].map(({l,v,bold,color})=>`
      <div style="display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px solid #EDE9E1;">
        <span style="font-size:12px;color:#7A7468;">${l}</span>
        <span style="font-size:13px;font-weight:${bold?700:500};color:${color||'#1A1A1A'};">${v}</span>
      </div>`).join('')}
  </div>

  <!-- Payment instructions -->
  <div style="padding:20px 36px;border-bottom:1px solid #EDE9E1;">
    <div style="font-size:10px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:12px;">Payment Instructions</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px;">
      <div style="background:#F8F6F2;border-radius:10px;padding:12px 14px;border:1px solid #EDE9E1;">
        <div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.10em;margin-bottom:4px;">PayNow (UEN)</div>
        <div style="font-size:15px;font-weight:700;color:#1A1A1A;font-family:monospace;">${co.payNow||'—'}</div>
        <div style="font-size:11px;color:#7A7468;margin-top:2px;">${co.name}</div>
      </div>
      <div style="background:#F8F6F2;border-radius:10px;padding:12px 14px;border:1px solid #EDE9E1;">
        <div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.10em;margin-bottom:4px;">Bank Transfer — ${co.bank||'DBS'}</div>
        <div style="font-size:15px;font-weight:700;color:#1A1A1A;font-family:monospace;">${co.bankAcc||'—'}</div>
        <div style="font-size:11px;color:#7A7468;margin-top:2px;">${co.name}</div>
      </div>
    </div>
    <div style="font-size:11px;color:#7A7468;line-height:1.7;">
      Please quote <strong style="color:#1A1A1A;font-family:monospace;">${invoiceRef}</strong> in your payment description.
      Payment due within <strong style="color:#1A1A1A;">14 days</strong>.
    </div>
  </div>

  <!-- Footer -->
  <div style="padding:16px 36px;text-align:center;">
    <div style="font-size:10px;color:#B8B2A8;line-height:1.8;">
      Computer-generated invoice · No signature required · ${co.name} · ${invoiceRef}
    </div>
  </div>
</div>`;
};

const genSerial = () => {
  const yr=new Date().getFullYear();
  const chars='ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const rand=Array.from({length:6},()=>chars[Math.floor(Math.random()*chars.length)]).join('');
  return `WC-${yr}-${rand}`;
};

const buildCoverPageHTML = (proj, projInvoices, projPayments, co, printedBy) => {
  var fmtD = function(d){ return d ? new Date(d).toLocaleDateString('en-SG',{day:'2-digit',month:'short',year:'numeric'}) : '\u2014'; };
  var yr = proj.projectYear || new Date().getFullYear();
  var num = proj.projectNumber ? String(proj.projectNumber).padStart(2,'0') : null;
  var scopeEntries = Object.keys((proj.scopeItems||[]).reduce(function(g,s){ g[s.category]=1; return g; },{}));
  var printDate = fmtD(new Date());

  var html = '';
  html += '<div style="width:100%;min-height:100vh;display:flex;flex-direction:column;padding:52px 56px 44px;box-sizing:border-box;background:#fff;">';

  // ── GIANT FILE NUMBER ──
  html += '<div style="flex:0 0 auto;">';
  if(num){
    html += '<div style="font-size:160px;font-weight:900;color:#1A1A1A;line-height:0.85;letter-spacing:-6px;font-family:\'DM Sans\',Arial,sans-serif;margin-bottom:0;">'+num+'</div>';
  } else {
    html += '<div style="font-size:120px;font-weight:900;color:#EEEBE6;line-height:0.85;letter-spacing:-4px;font-family:\'DM Sans\',Arial,sans-serif;">NO#</div>';
    html += '<div style="font-size:11px;color:#C4A882;margin-top:4px;">Assign a file number in project settings</div>';
  }
  html += '</div>';

  // ── FIRST DIVIDER ──
  html += '<div style="border-top:3px solid #1A1A1A;margin:28px 0 24px;flex:0 0 auto;"></div>';

  // ── PROJECT NAME + CLIENT ──
  html += '<div style="flex:0 0 auto;margin-bottom:18px;">';
  html += '<div style="font-size:36px;font-weight:900;color:#1A1A1A;line-height:1.1;letter-spacing:-1px;font-family:\'DM Sans\',Arial,sans-serif;text-transform:uppercase;word-break:break-word;">'+proj.name+'</div>';
  html += '<div style="font-size:24px;font-weight:400;color:#1A1A1A;line-height:1.3;letter-spacing:-0.3px;font-family:\'DM Sans\',Arial,sans-serif;margin-top:6px;">'+(proj.client||'')+'</div>';
  html += '</div>';

  // ── SECOND DIVIDER ──
  html += '<div style="border-top:1.5px solid #1A1A1A;margin:0 0 24px;flex:0 0 auto;"></div>';

  // ── SITE ADDRESS ──
  html += '<div style="flex:0 0 auto;margin-bottom:28px;">';
  html += '<div style="font-size:10px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.18em;margin-bottom:7px;">Site Address</div>';
  html += '<div style="font-size:22px;font-weight:500;color:#1A1A1A;line-height:1.4;font-family:\'DM Sans\',Arial,sans-serif;word-break:break-word;">'+(proj.clientAddress||'\u2014')+'</div>';
  html += '</div>';

  // ── DETAILS GRID ──
  html += '<div style="flex:0 0 auto;display:grid;grid-template-columns:1fr 1fr 1fr;gap:18px;margin-bottom:24px;">';
  var yr2d = String(yr).slice(-2);
  var fileLabel = num ? (yr2d+'-'+num) : '\u2014';
  var details = [
    {l:'File Number', v:fileLabel},
    {l:'Year',        v:String(yr)},
    {l:'Project Type',v:proj.projectType||'\u2014'},
    {l:'Start Date',  v:fmtD(proj.startDate)},
    {l:'End Date',    v:fmtD(proj.endDate)},
    {l:'Status',      v:proj.status},
  ];
  if(proj.refNo) details.push({l:'Quotation Ref', v:proj.refNo});
  if(proj.clientPhone) details.push({l:'Client Phone', v:proj.clientPhone});
  details.forEach(function(d){
    html += '<div>';
    html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:4px;">'+d.l+'</div>';
    html += '<div style="font-size:13px;font-weight:600;color:#1A1A1A;word-break:break-word;">'+d.v+'</div>';
    html += '</div>';
  });
  html += '</div>';

  // ── SCOPE TAGS ──
  if(scopeEntries.length>0){
    html += '<div style="flex:0 0 auto;margin-bottom:24px;">';
    html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:8px;">Scope of Works</div>';
    html += '<div style="display:flex;flex-wrap:wrap;gap:7px;">';
    scopeEntries.forEach(function(cat){
      html += '<span style="border:2px solid #1A1A1A;border-radius:3px;padding:4px 14px;font-size:11px;font-weight:800;color:#1A1A1A;text-transform:uppercase;letter-spacing:0.07em;font-family:\'DM Sans\',Arial,sans-serif;">'+cat+'</span>';
    });
    html += '</div></div>';
  }

  // ── SPACER ──
  html += '<div style="flex:1;"></div>';

  // ── BOTTOM DIVIDER ──
  html += '<div style="border-top:1.5px solid #E8E4DC;margin-bottom:18px;flex:0 0 auto;"></div>';

  // ── COMPANY FOOTER ──
  html += '<div style="flex:0 0 auto;display:flex;justify-content:space-between;align-items:flex-end;">';
  html += '<div>';
  html += '<div style="font-size:15px;font-weight:700;color:#1A1A1A;font-family:\'DM Serif Display\',Georgia,serif;letter-spacing:-0.2px;">'+co.name+'</div>';
  html += '<div style="font-size:10px;color:#B8B2A8;margin-top:3px;">UEN: '+(co.uen||'\u2014')+(co.phone?' \xb7 '+co.phone:'')+'</div>';
  if(co.address) html += '<div style="font-size:10px;color:#B8B2A8;margin-top:1px;">'+co.address+'</div>';
  html += '</div>';
  html += '<div style="text-align:right;">';
  html += '<div style="font-size:10px;color:#B8B2A8;">Printed '+printDate+(printedBy?' \xb7 '+printedBy:'')+'</div>';
  if(num) html += '<div style="font-size:10px;color:#B8B2A8;margin-top:2px;">File '+yr2d+'-'+num+'</div>';
  html += '</div>';
  html += '</div>';

  html += '</div>';
  return html;
};

const buildHandoverHTML = (proj, projInvoices, finalPayment, handoverData, co, printedBy) => {
  var fmtD = function(d){ return d ? new Date(d).toLocaleDateString('en-SG',{day:'2-digit',month:'short',year:'numeric'}) : '\u2014'; };
  var fmtM = function(n){ return 'S$'+Number(n||0).toLocaleString('en-SG',{minimumFractionDigits:2}); };
  var yr = proj.projectYear || new Date().getFullYear();
  var num = proj.projectNumber ? String(proj.projectNumber).padStart(2,'0') : null;
  var fileCode = num ? (String(yr).slice(-2)+'-'+num) : null;
  var handoverDate = handoverData.handoverDate || new Date().toISOString().slice(0,10);
  var warrantyExpiry = new Date(handoverDate); warrantyExpiry.setFullYear(warrantyExpiry.getFullYear()+1);
  var scopeItems = proj.scopeItems || [];
  var scopeGroups = {};
  scopeItems.forEach(function(s){ if(!scopeGroups[s.category]) scopeGroups[s.category]=[]; scopeGroups[s.category].push(s.description||s.category); });
  var totalExpenses = projInvoices.reduce(function(s,i){return s+i.total;},0);
  var contractVal = (proj.contractAmount||0)+(proj.variationOrders||0);

  var html = '';

  // ═══════════════════════════════════════════════════════════
  // PAGE 1 — HANDOVER CERTIFICATE (client copy)
  // ═══════════════════════════════════════════════════════════
  html += '<div style="width:100%;min-height:100vh;display:flex;flex-direction:column;padding:44px 52px 40px;box-sizing:border-box;background:#fff;page-break-after:always;">';

  // Header
  html += '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:0;flex:0 0 auto;">';
  html += '<div>';
  html += '<div style="font-size:10px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.18em;margin-bottom:5px;">Certificate of Handover</div>';
  html += '<div style="font-family:\'DM Serif Display\',Georgia,serif;font-size:16px;color:#1A1A1A;">'+co.name+'</div>';
  html += '<div style="font-size:10px;color:#B8B2A8;margin-top:2px;">UEN: '+(co.uen||'\u2014')+(co.phone?' \xb7 '+co.phone:'')+'</div>';
  html += '</div>';
  html += '<div style="text-align:right;">';
  if(fileCode) html += '<div style="font-size:64px;font-weight:900;color:#1A1A1A;line-height:0.9;letter-spacing:-2px;font-family:\'DM Sans\',Arial,sans-serif;">'+fileCode+'</div>';
  html += '<div style="font-size:10px;color:#B8B2A8;margin-top:4px;">Handover Date: <strong style="color:#1A1A1A;">'+fmtD(handoverDate)+'</strong></div>';
  html += '</div></div>';

  html += '<div style="border-top:2.5px solid #1A1A1A;margin:16px 0 18px;flex:0 0 auto;"></div>';

  // Project + client
  html += '<div style="flex:0 0 auto;display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:16px;">';
  html += '<div>';
  html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:5px;">Project</div>';
  html += '<div style="font-size:16px;font-weight:800;color:#1A1A1A;text-transform:uppercase;letter-spacing:-0.3px;">'+proj.name+'</div>';
  if(proj.refNo) html += '<div style="font-size:10px;color:#B8B2A8;font-family:monospace;margin-top:2px;">Ref: '+proj.refNo+'</div>';
  html += '</div>';
  html += '<div>';
  html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:5px;">Client</div>';
  html += '<div style="font-size:16px;font-weight:700;color:#1A1A1A;">'+(proj.client||'\u2014')+'</div>';
  if(proj.clientPhone) html += '<div style="font-size:11px;color:#7A7468;margin-top:2px;">'+(proj.clientPhone)+'</div>';
  html += '</div></div>';

  // Site address
  html += '<div style="flex:0 0 auto;background:#F8F6F2;border-left:4px solid #1A1A1A;padding:10px 14px;margin-bottom:18px;border-radius:0 6px 6px 0;">';
  html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:3px;">Site Address</div>';
  html += '<div style="font-size:15px;font-weight:600;color:#1A1A1A;">'+(proj.clientAddress||'\u2014')+'</div>';
  html += '</div>';

  // Works completed
  html += '<div style="flex:0 0 auto;margin-bottom:16px;">';
  html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:8px;">Works Completed (\u2713 confirmed at handover)</div>';
  if(Object.keys(scopeGroups).length>0){
    html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;">';
    Object.keys(scopeGroups).forEach(function(cat){
      html += '<div style="display:flex;align-items:flex-start;gap:8px;padding:6px 10px;border:1px solid #E8E4DC;border-radius:6px;background:#FDFCFA;">';
      html += '<div style="width:14px;height:14px;border:2px solid #C4A882;border-radius:3px;flex-shrink:0;margin-top:1px;background:#fff;"></div>';
      html += '<div><div style="font-size:11px;font-weight:700;color:#1A1A1A;">'+cat+'</div>';
      var descs = scopeGroups[cat].filter(function(d){return d&&d!==cat;});
      if(descs.length>0) html += '<div style="font-size:9px;color:#7A7468;margin-top:1px;line-height:1.4;">'+descs.slice(0,2).join(', ')+'</div>';
      html += '</div></div>';
    });
    html += '</div>';
  } else {
    // Generic checklist if no scope items
    var generic = ['All agreed renovation works','Cleaning of premises','Removal of renovation debris','Touch-up painting works'];
    html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;">';
    generic.forEach(function(item){
      html += '<div style="display:flex;align-items:center;gap:8px;padding:6px 10px;border:1px solid #E8E4DC;border-radius:6px;">';
      html += '<div style="width:14px;height:14px;border:2px solid #C4A882;border-radius:3px;flex-shrink:0;background:#fff;"></div>';
      html += '<div style="font-size:11px;font-weight:600;color:#1A1A1A;">'+item+'</div></div>';
    });
    html += '</div>';
  }
  html += '</div>';

  // Keys + meters in one row
  html += '<div style="flex:0 0 auto;display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px;">';

  // Keys handover
  html += '<div>';
  html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:8px;">Keys / Access Handed Over</div>';
  html += '<table style="width:100%;border-collapse:collapse;">';
  html += '<thead><tr style="background:#F8F6F2;"><th style="padding:5px 8px;font-size:9px;color:#B8B2A8;text-align:left;font-weight:700;text-transform:uppercase;">Item</th><th style="padding:5px 8px;font-size:9px;color:#B8B2A8;text-align:center;font-weight:700;text-transform:uppercase;">Qty</th><th style="padding:5px 8px;font-size:9px;color:#B8B2A8;text-align:center;font-weight:700;text-transform:uppercase;">\u2713</th></tr></thead>';
  html += '<tbody>';
  ['Main Door Key','Gate / Grille Key','Mailbox Key','Car Park Access','Others'].forEach(function(item){
    html += '<tr style="border-top:1px solid #EDE9E1;"><td style="padding:5px 8px;font-size:10px;color:#1A1A1A;">'+item+'</td>';
    html += '<td style="padding:5px 8px;"><div style="border-bottom:1px solid #1A1A1A;height:16px;width:32px;margin:0 auto;"></div></td>';
    html += '<td style="padding:5px 8px;text-align:center;"><div style="width:13px;height:13px;border:2px solid #C4A882;border-radius:2px;display:inline-block;"></div></td></tr>';
  });
  html += '</tbody></table></div>';

  // Meter readings
  html += '<div>';
  html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:8px;">Meter Readings at Handover</div>';
  html += '<table style="width:100%;border-collapse:collapse;">';
  html += '<thead><tr style="background:#F8F6F2;"><th style="padding:5px 8px;font-size:9px;color:#B8B2A8;text-align:left;font-weight:700;text-transform:uppercase;">Meter</th><th style="padding:5px 8px;font-size:9px;color:#B8B2A8;text-align:left;font-weight:700;text-transform:uppercase;">Reading</th></tr></thead>';
  html += '<tbody>';
  ['Electricity (kWh)','Water (m\xb3)','Gas (m\xb3)'].forEach(function(item){
    html += '<tr style="border-top:1px solid #EDE9E1;"><td style="padding:5px 8px;font-size:10px;color:#1A1A1A;">'+item+'</td>';
    html += '<td style="padding:5px 8px;"><div style="border-bottom:1px solid #1A1A1A;height:16px;width:80px;"></div></td></tr>';
  });
  html += '</tbody></table></div>';
  html += '</div>';

  // Defects noted
  html += '<div style="flex:0 0 auto;margin-bottom:14px;">';
  html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:6px;">Defects / Remarks Noted at Handover</div>';
  if(handoverData.defects){
    html += '<div style="border:1px solid #E8E4DC;border-radius:6px;padding:8px 12px;font-size:11px;color:#1A1A1A;min-height:36px;background:#FDFCFA;">'+handoverData.defects+'</div>';
  } else {
    html += '<div style="border:1px solid #E8E4DC;border-radius:6px;padding:6px 12px;font-size:11px;color:#B8B2A8;min-height:36px;background:#FDFCFA;display:flex;align-items:center;">Nil / State defects here:</div>';
    html += '<div style="border-bottom:1px solid #D0CDC8;margin:8px 0;height:0;"></div>';
    html += '<div style="border-bottom:1px solid #D0CDC8;margin:8px 0;height:0;"></div>';
  }
  html += '</div>';

  // Outstanding works
  if(handoverData.outstanding){
    html += '<div style="flex:0 0 auto;margin-bottom:14px;background:#FEF9F0;border:1px solid #F5D9A0;border-radius:6px;padding:10px 14px;">';
    html += '<div style="font-size:9px;font-weight:700;color:#9A6A00;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:5px;">\u26a0 Outstanding Works (to be completed by agreed date)</div>';
    html += '<div style="font-size:11px;color:#1A1A1A;">'+handoverData.outstanding+'</div>';
    html += '</div>';
  }

  // Warranty
  html += '<div style="flex:0 0 auto;background:#F0F7F3;border:1px solid #A8D5B8;border-radius:6px;padding:10px 14px;margin-bottom:14px;">';
  html += '<div style="font-size:9px;font-weight:700;color:#2D7A4F;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:4px;">Warranty Period</div>';
  html += '<div style="font-size:12px;color:#1A1A1A;">Workmanship warranty commences on <strong>'+fmtD(handoverDate)+'</strong> and expires on <strong>'+fmtD(warrantyExpiry)+'</strong> (12 months).</div>';
  html += '<div style="font-size:10px;color:#7A7468;margin-top:3px;">Subject to warranty terms. Excludes normal wear and tear, misuse, and third-party modifications.</div>';
  html += '</div>';

  // Final payment
  if(finalPayment && finalPayment.amount){
    html += '<div style="flex:0 0 auto;background:#F8F6F2;border:1.5px solid #1A1A1A;border-radius:6px;padding:10px 14px;margin-bottom:14px;">';
    html += '<div style="display:flex;justify-content:space-between;align-items:center;">';
    html += '<div>';
    html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:3px;">Final Payment Received</div>';
    html += '<div style="font-size:20px;font-weight:900;color:#1A1A1A;font-family:\'DM Sans\',Arial,sans-serif;">'+fmtM(finalPayment.amount)+'</div>';
    html += '</div>';
    html += '<div style="text-align:right;">';
    if(finalPayment.method) html += '<div style="font-size:11px;color:#7A7468;">'+finalPayment.method+'</div>';
    if(finalPayment.receiptNo) html += '<div style="font-size:10px;color:#B8B2A8;font-family:monospace;">Rcpt: '+finalPayment.receiptNo+'</div>';
    html += '</div></div></div>';
  }

  // Spacer
  html += '<div style="flex:1;min-height:12px;"></div>';

  // Declaration
  html += '<div style="flex:0 0 auto;border-top:1px solid #E8E4DC;padding-top:12px;margin-bottom:16px;">';
  html += '<div style="font-size:9px;color:#7A7468;line-height:1.6;text-align:center;">';
  html += 'By signing below, the client acknowledges that all renovation works have been completed to satisfaction, keys have been received, and accepts the premises in the condition described above. ';
  html += 'This document serves as the official handover record between <strong style="color:#1A1A1A;">'+co.name+'</strong> and the client.';
  html += '</div></div>';

  // Signature blocks
  html += '<div style="flex:0 0 auto;display:grid;grid-template-columns:1fr 1fr;gap:28px;">';
  ['Client', 'Contractor / '+co.name].forEach(function(party){
    html += '<div>';
    html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:20px;">Signed by — '+party+'</div>';
    html += '<div style="border-bottom:1.5px solid #1A1A1A;height:36px;margin-bottom:5px;"></div>';
    html += '<div style="font-size:9px;color:#B8B2A8;">Signature &amp; Date</div>';
    if(party.startsWith('Client')){
      html += '<div style="margin-top:10px;border-bottom:1.5px solid #1A1A1A;height:24px;margin-bottom:5px;"></div>';
      html += '<div style="font-size:9px;color:#B8B2A8;">Full Name &amp; NRIC / FIN</div>';
    } else {
      html += '<div style="margin-top:10px;font-size:10px;color:#1A1A1A;font-weight:600;">'+printedBy+'</div>';
      html += '<div style="font-size:9px;color:#B8B2A8;">Authorised Representative</div>';
    }
    html += '</div>';
  });
  html += '</div>';
  html += '</div>'; // end page 1

  // ═══════════════════════════════════════════════════════════
  // PAGE 2 — DEFECTS LIABILITY RECORD (company copy)
  // ═══════════════════════════════════════════════════════════
  html += '<div style="width:100%;min-height:100vh;display:flex;flex-direction:column;padding:44px 52px 40px;box-sizing:border-box;background:#fff;">';

  html += '<div style="display:flex;justify-content:space-between;align-items:flex-start;flex:0 0 auto;margin-bottom:0;">';
  html += '<div>';
  html += '<div style="font-size:10px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.18em;margin-bottom:4px;">Defects Liability Record — Company Copy</div>';
  html += '<div style="font-family:\'DM Serif Display\',Georgia,serif;font-size:15px;color:#1A1A1A;">'+co.name+'</div>';
  html += '</div>';
  if(fileCode) html += '<div style="font-size:48px;font-weight:900;color:#1A1A1A;line-height:0.9;letter-spacing:-2px;font-family:\'DM Sans\',Arial,sans-serif;">'+fileCode+'</div>';
  html += '</div>';

  html += '<div style="border-top:2.5px solid #1A1A1A;margin:14px 0 16px;flex:0 0 auto;"></div>';

  // Project info strip
  html += '<div style="flex:0 0 auto;display:grid;grid-template-columns:2fr 1fr 1fr;gap:14px;margin-bottom:16px;padding:10px 14px;background:#F8F6F2;border-radius:6px;">';
  html += '<div><div style="font-size:9px;color:#B8B2A8;font-weight:700;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:3px;">Project / Client</div>';
  html += '<div style="font-size:13px;font-weight:700;color:#1A1A1A;">'+proj.name+'</div>';
  html += '<div style="font-size:11px;color:#7A7468;">'+proj.client+(proj.clientAddress?' \xb7 '+proj.clientAddress:'')+'</div></div>';
  html += '<div><div style="font-size:9px;color:#B8B2A8;font-weight:700;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:3px;">Handover Date</div>';
  html += '<div style="font-size:12px;font-weight:600;color:#1A1A1A;">'+fmtD(handoverDate)+'</div></div>';
  html += '<div><div style="font-size:9px;color:#B8B2A8;font-weight:700;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:3px;">Warranty Expires</div>';
  html += '<div style="font-size:12px;font-weight:600;color:#C0392B;">'+fmtD(warrantyExpiry)+'</div></div>';
  html += '</div>';

  // Defects log table
  html += '<div style="flex:0 0 auto;margin-bottom:16px;">';
  html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:8px;">Defects / Snagging Log</div>';
  html += '<table style="width:100%;border-collapse:collapse;">';
  html += '<thead><tr style="background:#1A1A1A;"><th style="padding:7px 8px;font-size:9px;font-weight:700;color:#F8F6F2;text-align:center;width:28px;">#</th><th style="padding:7px 8px;font-size:9px;font-weight:700;color:#F8F6F2;text-align:left;">Description of Defect / Issue</th><th style="padding:7px 8px;font-size:9px;font-weight:700;color:#F8F6F2;text-align:center;width:80px;">Date Reported</th><th style="padding:7px 8px;font-size:9px;font-weight:700;color:#F8F6F2;text-align:center;width:80px;">Date Resolved</th><th style="padding:7px 8px;font-size:9px;font-weight:700;color:#F8F6F2;text-align:center;width:36px;">\u2713</th></tr></thead>';
  html += '<tbody>';
  for(var ri=1;ri<=12;ri++){
    html += '<tr style="border-bottom:1px solid #EDE9E1;"><td style="padding:6px 8px;text-align:center;font-size:10px;color:#B8B2A8;">'+ri+'</td>';
    html += '<td style="padding:6px 8px;"></td>';
    html += '<td style="padding:6px 8px;border-left:1px solid #EDE9E1;"></td>';
    html += '<td style="padding:6px 8px;border-left:1px solid #EDE9E1;"></td>';
    html += '<td style="padding:6px 8px;border-left:1px solid #EDE9E1;text-align:center;"><div style="width:13px;height:13px;border:1.5px solid #C4A882;border-radius:2px;display:inline-block;"></div></td></tr>';
  }
  html += '</tbody></table></div>';

  // Photo area
  html += '<div style="flex:0 0 auto;margin-bottom:16px;">';
  html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:8px;">Photo Evidence (attach photos below)</div>';
  html += '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;">';
  for(var pi=0;pi<4;pi++){
    html += '<div style="border:1.5px dashed #D0CDC8;border-radius:6px;height:70px;display:flex;align-items:center;justify-content:center;">';
    html += '<span style="font-size:9px;color:#D0CDC8;font-weight:600;">Photo '+(pi+1)+'</span>';
    html += '</div>';
  }
  html += '</div></div>';

  // Notes
  html += '<div style="flex:0 0 auto;margin-bottom:14px;">';
  html += '<div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:6px;">Follow-up Notes &amp; Actions</div>';
  for(var li=0;li<5;li++){
    html += '<div style="border-bottom:1px solid #E8E4DC;height:22px;margin-bottom:2px;"></div>';
  }
  html += '</div>';

  html += '<div style="flex:1;"></div>';

  // Footer p2
  html += '<div style="border-top:1px solid #E8E4DC;padding-top:10px;flex:0 0 auto;display:flex;justify-content:space-between;align-items:center;">';
  html += '<div style="font-size:9px;color:#B8B2A8;">'+co.name+(fileCode?' \xb7 File '+fileCode:'')+' \xb7 Defects Record</div>';
  html += '<div style="font-size:9px;color:#B8B2A8;">Page 2 of 2 \xb7 Company Copy — Not for Client</div>';
  html += '</div>';
  html += '</div>'; // end page 2

  return html;
};





const printDoc = (html, title='Document', fullPage=false) => {
  const fullHTML = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${title}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800;900&family=DM+Serif+Display&display=swap" rel="stylesheet">
  <style>
    *{margin:0;padding:0;box-sizing:border-box;}
    html,body{height:100%;}
    body{font-family:"DM Sans",-apple-system,sans-serif;font-size:13px;color:#1A1A1A;background:${fullPage?'#fff':'#F8F6F2'};${fullPage?'':'padding:36px;'}}
    @media print{
      html,body{height:100%;margin:0;padding:0;background:#fff;}
      @page{margin:${fullPage?'0':'12mm'};size:A4;}
      .no-print{display:none!important;}
    }
  </style>
</head>
<body>
  ${fullPage?html:`<div style="max-width:760px;margin:0 auto;">${html}</div>`}
  <div class="no-print" style="text-align:center;padding:20px;">
    <button onclick="window.print()" style="background:#1A1A1A;color:#F8F6F2;border:none;padding:11px 32px;border-radius:10px;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit;">
      Print / Save as PDF
    </button>
    <p style="margin-top:10px;font-size:11px;color:#B8B2A8;">Set paper to A4, margins to None for cover pages.</p>
  </div>
  <script>window.onload=function(){setTimeout(function(){window.print();},350);};<\/script>
</body>
</html>`;
  const blob = new Blob([fullHTML], {type:'text/html;charset=utf-8'});
  const url = URL.createObjectURL(blob);
  const w = window.open(url, '_blank');
  if (!w) {
    const a = document.createElement('a');
    a.href=url; a.download=title.replace(/[^a-z0-9]/gi,'_')+'.html';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
  }
  setTimeout(()=>URL.revokeObjectURL(url), 30000);
};

const buildReceiptHTML = (payment, proj, receiptNo, co) => `
<div style="max-width:620px;margin:0 auto;font-family:'DM Sans',-apple-system,sans-serif;border:1px solid #EDE9E1;border-radius:16px;overflow:hidden;">
  <!-- Header -->
  <div style="background:#1A1A1A;padding:28px 32px;display:flex;justify-content:space-between;align-items:flex-start;">
    <div>
      <div style="font-family:'DM Serif Display',Georgia,serif;font-size:20px;color:#F8F6F2;line-height:1.1;">${co.name}</div>
      <div style="font-size:11px;color:#7A7468;margin-top:4px;">UEN: ${co.uen||'—'}</div>
    </div>
    <div style="text-align:right;">
      <div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:4px;">Official Receipt</div>
      <div style="font-size:14px;font-weight:600;color:#C4A882;font-family:monospace;">${receiptNo}</div>
    </div>
  </div>
  <!-- Body -->
  <div style="padding:24px 32px;background:#fff;">
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px;padding-bottom:20px;border-bottom:1px solid #EDE9E1;">
      <div>
        <div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:6px;">Received From</div>
        <div style="font-size:14px;font-weight:600;color:#1A1A1A;">${proj?.client||'—'}</div>
        ${proj?.clientPhone?`<div style="font-size:11px;color:#7A7468;margin-top:2px;">${proj.clientPhone}</div>`:''}
        ${proj?.clientEmail?`<div style="font-size:11px;color:#7A7468;">${proj.clientEmail}</div>`:''}
      </div>
      <div style="text-align:right;">
        <div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:6px;">Transfer Date</div>
        <div style="font-size:13px;font-weight:600;color:#1A1A1A;">${payment.date?new Date(payment.date).toLocaleDateString('en-SG',{day:'2-digit',month:'long',year:'numeric'}):'—'}</div>
      </div>
    </div>
    <div style="background:#F8F6F2;border-radius:10px;padding:18px 20px;margin-bottom:20px;border:1px solid #EDE9E1;">
      <div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:12px;">Payment Details</div>
      <table style="width:100%;border-collapse:collapse;">
        <tr style="border-bottom:1px solid #EDE9E1;">
          <td style="padding:7px 0;font-size:11px;color:#7A7468;">Project</td>
          <td style="padding:7px 0;font-size:12px;font-weight:600;text-align:right;color:#1A1A1A;">${proj?.name||'—'}</td>
        </tr>
        ${proj?.clientAddress?`<tr style="border-bottom:1px solid #EDE9E1;"><td style="padding:7px 0;font-size:11px;color:#7A7468;">Project Address</td><td style="padding:7px 0;font-size:12px;font-weight:600;text-align:right;color:#1A1A1A;">${proj.clientAddress}</td></tr>`:''}
        <tr style="border-bottom:1px solid #EDE9E1;">
          <td style="padding:7px 0;font-size:11px;color:#7A7468;">Payment Type</td>
          <td style="padding:7px 0;font-size:12px;font-weight:600;text-align:right;color:#1A1A1A;">${payment.type} Payment</td>
        </tr>
        ${payment.paymentMethod?`<tr style="border-bottom:1px solid #EDE9E1;"><td style="padding:7px 0;font-size:11px;color:#7A7468;">Transfer Method</td><td style="padding:7px 0;font-size:12px;font-weight:600;text-align:right;color:#1A1A1A;">${payment.paymentMethod==='paynow'?'PayNow':payment.paymentMethod==='bank_transfer'?'Bank Transfer':'Cash'}</td></tr>`:''}
        ${payment.reference?`<tr style="border-bottom:1px solid #EDE9E1;"><td style="padding:7px 0;font-size:11px;color:#7A7468;">Transaction Ref.</td><td style="padding:7px 0;font-size:12px;font-weight:600;text-align:right;color:#1A1A1A;font-family:monospace;">${payment.reference}</td></tr>`:''}
        <tr>
          <td style="padding:12px 0 0;font-size:13px;font-weight:600;color:#1A1A1A;">Amount Received</td>
          <td style="padding:12px 0 0;font-size:22px;font-weight:700;color:#2D7A4F;text-align:right;font-family:'DM Serif Display',Georgia,serif;">S$${Number(payment.amount).toLocaleString('en-SG',{minimumFractionDigits:2})}</td>
        </tr>
      </table>
    </div>
    <div style="text-align:center;font-size:10px;color:#B8B2A8;line-height:1.8;padding-top:12px;border-top:1px solid #EDE9E1;">
      Computer-generated receipt · No signature required · ${co.name} · ${receiptNo}
    </div>
  </div>
</div>`;

const buildSOAHTML = (proj, projInv, projPay, co) => {
  const rev=(proj.contractAmount||0)+(proj.variationOrders||0);
  const totalPaid=projPay.filter(p=>p.status==='Received').reduce((s,p)=>s+p.amount,0);
  const outstanding=rev-totalPaid;
  const payRows=projPay.map(p=>`
    <tr style="border-bottom:1px solid #EDE9E1;">
      <td style="padding:7px 4px;font-size:11px;color:#1A1A1A;">${p.date?new Date(p.date).toLocaleDateString('en-SG',{day:'2-digit',month:'short',year:'numeric'}):'—'}</td>
      <td style="padding:7px 4px;font-size:11px;color:#1A1A1A;">${p.type} Payment</td>
      <td style="padding:7px 4px;font-size:11px;color:#2D7A4F;text-align:right;font-weight:700;">S$${Number(p.amount).toLocaleString('en-SG',{minimumFractionDigits:2})}</td>
      <td style="padding:7px 4px;font-size:10px;text-align:center;"><span style="background:${p.status==='Received'?'rgba(45,122,79,0.1)':'rgba(196,168,130,0.2)'};color:${p.status==='Received'?'#2D7A4F':'#8A6A3A'};padding:2px 8px;border-radius:20px;font-weight:600;">${p.status}</span></td>
    </tr>`).join('');
  return `
<div style="max-width:700px;margin:0 auto;font-family:'DM Sans',-apple-system,sans-serif;border:1px solid #EDE9E1;border-radius:16px;overflow:hidden;">
  <!-- Header -->
  <div style="background:#1A1A1A;padding:28px 32px;display:flex;justify-content:space-between;align-items:flex-start;">
    <div>
      <div style="font-family:'DM Serif Display',Georgia,serif;font-size:20px;color:#F8F6F2;">${co.name}</div>
      <div style="font-size:11px;color:#7A7468;margin-top:3px;">UEN: ${co.uen||'—'}</div>
    </div>
    <div style="text-align:right;">
      <div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.14em;margin-bottom:4px;">Statement of Account</div>
      <div style="font-size:11px;color:#7A7468;">As at ${new Date().toLocaleDateString('en-SG',{day:'2-digit',month:'long',year:'numeric'})}</div>
    </div>
  </div>
  <!-- Client + Project -->
  <div style="padding:24px 32px;display:grid;grid-template-columns:1fr 1fr;gap:24px;background:#FDFCFA;border-bottom:1px solid #EDE9E1;">
    <div>
      <div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:6px;">Client</div>
      <div style="font-size:14px;font-weight:600;color:#1A1A1A;">${proj.client}</div>
      ${proj.clientPhone?`<div style="font-size:11px;color:#7A7468;margin-top:2px;">${proj.clientPhone}</div>`:''}
      ${proj.clientEmail?`<div style="font-size:11px;color:#7A7468;">${proj.clientEmail}</div>`:''}
    </div>
    <div>
      <div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:6px;">Project</div>
      <div style="font-size:14px;font-weight:600;color:#1A1A1A;">${proj.name}</div>
      <div style="font-size:11px;color:#7A7468;margin-top:2px;">${proj.projectType||'Residential'} · ${proj.status}</div>
    </div>
  </div>
  <!-- Summary stats -->
  <div style="padding:20px 32px;display:grid;grid-template-columns:repeat(3,1fr);gap:12px;border-bottom:1px solid #EDE9E1;">
    ${[
      {l:'Contract Value',v:`S$${Number(rev).toLocaleString('en-SG',{minimumFractionDigits:2})}`,c:'#1A1A1A'},
      {l:'Total Received',v:`S$${Number(totalPaid).toLocaleString('en-SG',{minimumFractionDigits:2})}`,c:'#2D7A4F'},
      {l:'Outstanding',v:`S$${Number(outstanding).toLocaleString('en-SG',{minimumFractionDigits:2})}`,c:outstanding>0?'#C0392B':'#2D7A4F'},
    ].map(({l,v,c})=>`
      <div style="background:#F8F6F2;border-radius:10px;padding:12px 14px;border:1px solid #EDE9E1;">
        <div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.10em;margin-bottom:4px;">${l}</div>
        <div style="font-size:16px;font-weight:700;color:${c};font-family:'DM Serif Display',Georgia,serif;">${v}</div>
      </div>`).join('')}
  </div>
  <!-- Payment history -->
  <div style="padding:20px 32px;border-bottom:1px solid #EDE9E1;">
    <div style="font-size:10px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:10px;">Payment History</div>
    <table style="width:100%;border-collapse:collapse;">
      <thead><tr style="background:#F8F6F2;">
        <th style="padding:8px 4px;font-size:9px;font-weight:700;color:#B8B2A8;text-align:left;text-transform:uppercase;letter-spacing:0.08em;">Date</th>
        <th style="padding:8px 4px;font-size:9px;font-weight:700;color:#B8B2A8;text-align:left;text-transform:uppercase;letter-spacing:0.08em;">Description</th>
        <th style="padding:8px 4px;font-size:9px;font-weight:700;color:#B8B2A8;text-align:right;text-transform:uppercase;letter-spacing:0.08em;">Amount</th>
        <th style="padding:8px 4px;font-size:9px;font-weight:700;color:#B8B2A8;text-align:center;text-transform:uppercase;letter-spacing:0.08em;">Status</th>
      </tr></thead>
      <tbody>${payRows||'<tr><td colspan="4" style="padding:14px;text-align:center;font-size:12px;color:#B8B2A8;">No payments recorded</td></tr>'}</tbody>
    </table>
  </div>
  ${outstanding>0?`
  <div style="padding:16px 32px;border-bottom:1px solid #EDE9E1;background:#FEF4F3;">
    <div style="font-size:13px;font-weight:700;color:#C0392B;">Outstanding Balance: S$${Number(outstanding).toLocaleString('en-SG',{minimumFractionDigits:2})}</div>
    <div style="font-size:11px;color:#C0392B;margin-top:3px;opacity:0.8;">Please arrange payment at your earliest convenience.</div>
  </div>`:`
  <div style="padding:14px 32px;border-bottom:1px solid #EDE9E1;background:rgba(45,122,79,0.06);">
    <div style="font-size:13px;font-weight:700;color:#2D7A4F;">✓ Account fully settled. Thank you!</div>
  </div>`}
  <div style="padding:14px 32px;text-align:center;">
    <div style="font-size:10px;color:#B8B2A8;line-height:1.8;">Computer-generated statement · No signature required · ${co.name}</div>
  </div>
</div>`;
};

const buildWarrantyHTML = (proj, serial, co) => {
  const issued=new Date();
  const expiry=new Date(issued);expiry.setFullYear(expiry.getFullYear()+1);
  const fmt=d=>d.toLocaleDateString('en-SG',{day:'2-digit',month:'long',year:'numeric'});
  return `
<div style="max-width:640px;margin:0 auto;font-family:'DM Sans',-apple-system,sans-serif;border:1px solid #EDE9E1;border-radius:16px;overflow:hidden;page-break-inside:avoid;">
  <!-- Header -->
  <div style="background:#1A1A1A;padding:36px 32px;text-align:center;">
    <div style="font-size:10px;font-weight:700;color:#C4A882;text-transform:uppercase;letter-spacing:0.18em;margin-bottom:10px;">Certificate of</div>
    <div style="font-family:'DM Serif Display',Georgia,serif;font-size:28px;color:#F8F6F2;line-height:1.1;">Workmanship Warranty</div>
    <div style="font-size:12px;color:#7A7468;margin-top:6px;">${co.name}</div>
    <div style="display:inline-flex;align-items:center;gap:10px;background:rgba(196,168,130,0.12);border:1px solid #C4A882;border-radius:8px;padding:7px 18px;margin-top:14px;">
      <span style="font-size:10px;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.10em;">Serial No.</span>
      <span style="font-size:16px;font-weight:700;color:#C4A882;font-family:monospace;">${serial}</span>
    </div>
  </div>
  <!-- Body -->
  <div style="padding:28px 32px;background:#fff;">
    <div style="text-align:center;font-size:12px;color:#7A7468;margin-bottom:22px;line-height:1.7;">
      This certifies that <strong style="color:#1A1A1A;">${proj.client}</strong> has engaged ${co.name} to carry out interior renovation works for the project below.
    </div>
    <!-- Project details grid -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:22px;">
      ${[
        {l:'Project Name',v:proj.name},
        {l:'Project Type',v:proj.projectType||'Residential'},
        {l:'Warranty Issued',v:fmt(issued)},
        {l:'Warranty Expires',v:`<strong style="color:#C0392B;">${fmt(expiry)}</strong>`},
      ].map(({l,v})=>`
        <div style="background:#F8F6F2;border-radius:10px;padding:12px 14px;border:1px solid #EDE9E1;">
          <div style="font-size:9px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.10em;margin-bottom:4px;">${l}</div>
          <div style="font-size:13px;font-weight:600;color:#1A1A1A;">${v}</div>
        </div>`).join('')}
    </div>
    <!-- Coverage -->
    <div style="background:#F8F6F2;border:1px solid #EDE9E1;border-radius:10px;padding:18px 20px;margin-bottom:16px;">
      <div style="font-size:10px;font-weight:700;color:#B8B2A8;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:12px;">Warranty Coverage</div>
      ${[
        'All workmanship carried out by '+co.name+' and appointed subcontractors',
        'Carpentry works including cabinets, built-ins, and custom joinery',
        'Electrical works including wiring, switches, and fixtures installed by us',
        'Plumbing works including pipe-works and fittings installed by us',
        'Painting and surface finishes applied by our team',
        'General installation defects arising from poor workmanship',
      ].map(item=>`
        <div style="display:flex;align-items:flex-start;gap:10px;margin-bottom:7px;">
          <span style="color:#2D7A4F;font-weight:700;font-size:13px;flex-shrink:0;line-height:1.5;">✓</span>
          <span style="font-size:12px;color:#3D3D3D;line-height:1.5;">${item}</span>
        </div>`).join('')}
    </div>
    <!-- Exclusions -->
    <div style="background:#FEF4F3;border:1px solid rgba(192,57,43,0.2);border-radius:10px;padding:14px 18px;margin-bottom:16px;">
      <div style="font-size:10px;font-weight:700;color:#C0392B;text-transform:uppercase;letter-spacing:0.10em;margin-bottom:6px;">Exclusions</div>
      <div style="font-size:11px;color:#7A7468;line-height:1.6;">
        Not covered: damage from misuse, negligence, or accidents · normal wear and tear · water seepage not caused by our installation · furniture or appliances not installed by us · works modified by third parties without our consent.
      </div>
    </div>
    <!-- Claim contact -->
    <div style="background:#F8F6F2;border:1px solid #EDE9E1;border-radius:10px;padding:12px 18px;margin-bottom:16px;text-align:center;">
      <div style="font-size:11px;color:#7A7468;">
        To make a claim, contact ${co.name} and quote <strong style="color:#1A1A1A;font-family:monospace;">${serial}</strong>. All claims are subject to inspection.
      </div>
    </div>
    <!-- Footer -->
    <div style="text-align:center;padding-top:14px;border-top:1px dashed #EDE9E1;">
      <div style="font-size:10px;color:#B8B2A8;line-height:1.8;">
        Computer-generated certificate · No signature required · ${co.name} · Serial: ${serial}
      </div>
    </div>
  </div>
</div>`;
};

const THEMES = {
  light: {
    bg:'#F8F6F2', surface:'#FFFFFF', card:'#FFFFFF', sidebar:'#FFFFFF',
    border:'#E8E4DC', borderLight:'#EDE9E1',
    accent:'#1A1A1A', accentHover:'#333333', accentLight:'rgba(26,26,26,0.06)',
    success:'#2D7A4F', successLight:'rgba(45,122,79,0.08)',
    danger:'#C0392B', dangerLight:'rgba(192,57,43,0.08)',
    warning:'#9A6A00', warningLight:'rgba(154,106,0,0.08)',
    info:'#1A5FA8', infoLight:'rgba(26,95,168,0.08)',
    text:'#1A1A1A', secondary:'#3D3D3D', muted:'#7A7468', dim:'#B8B2A8',
    tan:'#C4A882', tanLight:'rgba(196,168,130,0.15)',
    shadow:'0 1px 4px rgba(0,0,0,0.06), 0 2px 12px rgba(0,0,0,0.04)',
    shadowMd:'0 4px 20px rgba(0,0,0,0.08)',
    shadowLg:'0 8px 40px rgba(0,0,0,0.10)',
  },
  dark: {
    bg:'#141412', surface:'#1E1D1B', card:'#1E1D1B', sidebar:'#1A1917',
    border:'#2E2C28', borderLight:'#252320',
    accent:'#F8F6F2', accentHover:'#E8E4DC', accentLight:'rgba(248,246,242,0.07)',
    success:'#3DAA6A', successLight:'rgba(61,170,106,0.12)',
    danger:'#E05A4A', dangerLight:'rgba(224,90,74,0.12)',
    warning:'#D4A030', warningLight:'rgba(212,160,48,0.12)',
    info:'#4D8FD6', infoLight:'rgba(77,143,214,0.12)',
    text:'#F0EDE8', secondary:'#C8C4BE', muted:'#7A7468', dim:'#4A4844',
    tan:'#C4A882', tanLight:'rgba(196,168,130,0.12)',
    shadow:'0 1px 4px rgba(0,0,0,0.3), 0 2px 12px rgba(0,0,0,0.2)',
    shadowMd:'0 4px 20px rgba(0,0,0,0.35)',
    shadowLg:'0 8px 40px rgba(0,0,0,0.4)',
  },
};

// T is set at runtime by App — starts with light theme
let T = THEMES.light;

// iStyle is also set at runtime (needs T)
let iStyle = {};
const makeIStyle = () => ({
  background: T.card,
  border:`1px solid ${T.borderLight}`,
  borderRadius:10, padding:'11px 14px', fontSize:16,
  color:T.text, outline:'none', width:'100%',
  fontFamily:'inherit', transition:'border-color 0.15s, box-shadow 0.15s',
});

// ── Mobile-responsive helpers ──────────────────────────────────────────────
// Hook: returns true when viewport ≤ 600px (iPhone-sized)
const useIsMobile = () => {
  const [isMobile, setIsMobile] = useState(() => window.innerWidth <= 600);
  useEffect(() => {
    const fn = () => setIsMobile(window.innerWidth <= 600);
    window.addEventListener('resize', fn);
    return () => window.removeEventListener('resize', fn);
  }, []);
  return isMobile;
};

// Responsive grid: single column on mobile, multi-column on desktop
// cols = number of columns on desktop (2 or 3)
const cols = (n, isMobile) => ({
  display: 'grid',
  gridTemplateColumns: isMobile ? '1fr' : `repeat(${n}, 1fr)`,
  gap: isMobile ? 10 : 12,
});

// Method selector: 2-col on mobile (icons + short labels), 3-col on desktop
const methodGrid = (isMobile) => ({
  display: 'grid',
  gridTemplateColumns: isMobile ? '1fr 1fr' : '1fr 1fr 1fr',
  gap: 8,
});
// ─────────────────────────────────────────────────────────────────────────────


const fmtTime = (iso) => iso ? new Date(iso).toLocaleTimeString('en-SG',{hour:'2-digit',minute:'2-digit',hour12:true}) : '—';
const fmtHours = (h) => (h != null && h > 0) ? `${Math.floor(h)}h ${Math.round((h%1)*60)}m` : '0h 0m';
const daysUntil = (dateStr) => {
  if(!dateStr) return 999;
  return Math.ceil((new Date(dateStr).setHours(23,59,59) - Date.now()) / 864e5);
};
const expiryColor = (days) => days <= 30 ? T.danger : days <= 60 ? T.warning : T.success;
const expiryLabel = (days) => days < 0 ? 'EXPIRED' : days === 0 ? 'Expires today' : `${days}d left`;

const calcComm = (proj, invoices) => {
  const exp = invoices.filter(i=>i.projectId===proj.id).reduce((s,i)=>s+i.total,0);
  const rev = (proj.contractAmount||0)+(proj.variationOrders||0);
  const gross = rev-exp;
  const dMethod = proj.designerCommMethod||proj.commissionMethod||'profit_pct';
  const pmMethod = proj.pmCommMethod||proj.commissionMethod||'profit_pct';
  let dComm=0,pmComm=0;
  if(dMethod==='profit_pct'){const r=parseFloat(proj.designerRate)||0;dComm=gross>0&&r>0?gross*r/100:0;}
  else if(dMethod==='project_sum'){const r=parseFloat(proj.designerRate)||0;dComm=rev>0&&r>0?rev*r/100:0;}
  else{dComm=parseFloat(proj.designerCommAmt)||0;}
  if(pmMethod==='profit_pct'){const r=parseFloat(proj.pmRate)||0;pmComm=gross>0&&r>0?gross*r/100:0;}
  else if(pmMethod==='project_sum'){const r=parseFloat(proj.pmRate)||0;pmComm=rev>0&&r>0?rev*r/100:0;}
  else{pmComm=parseFloat(proj.pmCommAmt)||0;}
  return {rev,exp,gross,dComm,pmComm,dMethod,pmMethod};
};

const COMM_METHOD_LABEL = {
  profit_pct:'% of Gross Profit',
  project_sum:'% of Total Project Sum',
  fixed_rate:'Fixed Rate',
};

const Badge = ({color,children,sm}) => (
  <span style={{
    background:`${color}12`,color,
    border:`1px solid ${color}28`,
    borderRadius:6,
    padding:sm?'2px 7px':'3px 10px',
    fontSize:sm?10:11,
    fontWeight:600,
    letterSpacing:'0.01em',
    whiteSpace:'nowrap',
  }}>{children}</span>
);

const Btn = ({variant='primary',size='md',onClick,children,disabled,loading,full,style:sx}) => {
  const [hov,setHov]=useState(false);
  const pad={sm:'7px 14px',md:'10px 18px',lg:'12px 24px'}[size];
  const fs={sm:12,md:13,lg:14}[size];
  const vs={
    primary:{background:hov?T.accentHover:T.accent,color:'#fff',boxShadow:hov?'0 4px 14px rgba(26,26,26,0.22)':'none'},
    secondary:{background:hov?T.borderLight:T.bg,color:T.secondary,border:`1px solid ${T.borderLight}`},
    ghost:{background:'transparent',color:T.muted},
    danger:{background:hov?'#FDEEEC':'#FEF4F3',color:T.danger,border:`1px solid #F5C8C4`},
  }[variant];
  return (
    <button onClick={onClick} disabled={disabled||loading}
      onMouseEnter={()=>setHov(true)} onMouseLeave={()=>setHov(false)}
      style={{...vs,padding:pad,fontSize:fs,borderRadius:10,border:vs.border||'none',fontWeight:600,
        cursor:disabled||loading?'not-allowed':'pointer',opacity:disabled||loading?0.38:1,
        display:'inline-flex',alignItems:'center',gap:6,letterSpacing:'0em',
        width:full?'100%':undefined,justifyContent:full?'center':undefined,
        fontFamily:'inherit',transition:'all 0.15s ease',...sx}}>
      {loading && <Loader2 size={12} style={{animation:'spin 1s linear infinite'}}/>}
      {children}
    </button>
  );
};

const Field = ({label,type='text',value,onChange,placeholder,as='input',options}) => (
  <div>
    {label&&<label style={{fontSize:13,fontWeight:500,color:T.muted,display:'block',marginBottom:7}}>{label}</label>}
    {as==='select'?(
      <select value={value} onChange={e=>onChange(e.target.value)}
        style={{...iStyle,appearance:'none',cursor:'pointer',
          backgroundImage:`url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8' fill='none'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%23AEAEB2' stroke-width='1.5' stroke-linecap='round'/%3E%3C/svg%3E")`,
          backgroundRepeat:'no-repeat',backgroundPosition:'right 13px center',paddingRight:36}}>
        {options?.map(o=><option key={o.v??o} value={o.v??o}>{o.l??o}</option>)}
      </select>
    ):(
      <input type={type} value={value} onChange={e=>onChange(e.target.value)}
        placeholder={placeholder} style={iStyle}/>
    )}
  </div>
);

const isMobileDevice = () => typeof window !== 'undefined' && window.innerWidth < 768;

// DropZone — wraps any upload trigger with drag-and-drop on desktop
// onDrop receives a File object
const DropZone = ({onDrop, accept='*', children, style={}}) => {
  const [dragging, setDragging] = useState(false);
  // Only disable on touch-only devices (no mouse)
  if(typeof window !== 'undefined' && !window.matchMedia('(hover: hover)').matches)
    return <div style={style}>{children}</div>;

  const checkType = (file) => {
    if(accept==='*') return true;
    const types = accept.split(',').map(s=>s.trim());
    return types.some(t=>{
      if(t==='application/pdf') return file.type==='application/pdf' || file.name.toLowerCase().endsWith('.pdf');
      if(t.startsWith('.')) return file.name.toLowerCase().endsWith(t);
      if(t.endsWith('/*')) return file.type.startsWith(t.replace('/*',''));
      return file.type===t;
    });
  };

  const handleDrop = e => {
    e.preventDefault(); e.stopPropagation();
    setDragging(false);
    const file = e.dataTransfer.files?.[0];
    if(file && checkType(file)) onDrop(file);
  };

  return (
    <div
      onDragOver={e=>{e.preventDefault();e.stopPropagation();setDragging(true);}}
      onDragEnter={e=>{e.preventDefault();e.stopPropagation();setDragging(true);}}
      onDragLeave={e=>{e.preventDefault();e.stopPropagation();setDragging(false);}}
      onDrop={handleDrop}
      style={{
        ...style,
        outline: dragging ? `2.5px dashed ${T.accent}` : '2.5px dashed transparent',
        outlineOffset: 2,
        borderRadius: 14,
        background: dragging ? `${T.accent}08` : 'transparent',
        transition: 'outline 0.12s, background 0.12s',
        position: 'relative',
      }}>
      {children}
      {dragging && (
        <div style={{
          position:'absolute', inset:0, borderRadius:12,
          background:`${T.accent}12`,
          display:'flex', alignItems:'center', justifyContent:'center',
          pointerEvents:'none', zIndex:10,
        }}>
          <div style={{background:'#fff', borderRadius:12, padding:'12px 22px',
            boxShadow:'0 8px 32px rgba(0,113,227,0.18)',
            display:'flex', alignItems:'center', gap:10}}>
            <Upload size={18} style={{color:T.accent}}/>
            <span style={{fontSize:14, fontWeight:700, color:T.accent}}>Drop to upload</span>
          </div>
        </div>
      )}
    </div>
  );
};

const Modal = ({title,onClose,children,wide}) => (
  <div style={{position:'fixed',inset:0,zIndex:100,display:'flex',alignItems:'flex-end',
    justifyContent:'center',background:'rgba(29,29,31,0.36)',backdropFilter:'blur(24px)',WebkitBackdropFilter:'blur(24px)'}}>
    <div style={{background:'rgba(255,255,255,0.97)',border:`1px solid ${T.borderLight}`,
      borderRadius:'20px 20px 0 0',
      width:'100%',maxWidth:wide?860:580,
      maxHeight:'92dvh',
      display:'flex',flexDirection:'column',
      boxShadow:'0 -8px 40px rgba(0,0,0,0.14)',
      paddingBottom:'env(safe-area-inset-bottom,0px)'}}>
      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',
        padding:'18px 20px 14px',borderBottom:`1px solid ${T.borderLight}`,flexShrink:0}}>
        <span style={{fontWeight:700,color:T.text,fontSize:17,letterSpacing:'-0.02em',flex:1,marginRight:12}}>{title}</span>
        <button type="button" onClick={onClose} style={{background:T.bg,border:`1px solid ${T.borderLight}`,cursor:'pointer',
          color:T.muted,display:'flex',padding:6,borderRadius:8,flexShrink:0}}><X size={15}/></button>
      </div>
      <div style={{overflowY:'auto',padding:'20px',flex:1,WebkitOverflowScrolling:'touch'}}>{children}</div>
    </div>
  </div>
);

// PDF viewer using PDF.js — renders pages as canvas, works inside sandboxed iframes
const PdfViewer = ({src}) => {
  const [pages, setPages] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const canvasRef = useRef();

  useEffect(()=>{
    if(!src) return;
    setLoading(true); setError(''); setPages([]);
    // Load PDF.js from CDN
    const loadPdf = async () => {
      try {
        // Dynamically import PDF.js if not already loaded
        if(!window.pdfjsLib){
          await new Promise((res,rej)=>{
            const s=document.createElement('script');
            s.src='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
            s.onload=res; s.onerror=rej;
            document.head.appendChild(s);
          });
          window.pdfjsLib.GlobalWorkerOptions.workerSrc=
            'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }
        const pdf = await window.pdfjsLib.getDocument({data: atob(src.split(',')[1])}).promise;
        setPages(Array.from({length:pdf.numPages},(_,i)=>i+1));
        setLoading(false);
        // Render first page
        renderPage(pdf, 1);
      } catch(e){
        setError('Could not render PDF. Try downloading it instead.');
        setLoading(false);
      }
    };

    const renderPage = async (pdf, pageNum) => {
      const page = await pdf.getPage(pageNum);
      const canvas = canvasRef.current;
      if(!canvas) return;
      const viewport = page.getViewport({scale: Math.min(2, (canvas.parentElement?.clientWidth||800) / page.getViewport({scale:1}).width)});
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({canvasContext: canvas.getContext('2d'), viewport}).promise;
    };

    // Store pdf ref for page navigation
    const renderPageByNum = async (num) => {
      if(!window.pdfjsLib) return;
      const pdf = await window.pdfjsLib.getDocument({data: atob(src.split(',')[1])}).promise;
      renderPage(pdf, num);
    };

    window._pdfRenderPage = renderPageByNum;
    loadPdf();
  },[src]);

  const goPage = (n) => {
    const next = Math.min(pages.length, Math.max(1, n));
    setCurrentPage(next);
    window._pdfRenderPage?.(next);
  };

  if(loading) return (
    <div style={{display:'flex',flexDirection:'column',alignItems:'center',gap:12,padding:40,color:'rgba(255,255,255,0.7)'}}>
      <div style={{width:36,height:36,border:'3px solid rgba(255,255,255,0.3)',borderTopColor:'#fff',borderRadius:'50%',animation:'spin 0.8s linear infinite'}}/>
      <span style={{fontSize:13}}>Loading PDF…</span>
    </div>
  );

  if(error) return (
    <div style={{color:'rgba(255,255,255,0.6)',fontSize:13,textAlign:'center',padding:40}}>{error}</div>
  );

  return (
    <div style={{display:'flex',flexDirection:'column',alignItems:'center',gap:12,width:'100%',maxWidth:960}}>
      {pages.length>1&&(
        <div style={{display:'flex',alignItems:'center',gap:12}}>
          <button type="button" onClick={()=>goPage(currentPage-1)} disabled={currentPage===1}
            style={{background:'rgba(255,255,255,0.15)',border:'1px solid rgba(255,255,255,0.25)',color:'#fff',borderRadius:8,padding:'5px 12px',cursor:'pointer',fontFamily:'inherit',fontSize:13,opacity:currentPage===1?0.4:1}}>
            ‹ Prev
          </button>
          <span style={{color:'rgba(255,255,255,0.7)',fontSize:13}}>Page {currentPage} of {pages.length}</span>
          <button type="button" onClick={()=>goPage(currentPage+1)} disabled={currentPage===pages.length}
            style={{background:'rgba(255,255,255,0.15)',border:'1px solid rgba(255,255,255,0.25)',color:'#fff',borderRadius:8,padding:'5px 12px',cursor:'pointer',fontFamily:'inherit',fontSize:13,opacity:currentPage===pages.length?0.4:1}}>
            Next ›
          </button>
        </div>
      )}
      <div style={{background:'#fff',borderRadius:12,overflow:'hidden',boxShadow:'0 20px 80px rgba(0,0,0,0.55)',maxWidth:'100%'}}>
        <canvas ref={canvasRef} style={{display:'block',maxWidth:'100%'}}/>
      </div>
    </div>
  );
};

const Lightbox = ({src,title,onClose}) => {
  const isPdf = src?.startsWith('data:application/pdf') || src?.toLowerCase().endsWith('.pdf');
  return (
    <div onClick={onClose} style={{position:'fixed',inset:0,zIndex:200,
      background:'rgba(0,0,0,0.88)',backdropFilter:'blur(28px)',WebkitBackdropFilter:'blur(28px)',
      display:'flex',alignItems:'center',justifyContent:'center',flexDirection:'column',gap:14,cursor:'zoom-out',
      overflowY:'auto',padding:'20px 16px'}}>
      <div onClick={e=>e.stopPropagation()} style={{display:'flex',alignItems:'center',justifyContent:'space-between',
        width:'100%',maxWidth:980,padding:'0 4px',flexShrink:0}}>
        <div style={{display:'flex',alignItems:'center',gap:10}}>
          <span style={{fontSize:18}}>{isPdf?'📄':'🖼️'}</span>
          <span style={{color:'rgba(255,255,255,0.7)',fontSize:13,fontWeight:500}}>{title}</span>
        </div>
        <div style={{display:'flex',gap:8}}>
          {isPdf&&(
            <a href={src} download={`${title||'document'}.pdf`}
              onClick={e=>e.stopPropagation()}
              style={{background:'rgba(255,255,255,0.12)',border:'1px solid rgba(255,255,255,0.2)',cursor:'pointer',
                color:'#fff',display:'flex',alignItems:'center',gap:6,padding:'7px 12px',borderRadius:9,
                textDecoration:'none',fontSize:12,fontWeight:600}}>
              <Download size={13}/>Download PDF
            </a>
          )}
          <button type="button" onClick={onClose} style={{background:'rgba(255,255,255,0.12)',border:'1px solid rgba(255,255,255,0.2)',cursor:'pointer',
            color:'#fff',display:'flex',padding:8,borderRadius:9}}><X size={15}/></button>
        </div>
      </div>
      <div onClick={e=>e.stopPropagation()} style={{width:'100%',cursor:'default',display:'flex',justifyContent:'center'}}>
        {isPdf
          ? <PdfViewer src={src}/>
          : <img src={src} alt={title}
              style={{maxWidth:'min(960px,94vw)',maxHeight:'82vh',borderRadius:14,objectFit:'contain',
                boxShadow:'0 20px 80px rgba(0,0,0,0.55)',cursor:'default'}}/>
        }
      </div>
    </div>
  );
};

const Stat = ({label,value,sub,Icon,color=T.accent}) => (
  <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:'22px 24px',boxShadow:T.shadow}}>
    <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:16}}>
      <span style={{fontSize:12,fontWeight:500,color:T.muted,letterSpacing:'0.01em'}}>{label}</span>
      <div style={{background:`${color}12`,color,borderRadius:12,width:34,height:34,
        display:'flex',alignItems:'center',justifyContent:'center'}}><Icon size={15}/></div>
    </div>
    <div style={{fontSize:26,fontWeight:700,color:T.text,letterSpacing:'-0.03em'}}>{value}</div>
    {sub&&<div style={{fontSize:12,color:T.dim,marginTop:6}}>{sub}</div>}
  </div>
);

// -- Avatar — profile photo or initial fallback --
const Avatar = ({photo, name, size=36, color=T.accent, onUpload, editable=false}) => {
  const fileRef = useRef();
  const bgColor = color;
  const initial = name ? name.charAt(0).toUpperCase() : '?';
  const fontSize = Math.round(size * 0.38);
  const handleFile = async (e) => {
    const f = e.target.files?.[0];
    if(!f || !onUpload) return;
    const r = new FileReader();
    r.onload = async ev => {
      const compressed = await compressPhoto(ev.target.result);
      onUpload(compressed);
    };
    r.readAsDataURL(f);
    e.target.value = '';
  };
  const triggerUpload = () => fileRef.current?.click();
  return (
    <div style={{position:'relative',display:'inline-block',flexShrink:0}}>
      {/* Clicking the whole circle also opens the picker when editable */}
      <div
        onClick={editable ? triggerUpload : undefined}
        style={{
          width:size, height:size, borderRadius:'50%',
          background: photo ? (photo.includes('dicebear') ? '#F5F5F7' : 'transparent') : `${bgColor}18`,
          border: `2px solid ${editable ? bgColor+'50' : photo ? T.borderLight : bgColor+'30'}`,
          display:'flex', alignItems:'center', justifyContent:'center',
          overflow:'hidden', flexShrink:0,
          cursor: editable ? 'pointer' : 'default',
          position:'relative',
        }}>
        {photo
          ? <img src={photo} alt={name} style={{width:'100%',height:'100%',objectFit:'cover'}}/>
          : <span style={{fontSize,fontWeight:700,color:bgColor,letterSpacing:'-0.01em'}}>{initial}</span>
        }
        {/* Hover overlay hint when editable */}
        {editable && (
          <div style={{position:'absolute',inset:0,borderRadius:'50%',
            background:'rgba(0,0,0,0.28)',display:'flex',alignItems:'center',justifyContent:'center',
            opacity:0,transition:'opacity 0.15s'}}
            onMouseEnter={e=>e.currentTarget.style.opacity='1'}
            onMouseLeave={e=>e.currentTarget.style.opacity='0'}>
            <Camera size={Math.round(size*0.28)} style={{color:'#fff'}}/>
          </div>
        )}
      </div>

      {editable && (
        <>
          {/* Camera badge — larger touch target */}
          <button
            type="button"
            onClick={triggerUpload}
            title="Upload profile photo"
            style={{
              position:'absolute', bottom:-3, right:-3,
              width:Math.max(22,Math.round(size*0.42)), height:Math.max(22,Math.round(size*0.42)),
              borderRadius:'50%', background:T.accent, border:'2.5px solid #fff',
              cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center',
              boxShadow:'0 2px 6px rgba(0,113,227,0.35)',
              zIndex:2,
            }}>
            <Camera size={Math.max(10,Math.round(size*0.22))} style={{color:'#fff'}}/>
          </button>
          <input ref={fileRef} type="file" accept="image/*" capture="environment"
            style={{display:'none', position:'absolute'}} onChange={handleFile}/>
        </>
      )}
    </div>
  );
};

const ConfirmDelete = ({matchValue,typeLabel,impact,onConfirm,onClose}) => {
  const [typed,setTyped]=useState('');
  const [shake,setShake]=useState(false);
  const matches=typed.trim().toLowerCase()===matchValue?.trim().toLowerCase();
  const tryConfirm=()=>{if(!matches){setShake(true);setTimeout(()=>setShake(false),550);return;}onConfirm();};
  return (
    <Modal title="" onClose={onClose}>
      <div style={{textAlign:'center',paddingBottom:4}}>
        <div style={{width:52,height:52,background:T.dangerLight,borderRadius:16,display:'flex',alignItems:'center',justifyContent:'center',margin:'0 auto 18px'}}>
          <AlertCircle size={24} style={{color:T.danger}}/>
        </div>
        <div style={{fontSize:19,fontWeight:700,color:T.text,letterSpacing:'-0.025em',marginBottom:8}}>Delete {typeLabel}?</div>
        <div style={{fontSize:13,color:T.muted,lineHeight:1.65}}>Moving <strong style={{color:T.secondary,fontFamily:'monospace',fontSize:12,background:T.bg,padding:'1px 5px',borderRadius:4}}>{matchValue}</strong> to Trash. Recoverable for 30 days.</div>
        {impact&&(<div style={{background:'#FFFBEB',border:'1px solid #FDE68A',borderRadius:10,padding:'10px 14px',margin:'14px 0',textAlign:'left',display:'flex',gap:10,alignItems:'flex-start'}}><Info size={14} style={{color:'#B45309',flexShrink:0,marginTop:1}}/><span style={{fontSize:12,color:'#78350F',lineHeight:1.55}}>{impact}</span></div>)}
        <div style={{margin:'18px 0 4px',textAlign:'left'}}>
          <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:7}}>Type <span style={{fontFamily:'monospace',fontWeight:700,color:T.text,background:T.bg,padding:'1px 6px',borderRadius:4,fontSize:12}}>{matchValue}</span> to confirm</label>
          <input autoFocus value={typed} onChange={e=>setTyped(e.target.value)} onKeyDown={e=>e.key==='Enter'&&tryConfirm()} placeholder={`Type ${matchValue}...`}
            style={{...iStyle,animation:shake?'shake 0.5s ease':'none',borderColor:typed&&!matches?T.danger:typed&&matches?T.success:T.borderLight,boxShadow:typed&&matches?`0 0 0 3px ${T.success}20`:typed&&!matches?`0 0 0 3px ${T.danger}10`:undefined}}/>
          {typed&&!matches&&<div style={{fontSize:12,color:T.danger,marginTop:5}}>Does not match — check spelling and try again.</div>}
          {typed&&matches&&<div style={{fontSize:12,color:T.success,marginTop:5}}>Confirmed — ready to proceed.</div>}
        </div>
      </div>
      <div style={{display:'flex',gap:10,justifyContent:'flex-end',marginTop:18}}>
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn variant="danger" onClick={tryConfirm} disabled={!matches}><Trash2 size={13}/>Move to Trash</Btn>
      </div>
    </Modal>
  );
};

const UndoToast = ({message,onUndo,onDismiss}) => {
  const [pct,setPct]=useState(100);
  const [vis,setVis]=useState(false);
  useEffect(()=>{
    setTimeout(()=>setVis(true),20);
    const start=Date.now(),dur=5000;
    const tick=setInterval(()=>{const el=Date.now()-start;setPct(Math.max(0,100-el/dur*100));if(el>=dur){clearInterval(tick);onDismiss();}},40);
    return()=>clearInterval(tick);
  },[]);
  return (
    <div style={{position:'fixed',bottom:28,left:'50%',transform:`translateX(-50%) translateY(${vis?0:16}px)`,opacity:vis?1:0,zIndex:300,background:'#1D1D1F',color:'#fff',borderRadius:16,padding:'13px 18px',boxShadow:'0 8px 32px rgba(0,0,0,0.25)',display:'flex',alignItems:'center',gap:14,minWidth:300,transition:'all 0.28s cubic-bezier(0.34,1.56,0.64,1)'}}>
      <RotateCcw size={14} style={{color:'#8E8E93',flexShrink:0}}/>
      <span style={{fontSize:13,fontWeight:500,flex:1}}>{message}</span>
      <button onClick={onUndo} style={{background:'rgba(255,255,255,0.14)',border:'none',borderRadius:8,padding:'5px 13px',color:'#fff',fontSize:12,fontWeight:600,cursor:'pointer',fontFamily:'inherit'}}>Undo</button>
      <button onClick={onDismiss} style={{background:'none',border:'none',cursor:'pointer',color:'#636366',display:'flex',padding:2}}><X size={13}/></button>
      <div style={{position:'absolute',bottom:0,left:0,right:0,height:3,background:'rgba(255,255,255,0.08)',borderRadius:'0 0 16px 16px',overflow:'hidden'}}>
        <div style={{height:'100%',width:`${pct}%`,background:T.accent,transition:'width 0.04s linear',borderRadius:'0 0 16px 16px'}}/>
      </div>
    </div>
  );
};

function TrashBin({trash,onRestore,onPermanentDelete,isSuperAdmin}){
  const [filter,setFilter]=useState('All');
  const [confirmPerm,setConfirmPerm]=useState(null);
  const daysLeft=d=>Math.max(0,Math.ceil((new Date(d).getTime()+365*24*60*60*1000-Date.now())/864e5));
  const TL={project:'Project',invoice:'Invoice',payment:'Payment',user:'User',staffClaim:'Expense Claim'};
  const TI={project:FolderOpen,invoice:Receipt,payment:CreditCard,user:Users,staffClaim:DollarSign};
  const TC={project:T.info,invoice:T.warning,payment:T.success,user:T.danger,staffClaim:'#7c3aed'};
  const counts={All:trash.length,project:0,invoice:0,payment:0,user:0,staffClaim:0};
  trash.forEach(t=>{if(counts[t._trashType]!==undefined)counts[t._trashType]++;});
  const shown=[...(filter==='All'?trash:trash.filter(t=>t._trashType===filter))].sort((a,b)=>new Date(b._deletedAt)-new Date(a._deletedAt));

  if(trash.length===0) return (
    <div style={{textAlign:'center',padding:'72px 0'}}>
      <div style={{width:60,height:60,background:T.bg,borderRadius:18,display:'flex',alignItems:'center',justifyContent:'center',margin:'0 auto 18px',border:`1px solid ${T.borderLight}`}}><Trash size={24} style={{color:T.dim}}/></div>
      <div style={{fontSize:17,fontWeight:600,color:T.secondary,marginBottom:6}}>Trash is empty</div>
      <div style={{fontSize:13,color:T.muted,maxWidth:320,margin:'0 auto',lineHeight:1.6}}>Deleted items are kept for 12 months. Only the Super Admin can permanently delete items.</div>
    </div>
  );

  return(
    <div style={{display:'flex',flexDirection:'column',gap:18}}>
      {!isSuperAdmin&&(
        <div style={{background:'rgba(109,40,217,0.06)',border:'1px solid rgba(109,40,217,0.18)',borderRadius:12,padding:'12px 16px',fontSize:13,color:'#6d28d9',display:'flex',alignItems:'center',gap:8}}>
          <Shield size={14} style={{flexShrink:0}}/>Items are kept for 12 months. Contact the Super Admin to permanently remove items.
        </div>
      )}
      <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
        {['All','project','invoice','payment','staffClaim','user'].map(f=>(
          <button key={f} onClick={()=>setFilter(f)} style={{padding:'6px 14px',borderRadius:20,
            border:`1px solid ${filter===f?T.text:T.borderLight}`,
            background:filter===f?T.text:'transparent',
            color:filter===f?'#F8F6F2':T.muted,
            fontSize:12,fontWeight:500,cursor:'pointer',fontFamily:'inherit',transition:'all 0.15s'}}>
            {f==='All'?'All':TL[f]} ({counts[f]})
          </button>
        ))}
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:10}}>
        {shown.map(item=>{
          const t=item._trashType,Icon=TI[t]||Trash,color=TC[t]||T.dim;
          const left=daysLeft(item._deletedAt);
          const urgent=left<=30;
          const label=t==='staffClaim'?`${item.submittedBy} — ${item.type||'Claim'}`:item.name||item.invoiceNo||`${item.type} Payment`||item.email||item.id;
          const sub=t==='project'?`Client: ${item.client}`:t==='invoice'?`${item.supplier} · ${fmtSGD(item.total)}`:t==='payment'?`${item.type} · ${fmtSGD(item.amount)}`:t==='staffClaim'?`${fmtSGD(item.amount)} · ${fmtDate(item.date)}`:item.email;
          return(
            <div key={item.id} style={{background:T.card,border:`1px solid ${urgent?'#FECACA':T.borderLight}`,borderRadius:14,padding:'14px 18px',boxShadow:T.shadow,display:'flex',alignItems:'center',gap:14}}>
              <div style={{width:38,height:38,borderRadius:12,background:`${color}12`,display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0}}><Icon size={16} style={{color}}/></div>
              <div style={{flex:1,minWidth:0}}>
                <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:2,flexWrap:'wrap'}}>
                  <span style={{fontSize:14,fontWeight:600,color:T.text,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{label}</span>
                  <Badge color={color} sm>{TL[t]}</Badge>
                </div>
                <div style={{fontSize:12,color:T.muted}}>{sub}</div>
                <div style={{fontSize:11,color:urgent?T.warning:T.dim,marginTop:3}}>
                  Deleted {fmtDate(item._deletedAt)} · {left>0?`${left} days left`:'Expires soon'}
                </div>
              </div>
              <div style={{display:'flex',gap:8,flexShrink:0}}>
                <Btn variant="secondary" size="sm" onClick={()=>onRestore(item)}><RotateCcw size={12}/>Restore</Btn>
                {isSuperAdmin&&(
                  <button onClick={()=>setConfirmPerm(item)}
                    style={{background:T.dangerLight,border:'none',borderRadius:10,padding:'6px 8px',cursor:'pointer',color:T.danger,display:'flex',alignItems:'center'}}>
                    <Trash2 size={13}/>
                  </button>
                )}
              </div>
            </div>
          );
        })}
      </div>

      {confirmPerm&&(
        <Modal title="Permanently Delete?" onClose={()=>setConfirmPerm(null)}>
          <div style={{textAlign:'center',padding:'8px 0'}}>
            <div style={{width:48,height:48,background:T.dangerLight,borderRadius:14,display:'flex',alignItems:'center',justifyContent:'center',margin:'0 auto 14px'}}><Trash2 size={20} style={{color:T.danger}}/></div>
            <div style={{fontSize:15,fontWeight:700,color:T.text,marginBottom:8}}>This cannot be undone</div>
            <div style={{fontSize:13,color:T.muted,lineHeight:1.6,marginBottom:20}}>
              <strong style={{color:T.text}}>{confirmPerm.name||confirmPerm.invoiceNo||confirmPerm.email||'This item'}</strong> will be permanently removed with no recovery possible.
            </div>
            <div style={{display:'flex',gap:10,justifyContent:'flex-end'}}>
              <Btn variant="secondary" onClick={()=>setConfirmPerm(null)}>Cancel</Btn>
              <Btn variant="danger" onClick={()=>{onPermanentDelete(confirmPerm.id);setConfirmPerm(null);}}>
                <Trash2 size={13}/>Delete Permanently
              </Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

function Dashboard({projects,invoices,payments,widgets=[],siteWorkers=[],onlinePresence=[],activeUserId,notices=[],setNotices,isAdmin}){
  const totRev = useMemo(()=>projects.reduce((s,p)=>s+p.contractAmount+(p.variationOrders||0),0),[projects]);
  const totExp = useMemo(()=>invoices.reduce((s,i)=>s+i.total,0),[invoices]);
  const totRecv = useMemo(()=>payments.filter(p=>p.status==='Received').reduce((s,p)=>s+p.amount,0),[payments]);
  const pendAmt = useMemo(()=>invoices.filter(i=>i.status==='Pending').reduce((s,i)=>s+i.total,0),[invoices]);
  const grossP = totRev - totExp;
  const margin = totRev>0?grossP/totRev*100:0;

  const catBreak = useMemo(()=>{
    const m={};
    invoices.forEach(i=>{m[i.category]=(m[i.category]||0)+i.total;});
    return Object.entries(m).sort((a,b)=>b[1]-a[1]);
  },[invoices]);

  const recent = [...invoices].sort((a,b)=>new Date(b.invoiceDate)-new Date(a.invoiceDate)).slice(0,8);

  const [noticeForm,setNoticeForm]=useState(null); // null=closed, {}=new, {id,...}=editing
  const [noticeText,setNoticeText]=useState('');
  const [noticePriority,setNoticePriority]=useState('normal');
  const [noticePinned,setNoticePinned]=useState(false);

  const PRIO_CLR={urgent:T.danger,high:T.warning,normal:T.info};
  const PRIO_BG={urgent:T.dangerLight,high:T.warningLight,normal:T.infoLight};
  const PRIO_LABEL={urgent:'Urgent',high:'High',normal:'Normal'};

  const sortedNotices=[...notices].sort((a,b)=>{
    if(a.pinned!==b.pinned) return a.pinned?-1:1;
    const po={urgent:0,high:1,normal:2};
    if((po[a.priority]||2)!==(po[b.priority]||2)) return (po[a.priority]||2)-(po[b.priority]||2);
    return new Date(b.postedAt)-new Date(a.postedAt);
  });

  const openNew=()=>{setNoticeText('');setNoticePriority('normal');setNoticePinned(false);setNoticeForm({});};
  const openEdit=(n)=>{setNoticeText(n.text);setNoticePriority(n.priority||'normal');setNoticePinned(!!n.pinned);setNoticeForm(n);};
  const saveNotice=()=>{
    if(!noticeText.trim()) return;
    let updated;
    if(noticeForm.id){
      updated=notices.map(n=>n.id===noticeForm.id?{...n,text:noticeText.trim(),priority:noticePriority,pinned:noticePinned,editedAt:new Date().toISOString()}:n);
    } else {
      updated=[{id:uid(),text:noticeText.trim(),priority:noticePriority,pinned:noticePinned,postedAt:new Date().toISOString()}, ...notices];
    }
    setNotices(updated);
    setNoticeForm(null);
  };
  const deleteNotice=(id)=>setNotices(notices.filter(n=>n.id!==id));
  const togglePin=(id)=>setNotices(notices.map(n=>n.id===id?{...n,pinned:!n.pinned}:n));

  return (
    <div style={{display:'flex',flexDirection:'column',gap:20}}>
      {/* ── Notice Board ── */}
      <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:16,overflow:'hidden',boxShadow:T.shadow}}>
        <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'14px 20px',borderBottom:`1px solid ${T.borderLight}`}}>
          <div style={{display:'flex',alignItems:'center',gap:10}}>
            <div style={{width:32,height:32,borderRadius:10,background:T.tanLight,display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0}}>
              <Bell size={15} style={{color:T.tan}}/>
            </div>
            <div>
              <div style={{fontSize:14,fontWeight:700,color:T.text}}>Notice Board</div>
              <div style={{fontSize:11,color:T.muted}}>{notices.length===0?'No notices posted':`${notices.length} notice${notices.length!==1?'s':''}`}</div>
            </div>
          </div>
          {isAdmin&&(
            <button onClick={openNew} style={{display:'flex',alignItems:'center',gap:6,background:T.text,color:T.card,border:'none',borderRadius:10,padding:'7px 14px',fontSize:12,fontWeight:600,cursor:'pointer',fontFamily:'inherit'}}>
              <Plus size={13}/>Post Notice
            </button>
          )}
        </div>

        {/* Post / edit form */}
        {noticeForm!==null&&(
          <div style={{padding:'16px 20px',borderBottom:`1px solid ${T.borderLight}`,background:T.bg}}>
            <textarea
              value={noticeText}
              onChange={e=>setNoticeText(e.target.value)}
              placeholder="Write your notice here…"
              rows={3}
              style={{width:'100%',background:T.card,border:`1px solid ${T.border}`,borderRadius:10,padding:'10px 13px',fontSize:13,color:T.text,outline:'none',resize:'vertical',fontFamily:'inherit',lineHeight:1.55,marginBottom:12}}
            />
            <div style={{display:'flex',alignItems:'center',gap:10,flexWrap:'wrap'}}>
              <select value={noticePriority} onChange={e=>setNoticePriority(e.target.value)}
                style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:8,padding:'6px 10px',fontSize:12,color:T.text,fontFamily:'inherit',cursor:'pointer'}}>
                <option value="normal">Normal</option>
                <option value="high">High Priority</option>
                <option value="urgent">Urgent</option>
              </select>
              <label style={{display:'flex',alignItems:'center',gap:6,fontSize:12,color:T.muted,cursor:'pointer'}}>
                <input type="checkbox" checked={noticePinned} onChange={e=>setNoticePinned(e.target.checked)} style={{accentColor:T.tan}}/>
                Pin to top
              </label>
              <div style={{marginLeft:'auto',display:'flex',gap:8}}>
                <button onClick={()=>setNoticeForm(null)} style={{background:'transparent',border:`1px solid ${T.border}`,borderRadius:8,padding:'6px 14px',fontSize:12,color:T.muted,cursor:'pointer',fontFamily:'inherit'}}>Cancel</button>
                <button onClick={saveNotice} disabled={!noticeText.trim()} style={{background:T.text,color:T.card,border:'none',borderRadius:8,padding:'6px 16px',fontSize:12,fontWeight:600,cursor:noticeText.trim()?'pointer':'not-allowed',opacity:noticeText.trim()?1:0.45,fontFamily:'inherit'}}>
                  {noticeForm.id?'Save Changes':'Post'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Notices list */}
        {sortedNotices.length===0?(
          <div style={{padding:'28px 20px',textAlign:'center',color:T.dim,fontSize:13}}>No notices yet{isAdmin?' — post one above':''}</div>
        ):(
          <div style={{display:'flex',flexDirection:'column'}}>
            {sortedNotices.map((n,idx)=>(
              <div key={n.id} style={{display:'flex',gap:12,padding:'13px 20px',borderBottom:idx<sortedNotices.length-1?`1px solid ${T.borderLight}`:'none',alignItems:'flex-start'}}>
                <div style={{width:4,flexShrink:0,borderRadius:4,alignSelf:'stretch',background:PRIO_CLR[n.priority||'normal'],minHeight:36}}/>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:4,flexWrap:'wrap'}}>
                    <Badge color={PRIO_CLR[n.priority||'normal']} sm>{PRIO_LABEL[n.priority||'normal']}</Badge>
                    {n.pinned&&<Badge color={T.tan} sm>Pinned</Badge>}
                    <span style={{fontSize:11,color:T.dim,marginLeft:'auto'}}>
                      {n.editedAt?`Edited ${fmtDate(n.editedAt)}`:`${fmtDate(n.postedAt)}`}
                    </span>
                  </div>
                  <div style={{fontSize:13,color:T.text,lineHeight:1.6,whiteSpace:'pre-wrap'}}>{n.text}</div>
                </div>
                {isAdmin&&(
                  <div style={{display:'flex',gap:4,flexShrink:0}}>
                    <button title={n.pinned?'Unpin':'Pin'} onClick={()=>togglePin(n.id)}
                      style={{background:'transparent',border:'none',cursor:'pointer',color:n.pinned?T.tan:T.dim,padding:4,borderRadius:6,display:'flex'}}>
                      <Star size={13} fill={n.pinned?T.tan:'none'}/>
                    </button>
                    <button title="Edit" onClick={()=>openEdit(n)}
                      style={{background:'transparent',border:'none',cursor:'pointer',color:T.muted,padding:4,borderRadius:6,display:'flex'}}>
                      <Edit3 size={13}/>
                    </button>
                    <button title="Delete" onClick={()=>deleteNotice(n.id)}
                      style={{background:'transparent',border:'none',cursor:'pointer',color:T.danger,padding:4,borderRadius:6,display:'flex'}}>
                      <Trash2 size={13}/>
                    </button>
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Worker document expiry alerts */}
      {(()=>{
        const alerts=[];
        siteWorkers.filter(w=>w.status==='Active').forEach(w=>{
          const wpd=daysUntil(w.workPassExpiry);
          if(wpd<=60) alerts.push({name:w.name,doc:`Work Pass (${w.workPassNo})`,days:wpd});
          (w.certificates||[]).forEach(c=>{
            if(!c.expiryDate)return;
            const cd=daysUntil(c.expiryDate);
            if(cd<=60) alerts.push({name:w.name,doc:c.name,days:cd});
          });
        });
        if(alerts.length===0)return null;
        return (
          <div style={{background:T.dangerLight,border:`1px solid ${T.danger}30`,borderRadius:14,padding:'14px 18px'}}>
            <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:10}}>
              <AlertTriangle size={15} style={{color:T.danger}}/>
              <span style={{fontSize:13,fontWeight:700,color:T.danger}}>{alerts.length} Worker Document{alerts.length>1?'s':''} Expiring Soon</span>
            </div>
            <div style={{display:'flex',flexWrap:'wrap',gap:8}}>
              {alerts.sort((a,b)=>a.days-b.days).map((al,i)=>(
                <div key={i} style={{background:T.card,borderRadius:8,padding:'6px 12px',
                  border:`1px solid ${expiryColor(al.days)}30`,display:'flex',alignItems:'center',gap:8}}>
                  <span style={{fontSize:12,fontWeight:600,color:T.text}}>{al.name}</span>
                  <span style={{fontSize:11,color:T.muted}}>.</span>
                  <span style={{fontSize:12,color:T.muted}}>{al.doc}</span>
                  <Badge color={expiryColor(al.days)} sm>{expiryLabel(al.days)}</Badge>
                </div>
              ))}
            </div>
          </div>
        );
      })()}

      {/* Overdue payment alerts */}
      {(()=>{
        const overdueProjects=projects.filter(p=>{
          if(p.archived||p.status==='Cancelled')return false;
          const contractVal=(p.contractAmount||0)+(p.variationOrders||0);
          const recv=payments.filter(py=>py.projectId===p.id&&py.status==='Received').reduce((s,py)=>s+py.amount,0);
          const outstanding=contractVal-recv;
          const endDate=p.endDate?new Date(p.endDate):null;
          const isOverdue=endDate&&endDate<new Date()&&outstanding>0;
          return isOverdue;
        });
        if(overdueProjects.length===0)return null;
        return(
          <div style={{background:'#FEF9EC',border:'1px solid #F5DFA0',borderRadius:14,padding:'14px 18px'}}>
            <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:10}}>
              <Bell size={15} style={{color:'#B7860A'}}/>
              <span style={{fontSize:13,fontWeight:700,color:'#78350F'}}>{overdueProjects.length} Project{overdueProjects.length>1?'s':''} Overdue — Outstanding Client Payment</span>
            </div>
            <div style={{display:'flex',flexDirection:'column',gap:6}}>
              {overdueProjects.map(p=>{
                const contractVal=(p.contractAmount||0)+(p.variationOrders||0);
                const recv=payments.filter(py=>py.projectId===p.id&&py.status==='Received').reduce((s,py)=>s+py.amount,0);
                return(
                  <div key={p.id} style={{background:T.card,borderRadius:8,padding:'8px 12px',border:'1px solid #F5DFA0',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                    <div>
                      <span style={{fontSize:13,fontWeight:600,color:T.text}}>{p.name}</span>
                      <span style={{fontSize:12,color:T.muted,marginLeft:8}}>{p.client}</span>
                    </div>
                    <span style={{fontSize:13,fontWeight:700,color:'#B7860A'}}>{fmtSGD(contractVal-recv)} outstanding</span>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })()}

      {/* Who's Online */}
      {onlinePresence.length>0&&(
        <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:16,padding:'14px 18px',boxShadow:T.shadow}}>
          <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:12}}>
            <div style={{display:'flex',alignItems:'center',gap:8}}>
              <div style={{width:8,height:8,borderRadius:'50%',background:T.success,boxShadow:`0 0 0 3px ${T.successLight}`}}/>
              <span style={{fontSize:13,fontWeight:700,color:T.text}}>Online Now</span>
              <span style={{fontSize:12,color:T.muted}}>{onlinePresence.length} {onlinePresence.length===1?'person':'people'}</span>
            </div>
            <span style={{fontSize:11,color:T.dim}}>Updates every 20s</span>
          </div>
          <div style={{display:'flex',flexWrap:'wrap',gap:10}}>
            {onlinePresence.map(p=>{
              const isMe=p.userId===activeUserId;
              const tabLabel=ALL_NAV.find(n=>n.id===p.currentTab)?.label||p.currentTab;
              const minsAgo=Math.floor((Date.now()-new Date(p.lastSeen).getTime())/60000);
              return (
                <div key={p.userId} style={{display:'flex',alignItems:'center',gap:10,
                  background:isMe?T.accentLight:T.bg,
                  border:`1px solid ${isMe?T.accent+'30':T.borderLight}`,
                  borderRadius:12,padding:'8px 12px',minWidth:160}}>
                  <div style={{position:'relative',flexShrink:0}}>
                    <Avatar photo={p.photo} name={p.name} size={32} color={ROLE_CLR[p.role]||T.accent}/>
                    <div style={{position:'absolute',bottom:-1,right:-1,width:9,height:9,borderRadius:'50%',
                      background:T.success,border:'2px solid #fff'}}/>
                  </div>
                  <div style={{minWidth:0}}>
                    <div style={{fontSize:13,fontWeight:600,color:T.text,display:'flex',alignItems:'center',gap:5}}>
                      {p.name}{isMe&&<span style={{fontSize:10,color:T.accent,fontWeight:700}}>(you)</span>}
                    </div>
                    <div style={{fontSize:10,color:ROLE_CLR[p.role]||T.muted,fontWeight:600}}>{ROLE_LABEL[p.role]}</div>
                    <div style={{fontSize:10,color:T.dim,marginTop:1}}>
                      {tabLabel&&`on ${tabLabel}`}
                      {minsAgo>0&&` · ${minsAgo}m ago`}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* ── Section header helper ───────────────────────────────── */}
      {widgets.includes('stats')&&(
      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(175px,1fr))',gap:12}}>
        {[
          {label:'Total Revenue',value:fmtSGD(totRev),sub:`${projects.length} projects`,Icon:TrendingUp,color:T.success},
          {label:'Total Expenses',value:fmtSGD(totExp),sub:`${invoices.length} invoices`,Icon:Receipt,color:T.danger},
          {label:'Gross Profit',value:fmtSGD(grossP),sub:`${margin.toFixed(1)}% margin`,Icon:DollarSign,color:'#8A6A3A'},
          {label:'Payments Collected',value:fmtSGD(totRecv),sub:`${fmtSGD(pendAmt)} outstanding`,Icon:CreditCard,color:T.info},
        ].map(({label,value,sub,Icon,color})=>(
          <div key={label} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:16,padding:'18px 20px',boxShadow:T.shadow}}>
            <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:12}}>
              <span style={{fontSize:12,fontWeight:500,color:T.muted}}>{label}</span>
              <div style={{width:32,height:32,borderRadius:9,background:color+'18',display:'flex',alignItems:'center',justifyContent:'center'}}>
                <Icon size={15} style={{color}}/>
              </div>
            </div>
            <div style={{fontSize:22,fontWeight:700,color:T.text,letterSpacing:'-0.03em'}}>{value}</div>
            <div style={{fontSize:11,color:T.dim,marginTop:4}}>{sub}</div>
          </div>
        ))}
      </div>
      )}

      {/* ══ FINANCIAL OVERVIEW ══════════════════════════════════════ */}
      {(widgets.includes('budget')||widgets.includes('catbreak'))&&(
        <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:-8}}>
          <div style={{width:3,height:14,background:T.tan,borderRadius:2}}/>
          <span style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:'uppercase',letterSpacing:'0.12em'}}>Financial Overview</span>
          <div style={{flex:1,height:1,background:T.borderLight}}/>
        </div>
      )}

      {(widgets.includes('budget')||widgets.includes('catbreak'))&&(()=>{
        const [chartRange,setChartRange]=useState('12m');
        const now=new Date();
        const numMonths=chartRange==='6m'?6:chartRange==='24m'?24:12;
        const months=[];
        for(let m=numMonths-1;m>=0;m--){
          const d=new Date(now.getFullYear(),now.getMonth()-m,1);
          const key=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
          const label=numMonths>12?d.toLocaleString('en-SG',{month:'short',year:'2-digit'}):d.toLocaleString('en-SG',{month:'short'});
          const rev=payments.filter(py=>py.status==='Received'&&py.date?.startsWith(key)).reduce((s,py)=>s+py.amount,0);
          const exp=invoices.filter(i=>i.invoiceDate?.startsWith(key)).reduce((s,i)=>s+i.total,0);
          months.push({label,rev,exp});
        }
        const maxVal=Math.max(...months.map(m=>Math.max(m.rev,m.exp)),1);
        const niceMax=(v)=>{const e=Math.pow(10,Math.floor(Math.log10(v||1)));const f=v/e;return(f<=1?1:f<=2?2:f<=5?5:10)*e;};
        const yMax=niceMax(maxVal);
        const yTicks=[0.25,0.5,0.75,1].map(f=>Math.round(yMax*f));
        const fmtY=(v)=>v>=1000?`$${(v/1000).toFixed(v>=10000?0:1)}k`:`$${v}`;
        const PAD_L=42,PAD_B=22,PAD_T=8,PAD_R=8,W=460,H=130;
        const chartW=W-PAD_L-PAD_R,chartH=H-PAD_B-PAD_T;
        const px=(i)=>PAD_L+(i/(months.length-1||1))*chartW;
        const py=(v)=>PAD_T+chartH-(v/yMax)*chartH;
        const revPath=months.map((m,i)=>`${i===0?'M':'L'}${px(i)},${py(m.rev)}`).join(' ');
        const expPath=months.map((m,i)=>`${i===0?'M':'L'}${px(i)},${py(m.exp)}`).join(' ');
        const revArea=`${revPath} L${px(months.length-1)},${PAD_T+chartH} L${PAD_L},${PAD_T+chartH} Z`;
        const expArea=`${expPath} L${px(months.length-1)},${PAD_T+chartH} L${PAD_L},${PAD_T+chartH} Z`;
        const labelStep=numMonths<=6?1:numMonths<=12?1:numMonths<=18?2:3;
        const statusCounts={Planning:0,Active:0,'In Progress':0,Completed:0,Cancelled:0};
        projects.forEach(p=>{if(statusCounts[p.status]!==undefined)statusCounts[p.status]++;else statusCounts['Active']++;});
        const donutData=Object.entries(statusCounts).filter(([,v])=>v>0).map(([k,v])=>({name:k,value:v}));
        const DONUT_COLORS={Planning:'#C4A882',Active:'#1A1A1A','In Progress':'#8A6A3A',Completed:'#2D7A4F',Cancelled:'#AEAEB2'};
        const dTotal=donutData.reduce((s,d)=>s+d.value,0)||1;
        let cumA=-Math.PI/2;
        const donutSegs=donutData.map(d=>{
          const angle=(d.value/dTotal)*Math.PI*2;
          const x1=Math.cos(cumA)*38+50,y1=Math.sin(cumA)*38+50;
          cumA+=angle;
          const x2=Math.cos(cumA)*38+50,y2=Math.sin(cumA)*38+50;
          return {...d,path:`M50,50 L${x1},${y1} A38,38 0 ${angle>Math.PI?1:0},1 ${x2},${y2} Z`};
        });
        return (
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(280px,1fr))',gap:14}}>
            {widgets.includes('budget')&&(
            <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:'22px 24px',boxShadow:T.shadow,gridColumn:'span 2'}}>
              <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:16,flexWrap:'wrap',gap:10}}>
                <div>
                  <div style={{fontSize:14,fontWeight:600,color:T.text}}>Revenue Overview</div>
                  <div style={{fontSize:11,color:T.muted,marginTop:2}}>Payments collected vs expenses</div>
                </div>
                <div style={{display:'flex',gap:6,alignItems:'center',flexWrap:'wrap'}}>
                  {[{id:'6m',label:'6M'},{id:'12m',label:'12M'},{id:'24m',label:'2Y'}].map(o=>(
                    <button key={o.id} onClick={()=>setChartRange(o.id)}
                      style={{padding:'4px 10px',borderRadius:7,border:`1px solid ${T.borderLight}`,
                        background:chartRange===o.id?T.text:'transparent',color:chartRange===o.id?'#F8F6F2':T.muted,
                        fontSize:11,fontWeight:600,cursor:'pointer',fontFamily:'inherit',transition:'all .12s'}}>
                      {o.label}
                    </button>
                  ))}
                  <div style={{marginLeft:8,display:'flex',gap:12}}>
                    {[{color:T.text,label:'Collected'},{color:T.tan,label:'Expenses'}].map(({color,label})=>(
                      <div key={label} style={{display:'flex',alignItems:'center',gap:5}}>
                        <div style={{width:14,height:2,background:color,borderRadius:2}}/>
                        <span style={{fontSize:10,color:T.muted}}>{label}</span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
              <svg viewBox={`0 0 ${W} ${H}`} style={{width:'100%',height:'auto',overflow:'visible'}}>
                <defs>
                  <linearGradient id="rg" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={T.text} stopOpacity="0.10"/><stop offset="100%" stopColor={T.text} stopOpacity="0"/></linearGradient>
                  <linearGradient id="eg" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={T.tan} stopOpacity="0.22"/><stop offset="100%" stopColor={T.tan} stopOpacity="0"/></linearGradient>
                </defs>
                {yTicks.map(v=>(<g key={v}><line x1={PAD_L} y1={py(v)} x2={W-PAD_R} y2={py(v)} stroke={T.borderLight} strokeWidth="1"/><text x={PAD_L-4} y={py(v)+3} textAnchor="end" fontSize="8" fill={T.dim}>{fmtY(v)}</text></g>))}
                <path d={expArea} fill="url(#eg)"/><path d={revArea} fill="url(#rg)"/>
                <path d={expPath} fill="none" stroke={T.tan} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/>
                <path d={revPath} fill="none" stroke={T.text} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/>
                {months.map((m,i)=>(<g key={i}>{i%labelStep===0&&<text x={px(i)} y={H-6} textAnchor="middle" fontSize="8" fill={T.dim}>{m.label}</text>}<circle cx={px(i)} cy={py(m.rev)} r="2.5" fill={T.card} stroke={T.text} strokeWidth="1.5"/><circle cx={px(i)} cy={py(m.exp)} r="2.5" fill={T.card} stroke={T.tan} strokeWidth="1.5"/></g>))}
              </svg>
              <div style={{marginTop:14,borderTop:`1px solid ${T.borderLight}`,paddingTop:12,display:'flex',flexDirection:'column',gap:7}}>
                <div style={{fontSize:11,fontWeight:600,color:T.muted,marginBottom:2}}>Budget utilization by project</div>
                {projects.filter(p=>p.status!=='Cancelled').slice(0,5).map(p=>{
                  const exp=invoices.filter(i=>i.projectId===p.id).reduce((s,i)=>s+i.total,0);
                  const rev=p.contractAmount+(p.variationOrders||0);
                  const pct=rev>0?Math.min(exp/rev*100,100):0;
                  const over=pct>85&&p.status!=='Completed';
                  return (<div key={p.id}><div style={{display:'flex',justifyContent:'space-between',fontSize:11,marginBottom:2}}><span style={{color:T.text,fontWeight:500,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:'60%'}}>{p.name}</span><span style={{color:over?T.danger:T.muted,fontWeight:600,flexShrink:0}}>{pct.toFixed(0)}%</span></div><div style={{height:3,background:T.bg,borderRadius:2,overflow:'hidden'}}><div style={{height:'100%',width:`${pct}%`,background:over?T.danger:T.tan,borderRadius:2,transition:'width .6s ease'}}/></div></div>);
                })}
                {projects.length===0&&<div style={{color:T.dim,fontSize:11,textAlign:'center',padding:'6px 0'}}>No projects yet</div>}
              </div>
            </div>
            )}
            {widgets.includes('catbreak')&&(
            <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:'22px 24px',boxShadow:T.shadow}}>
              <div style={{fontSize:14,fontWeight:600,color:T.text,marginBottom:4}}>Project Status</div>
              <div style={{fontSize:11,color:T.muted,marginBottom:14}}>{projects.length} total projects</div>
              <div style={{display:'flex',alignItems:'center',gap:14,marginBottom:16}}>
                <svg viewBox="0 0 100 100" width="80" height="80" style={{flexShrink:0}}>
                  {donutSegs.length>0?donutSegs.map((d,i)=><path key={i} d={d.path} fill={DONUT_COLORS[d.name]||T.dim}/>):<circle cx="50" cy="50" r="38" fill={T.borderLight}/>}
                  <circle cx="50" cy="50" r="24" fill={T.card}/><text x="50" y="54" textAnchor="middle" fontSize="14" fontWeight="700" fill={T.text}>{projects.length}</text>
                </svg>
                <div style={{display:'flex',flexDirection:'column',gap:5,flex:1}}>
                  {donutSegs.map(d=>(<div key={d.name} style={{display:'flex',alignItems:'center',justifyContent:'space-between'}}><div style={{display:'flex',alignItems:'center',gap:6}}><div style={{width:7,height:7,borderRadius:2,background:DONUT_COLORS[d.name]||T.dim,flexShrink:0}}/><span style={{fontSize:11,color:T.muted}}>{d.name}</span></div><span style={{fontSize:11,fontWeight:600,color:T.text}}>{d.value}</span></div>))}
                </div>
              </div>
              <div style={{borderTop:`1px solid ${T.borderLight}`,paddingTop:12}}>
                <div style={{fontSize:11,fontWeight:600,color:T.muted,marginBottom:8}}>Expenses by category</div>
                {catBreak.slice(0,5).map(([cat,amt])=>{
                  const pct=totExp>0?amt/totExp*100:0;
                  return (<div key={cat} style={{marginBottom:7}}><div style={{display:'flex',justifyContent:'space-between',fontSize:10,marginBottom:2}}><span style={{color:T.muted,display:'flex',alignItems:'center',gap:4}}><span style={{width:5,height:5,borderRadius:'50%',background:CAT_CLR[cat]||T.tan,display:'inline-block'}}/>{cat}</span><span style={{color:T.text,fontWeight:600}}>{fmtSGD(amt)}</span></div><div style={{height:3,background:T.bg,borderRadius:2,overflow:'hidden'}}><div style={{height:'100%',width:`${pct}%`,background:CAT_CLR[cat]||T.tan,borderRadius:2}}/></div></div>);
                })}
                {catBreak.length===0&&<div style={{color:T.dim,fontSize:11,textAlign:'center',padding:'6px 0'}}>No expenses yet</div>}
              </div>
            </div>
            )}
          </div>
        );
      })()}


      {/* ══ PROJECT HEALTH ══════════════════════════════════════════ */}
      {(widgets.includes('aging')||widgets.includes('margin')||widgets.includes('gantt')||widgets.includes('collection'))&&(
        <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:-8}}>
          <div style={{width:3,height:14,background:'#1A5FA8',borderRadius:2}}/>
          <span style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:'uppercase',letterSpacing:'0.12em'}}>Project Health</span>
          <div style={{flex:1,height:1,background:T.borderLight}}/>
        </div>
      )}

      {(widgets.includes('aging')||widgets.includes('margin'))&&(
      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(280px,1fr))',gap:14}}>
        {widgets.includes('aging')&&(()=>{
          const now=Date.now();
          const unpaid=invoices.filter(i=>i.status==='Pending'||i.status==='Overdue');
          const buckets=[
            {label:'0–30 days',min:0,max:30,color:'#2D7A4F'},
            {label:'31–60 days',min:31,max:60,color:T.tan},
            {label:'61–90 days',min:61,max:90,color:T.warning},
            {label:'90+ days',min:91,max:Infinity,color:T.danger},
          ];
          const aged=buckets.map(b=>{
            const items=unpaid.filter(i=>{const ageDays=Math.floor((now-new Date(i.invoiceDate))/864e5);return ageDays>=b.min&&ageDays<=b.max;});
            return {...b,count:items.length,total:items.reduce((s,i)=>s+i.total,0)};
          });
          const grandTotal=aged.reduce((s,b)=>s+b.total,0)||1;
          return (
            <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:'22px 24px',boxShadow:T.shadow}}>
              <div style={{fontSize:14,fontWeight:600,color:T.text,marginBottom:4}}>Invoice Aging</div>
              <div style={{fontSize:11,color:T.muted,marginBottom:16}}>{fmtSGD(grandTotal)} unpaid — by age</div>
              <div style={{display:'flex',flexDirection:'column',gap:10}}>
                {aged.map(b=>(<div key={b.label}><div style={{display:'flex',justifyContent:'space-between',fontSize:12,marginBottom:4}}><span style={{color:T.text,fontWeight:500,display:'flex',alignItems:'center',gap:6}}><span style={{width:8,height:8,borderRadius:2,background:b.color,display:'inline-block'}}/>{b.label}{b.count>0&&<span style={{fontSize:10,color:T.dim}}>({b.count})</span>}</span><span style={{fontWeight:700,color:b.total>0?b.color:T.dim}}>{b.total>0?fmtSGD(b.total):'—'}</span></div><div style={{height:6,background:T.bg,borderRadius:3,overflow:'hidden'}}><div style={{height:'100%',width:`${(b.total/grandTotal)*100}%`,background:b.color,borderRadius:3}}/></div></div>))}
              </div>
              {unpaid.length===0&&<div style={{textAlign:'center',color:T.success,fontSize:13,fontWeight:600,marginTop:8}}>✓ No outstanding invoices</div>}
            </div>
          );
        })()}
        {widgets.includes('margin')&&(()=>{
          const projData=projects.filter(p=>p.status!=='Cancelled').map(p=>{
            const rev=(p.contractAmount||0)+(p.variationOrders||0);
            const cost=invoices.filter(i=>i.projectId===p.id).reduce((s,i)=>s+i.total,0);
            const margin=rev>0?((rev-cost)/rev)*100:0;
            return {name:p.name,margin,rev,cost};
          }).sort((a,b)=>b.margin-a.margin).slice(0,8);
          const maxM=Math.max(...projData.map(p=>Math.abs(p.margin)),1);
          return (
            <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:'22px 24px',boxShadow:T.shadow}}>
              <div style={{fontSize:14,fontWeight:600,color:T.text,marginBottom:4}}>Profit Margin by Project</div>
              <div style={{fontSize:11,color:T.muted,marginBottom:16}}>Revenue minus cost ÷ revenue</div>
              <div style={{display:'flex',flexDirection:'column',gap:8}}>
                {projData.map(p=>{
                  const pct=Math.abs(p.margin);
                  const color=p.margin>=20?'#2D7A4F':p.margin>=0?T.tan:T.danger;
                  return (<div key={p.name}><div style={{display:'flex',justifyContent:'space-between',fontSize:11,marginBottom:3}}><span style={{color:T.text,fontWeight:500,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:'65%'}}>{p.name}</span><span style={{fontWeight:700,color,flexShrink:0}}>{p.margin>=0?'+':''}{p.margin.toFixed(1)}%</span></div><div style={{height:5,background:T.bg,borderRadius:3,overflow:'hidden'}}><div style={{height:'100%',width:`${(pct/maxM)*100}%`,background:color,borderRadius:3}}/></div></div>);
                })}
                {projData.length===0&&<div style={{textAlign:'center',color:T.dim,fontSize:12,padding:'12px 0'}}>No project data yet</div>}
              </div>
            </div>
          );
        })()}
      </div>
      )}

      {widgets.includes('gantt')&&(()=>{
        const active=projects.filter(p=>p.status!=='Cancelled'&&(p.startDate||p.endDate)).slice(0,10);
        if(!active.length) return null;
        const now=new Date();
        const allDates=[...active.flatMap(p=>[p.startDate,p.endDate].filter(Boolean)).map(d=>new Date(d))];
        const minD=new Date(Math.min(...allDates.map(d=>d.getTime()),now.getTime()));minD.setDate(1);
        const maxD=new Date(Math.max(...allDates.map(d=>d.getTime()),now.getTime()));maxD.setMonth(maxD.getMonth()+1,1);
        const span=maxD-minD||1;
        const toX=(d)=>((new Date(d)-minD)/span)*100;
        const nowX=((now-minD)/span)*100;
        const ST_C={Active:'#1A1A1A','In Progress':'#8A6A3A',Planning:T.tan,Completed:'#2D7A4F'};
        const marks=[];
        const md=new Date(minD);
        while(md<=maxD){marks.push({x:toX(md),label:md.toLocaleString('en-SG',{month:'short',year:'2-digit'})});md.setMonth(md.getMonth()+1);}
        return (
          <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:'22px 24px',boxShadow:T.shadow,overflowX:'auto'}}>
            <div style={{fontSize:14,fontWeight:600,color:T.text,marginBottom:4}}>Project Timeline</div>
            <div style={{fontSize:11,color:T.muted,marginBottom:14}}>Start to end — red line is today</div>
            <div style={{minWidth:360}}>
              <div style={{display:'flex',marginLeft:120,position:'relative',height:18,marginBottom:6}}>
                {marks.map((m,i)=><div key={i} style={{position:'absolute',left:`${m.x}%`,fontSize:9,color:T.dim,transform:'translateX(-50%)',whiteSpace:'nowrap'}}>{m.label}</div>)}
              </div>
              {active.map(p=>{
                const s=p.startDate?toX(p.startDate):0;
                const e=p.endDate?toX(p.endDate):nowX;
                const w=Math.max(1,e-s);
                const color=ST_C[p.status]||T.dim;
                return (
                  <div key={p.id} style={{display:'flex',alignItems:'center',gap:8,marginBottom:7}}>
                    <div style={{width:114,flexShrink:0,fontSize:11,fontWeight:500,color:T.text,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',textAlign:'right',paddingRight:6}}>{p.name}</div>
                    <div style={{flex:1,position:'relative',height:18,background:T.bg,borderRadius:4,overflow:'hidden'}}>
                      <div style={{position:'absolute',left:`${Math.min(s,99)}%`,width:`${Math.min(w,100-Math.min(s,99))}%`,height:'100%',background:color,borderRadius:4,opacity:0.85,minWidth:4}}/>
                      <div style={{position:'absolute',left:`${Math.min(Math.max(nowX,0),100)}%`,top:0,bottom:0,width:2,background:T.danger,opacity:0.9}}/>
                    </div>
                    <Badge color={color} sm>{p.status}</Badge>
                  </div>
                );
              })}
              <div style={{display:'flex',alignItems:'center',gap:6,marginTop:8,marginLeft:120}}>
                <div style={{width:12,height:2,background:T.danger,opacity:0.9,borderRadius:1}}/>
                <span style={{fontSize:10,color:T.dim}}>Today</span>
              </div>
            </div>
          </div>
        );
      })()}

      {widgets.includes('collection')&&(()=>{
        const projData=projects.filter(p=>p.status!=='Cancelled'&&p.contractAmount>0).map(p=>{
          const rev=(p.contractAmount||0)+(p.variationOrders||0);
          const collected=payments.filter(py=>py.projectId===p.id&&py.status==='Received').reduce((s,py)=>s+py.amount,0);
          const pct=Math.min(100,(collected/rev)*100);
          return {name:p.name,pct,collected,rev};
        }).sort((a,b)=>b.pct-a.pct).slice(0,8);
        return (
          <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:'22px 24px',boxShadow:T.shadow}}>
            <div style={{fontSize:14,fontWeight:600,color:T.text,marginBottom:4}}>Collection Rate by Project</div>
            <div style={{fontSize:11,color:T.muted,marginBottom:16}}>% of contract value collected</div>
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(260px,1fr))',gap:8}}>
              {projData.map(p=>{
                const color=p.pct>=100?'#2D7A4F':p.pct>=50?T.tan:T.danger;
                return (<div key={p.name}><div style={{display:'flex',justifyContent:'space-between',fontSize:11,marginBottom:3}}><span style={{color:T.text,fontWeight:500,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:'65%'}}>{p.name}</span><span style={{fontWeight:700,color,flexShrink:0}}>{p.pct.toFixed(0)}%</span></div><div style={{height:5,background:T.bg,borderRadius:3,overflow:'hidden',marginBottom:2}}><div style={{height:'100%',width:`${p.pct}%`,background:color,borderRadius:3}}/></div><div style={{fontSize:9,color:T.dim}}>{fmtSGD(p.collected)} of {fmtSGD(p.rev)}</div></div>);
              })}
              {projData.length===0&&<div style={{textAlign:'center',color:T.dim,fontSize:12,padding:'12px 0'}}>No project data yet</div>}
            </div>
          </div>
        );
      })()}

      {/* ══ OPERATIONS ══════════════════════════════════════════════ */}
      {(widgets.includes('suppliers')||widgets.includes('attendance'))&&(
        <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:-8}}>
          <div style={{width:3,height:14,background:'#7c3aed',borderRadius:2}}/>
          <span style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:'uppercase',letterSpacing:'0.12em'}}>Operations</span>
          <div style={{flex:1,height:1,background:T.borderLight}}/>
        </div>
      )}

      {widgets.includes('suppliers')&&(()=>{
        const supMap={};
        invoices.forEach(i=>{supMap[i.supplier]=(supMap[i.supplier]||0)+i.total;});
        const top=Object.entries(supMap).sort((a,b)=>b[1]-a[1]).slice(0,8);
        const maxSpend=top[0]?.[1]||1;
        const COLORS=['#1A1A1A','#C4A882','#2D7A4F','#8A6A3A','#9A6A00','#1A5FA8','#7c3aed','#B8B2A8'];
        return (
          <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:'22px 24px',boxShadow:T.shadow}}>
            <div style={{fontSize:14,fontWeight:600,color:T.text,marginBottom:4}}>Top Suppliers by Spend</div>
            <div style={{fontSize:11,color:T.muted,marginBottom:16}}>Total invoiced amount per supplier</div>
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(240px,1fr))',gap:10}}>
              {top.map(([sup,amt],i)=>(<div key={sup}><div style={{display:'flex',justifyContent:'space-between',fontSize:11,marginBottom:3}}><span style={{color:T.text,fontWeight:500,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:'65%',display:'flex',alignItems:'center',gap:6}}><span style={{width:8,height:8,borderRadius:2,background:COLORS[i]||T.dim,display:'inline-block',flexShrink:0}}/>{sup}</span><span style={{fontWeight:700,color:T.text,flexShrink:0}}>{fmtSGD(amt)}</span></div><div style={{height:5,background:T.bg,borderRadius:3,overflow:'hidden'}}><div style={{height:'100%',width:`${(amt/maxSpend)*100}%`,background:COLORS[i]||T.dim,borderRadius:3}}/></div></div>))}
              {top.length===0&&<div style={{textAlign:'center',color:T.dim,fontSize:12,padding:'12px 0'}}>No invoice data yet</div>}
            </div>
          </div>
        );
      })()}

      {widgets.includes('attendance')&&(()=>{
        if(!siteWorkers.length) return null;
        const today=new Date();
        const days=[];
        for(let d=29;d>=0;d--){const dt=new Date(today);dt.setDate(today.getDate()-d);days.push(dt.toISOString().slice(0,10));}
        const dayLabels=days.map(d=>{const dt=new Date(d);return dt.getDate()===1||d===days[0]?dt.toLocaleString('en-SG',{day:'numeric',month:'short'}):dt.getDate()%5===0?String(dt.getDate()):'';});
        const active=siteWorkers.filter(w=>w.status==='Active').slice(0,8);
        const getStatus=(worker,day)=>{const rec=attendance.filter(a=>a.workerId===worker.id&&a.date===day);if(!rec.length)return 'absent';return rec.some(a=>a.type==='out')?'full':'partial';};
        const CELL={full:'#2D7A4F',partial:T.tan,absent:T.bg};
        return (
          <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:'22px 24px',boxShadow:T.shadow}}>
            <div style={{fontSize:14,fontWeight:600,color:T.text,marginBottom:4}}>Worker Attendance</div>
            <div style={{fontSize:11,color:T.muted,marginBottom:14}}>Last 30 days — green full day · tan partial · grey absent</div>
            <div style={{overflowX:'auto'}}><div style={{minWidth:500}}>
              <div style={{display:'flex',marginLeft:90,marginBottom:4}}>{days.map((d,i)=>(<div key={d} style={{flex:1,fontSize:8,color:T.dim,textAlign:'center',minWidth:0,overflow:'hidden',whiteSpace:'nowrap'}}>{dayLabels[i]}</div>))}</div>
              {active.map(w=>(<div key={w.id} style={{display:'flex',alignItems:'center',marginBottom:4}}><div style={{width:84,fontSize:11,color:T.text,fontWeight:500,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',paddingRight:6,flexShrink:0}}>{w.name}</div><div style={{display:'flex',flex:1,gap:2}}>{days.map(day=>(<div key={day} style={{flex:1,aspectRatio:'1',borderRadius:2,background:CELL[getStatus(w,day)],border:`1px solid ${T.borderLight}`,minWidth:0}}/>))}</div></div>))}
              <div style={{display:'flex',gap:14,marginTop:10,marginLeft:90}}>{[{color:'#2D7A4F',l:'Full day'},{color:T.tan,l:'Partial'},{color:T.bg,l:'Absent'}].map(({color,l})=>(<div key={l} style={{display:'flex',alignItems:'center',gap:4}}><div style={{width:10,height:10,borderRadius:2,background:color,border:`1px solid ${T.borderLight}`}}/><span style={{fontSize:10,color:T.dim}}>{l}</span></div>))}</div>
            </div></div>
          </div>
        );
      })()}

      {/* ══ ACTIVITY ════════════════════════════════════════════════ */}
      {widgets.includes('recent')&&(
        <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:-8}}>
          <div style={{width:3,height:14,background:T.dim,borderRadius:2}}/>
          <span style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:'uppercase',letterSpacing:'0.12em'}}>Recent Activity</span>
          <div style={{flex:1,height:1,background:T.borderLight}}/>
        </div>
      )}

      {widgets.includes('recent')&&(
      <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow}}>
        <div style={{fontWeight:600,color:T.text,fontSize:14,marginBottom:16}}>Recent Supplier Invoices</div>
        <div style={{overflowX:'auto'}}>
          <table style={{width:'100%',fontSize:13,borderCollapse:'collapse'}}>
            <thead>
              <tr style={{color:T.dim,fontSize:11,fontWeight:700,textTransform:'uppercase',letterSpacing:'0.06em'}}>
                {['Supplier','Invoice #','Project','Category','Date','Total','Status'].map(h=>(
                  <th key={h} style={{textAlign:['Total','Status'].includes(h)?'right':'left',paddingBottom:10,paddingRight:14,whiteSpace:'nowrap'}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {recent.map(inv=>{
                const proj=projects.find(p=>p.id===inv.projectId);
                return (
                  <tr key={inv.id} style={{borderTop:`1px solid ${T.borderLight}`}}>
                    <td style={{padding:'10px 14px 10px 0',color:T.text,fontWeight:600}}>{inv.supplier}</td>
                    <td style={{padding:'10px 14px 10px 0',color:T.dim,fontFamily:'monospace',fontSize:12}}>{inv.invoiceNo}</td>
                    <td style={{padding:'10px 14px 10px 0',color:T.muted,fontSize:12,maxWidth:130,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{proj?.name||'—'}</td>
                    <td style={{padding:'10px 14px 10px 0'}}><Badge color={CAT_CLR[inv.category]||T.dim}>{inv.category}</Badge></td>
                    <td style={{padding:'10px 14px 10px 0',color:T.dim}}>{fmtDate(inv.invoiceDate)}</td>
                    <td style={{padding:'10px 14px 10px 0',textAlign:'right',color:T.text,fontWeight:700}}>{fmtSGD(inv.total)}</td>
                    <td style={{padding:'10px 0',textAlign:'right'}}><Badge color={ST_CLR[inv.status]}>{inv.status}</Badge></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {recent.length===0&&<div style={{color:T.dim,fontSize:13,textAlign:'center',padding:32}}>No invoices yet</div>}
        </div>
      </div>
      )}

      {widgets.length===0&&(
        <div style={{color:T.dim,fontSize:13,textAlign:'center',padding:48,background:T.card,borderRadius:14,border:`1px solid ${T.borderLight}`}}>
          No dashboard widgets enabled. Contact your admin to adjust visibility settings.
        </div>
      )}

    </div>
  );
}

function Projects({projects,setProjects,invoices,payments,isAdmin,onSoftDelete,onShowToast,users,acctSettings,logAction=()=>{},activeUser=null}){
  const [modal,setModal]=useState(null);
  const [search,setSearch]=useState('');
  const [sfilt,setSfilt]=useState('Active');
  const [confirmClose,setConfirmClose]=useState(null);
  const [invoiceModal,setInvoiceModal]=useState(null);
  const [invForm,setInvForm]=useState({paymentType:'Progress',amount:'',description:'',date:new Date().toISOString().slice(0,10)});
  const iif=k=>v=>setInvForm(p=>({...p,[k]:v}));
  const [deleteTarget,setDeleteTarget]=useState(null);
  const [quoteOcr,setQuoteOcr]=useState({loading:false,done:false,err:''});
  const [quoteCompressing,setQuoteCompressing]=useState(false);
  const [voOcr,setVoOcr]=useState({loading:false,done:false,err:''});
  const [docViewer,setDocViewer]=useState(null); // {title, data, filename}
  const quoteFileRef=useRef();
  const voFileRef=useRef();
  const blank={id:'',name:'',client:'',clientEmail:'',clientPhone:'',clientAddress:'',contractAmount:'',
    refNo:'',quotationFilename:'',quotationFile:null,
    voList:[],scopeItems:[],
    designer:'',pm:'',
    startDate:'',endDate:'',status:'Planning',variationOrders:'0',
    projectNumber:null,projectYear:null,
    designerCommMethod:'profit_pct',designerRate:'0',designerCommAmt:'0',
    pmCommMethod:'profit_pct',pmRate:'0',pmCommAmt:'0',
    commissionPaid:false,commissionPaidAt:null,
    projectType:'Residential',createdAt:new Date().toISOString(),archived:false};
  const [form,setForm]=useState(blank);
  const ff=k=>v=>setForm(p=>({...p,[k]:v}));

  const [handoverTarget,setHandoverTarget]=useState(null); // project being handed over
  const [handoverForm,setHandoverForm]=useState({
    handoverDate:new Date().toISOString().slice(0,10),
    defects:'',
    outstanding:'',
    finalPaymentAmount:'',
    finalPaymentMethod:'PayNow',
    finalPaymentReceiptNo:'',
  });

  const extractQuotation=async(file)=>{
    const apiKey=(acctSettings?.anthropicApiKey||'').trim();
    ff('quotationFilename')(file.name);
    // Compress BEFORE setting quotationFile so the save button never fires with a raw oversized file
    setQuoteCompressing(true);
    const fileData=await toB64(file);
    const compressed=await compressForSync(file);
    // Use the compressed version if available, otherwise fall back to raw (images already small)
    const dataUri=compressed||`data:${file.type};base64,${fileData}`;
    ff('quotationFile')(dataUri);
    setQuoteCompressing(false);

    if(!apiKey){
      setQuoteOcr({loading:false,done:false,err:'AI OCR not configured. File saved. Go to System → set Anthropic API key to enable auto-fill.'});
      return;
    }
    setQuoteOcr({loading:true,done:false,err:''});
    try{
      const isImg=file.type.startsWith('image/');
      const part=isImg
        ?{type:'image',source:{type:'base64',media_type:file.type,data:fileData}}
        :{type:'document',source:{type:'base64',media_type:'application/pdf',data:fileData}};
      const res=await fetch('https://api.anthropic.com/v1/messages',{
        method:'POST',
        headers:{'Content-Type':'application/json','x-api-key':apiKey,'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},
        body:JSON.stringify({
          model:'claude-opus-4-5',max_tokens:1500,
          messages:[{role:'user',content:[part,{type:'text',text:'Extract from this renovation quotation. Return ONLY valid JSON:\n{"clientName":"","clientPhone":"","clientAddress":"","refNo":"","contractAmount":0,"quotationDate":"YYYY-MM-DD","scopeItems":[{"category":"","description":"","amount":0}]}\nRules: contractAmount is the Final Total numeric value. quotationDate is the document date / quotation date — this is the project start date, look for DATE:, issue date, or any date at the top of the document, format as YYYY-MM-DD. scopeItems = each main work section with its total amount. category must be one of: Carpentry|Electrical|Plumbing|Painting|Lighting|Furniture|Appliances|Aircon|Miscellaneous|Preliminaries.'}]}]
        })
      });
      const data=await res.json();
      if(data.error) throw new Error(data.error.message);
      const parsed=JSON.parse((data.content||[]).map(c=>c.text||'').join('').replace(/```json|```/g,'').trim());
      setForm(prev=>({
        ...prev,
        client:parsed.clientName||prev.client,
        clientPhone:parsed.clientPhone||prev.clientPhone,
        clientAddress:parsed.clientAddress||prev.clientAddress,
        refNo:parsed.refNo||prev.refNo,
        name:prev.name||(parsed.refNo?`Job ${parsed.refNo}`:''),
        contractAmount:parsed.contractAmount?String(parsed.contractAmount):prev.contractAmount,
        startDate:parsed.quotationDate||prev.startDate,
        scopeItems:Array.isArray(parsed.scopeItems)?parsed.scopeItems:[],
      }));
      setQuoteOcr({loading:false,done:true,err:''});
    }catch(e){
      setQuoteOcr({loading:false,done:false,err:`Could not read quotation: ${e.message||'Please fill in manually.'}`});
    }
  };

  const extractVO=async(file)=>{
    const apiKey=(acctSettings?.anthropicApiKey||'').trim();
    // Read full file for OCR, compress separately for storage
    const fileData=await toB64(file);
    const rawDataUri=`data:${file.type};base64,${fileData}`;
    setVoOcr({loading:true,done:false,err:''});
    // Start compression in background
    const compressedPromise=compressForSync(file);
    const voEntry={id:uid(),filename:file.name,file:rawDataUri,amount:0,description:'',date:new Date().toISOString().slice(0,10),uploadedAt:new Date().toISOString(),confirmed:false};

    if(!apiKey){
      const compressed=await compressedPromise;
      setForm(prev=>({...prev,voList:[...(prev.voList||[]),{...voEntry,file:compressed||rawDataUri,confirmed:true}]}));
      setVoOcr({loading:false,done:false,err:'File saved. Set Anthropic API key for auto-fill.'});
      return;
    }
    try{
      const isImg=file.type.startsWith('image/');
      const part=isImg
        ?{type:'image',source:{type:'base64',media_type:file.type,data:fileData}}
        :{type:'document',source:{type:'base64',media_type:'application/pdf',data:fileData}};
      const res=await fetch('https://api.anthropic.com/v1/messages',{
        method:'POST',
        headers:{'Content-Type':'application/json','x-api-key':apiKey,'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},
        body:JSON.stringify({
          model:'claude-opus-4-5',max_tokens:600,
          messages:[{role:'user',content:[part,{type:'text',text:'Extract from this Variation Order document. Return ONLY valid JSON:\n{"amount":0,"description":"","voNo":"","date":"YYYY-MM-DD"}\nRules: amount=VO total amount (numeric). description=brief summary of additional works (max 120 chars). voNo=VO reference number if any. date=document date YYYY-MM-DD. Use 0 or empty string if not found.'}]}]
        })
      });
      const data=await res.json();
      if(data.error) throw new Error(data.error.message);
      const parsed=JSON.parse((data.content||[]).map(c=>c.text||'').join('').replace(/```json|```/g,'').trim());
      const compressed=await compressedPromise;
      const enriched={...voEntry,file:compressed||rawDataUri,amount:parsed.amount||0,description:parsed.description||'',date:parsed.date||voEntry.date,voNo:parsed.voNo||''};
      setForm(prev=>({...prev,voList:[...(prev.voList||[]),enriched]}));
      setVoOcr({loading:false,done:true,err:''});
    }catch(e){
      const compressed=await compressedPromise;
      setForm(prev=>({...prev,voList:[...(prev.voList||[]),{...voEntry,file:compressed||rawDataUri}]}));
      setVoOcr({loading:false,done:false,err:`Could not read VO: ${e.message||'File saved, please fill in manually.'}`});
    }
  };

  const canClose=(p)=>{
    const rev=p.contractAmount+(p.variationOrders||0);
    const recv=payments.filter(py=>py.projectId===p.id&&py.status==='Received').reduce((s,py)=>s+py.amount,0);
    const allPaid=invoices.filter(i=>i.projectId===p.id).every(i=>i.status==='Paid');
    const fullyCollected=recv>=rev;
    return allPaid&&fullyCollected;
  };

  const closeProject=(id)=>{
    const p=projects.find(p=>p.id===id);
    const upd=projects.map(p=>p.id===id?{...p,status:'Completed',archived:true,archivedAt:new Date().toISOString()}:p);
    setProjects(upd);saveProjects(upd);setConfirmClose(null);
    logAction('CLOSE_PROJECT',`Closed & archived project: ${p?.name||id}`);
  };

  const unarchive=(id)=>{
    const p=projects.find(pr=>pr.id===id);
    const upd=projects.map(pr=>pr.id===id?{...pr,archived:false,status:'In Progress'}:pr);
    setProjects(upd);saveProjects(upd);
    logAction('REOPEN_PROJECT',`Reopened project: ${p?.name||id}`);
  };


  const sorted=useMemo(()=>[...projects].sort((a,b)=>{
    // Sort by start date desc (most recent first), fall back to createdAt
    const da=new Date(a.startDate||a.createdAt||0);
    const db=new Date(b.startDate||b.createdAt||0);
    return db-da;
  }),[projects]);

  const filtered=useMemo(()=>sorted.filter(p=>{
    const matchSearch=(p.name+p.client).toLowerCase().includes(search.toLowerCase());
    const matchStatus=sfilt==='All'||(sfilt==='Active'&&!p.archived)||(sfilt==='Archived'&&p.archived)||(sfilt===p.status);
    return matchSearch&&matchStatus;
  }),[sorted,search,sfilt]);

  const save_=()=>{
    if(!form.name||!form.client)return;
    // Auto-compute variationOrders from voList
    const voTotal=(form.voList||[]).reduce((s,v)=>s+(parseFloat(v.amount)||0),0);
    const d={...form,
      contractAmount:parseFloat(form.contractAmount)||0,
      variationOrders:voTotal||parseFloat(form.variationOrders)||0,
      designerRate:parseFloat(form.designerRate)||0,
      pmRate:parseFloat(form.pmRate)||0,
      designerCommAmt:parseFloat(form.designerCommAmt)||0,
      pmCommAmt:parseFloat(form.pmCommAmt)||0,
      createdAt:form.createdAt||new Date().toISOString()};
    const isNew=modal==='new';
    const newId=uid();
    let finalD={...d};
    if(isNew){
      const yr=new Date().getFullYear();
      const sameYear=projects.filter(p=>(p.projectYear||(p.createdAt?new Date(p.createdAt).getFullYear():yr))===yr);
      const maxNum=sameYear.reduce((m,p)=>Math.max(m,p.projectNumber||0),0);
      finalD.projectNumber=maxNum+1;
      finalD.projectYear=yr;
    }
    const upd=isNew?[...projects,{...finalD,id:newId}]:projects.map(p=>p.id===d.id?d:p);
    setProjects(upd);saveProjects(upd);setModal(null);setQuoteCompressing(false);
    logAction(isNew?'CREATE_PROJECT':'EDIT_PROJECT',`${isNew?'Created':'Edited'} project: ${d.name} (Client: ${d.client}${d.contractAmount?', $'+Number(d.contractAmount).toLocaleString('en-SG'):''})`);
  };
  const del_=(id)=>{
    const proj=projects.find(p=>p.id===id);
    if(!proj)return;

    if(!isAdmin){return;}
    setDeleteTarget(proj);
  };

  const confirmSoftDelete=(proj)=>{
    const invCount=invoices.filter(i=>i.projectId===proj.id).length;
    const trashItem={...proj,_trashType:'project',_deletedAt:new Date().toISOString()};
    onSoftDelete(trashItem);
    const upd=projects.filter(p=>p.id!==proj.id);
    setProjects(upd);saveProjects(upd);
    setDeleteTarget(null);
    onShowToast(`"${proj.name}" moved to Trash`,()=>{
      setProjects(prev=>[...prev,proj]);saveProjects([...projects,proj]);
    });
  };

  const projToClose=confirmClose?projects.find(p=>p.id===confirmClose):null;
  const activeCount=projects.filter(p=>!p.archived).length;
  const archivedCount=projects.filter(p=>p.archived).length;

  return (
    <div style={{display:'flex',flexDirection:'column',gap:16}}>
      <div style={{display:'flex',gap:10,flexWrap:'wrap'}}>
        <div style={{flex:1,minWidth:180,position:'relative'}}>
          <Search size={13} style={{position:'absolute',left:11,top:'50%',transform:'translateY(-50%)',color:T.dim}}/>
          <input value={search} onChange={e=>setSearch(e.target.value)}
            placeholder="Search projects or clients..." style={{...iStyle,paddingLeft:33}}/>
        </div>
        <div style={{display:'flex',background:T.bg,borderRadius:10,border:`1px solid ${T.borderLight}`,overflow:'hidden'}}>
          {[
            {v:'Active',l:`Active (${activeCount})`},
            {v:'Archived',l:`Archived (${archivedCount})`},
            {v:'All',l:'All'},
          ].map(({v,l})=>(
            <button key={v} onClick={()=>setSfilt(v)}
              style={{padding:'7px 14px',border:'none',cursor:'pointer',fontFamily:'inherit',fontSize:12,fontWeight:600,
                background:sfilt===v?T.text:'transparent',color:sfilt===v?'#F8F6F2':T.muted,transition:'all .15s'}}>
              {l}
            </button>
          ))}
        </div>
        <select value={['All','Active','Archived'].includes(sfilt)?'':sfilt}
          onChange={e=>e.target.value&&setSfilt(e.target.value)}
          style={{...iStyle,width:'auto',minWidth:120}}>
          <option value="" style={{}}>By status…</option>
          {PROJ_STATUSES.map(s=><option key={s} value={s} style={{}}>{s}</option>)}
        </select>
        <Btn onClick={()=>{setForm({...blank,createdAt:new Date().toISOString()});setModal('new');}}><Plus size={13}/>New Project</Btn>
      </div>

      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(min(350px,100%),1fr))',gap:14}}>
        {filtered.map(p=>{
          const exp=invoices.filter(i=>i.projectId===p.id).reduce((s,i)=>s+i.total,0);
          const rev=p.contractAmount+(p.variationOrders||0);
          const recv=payments.filter(py=>py.projectId===p.id&&py.status==='Received').reduce((s,py)=>s+py.amount,0);
          const gross=rev-exp;
          const margin=rev>0?gross/rev*100:0;
          const pctBudget=rev>0?Math.min(exp/rev*100,100):0;
          const {dComm:designerComm,pmComm}=calcComm(p,invoices);
          const closeable=canClose(p)&&!p.archived&&isAdmin;
          return (
            <div key={p.id} style={{background:T.card,
              border:`1px solid ${p.archived?'rgba(255,255,255,0.04)':T.border}`,
              borderRadius:14,padding:20,opacity:p.archived?0.7:1,
              position:'relative',overflow:'hidden'}}>
              {p.archived&&(
                <div style={{position:'absolute',top:10,right:10,
                  background:'rgba(100,116,139,0.15)',color:T.dim,
                  fontSize:10,fontWeight:700,padding:'2px 8px',borderRadius:4,
                  textTransform:'uppercase',letterSpacing:'0.07em'}}>
                  Archived
                </div>
              )}
              {/* ── Top: name + info + edit/delete only ── */}
              <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:10}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:'flex',alignItems:'center',gap:7,marginBottom:3,flexWrap:'wrap'}}>
                    {p.projectNumber&&(
                      <span style={{background:T.text,color:'#F8F6F2',borderRadius:5,padding:'1px 7px',
                        fontSize:11,fontWeight:700,fontFamily:'monospace',flexShrink:0,letterSpacing:'0.04em'}}>
                        {String(p.projectYear||new Date().getFullYear()).slice(-2)}-{String(p.projectNumber).padStart(2,'0')}
                      </span>
                    )}
                    <div style={{fontSize:15,fontWeight:700,color:T.text}}>{p.name}</div>
                  </div>
                  <div style={{fontSize:12,color:T.muted}}>{p.client}{p.clientPhone&&` · ${p.clientPhone}`}</div>
                  {p.clientAddress&&<div style={{fontSize:11,color:T.dim,marginTop:2}}>{p.clientAddress}</div>}
                  {p.refNo&&<div style={{fontSize:10,color:T.dim,marginTop:2,fontFamily:'monospace'}}>Ref: {p.refNo}</div>}
                  {(p.designer||p.pm)&&<div style={{fontSize:11,color:T.dim,marginTop:2}}>{p.designer}{p.pm&&p.designer&&' · '}{p.pm}</div>}
                </div>
                {/* Edit / delete icons only */}
                <div style={{display:'flex',gap:4,alignItems:'center',flexShrink:0,marginLeft:8}}>
                  {!p.archived&&(
                    <button onClick={()=>{setForm({...p,contractAmount:String(p.contractAmount),variationOrders:String(p.variationOrders??0),designerRate:String(p.designerRate??0),pmRate:String(p.pmRate??0)});setModal('edit');}}
                      style={{background:'none',border:'none',cursor:'pointer',color:T.dim,display:'flex',padding:4,borderRadius:6}}>
                      <Edit3 size={13}/>
                    </button>
                  )}
                  {p.archived?(
                    <button onClick={()=>unarchive(p.id)}
                      style={{background:T.infoLight,border:'none',cursor:'pointer',color:T.info,
                        display:'flex',padding:'3px 8px',borderRadius:5,fontSize:11,fontWeight:600,alignItems:'center',gap:3}}>
                      Restore
                    </button>
                  ):isAdmin?(
                    <button onClick={()=>del_(p.id)}
                      style={{background:'none',border:'none',cursor:'pointer',color:T.dim,display:'flex',padding:4,borderRadius:6}}>
                      <Trash2 size={13}/>
                    </button>
                  ):null}
                </div>
              </div>
              {/* ── Action row: status + Cover + Handover ── */}
              <div style={{display:'flex',alignItems:'center',gap:7,marginBottom:12,flexWrap:'wrap'}}>
                <Badge color={ST_CLR[p.status]}>{p.status}</Badge>
                <button
                  title="Print filing cover page"
                  onClick={()=>{
                    const projInv=invoices.filter(i=>i.projectId===p.id);
                    const projPay=payments.filter(py=>py.projectId===p.id);
                    printDoc(buildCoverPageHTML(p,projInv,projPay,getCo(acctSettings),activeUser?.name),`Cover — ${p.name}`,true);
                  }}
                  style={{background:T.tanLight,border:`1px solid ${T.tan}`,cursor:'pointer',color:T.text,
                    display:'flex',padding:'4px 10px',borderRadius:6,fontSize:11,fontWeight:600,
                    alignItems:'center',gap:4,fontFamily:'inherit'}}>
                  <FileSpreadsheet size={11}/>Cover
                </button>
                {!p.archived&&p.status!=='Cancelled'&&(
                <button
                  title="Print handover form"
                  onClick={()=>{
                    setHandoverTarget(p);
                    setHandoverForm({
                      handoverDate:new Date().toISOString().slice(0,10),
                      defects:'',outstanding:'',
                      finalPaymentAmount:'',finalPaymentMethod:'PayNow',finalPaymentReceiptNo:'',
                    });
                  }}
                  style={{background:T.successLight,border:`1px solid ${T.success}55`,
                    cursor:'pointer',color:T.success,
                    display:'flex',padding:'4px 10px',borderRadius:6,fontSize:11,fontWeight:600,
                    alignItems:'center',gap:4,fontFamily:'inherit'}}>
                  <CheckCircle size={11}/>Handover
                </button>
                )}
              </div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:8,marginBottom:12}}>
                {[
                  {l:'Contract',v:fmtSGD(rev),c:T.text},
                  {l:'Expenses',v:fmtSGD(exp),c:T.danger},
                  {l:'Margin',v:`${margin.toFixed(1)}%`,c:margin>20?T.success:margin>0?T.accent:T.danger},
                ].map(({l,v,c})=>(
                  <div key={l} style={{background:T.bg,borderRadius:10,padding:'10px 12px'}}>
                    <div style={{fontSize:10,color:T.dim,marginBottom:3,textTransform:'uppercase',letterSpacing:'0.05em'}}>{l}</div>
                    <div style={{fontSize:13,fontWeight:700,color:c}}>{v}</div>
                  </div>
                ))}
              </div>
              {/* ── Payment collection bar ── */}
              {(()=>{
                const collected=payments.filter(py=>py.projectId===p.id&&py.status==='Received').reduce((s,py)=>s+py.amount,0);
                const pctCollected=rev>0?Math.min(collected/rev*100,100):0;
                const invCount=invoices.filter(i=>i.projectId===p.id).length;
                const paidInvCount=invoices.filter(i=>i.projectId===p.id&&i.status==='Paid').length;
                const daysLeft=p.endDate?Math.ceil((new Date(p.endDate)-new Date())/864e5):null;
                const overdue=daysLeft!==null&&daysLeft<0&&p.status!=='Completed'&&!p.archived;
                return (
                  <div style={{marginBottom:10}}>
                    {/* Progress bars */}
                    <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginBottom:8}}>
                      <div>
                        <div style={{display:'flex',justifyContent:'space-between',fontSize:9,color:T.dim,marginBottom:3,textTransform:'uppercase',letterSpacing:'0.06em'}}>
                          <span>Budget used</span><span style={{color:pctBudget>85?T.danger:T.muted,fontWeight:600}}>{pctBudget.toFixed(0)}%</span>
                        </div>
                        <div style={{height:4,background:T.bg,borderRadius:2,overflow:'hidden'}}>
                          <div style={{height:'100%',width:`${pctBudget}%`,background:pctBudget>85?T.danger:T.tan,borderRadius:2,transition:'width .5s'}}/>
                        </div>
                      </div>
                      <div>
                        <div style={{display:'flex',justifyContent:'space-between',fontSize:9,color:T.dim,marginBottom:3,textTransform:'uppercase',letterSpacing:'0.06em'}}>
                          <span>Collected</span><span style={{color:pctCollected>=100?T.success:T.muted,fontWeight:600}}>{pctCollected.toFixed(0)}%</span>
                        </div>
                        <div style={{height:4,background:T.bg,borderRadius:2,overflow:'hidden'}}>
                          <div style={{height:'100%',width:`${pctCollected}%`,background:pctCollected>=100?T.success:T.info,borderRadius:2,transition:'width .5s'}}/>
                        </div>
                      </div>
                    </div>
                    {/* Meta row */}
                    <div style={{display:'flex',gap:12,fontSize:11,color:T.dim,flexWrap:'wrap',alignItems:'center'}}>
                      <span>{fmtDate(p.startDate)}{p.endDate&&` → ${fmtDate(p.endDate)}`}</span>
                      {invCount>0&&<span style={{color:T.muted}}>{paidInvCount}/{invCount} inv paid</span>}
                      {daysLeft!==null&&!p.archived&&p.status!=='Completed'&&(
                        <span style={{color:overdue?T.danger:daysLeft<=14?T.warning:T.dim,fontWeight:overdue?700:400}}>
                          {overdue?`${Math.abs(daysLeft)}d overdue`:`${daysLeft}d left`}
                        </span>
                      )}
                      {designerComm>0&&<span>D: <span style={{color:T.accent,fontWeight:600}}>{fmtSGD(designerComm)}</span></span>}
                      {pmComm>0&&<span>PM: <span style={{color:T.info,fontWeight:600}}>{fmtSGD(pmComm)}</span></span>}
                    </div>
                  </div>
                );
              })()}

              {/* ── Bottom actions: all buttons in one tidy row ── */}
              <div style={{borderTop:`1px solid ${T.borderLight}`,paddingTop:10,display:'flex',gap:6,flexWrap:'wrap',alignItems:'center'}}>
                {!p.archived&&(
                  <button onClick={()=>{setInvoiceModal(p.id);setInvForm({paymentType:'Progress',amount:'',description:'',date:new Date().toISOString().slice(0,10)});}}
                    style={{background:T.accentLight,border:`1px solid ${T.borderLight}`,cursor:'pointer',color:T.text,
                      display:'flex',padding:'5px 10px',borderRadius:7,fontSize:11,fontWeight:700,fontFamily:'inherit',alignItems:'center',gap:4}}>
                    <Receipt size={11}/>Generate Invoice
                  </button>
                )}
                {p.quotationFile&&(
                  <>
                    <button type="button" onClick={()=>setDocViewer({title:`Quotation — ${p.quotationFilename||'document'}`,data:p.quotationFile,filename:p.quotationFilename||'quotation'})}
                      style={{background:T.accentLight,border:`1px solid ${T.borderLight}`,borderRadius:7,padding:'5px 10px',cursor:'pointer',fontSize:11,color:T.text,fontFamily:'inherit',fontWeight:600,display:'inline-flex',alignItems:'center',gap:4}}>
                      <ZoomIn size={10}/>View Quotation
                    </button>
                    <a href={p.quotationFile} download={p.quotationFilename||'quotation'}
                      style={{background:T.bg,border:`1px solid ${T.borderLight}`,borderRadius:7,padding:'5px 10px',cursor:'pointer',fontSize:11,color:T.muted,fontFamily:'inherit',fontWeight:600,display:'inline-flex',alignItems:'center',gap:4,textDecoration:'none'}}>
                      <Download size={10}/>Download
                    </a>
                  </>
                )}
                {(p.voList||[]).map((vo,i)=>(
                  vo.file?(
                    <button key={vo.id} type="button" onClick={()=>setDocViewer({title:`VO ${i+1} — ${vo.filename||'document'}`,data:vo.file,filename:vo.filename||`VO${i+1}`})}
                      style={{background:T.infoLight,border:`1px solid ${T.borderLight}`,borderRadius:7,padding:'5px 10px',cursor:'pointer',fontSize:11,color:T.info,fontFamily:'inherit',fontWeight:600,display:'inline-flex',alignItems:'center',gap:4}}>
                      <ZoomIn size={10}/>VO {i+1}{vo.voNo?` #${vo.voNo}`:''}
                    </button>
                  ):null
                ))}
                {closeable&&(
                  <button onClick={()=>setConfirmClose(p.id)}
                    style={{background:T.successLight,border:`1px solid rgba(45,122,79,0.3)`,
                      cursor:'pointer',color:T.success,display:'flex',padding:'5px 10px',borderRadius:7,
                      fontSize:11,fontWeight:700,fontFamily:'inherit',alignItems:'center',gap:4,marginLeft:'auto'}}>
                    <CheckCircle size={11}/>Close Project
                  </button>
                )}
              </div>
              {p.archived&&p.archivedAt&&(
                <div style={{marginTop:8,fontSize:11,color:T.dim}}>
                  Closed & archived {fmtDate(p.archivedAt)}
                </div>
              )}
            </div>
          );
        })}
        {filtered.length===0&&<div style={{color:T.dim,fontSize:13,padding:24,gridColumn:'1/-1',textAlign:'center'}}>No projects found</div>}
      </div>

      {/* Document viewer modal */}
      {docViewer&&(
        <Modal title={docViewer.title} onClose={()=>setDocViewer(null)} wide>
          <div style={{display:'flex',flexDirection:'column',gap:12}}>
            <div style={{textAlign:'right'}}>
              <a href={docViewer.data} download={docViewer.filename}
                style={{display:'inline-flex',alignItems:'center',gap:6,padding:'7px 16px',borderRadius:9,background:T.accent,color:'#fff',fontFamily:'inherit',fontSize:13,fontWeight:600,textDecoration:'none'}}>
                <Download size={14}/>Download {docViewer.filename}
              </a>
            </div>
            <Lightbox src={docViewer.data} title={docViewer.title} onClose={()=>setDocViewer(null)}/>
          </div>
        </Modal>
      )}

      {/* Close Project Confirmation */}
      {confirmClose&&projToClose&&(
        <Modal title="Close & Archive Project" onClose={()=>setConfirmClose(null)}>
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <div style={{background:T.successLight,border:'1px solid rgba(16,185,129,0.25)',
              borderRadius:12,padding:14}}>
              <div style={{fontSize:13,fontWeight:700,color:T.success,marginBottom:6}}>✓ Ready to close</div>
              <div style={{fontSize:12,color:T.muted}}>All invoices are paid and full payment has been collected from the client.</div>
            </div>
            <div style={{fontSize:13,color:T.text}}>
              You are about to close <strong style={{color:T.accent}}>{projToClose.name}</strong>.<br/><br/>
              The project will be marked <strong>Completed</strong> and moved to the Archived folder. It can be restored at any time.
            </div>
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:8,background:T.bg,borderRadius:12,padding:12}}>
              {[
                {l:'Contract Value',v:fmtSGD(projToClose.contractAmount+(projToClose.variationOrders||0))},
                {l:'Total Collected',v:fmtSGD(payments.filter(py=>py.projectId===projToClose.id&&py.status==='Received').reduce((s,py)=>s+py.amount,0))},
                {l:'Total Invoices',v:invoices.filter(i=>i.projectId===projToClose.id).length},
                {l:'Total Expenses',v:fmtSGD(invoices.filter(i=>i.projectId===projToClose.id).reduce((s,i)=>s+i.total,0))},
              ].map(({l,v})=>(
                <div key={l}>
                  <div style={{fontSize:11,color:T.dim}}>{l}</div>
                  <div style={{fontSize:13,fontWeight:700,color:T.text}}>{v}</div>
                </div>
              ))}
            </div>
          </div>
          <div style={{display:'flex',justifyContent:'flex-end',gap:10,marginTop:22}}>
            <Btn variant="secondary" onClick={()=>setConfirmClose(null)}>Cancel</Btn>
            <Btn onClick={()=>closeProject(confirmClose)}><CheckCircle size={13}/>Close & Archive</Btn>
          </div>
        </Modal>
      )}

      {(modal==='new'||modal==='edit')&&(
        <Modal title={modal==='new'?'Create New Project':'Edit Project'} onClose={()=>setModal(null)} wide>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:14}}>

            {/* Quotation Upload */}
            {(modal==='new'||modal==='edit')&&(
              <div style={{gridColumn:'1/-1'}}>
                <div style={{fontSize:12,fontWeight:600,color:T.muted,marginBottom:8,textTransform:'uppercase',letterSpacing:'0.06em'}}>
                  Signed Quotation
                </div>
                <input ref={quoteFileRef} type="file" accept="application/pdf,image/*"
                  style={{display:'none'}}
                  onChange={e=>{const f=e.target.files?.[0];if(f)extractQuotation(f);e.target.value='';}}/>
                <DropZone accept="application/pdf,image/*" onDrop={f=>extractQuotation(f)}>
                {form.quotationFile?(
                  <div style={{background:T.successLight,border:`1px solid ${T.success}30`,borderRadius:12,padding:'12px 16px'}}>
                    <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:12}}>
                      <div style={{display:'flex',alignItems:'center',gap:10}}>
                        <CheckCircle size={16} style={{color:T.success,flexShrink:0}}/>
                        <div>
                          <div style={{fontSize:13,fontWeight:700,color:T.success}}>
                            {quoteOcr.done?'Quotation read by AI':'Quotation saved'}
                          </div>
                          <div style={{fontSize:11,color:T.muted,marginTop:1}}>{form.quotationFilename}</div>
                        </div>
                      </div>
                      <div style={{display:'flex',gap:8,alignItems:'center',flexShrink:0}}>
                        <button type="button" onClick={()=>setDocViewer({title:`Quotation — ${form.quotationFilename}`,data:form.quotationFile,filename:form.quotationFilename})}
                          style={{background:T.accentLight,border:'none',borderRadius:8,padding:'5px 11px',cursor:'pointer',fontSize:12,color:T.accent,fontFamily:'inherit',fontWeight:600,display:'flex',alignItems:'center',gap:5}}>
                          <ZoomIn size={11}/>View
                        </button>
                        <a href={form.quotationFile} download={form.quotationFilename}
                          style={{background:T.bg,border:`1px solid ${T.borderLight}`,borderRadius:8,padding:'5px 11px',cursor:'pointer',fontSize:12,color:T.muted,fontFamily:'inherit',fontWeight:600,display:'flex',alignItems:'center',gap:5,textDecoration:'none'}}>
                          <Download size={11}/>Download
                        </a>
                        <button type="button" onClick={()=>{setForm(p=>({...p,quotationFile:null,quotationFilename:''}));setQuoteOcr({loading:false,done:false,err:''}); quoteFileRef.current?.click();}}
                          style={{background:'none',border:`1px solid ${T.borderLight}`,borderRadius:8,padding:'5px 10px',cursor:'pointer',fontSize:12,color:T.muted,fontFamily:'inherit'}}>
                          Replace
                        </button>
                      </div>
                    </div>
                    {quoteOcr.done&&<div style={{fontSize:11,color:T.muted,marginTop:6}}>Fields auto-filled below — review and adjust as needed</div>}
                    {quoteOcr.err&&<div style={{fontSize:11,color:T.warning,marginTop:6}}>{quoteOcr.err}</div>}
                  </div>
                ):(
                  <button type="button" onClick={()=>quoteFileRef.current?.click()} disabled={quoteOcr.loading}
                    style={{width:'100%',background:T.accentLight,border:`2px dashed ${T.accent}50`,borderRadius:14,padding:'18px',
                      cursor:quoteOcr.loading?'wait':'pointer',fontFamily:'inherit',display:'flex',alignItems:'center',justifyContent:'center',gap:12}}>
                    {quoteOcr.loading
                      ?<><Loader2 size={18} style={{color:T.accent,animation:'spin 1s linear infinite'}}/><div style={{textAlign:'left'}}><div style={{fontSize:13,fontWeight:700,color:T.accent}}>Reading quotation…</div><div style={{fontSize:11,color:T.muted,marginTop:2}}>AI extracting client info and scope items</div></div></>
                      :<><div style={{width:40,height:40,background:T.card,borderRadius:10,display:'flex',alignItems:'center',justifyContent:'center',boxShadow:T.shadow,flexShrink:0}}><Upload size={18} style={{color:T.accent}}/></div><div style={{textAlign:'left'}}><div style={{fontSize:13,fontWeight:700,color:T.accent}}>Upload Signed Quotation (PDF or image)</div><div style={{fontSize:11,color:T.muted,marginTop:2}}>AI auto-fills client name, address, contract amount and scope items</div></div></>
                    }
                  </button>
                )}
                {quoteOcr.err&&!form.quotationFile&&<div style={{marginTop:8,fontSize:12,color:T.danger}}>{quoteOcr.err}</div>}

                {/* Scope items preview */}
                {(form.scopeItems||[]).length>0&&(
                  <div style={{marginTop:10,background:T.bg,border:`1px solid ${T.borderLight}`,borderRadius:10,padding:'10px 14px'}}>
                    <div style={{fontSize:11,fontWeight:600,color:T.muted,marginBottom:8,textTransform:'uppercase',letterSpacing:'0.06em'}}>Scope Items from Quotation</div>
                    {(form.scopeItems||[]).map((s,i)=>(
                      <div key={i} style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'5px 0',borderBottom:i<form.scopeItems.length-1?`1px solid ${T.borderLight}`:'none'}}>
                        <div style={{display:'flex',alignItems:'center',gap:8}}>
                          <Badge color={CAT_CLR[s.category]||T.dim} sm>{s.category}</Badge>
                          <span style={{fontSize:12,color:T.muted,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:260}}>{s.description}</span>
                        </div>
                        <span style={{fontSize:13,fontWeight:700,color:T.text,flexShrink:0,marginLeft:8}}>{fmtSGD(s.amount)}</span>
                      </div>
                    ))}
                  </div>
                )}
                </DropZone>
                <div style={{height:1,background:T.borderLight,margin:'10px 0'}}/>

                {/* VO Upload Section */}
                <div style={{fontSize:12,fontWeight:600,color:T.muted,marginBottom:8,textTransform:'uppercase',letterSpacing:'0.06em'}}>
                  Variation Orders (VO)
                </div>
                <input ref={voFileRef} type="file" accept="application/pdf,image/*" style={{display:'none'}}
                  onChange={e=>{const f=e.target.files?.[0];if(f)extractVO(f);e.target.value='';}}/>

                {/* Existing VOs */}
                {(form.voList||[]).length>0&&(
                  <div style={{display:'flex',flexDirection:'column',gap:8,marginBottom:10}}>
                    {(form.voList||[]).map((vo,idx)=>(
                      <div key={vo.id} style={{background:T.bg,border:`1px solid ${T.borderLight}`,borderRadius:12,padding:'12px 14px'}}>
                        <div style={{display:'flex',alignItems:'flex-start',gap:10}}>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:4,flexWrap:'wrap'}}>
                              <span style={{fontSize:12,fontWeight:700,color:T.accent}}>VO {idx+1}</span>
                              {vo.voNo&&<span style={{fontSize:11,color:T.muted}}>#{vo.voNo}</span>}
                              <span style={{fontSize:11,color:T.dim}}>{fmtDate(vo.date)}</span>
                            </div>
                            {/* Editable fields for VO */}
                            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(180px,1fr))',gap:8,marginBottom:6}}>
                              <div>
                                <label style={{fontSize:10,color:T.dim,display:'block',marginBottom:3}}>VO Amount (S$)</label>
                                <input type="number" value={vo.amount||''} onChange={e=>{
                                  const v=parseFloat(e.target.value)||0;
                                  setForm(prev=>({...prev,voList:prev.voList.map((x,i)=>i===idx?{...x,amount:v}:x)}));
                                }} style={{...iStyle,padding:'6px 10px',fontSize:12}}/>
                              </div>
                              <div>
                                <label style={{fontSize:10,color:T.dim,display:'block',marginBottom:3}}>Description</label>
                                <input type="text" value={vo.description||''} onChange={e=>{
                                  const v=e.target.value;
                                  setForm(prev=>({...prev,voList:prev.voList.map((x,i)=>i===idx?{...x,description:v}:x)}));
                                }} placeholder="e.g. Additional electrical points" style={{...iStyle,padding:'6px 10px',fontSize:12}}/>
                              </div>
                            </div>
                            <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
                              {vo.file&&(
                                <>
                                  <button type="button" onClick={()=>setDocViewer({title:`VO ${idx+1} — ${vo.filename}`,data:vo.file,filename:vo.filename})}
                                    style={{background:T.accentLight,border:'none',borderRadius:7,padding:'3px 9px',cursor:'pointer',fontSize:11,color:T.accent,fontFamily:'inherit',fontWeight:600,display:'flex',alignItems:'center',gap:4}}>
                                    <ZoomIn size={10}/>View
                                  </button>
                                  <a href={vo.file} download={vo.filename}
                                    style={{background:T.bg,border:`1px solid ${T.borderLight}`,borderRadius:7,padding:'3px 9px',cursor:'pointer',fontSize:11,color:T.muted,fontFamily:'inherit',fontWeight:600,display:'flex',alignItems:'center',gap:4,textDecoration:'none'}}>
                                    <Download size={10}/>Download
                                  </a>
                                </>
                              )}
                              <button type="button" onClick={()=>setForm(prev=>({...prev,voList:prev.voList.filter((_,i)=>i!==idx)}))}
                                style={{background:'none',border:`1px solid ${T.danger}40`,borderRadius:7,padding:'3px 9px',cursor:'pointer',fontSize:11,color:T.danger,fontFamily:'inherit',display:'flex',alignItems:'center',gap:4}}>
                                <Trash2 size={10}/>Remove
                              </button>
                            </div>
                          </div>
                          <div style={{textAlign:'right',flexShrink:0}}>
                            <div style={{fontSize:15,fontWeight:800,color:T.accent}}>{fmtSGD(vo.amount||0)}</div>
                          </div>
                        </div>
                      </div>
                    ))}
                    <div style={{textAlign:'right',fontSize:12,fontWeight:700,color:T.accent,paddingRight:4}}>
                      Total VO: {fmtSGD((form.voList||[]).reduce((s,v)=>s+(v.amount||0),0))}
                    </div>
                  </div>
                )}

                {/* Add VO button */}
                <button type="button" onClick={()=>voFileRef.current?.click()} disabled={voOcr.loading}
                  style={{width:'100%',background:voOcr.loading?T.bg:'rgba(8,100,200,0.05)',border:`2px dashed ${T.info}40`,borderRadius:12,padding:'12px',
                    cursor:voOcr.loading?'wait':'pointer',fontFamily:'inherit',display:'flex',alignItems:'center',justifyContent:'center',gap:10,transition:'all 0.15s'}}>
                  {voOcr.loading
                    ?<><Loader2 size={15} style={{color:T.info,animation:'spin 1s linear infinite'}}/><span style={{fontSize:12,color:T.info,fontWeight:600}}>Reading VO with AI…</span></>
                    :<><Plus size={15} style={{color:T.info}}/><span style={{fontSize:12,color:T.info,fontWeight:600}}>Add Variation Order (upload PDF or image)</span></>
                  }
                </button>
                {voOcr.err&&<div style={{marginTop:6,fontSize:11,color:T.warning}}>{voOcr.err}</div>}
                <div style={{height:1,background:T.borderLight,margin:'10px 0'}}/>
              </div>
            )}

            {/* Ref / Quotation No */}
            <div>
              <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>
                File No. <span style={{fontSize:10,color:T.dim}}>(auto-assigned, editable)</span>
              </label>
              <div style={{display:'flex',gap:8,alignItems:'center'}}>
                <div style={{background:T.bg,border:`1px solid ${T.borderLight}`,borderRadius:10,
                  padding:'11px 14px',fontFamily:'monospace',fontSize:15,fontWeight:700,
                  color:form.projectNumber?T.text:T.dim,minWidth:60,textAlign:'center',flexShrink:0}}>
                  {form.projectNumber?String(form.projectNumber).padStart(2,'0'):'—'}
                </div>
                <input type="number" min="1" max="999"
                  value={form.projectNumber||''}
                  onChange={e=>ff('projectNumber')(parseInt(e.target.value)||null)}
                  placeholder="Auto"
                  style={{...iStyle,width:90,flexShrink:0}}/>
                <span style={{fontSize:11,color:T.dim,lineHeight:1.4}}>
                  {form.projectYear||new Date().getFullYear()} · resets each year
                </span>
              </div>
            </div>
            <Field label="Quotation / Ref No." value={form.refNo||''} onChange={ff('refNo')} placeholder="e.g. Q2604-NS-01"/>
            <Field label="Project Type" value={form.projectType||'Residential'} onChange={ff('projectType')} as="select" options={PROJ_TYPES.map(t=>({v:t,l:t}))}/>

            <div style={{gridColumn:'1/-1'}}><Field label="Project Name" value={form.name} onChange={ff('name')} placeholder="e.g. Choa Chu Kang Crescent Painting & Aircon"/></div>
            <Field label="Client Name" value={form.client} onChange={ff('client')} placeholder="e.g. Mr. & Mrs. Tan Wei Ming"/>
            <Field label="Client Phone" value={form.clientPhone||''} onChange={ff('clientPhone')} placeholder="+65 9123 4567"/>
            <div style={{gridColumn:'1/-1'}}><Field label="Job Site / Client Address" value={form.clientAddress||''} onChange={ff('clientAddress')} placeholder="e.g. Blk 657 Choa Chu Kang Crescent #06-45 S680657"/></div>
            <Field label="Client Email" type="email" value={form.clientEmail||''} onChange={ff('clientEmail')} placeholder="client@email.com"/>
            <Field label="Contract Amount (S$)" type="number" value={form.contractAmount} onChange={ff('contractAmount')} placeholder="0"/>
            {/* Designer — any active staff can do design work */}
            <div>
              <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Designer Assigned</label>
              <select value={form.designer} onChange={e=>ff('designer')(e.target.value)} style={{...iStyle,appearance:'none',
                backgroundImage:`url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8' fill='none'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%23AEAEB2' stroke-width='1.5' stroke-linecap='round'/%3E%3C/svg%3E")`,
                backgroundRepeat:'no-repeat',backgroundPosition:'right 13px center',paddingRight:36}}>
                <option value="">— Select designer —</option>
                {users.filter(u=>u.active&&u.role!=='site_worker').map(u=>(
                  <option key={u.id} value={u.name}>{u.name} ({ROLE_LABEL[u.role]})</option>
                ))}
              </select>
            </div>
            {/* Project Manager — any active staff can do PM work */}
            <div>
              <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Project Manager</label>
              <select value={form.pm} onChange={e=>ff('pm')(e.target.value)} style={{...iStyle,appearance:'none',
                backgroundImage:`url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8' fill='none'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%23AEAEB2' stroke-width='1.5' stroke-linecap='round'/%3E%3C/svg%3E")`,
                backgroundRepeat:'no-repeat',backgroundPosition:'right 13px center',paddingRight:36}}>
                <option value="">— Select PM —</option>
                {users.filter(u=>u.active&&u.role!=='site_worker').map(u=>(
                  <option key={u.id} value={u.name}>{u.name} ({ROLE_LABEL[u.role]})</option>
                ))}
              </select>
            </div>
            <Field label="Start Date" type="date" value={form.startDate} onChange={ff('startDate')}/>
            {/* Completion date only shown when editing an existing project */}
            {modal==='edit'&&<Field label="Completion Date" type="date" value={form.endDate} onChange={ff('endDate')}/>}
            <Field label="Status" value={form.status} onChange={ff('status')} as="select" options={PROJ_STATUSES.map(s=>({v:s,l:s}))}/>
            <Field label="Variation Orders (S$)" type="number" value={form.variationOrders} onChange={ff('variationOrders')} placeholder="0"/>
            <div style={{gridColumn:'1/-1',display:'flex',flexDirection:'column',gap:14}}>
              {/* Designer commission */}
              <div style={{background:T.bg,borderRadius:12,padding:16,border:`1px solid ${T.borderLight}`}}>
                <div style={{fontSize:12,fontWeight:600,color:T.muted,marginBottom:10}}>Designer Commission Method</div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:8,marginBottom:12}}>
                  {[
                    {v:'profit_pct',l:'% of Gross Profit',desc:'e.g. 5% of project profit'},
                    {v:'project_sum',l:'% of Project Sum',desc:'e.g. 3% of total contract'},
                    {v:'fixed_rate',l:'Fixed Rate',desc:'Fixed S$ amount'},
                  ].map(({v,l,desc})=>(
                    <button key={v} onClick={()=>ff('designerCommMethod')(v)}
                      style={{padding:'8px 12px',borderRadius:10,textAlign:'left',border:`1px solid ${form.designerCommMethod===v?T.accent:T.borderLight}`,
                        background:form.designerCommMethod===v?T.accentLight:'transparent',cursor:'pointer',fontFamily:'inherit',transition:'all 0.15s'}}>
                      <div style={{fontSize:12,fontWeight:600,color:form.designerCommMethod===v?T.accent:T.text}}>{l}</div>
                      <div style={{fontSize:10,color:T.dim,marginTop:1}}>{desc}</div>
                    </button>
                  ))}
                </div>
                {(form.designerCommMethod==='profit_pct'||form.designerCommMethod==='project_sum')
                  ? <Field label={`Designer % (0 = no commission) — of ${form.designerCommMethod==='profit_pct'?'gross profit':'total project sum'}`} type="number" value={form.designerRate} onChange={ff('designerRate')} placeholder="0"/>
                  : <Field label="Designer Fixed Amount (S$, 0 = no commission)" type="number" value={form.designerCommAmt} onChange={ff('designerCommAmt')} placeholder="0"/>}
              </div>
              {/* PM commission */}
              <div style={{background:T.bg,borderRadius:12,padding:16,border:`1px solid ${T.borderLight}`}}>
                <div style={{fontSize:12,fontWeight:600,color:T.muted,marginBottom:10}}>Project Manager Commission Method</div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:8,marginBottom:12}}>
                  {[
                    {v:'profit_pct',l:'% of Gross Profit',desc:'e.g. 3% of project profit'},
                    {v:'project_sum',l:'% of Project Sum',desc:'e.g. 2% of total contract'},
                    {v:'fixed_rate',l:'Fixed Rate',desc:'Fixed S$ amount'},
                  ].map(({v,l,desc})=>(
                    <button key={v} onClick={()=>ff('pmCommMethod')(v)}
                      style={{padding:'8px 12px',borderRadius:10,textAlign:'left',border:`1px solid ${form.pmCommMethod===v?T.info:T.borderLight}`,
                        background:form.pmCommMethod===v?T.infoLight:'transparent',cursor:'pointer',fontFamily:'inherit',transition:'all 0.15s'}}>
                      <div style={{fontSize:12,fontWeight:600,color:form.pmCommMethod===v?T.info:T.text}}>{l}</div>
                      <div style={{fontSize:10,color:T.dim,marginTop:1}}>{desc}</div>
                    </button>
                  ))}
                </div>
                {(form.pmCommMethod==='profit_pct'||form.pmCommMethod==='project_sum')
                  ? <Field label={`PM % (0 = no commission) — of ${form.pmCommMethod==='profit_pct'?'gross profit':'total project sum'}`} type="number" value={form.pmRate} onChange={ff('pmRate')} placeholder="0"/>
                  : <Field label="PM Fixed Amount (S$, 0 = no commission)" type="number" value={form.pmCommAmt} onChange={ff('pmCommAmt')} placeholder="0"/>}
              </div>
            </div>
          </div>
          <div style={{display:'flex',justifyContent:'flex-end',gap:10,marginTop:22}}>
            <Btn variant="secondary" onClick={()=>setModal(null)}>Cancel</Btn>
            <Btn onClick={save_} disabled={quoteCompressing} loading={quoteCompressing}>{quoteCompressing?'Compressing file…':modal==='new'?'Create Project':'Save Changes'}</Btn>
          </div>
        </Modal>
      )}

      {deleteTarget&&(
        <ConfirmDelete
          matchValue={deleteTarget.name}
          typeLabel="project"
          impact={`Also removes ${invoices.filter(i=>i.projectId===deleteTarget.id).length} invoice(s) and ${payments.filter(p=>p.projectId===deleteTarget.id).length} payment record(s) linked to this project.`}
          onConfirm={()=>confirmSoftDelete(deleteTarget)}
          onClose={()=>setDeleteTarget(null)}
        />
      )}

      {/* ── Handover Modal ── */}
      {handoverTarget&&(
        <Modal title={`Handover — ${handoverTarget.name}`} onClose={()=>setHandoverTarget(null)} wide>
          <div style={{display:'flex',flexDirection:'column',gap:16}}>
            {/* Info banner */}
            <div style={{background:'rgba(45,122,79,0.07)',border:`1px solid ${T.success}30`,borderRadius:12,padding:'12px 16px',display:'flex',gap:12,alignItems:'flex-start'}}>
              <CheckCircle size={16} style={{color:T.success,flexShrink:0,marginTop:2}}/>
              <div style={{fontSize:13,color:T.text,lineHeight:1.5}}>
                This generates a <strong>2-page handover document</strong> — Page 1 (client signs &amp; keeps) and Page 2 (defects record for your file). Fill in the details below before printing.
              </div>
            </div>

            {/* Handover date */}
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(180px,1fr))',gap:12}}>
              <div>
                <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Handover Date *</label>
                <input type="date" value={handoverForm.handoverDate}
                  onChange={e=>setHandoverForm(f=>({...f,handoverDate:e.target.value}))}
                  style={{...iStyle}}/>
              </div>
              <div style={{background:T.bg,borderRadius:10,padding:'10px 14px',border:`1px solid ${T.borderLight}`}}>
                <div style={{fontSize:10,color:T.dim,marginBottom:2}}>Warranty expires</div>
                <div style={{fontSize:13,fontWeight:700,color:T.danger}}>
                  {handoverForm.handoverDate?(()=>{const d=new Date(handoverForm.handoverDate);d.setFullYear(d.getFullYear()+1);return d.toLocaleDateString('en-SG',{day:'2-digit',month:'short',year:'numeric'});})():'—'}
                </div>
              </div>
            </div>

            {/* Final payment */}
            <div>
              <div style={{fontSize:12,fontWeight:600,color:T.text,marginBottom:10,display:'flex',alignItems:'center',gap:6}}>
                <div style={{width:4,height:14,background:T.tan,borderRadius:2}}/>Final Payment Details
              </div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(150px,1fr))',gap:10}}>
                <div>
                  <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Amount (S$)</label>
                  <input type="number" value={handoverForm.finalPaymentAmount}
                    onChange={e=>setHandoverForm(f=>({...f,finalPaymentAmount:e.target.value}))}
                    placeholder="0.00" style={{...iStyle}}/>
                </div>
                <div>
                  <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Payment Method</label>
                  <select value={handoverForm.finalPaymentMethod}
                    onChange={e=>setHandoverForm(f=>({...f,finalPaymentMethod:e.target.value}))}
                    style={{...iStyle,appearance:'none'}}>
                    {['PayNow','Bank Transfer','Cash','Cheque','Others'].map(m=><option key={m} value={m}>{m}</option>)}
                  </select>
                </div>
                <div>
                  <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Receipt No.</label>
                  <input value={handoverForm.finalPaymentReceiptNo}
                    onChange={e=>setHandoverForm(f=>({...f,finalPaymentReceiptNo:e.target.value}))}
                    placeholder="RCP-2026-xxxxx" style={{...iStyle}}/>
                </div>
              </div>
            </div>

            {/* Defects */}
            <div>
              <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>
                Defects / Remarks at Handover <span style={{color:T.dim,fontWeight:400}}>(leave blank for "Nil")</span>
              </label>
              <textarea value={handoverForm.defects}
                onChange={e=>setHandoverForm(f=>({...f,defects:e.target.value}))}
                placeholder="e.g. Minor paint touch-up on bedroom wall to be completed within 7 days."
                rows={3}
                style={{...iStyle,resize:'vertical',fontFamily:'inherit',lineHeight:1.5}}/>
            </div>

            {/* Outstanding works */}
            <div>
              <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>
                Outstanding Works <span style={{color:T.dim,fontWeight:400}}>(leave blank if none)</span>
              </label>
              <textarea value={handoverForm.outstanding}
                onChange={e=>setHandoverForm(f=>({...f,outstanding:e.target.value}))}
                placeholder="e.g. Supply of kitchen cabinet handles — to complete by 15 Jun 2026."
                rows={2}
                style={{...iStyle,resize:'vertical',fontFamily:'inherit',lineHeight:1.5}}/>
            </div>

            {/* Actions */}
            <div style={{display:'flex',justifyContent:'flex-end',gap:10,paddingTop:4}}>
              <Btn variant="secondary" onClick={()=>setHandoverTarget(null)}>Cancel</Btn>
              <Btn onClick={()=>{
                const projInv=invoices.filter(i=>i.projectId===handoverTarget.id);
                const finalPay=handoverForm.finalPaymentAmount?{
                  amount:parseFloat(handoverForm.finalPaymentAmount)||0,
                  method:handoverForm.finalPaymentMethod,
                  receiptNo:handoverForm.finalPaymentReceiptNo,
                }:null;
                printDoc(
                  buildHandoverHTML(handoverTarget,projInv,finalPay,handoverForm,getCo(acctSettings),activeUser?.name||''),
                  `Handover — ${handoverTarget.name}`,
                  true
                );
                setHandoverTarget(null);
              }}>
                <CheckCircle size={13}/>Print Handover Form
              </Btn>
            </div>
          </div>
        </Modal>
      )}

      {invoiceModal&&(()=>{
        const proj=projects.find(p=>p.id===invoiceModal);
        if(!proj)return null;
        const contractVal=(proj.contractAmount||0)+(proj.variationOrders||0);
        const recv=payments.filter(py=>py.projectId===proj.id&&py.status==='Received').reduce((s,py)=>s+py.amount,0);
        const outstanding=contractVal-recv;
        const invRef=`INV-${new Date().getFullYear()}-${String(Date.now()).slice(-5)}`;
        return (
          <Modal title={`Generate Invoice — ${proj.name}`} onClose={()=>setInvoiceModal(null)} wide>
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:20}}>
              <div style={{display:'flex',flexDirection:'column',gap:14}}>
                <div style={{background:T.bg,borderRadius:12,padding:14,border:`1px solid ${T.borderLight}`}}>
                  <div style={{fontSize:11,fontWeight:600,color:T.muted,marginBottom:8}}>Project Summary</div>
                  {[
                    {l:'Contract Value',v:fmtSGD(contractVal)},
                    {l:'Received to Date',v:fmtSGD(recv)},
                    {l:'Outstanding Balance',v:fmtSGD(outstanding),c:outstanding>0?T.danger:T.success},
                  ].map(({l,v,c})=>(
                    <div key={l} style={{display:'flex',justifyContent:'space-between',padding:'5px 0',borderBottom:`1px solid ${T.borderLight}`}}>
                      <span style={{fontSize:12,color:T.muted}}>{l}</span>
                      <span style={{fontSize:13,fontWeight:700,color:c||T.text}}>{v}</span>
                    </div>
                  ))}
                </div>
                <Field label="Payment Type" value={invForm.paymentType} onChange={iif('paymentType')} as="select"
                  options={[{v:'Deposit',l:'Deposit Payment'},{v:'Progress',l:'Progress Payment'},{v:'Final',l:'Final Payment'}]}/>
                <Field label="Invoice Date" type="date" value={invForm.date} onChange={iif('date')}/>
                <Field label="Amount to Invoice (S$)" type="number" value={invForm.amount} onChange={iif('amount')} placeholder="0"/>
                <div>
                  <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Description (optional)</label>
                  <textarea value={invForm.description} onChange={e=>iif('description')(e.target.value)}
                    placeholder={`${invForm.paymentType} payment for renovation and fitting-out works at ${proj.name}`}
                    rows={3} style={{...iStyle,resize:'vertical'}}/>
                </div>
              </div>
              <div style={{background:T.bg,borderRadius:12,padding:16,border:`1px solid ${T.borderLight}`}}>
                <div style={{fontSize:11,fontWeight:600,color:T.muted,marginBottom:12}}>Invoice Preview</div>
                <div style={{fontSize:13,color:T.text,lineHeight:1.7}}>
                  <strong style={{color:T.accent}}>{acctSettings?.companyName||''}</strong><br/>
                  <span style={{fontSize:11,color:T.dim}}>Invoice {invRef}</span><br/><br/>
                  <strong>Bill to:</strong> {proj.client}<br/>
                  <strong>Project:</strong> {proj.name}<br/>
                  <strong>Type:</strong> {invForm.paymentType} Payment<br/>
                  <strong>Amount:</strong> <span style={{color:T.accent,fontWeight:700}}>{invForm.amount?fmtSGD(parseFloat(invForm.amount)):'—'}</span><br/>
                  <strong>Balance after:</strong> {invForm.amount?fmtSGD(Math.max(0,outstanding-parseFloat(invForm.amount||0))):'—'}<br/><br/>
                  <span style={{fontSize:11,color:T.dim}}>Computer-generated . No signature required</span>
                </div>
              </div>
            </div>
            <div style={{display:'flex',justifyContent:'flex-end',gap:10,marginTop:22}}>
              <Btn variant="secondary" onClick={()=>setInvoiceModal(null)}>Cancel</Btn>
              <Btn disabled={!invForm.amount||parseFloat(invForm.amount)<=0}
                onClick={()=>{
                  printDoc(buildClientInvoiceHTML(proj,payments,invRef,invForm.date,
                    invForm.description,parseFloat(invForm.amount),invForm.paymentType,getCo(acctSettings)),
                    `Invoice ${invRef}`);
                  setInvoiceModal(null);
                }}>
                <Receipt size={13}/>Generate & Print
              </Btn>
            </div>
          </Modal>
        );
      })()}
    </div>
  );
}

function Invoices({invoices,setInvoices,projects,isAdmin,onSoftDelete,onShowToast,invoiceBatches=[],setInvoiceBatches,acctSettings={},logAction=()=>{}}){
  const [modal,setModal]=useState(false);
  const [search,setSearch]=useState('');
  const [pfilt,setPfilt]=useState('All');
  const [cfilt,setCfilt]=useState('All');
  const [deleteTarget,setDeleteTarget]=useState(null);
  const [err,setErr]=useState('');
  const blank={projectId:'',supplier:'',invoiceNo:'',invoiceDate:'',
    subtotal:'',gst:'',total:'',category:'Carpentry',status:'Pending',
    paymentRecords:[]};  // array of {id,amount,date,method,reference,notes}
  const [form,setForm]=useState(blank);
  const [ocr,setOcr]=useState({loading:false,done:false,err:''});
  const [dup,setDup]=useState('');
  const [preview,setPreview]=useState(null);
  const [payModal,setPayModal]=useState(null);
  const [selected,setSelected]=useState(new Set()); // selected invoice ids for grouping
  const [batchModal,setBatchModal]=useState(false); // create batch modal
  const [batchName,setBatchName]=useState('');
  const [batchPayModal,setBatchPayModal]=useState(null);
  const [batchPayForm,setBatchPayForm]=useState({amount:'',date:new Date().toISOString().slice(0,10),method:'paynow',reference:'',notes:'',proofImage:null});
  const [hidePaidBatches,setHidePaidBatches]=useState(true);
  const [payForm,setPayForm]=useState({amount:'',date:new Date().toISOString().slice(0,10),method:'paynow',reference:'',notes:'',proofImage:null});
  const [proofOcrLoading,setProofOcrLoading]=useState(false);
  const payProofRef=useRef();
  const batchProofRef=useRef();

  // Shared OCR for payment proof screenshots (batch + individual)
  const extractProof=async(file, setFormFn)=>{
    const apiKey=(acctSettings?.anthropicApiKey||'').trim();
    // Compress image for Firebase sync, show original immediately
    const reader=new FileReader();
    reader.onload=ev=>setFormFn(p=>({...p,proofImage:ev.target.result}));
    reader.readAsDataURL(file);
    compressForSync(file).then(c=>{if(c)setFormFn(p=>({...p,proofImage:c}));});
    if(!apiKey) return;
    setProofOcrLoading(true);
    try{
      const b64=await toB64(file);
      const isPDF=file.type==='application/pdf';
      const part=isPDF
        ?{type:'document',source:{type:'base64',media_type:'application/pdf',data:b64}}
        :{type:'image',source:{type:'base64',media_type:file.type,data:b64}};
      const res=await fetch('https://api.anthropic.com/v1/messages',{
        method:'POST',
        headers:{'Content-Type':'application/json','x-api-key':apiKey,
          'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},
        body:JSON.stringify({
          model:'claude-opus-4-5',max_tokens:300,
          messages:[{role:'user',content:[part,{type:'text',
            text:'Extract from this payment screenshot/receipt. Return ONLY valid JSON:\n{"date":"YYYY-MM-DD","method":"paynow|bank_transfer|cash","reference":"","amount":0}\nRules: date=transaction date. method=paynow if PayNow/QR/UEN, bank_transfer if bank/FAST/GIRO/transfer, cash if cash. reference=PayNow number/UEN/account/transaction ref. amount=total amount paid (numeric, 0 if unclear). Empty string if not found.'}]}]
        })
      });
      const data=await res.json();
      if(data.error) throw new Error(data.error.message);
      const txt=(data.content||[]).map(c=>c.text||'').join('').replace(/```json|```/g,'').trim();
      const parsed=JSON.parse(txt);
      setFormFn(prev=>({
        ...prev,
        date:parsed.date||prev.date,
        method:parsed.method||prev.method,
        reference:parsed.reference||prev.reference,
        amount:parsed.amount&&!prev.amount?String(parsed.amount):prev.amount,
      }));
    }catch(e){console.warn('Proof OCR:',e);}
    finally{setProofOcrLoading(false);}
  };

  // Look up supplier's last used payment details from all invoices + batches
  const getSupplierHistory=(supplier)=>{
    const allRecords=[];
    // From individual invoice payment records
    invoices.filter(i=>i.supplier===supplier).forEach(inv=>{
      (inv.paymentRecords||[]).forEach(r=>allRecords.push(r));
    });
    // From batch payment records where supplier matches
    (invoiceBatches||[]).filter(b=>b.supplier===supplier).forEach(batch=>{
      (batch.paymentRecords||[]).forEach(r=>allRecords.push(r));
    });
    if(!allRecords.length) return null;
    // Return most recent record
    return allRecords.sort((a,b)=>b.recordedAt.localeCompare(a.recordedAt))[0];
  };
  const fileRef=useRef();
  const camRef=useRef();
  const proofRef=useRef();
  const [proofTarget,setProofTarget]=useState(null);
  const [lightbox,setLightbox]=useState(null);

  const ff=k=>v=>{
    setForm(prev=>{
      const n={...prev,[k]:v};
      if(k==='subtotal'){
        const s=parseFloat(v)||0;
        const g=parseFloat(prev.gst)||0;

        n.total=(s+g).toFixed(2);
      }
      if(k==='gst'){
        const s=parseFloat(prev.subtotal)||0;
        const g=parseFloat(v)||0;
        n.total=(s+g).toFixed(2);
      }
      return n;
    });
    if(k==='invoiceNo'){
      const d=invoices.find(i=>i.invoiceNo.toLowerCase()===v.toLowerCase()&&v.length>2);
      setDup(d?`Duplicate detected: "${v}" from ${d.supplier}`:'');
    }
  };

  const doOCR=async(file)=>{
    const apiKey=(acctSettings?.anthropicApiKey||'').trim();
    // Show original immediately for display while compressing in background
    const rdr=new FileReader();rdr.onload=e=>setPreview(e.target.result);rdr.readAsDataURL(file);
    // Compress for Firebase sync (replaces preview once done)
    compressForSync(file).then(compressed=>{if(compressed)setPreview(compressed);});
    if(!apiKey){
      setOcr({loading:false,done:false,err:'AI OCR not configured. Go to System → set your Anthropic API key, or fill in manually.'});
      return;
    }
    setOcr({loading:true,done:false,err:''});
    try{
      const data64=await toB64(file);
      const isPDF=file.type==='application/pdf';
      const part=isPDF
        ?{type:'document',source:{type:'base64',media_type:'application/pdf',data:data64}}
        :{type:'image',source:{type:'base64',media_type:file.type,data:data64}};
      const res=await fetch('https://api.anthropic.com/v1/messages',{
        method:'POST',
        headers:{'Content-Type':'application/json','x-api-key':apiKey,'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},
        body:JSON.stringify({
          model:'claude-opus-4-5',max_tokens:800,
          messages:[{role:'user',content:[part,{type:'text',text:'Extract invoice details. Return ONLY valid JSON, no markdown fences, no extra text:\n{"supplier":"company name","invoiceNo":"invoice/ref number","invoiceDate":"YYYY-MM-DD","subtotal":0,"gst":0,"total":0,"category":"one of: Carpentry|Electrical|Plumbing|Painting|Lighting|Furniture|Appliances|Miscellaneous"}\nRules: Singapore GST=9%. Infer category from supplier name or line items. If date missing use empty string. Compute missing fields from available ones.'}]}]
        })
      });
      const d=await res.json();
      if(d.error) throw new Error(d.error.message);
      const txt=d.content?.filter(c=>c.type==='text').map(c=>c.text).join('');
      const parsed=JSON.parse(txt.replace(/```json|```/g,'').trim());
      setForm(p=>({...p,
        supplier:parsed.supplier||p.supplier,
        invoiceNo:parsed.invoiceNo||p.invoiceNo,
        invoiceDate:parsed.invoiceDate||p.invoiceDate,
        subtotal:parsed.subtotal?String(parsed.subtotal):p.subtotal,
        gst:parsed.gst?String(parsed.gst):p.gst,
        total:parsed.total?String(parsed.total):p.total,
        category:parsed.category||p.category,
      }));
      if(parsed.invoiceNo){
        const d2=invoices.find(i=>i.invoiceNo.toLowerCase()===parsed.invoiceNo.toLowerCase());
        setDup(d2?`Duplicate detected: "${parsed.invoiceNo}" from ${d2.supplier}`:'');
      }
      setOcr({loading:false,done:true,err:''});
    }catch(e){setOcr({loading:false,done:false,err:`OCR failed: ${e.message||'Please enter details manually.'}`});}
  };

  const save_=()=>{
    if(!form.projectId||!form.supplier||!form.invoiceNo)return;
    if(!preview){setErr('Please upload or photograph the invoice document before saving.');return;}
    if(dup&&!window.confirm('Duplicate invoice number detected. Save anyway?'))return;
    const inv={...form,id:uid(),
      proofImage:preview||null,
      subtotal:parseFloat(form.subtotal)||0,
      gst:parseFloat(form.gst)||0,
      total:parseFloat(form.total)||0};
    const upd=[...invoices,inv];
    setInvoices(upd);saveInvoices(upd);
    const projName=projects.find(p=>p.id===inv.projectId)?.name||'Unknown project';
    logAction('CREATE_INVOICE',`Captured invoice ${inv.invoiceNo||''} from ${inv.supplier} — ${inv.category} — $${inv.total.toLocaleString('en-SG',{minimumFractionDigits:2})} (${projName})`);
    setModal(false);setPreview(null);setOcr({loading:false,done:false,err:''});setDup('');
  };

  // Auto-calculate invoice status from paymentRecords
  const calcInvStatus=(inv,records)=>{
    const total=parseFloat(inv.total)||0;
    if(!records||records.length===0) return 'Pending';
    const paid=records.reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
    if(paid<=0) return 'Pending';
    if(paid>=total) return 'Paid';
    return 'Partial';
  };

  const updStatus=(id,s)=>{
    const u=invoices.map(i=>i.id===id?{...i,status:s}:i);setInvoices(u);saveInvoices(u);
  };

  const recordPayment=()=>{
    const inv=invoices.find(i=>i.id===payModal);
    if(!inv||!payForm.amount||parseFloat(payForm.amount)<=0) return;
    const newRecord={id:uid(),amount:parseFloat(payForm.amount),date:payForm.date,
      method:payForm.method,reference:payForm.reference||'',notes:payForm.notes||'',
      proofImage:payForm.proofImage||null,
      recordedAt:new Date().toISOString()};
    const records=[...(inv.paymentRecords||[]),newRecord];
    const newStatus=calcInvStatus(inv,records);
    const upd=invoices.map(i=>i.id===payModal?{...i,paymentRecords:records,status:newStatus}:i);
    setInvoices(upd);saveInvoices(upd);
    const inv2=invoices.find(i=>i.id===payModal);
    logAction('PAY_INVOICE',`Recorded $${parseFloat(payForm.amount).toLocaleString('en-SG',{minimumFractionDigits:2})} payment for invoice ${inv2?.invoiceNo||''} from ${inv2?.supplier||''} (${newStatus})`);
    setPayModal(null);
    setPayForm({amount:'',date:new Date().toISOString().slice(0,10),method:'paynow',reference:'',notes:'',proofImage:null});
  };

  const deletePayRecord=(invId,recId)=>{
    const inv=invoices.find(i=>i.id===invId);
    if(!inv) return;
    const records=(inv.paymentRecords||[]).filter(r=>r.id!==recId);
    const newStatus=calcInvStatus(inv,records);
    const upd=invoices.map(i=>i.id===invId?{...i,paymentRecords:records,status:newStatus}:i);
    setInvoices(upd);saveInvoices(upd);
  };

  const addPayProof=(id,dataUrl)=>{
    const u=invoices.map(i=>i.id===id?{...i,paymentProof:dataUrl}:i);setInvoices(u);saveInvoices(u);
  };

  // --- Invoice Batch helpers ---
  const createBatch=()=>{
    if(selected.size<2||!batchName.trim()) return;
    const ids=[...selected];
    const invList=ids.map(id=>invoices.find(i=>i.id===id)).filter(Boolean);
    const combined=invList.reduce((s,i)=>s+(parseFloat(i.total)||0),0);

    // Capture any deposits/partial payments already made on individual invoices
    // These are counted toward the batch total so user only pays the remaining balance
    const preExistingPaid=invList.reduce((s,inv)=>
      s+(inv.paymentRecords||[]).reduce((ps,r)=>ps+(parseFloat(r.amount)||0),0)
    ,0);

    const newBatch={id:uid(),label:batchName.trim(),
      projectId:invList[0]?.projectId||'',supplier:invList[0]?.supplier||'',
      invoiceIds:ids,totalAmount:combined,
      preExistingPaid,  // deposits/partials paid before batch was created
      paymentRecords:[],
      createdAt:new Date().toISOString()};
    const upd=[...(invoiceBatches||[]),newBatch];
    setInvoiceBatches(upd);saveS('invoiceBatches',upd);
    setSelected(new Set());setBatchModal(false);setBatchName('');
  };

  // Total effective paid = pre-existing individual payments + batch payment records
  const getBatchTotalPaid=(batch)=>{
    const batchPaid=(batch.paymentRecords||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
    return batchPaid+(batch.preExistingPaid||0);
  };

  const getBatchStatus=(batch)=>{
    const totalPaid=getBatchTotalPaid(batch);
    if(totalPaid<=0) return 'Pending';
    if(totalPaid>=batch.totalAmount) return 'Paid';
    return 'Partial';
  };

  // Outstanding = full combined total minus everything already paid (deposits + batch payments)
  const getBatchOutstanding=(batch)=>
    Math.max(0, batch.totalAmount - getBatchTotalPaid(batch));

  const recordBatchPayment=()=>{
    const batch=(invoiceBatches||[]).find(b=>b.id===batchPayModal);
    if(!batch||!batchPayForm.amount||parseFloat(batchPayForm.amount)<=0) return;
    const newRec={id:uid(),amount:parseFloat(batchPayForm.amount),date:batchPayForm.date,
      method:batchPayForm.method,reference:batchPayForm.reference||'',notes:batchPayForm.notes||'',
      proofImage:batchPayForm.proofImage||null,
      recordedAt:new Date().toISOString()};
    const records=[...(batch.paymentRecords||[]),newRec];
    const updatedBatch={...batch,paymentRecords:records};
    const newStatus=getBatchStatus(updatedBatch);

    // Total paid via the batch channel only (not pre-existing individual payments)
    const batchChannelPaid=records.reduce((s,r)=>s+(parseFloat(r.amount)||0),0);

    // Distribute batch payments across invoices in order,
    // giving each invoice credit for what it already had paid individually
    let batchRemaining=batchChannelPaid;
    const statusMap={};

    for(const id of batch.invoiceIds){
      const inv=invoices.find(i=>i.id===id);
      if(!inv) continue;
      const invTotal=parseFloat(inv.total)||0;
      // What this invoice already has paid via its own paymentRecords
      const ownPaid=(inv.paymentRecords||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
      // How much more this invoice still needs
      const stillNeeds=Math.max(0, invTotal-ownPaid);
      // Allocate from the batch channel
      const batchAllocated=Math.min(batchRemaining, stillNeeds);
      batchRemaining=Math.max(0, batchRemaining-batchAllocated);
      // Total covered = own payments + batch allocation
      const totalCovered=ownPaid+batchAllocated;
      if(totalCovered<=0) statusMap[id]='Pending';
      else if(totalCovered>=invTotal-0.001) statusMap[id]='Paid'; // 0.001 tolerance for float math
      else statusMap[id]='Partial';
    }

    // If the batch is fully paid, force ALL invoices in batch to Paid
    // (handles float rounding edge cases)
    if(newStatus==='Paid'){
      for(const id of batch.invoiceIds) statusMap[id]='Paid';
    }

    const updInvoices=invoices.map(inv=>{
      if(!batch.invoiceIds.includes(inv.id)) return inv;
      return {...inv, status: statusMap[inv.id]||inv.status};
    });

    const updBatches=(invoiceBatches||[]).map(b=>b.id===batchPayModal?{...updatedBatch,status:newStatus}:b);
    setInvoices(updInvoices);saveInvoices(updInvoices);
    setInvoiceBatches(updBatches);saveS('invoiceBatches',updBatches);
    setBatchPayModal(null);
    setBatchPayForm({amount:'',date:new Date().toISOString().slice(0,10),method:'paynow',reference:'',notes:'',proofImage:null});
  };

  const deleteBatch=(batchId)=>{
    const upd=(invoiceBatches||[]).filter(b=>b.id!==batchId);
    setInvoiceBatches(upd);saveS('invoiceBatches',upd);
  };

  // Ungroup — dissolve the batch but keep all invoices and their payment records intact
  const ungroupBatch=(batchId)=>{
    // Reset individual invoice statuses based on their own paymentRecords only
    const batch=(invoiceBatches||[]).find(b=>b.id===batchId);
    if(!batch) return;
    const updInvoices=invoices.map(inv=>{
      if(!batch.invoiceIds.includes(inv.id)) return inv;
      // Recalculate status from own paymentRecords (batch payments won't apply)
      const ownPaid=(inv.paymentRecords||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
      const status=ownPaid<=0?'Pending':ownPaid>=inv.total?'Paid':'Partial';
      return {...inv,status};
    });
    setInvoices(updInvoices);saveInvoices(updInvoices);
    const upd=(invoiceBatches||[]).filter(b=>b.id!==batchId);
    setInvoiceBatches(upd);saveS('invoiceBatches',upd);
    onShowToast(`Payment batch ungrouped — invoices are now independent`);
  };

  const deleteBatchPayRecord=(batchId,recId)=>{
    const upd=(invoiceBatches||[]).map(b=>b.id===batchId
      ?{...b,paymentRecords:(b.paymentRecords||[]).filter(r=>r.id!==recId)}:b);
    setInvoiceBatches(upd);saveS('invoiceBatches',upd);
  };

  const del_=(id)=>{
    if(!isAdmin){return;}
    const inv=invoices.find(i=>i.id===id);
    if(inv)setDeleteTarget(inv);
  };

  const confirmSoftDelete=(inv)=>{
    onSoftDelete({...inv,_trashType:'invoice',_deletedAt:new Date().toISOString()});
    const updInvoices=invoices.filter(i=>i.id!==inv.id);
    setInvoices(updInvoices);saveInvoices(updInvoices);

    // Auto-ungroup any batch that contained this invoice
    const affectedBatches=(invoiceBatches||[]).filter(b=>b.invoiceIds.includes(inv.id));
    if(affectedBatches.length>0){
      const updBatches=(invoiceBatches||[]).filter(b=>!b.invoiceIds.includes(inv.id));
      setInvoiceBatches(updBatches);saveS('invoiceBatches',updBatches);
    }

    setDeleteTarget(null);
    const hadBatch=affectedBatches.length>0;
    onShowToast(
      `Invoice ${inv.invoiceNo} moved to Trash${hadBatch?' — payment batch was ungrouped':''}`,
      ()=>{
        setInvoices(prev=>[...prev,inv]);
        saveInvoices([...updInvoices,inv]);
      }
    );
  };

  const filt=invoices.filter(i=>
    (i.supplier+i.invoiceNo).toLowerCase().includes(search.toLowerCase())&&
    (pfilt==='All'||i.projectId===pfilt)&&
    (cfilt==='All'||i.category===cfilt)
  );
  const activeInv=filt.filter(i=>i.status!=='Paid');
  const archivedInv=filt.filter(i=>i.status==='Paid');
  const pendAmt=invoices.filter(i=>i.status==='Pending'||i.status==='Partial').reduce((s,i)=>s+i.total,0);
  const [showArchived,setShowArchived]=useState(false);
  const [editTarget,setEditTarget]=useState(null);
  const [editForm,setEditForm]=useState({});
  const ef=k=>v=>setEditForm(p=>({...p,[k]:v}));
  const saveEdit=()=>{
    const upd=invoices.map(i=>i.id===editTarget?{...i,...editForm,
      subtotal:parseFloat(editForm.subtotal)||0,gst:parseFloat(editForm.gst)||0,
      total:parseFloat(editForm.total)||0}:i);
    setInvoices(upd);saveInvoices(upd);setEditTarget(null);
  };
  const groupByProject=list=>{const m={};list.forEach(inv=>{if(!m[inv.projectId])m[inv.projectId]=[];m[inv.projectId].push(inv);});return m;};
  const activeByProject=groupByProject(activeInv);

  const InvRow=({inv,isArchived})=>{
    const proj=projects.find(p=>p.id===inv.projectId);
    const inBatch=(invoiceBatches||[]).some(b=>b.invoiceIds.includes(inv.id));
    const isSelected=selected.has(inv.id);
    const isClaimInv=!!inv._fromClaim; // auto-generated from staff claim approval
    return(
      <tr style={{borderTop:`1px solid ${T.borderLight}`,opacity:isArchived?0.75:1,
        background:isSelected?T.accentLight:isClaimInv?'rgba(0,113,227,0.02)':'transparent',transition:'background 0.12s'}}>
        <td style={{padding:'10px 14px',width:32}}>
          {isAdmin&&!inBatch&&!isArchived&&(
            <input type="checkbox" checked={isSelected}
              onChange={e=>{const s=new Set(selected);e.target.checked?s.add(inv.id):s.delete(inv.id);setSelected(s);}}
              style={{width:15,height:15,cursor:'pointer',accentColor:T.accent}}/>
          )}
          {inBatch&&<span title="Part of a payment batch" style={{fontSize:14}}>🔗</span>}
        </td>
        <td style={{padding:'10px 14px',color:T.text,fontWeight:600}}>
          <div>{inv.supplier}</div>
          {isClaimInv&&<div style={{fontSize:10,color:T.accent,fontWeight:600,marginTop:2}}>📋 From Staff Claim</div>}
        </td>
        <td style={{padding:'10px 14px',color:T.dim,fontFamily:'monospace',fontSize:11}}>{inv.invoiceNo}</td>
        <td style={{padding:'10px 14px'}}><Badge color={CAT_CLR[inv.category]||T.dim}>{inv.category}</Badge></td>
        <td style={{padding:'10px 14px',color:T.dim,whiteSpace:'nowrap'}}>{fmtDate(inv.invoiceDate)}</td>
        <td style={{padding:'10px 14px',textAlign:'right',color:T.muted}}>{fmtSGD(inv.subtotal)}</td>
        <td style={{padding:'10px 14px',textAlign:'right',color:T.dim,fontSize:12}}>{inv.gst>0?fmtSGD(inv.gst):'—'}</td>
        <td style={{padding:'10px 14px',textAlign:'right',color:T.text,fontWeight:700}}>
          <div>{fmtSGD(inv.total)}</div>
          {(()=>{
            const paid=(inv.paymentRecords||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
            const outstanding=inv.total-paid;
            if(paid>0&&outstanding>0) return <div style={{fontSize:10,color:T.danger,marginTop:2}}>Bal: {fmtSGD(outstanding)}</div>;
            return null;
          })()}
        </td>
        <td style={{padding:'10px 14px'}}>
          <div style={{display:'flex',flexDirection:'column',gap:3}}>
            <Badge color={ST_CLR[inv.status]||T.muted}>{inv.status}</Badge>
            {(()=>{
              const paid=(inv.paymentRecords||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
              if(paid>0){
                const lastRec=(inv.paymentRecords||[]).slice(-1)[0];
                const icons={paynow:'📱',bank_transfer:'🏦',cash:'💵'};
                return <div style={{fontSize:10,color:T.muted,display:'flex',alignItems:'center',gap:3}}>
                  {icons[lastRec?.method]||'💰'} {fmtSGD(paid)} paid
                </div>;
              }
              return null;
            })()}
          </div>
        </td>
        <td style={{padding:'10px 14px'}}>
          <div style={{display:'flex',alignItems:'center',gap:5,flexWrap:'wrap'}}>
            {isAdmin&&inv.status!=='Paid'&&(
              <button title="Record Payment" onClick={()=>{
                const outstanding=inv.total-(inv.paymentRecords||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
                const hist=getSupplierHistory(inv.supplier);
                setPayModal(inv.id);
                setPayForm({amount:String(Math.round(outstanding*100)/100),date:new Date().toISOString().slice(0,10),
                  method:hist?.method||'paynow',reference:hist?.reference||'',notes:'',proofImage:null,_histFound:!!hist});
              }}
                style={{background:T.successLight,border:`1px solid ${T.success}30`,cursor:'pointer',color:T.success,display:'flex',padding:'3px 8px',borderRadius:6,alignItems:'center',gap:3,fontSize:11,fontFamily:'inherit',fontWeight:600}}>
                <CreditCard size={11}/>Pay</button>
            )}
            {(inv.paymentRecords||[]).length>0&&(
              <button title="View payment history" onClick={()=>setLightbox({type:'payments',inv})}
                style={{background:T.bg,border:`1px solid ${T.borderLight}`,cursor:'pointer',color:T.muted,display:'flex',padding:'3px 7px',borderRadius:6,alignItems:'center',gap:3,fontSize:11,fontFamily:'inherit'}}>
                <Receipt size={11}/>{(inv.paymentRecords||[]).length}</button>
            )}
            {inv.proofImage&&<button title="View invoice doc" onClick={()=>setLightbox({type:'img',src:inv.proofImage,title:`Invoice — ${inv.invoiceNo}`})}
              style={{background:T.infoLight,border:'none',cursor:'pointer',color:T.info,display:'flex',padding:'3px 7px',borderRadius:6,alignItems:'center',gap:3,fontSize:11,fontFamily:'inherit'}}>
              {inv.proofImage?.startsWith('data:application/pdf')?<><FileSpreadsheet size={11}/>PDF</>:<><ZoomIn size={11}/>Doc</>}</button>}
            {/* Edit — available on ALL invoices including paid/archived */}
            {isAdmin&&<button title="Edit invoice" onClick={()=>{
              setEditTarget(inv.id);
              setEditForm({supplier:inv.supplier,invoiceNo:inv.invoiceNo,invoiceDate:inv.invoiceDate,
                subtotal:String(inv.subtotal),gst:String(inv.gst),total:String(inv.total),
                category:inv.category,status:inv.status,
                paymentMethod:inv.paymentMethod||'',paidDate:inv.paidDate||''});
            }}
              style={{background:T.bg,border:`1px solid ${T.borderLight}`,cursor:'pointer',color:T.muted,display:'flex',padding:'3px 7px',borderRadius:6,alignItems:'center',gap:3,fontSize:11,fontFamily:'inherit'}}>
              <Edit3 size={11}/>Edit</button>}
            {/* Delete — available on ALL invoices including paid ones */}
            {isAdmin&&<button onClick={()=>del_(inv.id)} title={isClaimInv?'Delete this claim-generated entry':'Delete invoice'}
              style={{background:isClaimInv?T.dangerLight:'none',border:isClaimInv?`1px solid ${T.danger}30`:'none',cursor:'pointer',color:T.danger,display:'flex',padding:'3px 6px',borderRadius:6,alignItems:'center',gap:3,fontSize:11,fontFamily:'inherit'}}>
              <Trash2 size={13}/>{isClaimInv&&<span>Delete</span>}</button>}
          </div>
        </td>
      </tr>
    );
  };

  const InvTable=({invList,isArchived=false})=>(
    <div style={{overflowX:'auto'}}>
      <table style={{width:'100%',fontSize:13,borderCollapse:'collapse'}}>
        <thead>
          <tr style={{color:T.dim,fontSize:11,fontWeight:600,textTransform:'uppercase',letterSpacing:'0.06em',background:T.bg}}>
            <th style={{padding:'10px 14px',width:32}}/>
            {['Supplier','Invoice #','Category','Date','Subtotal','GST','Total','Status',''].map(h=>(
              <th key={h} style={{textAlign:['Subtotal','GST','Total'].includes(h)?'right':'left',padding:'10px 14px',whiteSpace:'nowrap'}}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>{invList.map(inv=><InvRow key={inv.id} inv={inv} isArchived={isArchived}/>)}</tbody>
      </table>
    </div>
  );

  return (
    <div style={{display:'flex',flexDirection:'column',gap:16}}>
      {pendAmt>0&&(
        <div style={{background:T.warningLight,border:'1px solid rgba(183,86,10,0.2)',borderRadius:12,padding:'10px 16px',display:'flex',alignItems:'center',gap:10}}>
          <Bell size={13} style={{color:T.warning,flexShrink:0}}/>
          <span style={{fontSize:13,color:T.warning}}><strong>{fmtSGD(pendAmt)}</strong> awaiting payment — {invoices.filter(i=>i.status==='Pending'||i.status==='Partial').length} invoices</span>
        </div>
      )}

      {/* Selection bar — appears when invoices are checked */}
      {selected.size>0&&isAdmin&&(
        <div style={{background:T.accentLight,border:`1px solid ${T.accent}30`,borderRadius:12,padding:'10px 16px',
          display:'flex',alignItems:'center',justifyContent:'space-between',gap:12}}>
          <div style={{display:'flex',alignItems:'center',gap:10}}>
            <span style={{fontSize:13,fontWeight:700,color:T.accent}}>{selected.size} invoice{selected.size!==1?'s':''} selected</span>
            <span style={{fontSize:12,color:T.muted}}>·</span>
            <span style={{fontSize:13,color:T.accent,fontWeight:600}}>
              Combined: {fmtSGD([...selected].reduce((s,id)=>s+(invoices.find(i=>i.id===id)?.total||0),0))}
            </span>
          </div>
          <div style={{display:'flex',gap:8}}>
            <button type="button" onClick={()=>setSelected(new Set())}
              style={{background:'none',border:`1px solid ${T.borderLight}`,borderRadius:8,padding:'6px 12px',cursor:'pointer',fontFamily:'inherit',fontSize:12,color:T.muted}}>
              Clear
            </button>
            <Btn onClick={()=>{
              const invList=[...selected].map(id=>invoices.find(i=>i.id===id)).filter(Boolean);
              const supplier=invList[0]?.supplier||'';
              setBatchName(`${supplier} — Combined Payment`);
              setBatchModal(true);
            }}>
              <Receipt size={13}/>Group for Payment Batch
            </Btn>
          </div>
        </div>
      )}

      <div style={{display:'flex',gap:10,flexWrap:'wrap'}}>
        <div style={{flex:1,minWidth:180,position:'relative'}}>
          <Search size={13} style={{position:'absolute',left:11,top:'50%',transform:'translateY(-50%)',color:T.dim}}/>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search supplier or invoice #..." style={{...iStyle,paddingLeft:33}}/>
        </div>
        <select value={pfilt} onChange={e=>setPfilt(e.target.value)} style={{...iStyle,width:'auto',minWidth:145}}>
          <option value="All">All Projects</option>
          {projects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
        <select value={cfilt} onChange={e=>setCfilt(e.target.value)} style={{...iStyle,width:'auto'}}>
          <option value="All">All Categories</option>
          {CATS.map(c=><option key={c} value={c}>{c}</option>)}
        </select>
        <Btn onClick={()=>{setForm(blank);setPreview(null);setOcr({loading:false,done:false,err:''});setDup('');setModal(true);}}><Plus size={13}/>Capture Invoice</Btn>
      </div>

      {Object.keys(activeByProject).length===0&&(
        <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:'36px',textAlign:'center',color:T.dim,boxShadow:T.shadow,fontSize:13}}>
          No active invoices — all invoices are paid or archived
        </div>
      )}

      {/* Payment Batches */}
      {(invoiceBatches||[]).length>0&&(
        <div style={{display:'flex',flexDirection:'column',gap:10}}>
          <div style={{fontSize:13,fontWeight:700,color:T.text,display:'flex',alignItems:'center',gap:8,flexWrap:'wrap'}}>
            <Receipt size={14} style={{color:T.accent}}/>Payment Batches
            <span style={{fontSize:11,color:T.muted,fontWeight:400}}>Grouped invoices paid together</span>
            <button onClick={()=>setHidePaidBatches(h=>!h)}
              style={{marginLeft:'auto',display:'flex',alignItems:'center',gap:5,background:'transparent',border:`1px solid ${T.borderLight}`,borderRadius:8,padding:'4px 10px',fontSize:11,fontWeight:500,color:T.muted,cursor:'pointer',fontFamily:'inherit'}}>
              {hidePaidBatches?<Eye size={11}/>:<EyeOff size={11}/>}
              {hidePaidBatches?'Show paid':'Hide paid'}
              {hidePaidBatches&&(invoiceBatches||[]).filter(b=>getBatchStatus(b)==='Paid').length>0&&(
                <span style={{background:T.success,color:'#fff',borderRadius:10,padding:'1px 6px',fontSize:10,fontWeight:700}}>
                  {(invoiceBatches||[]).filter(b=>getBatchStatus(b)==='Paid').length}
                </span>
              )}
            </button>
          </div>
          {(invoiceBatches||[]).filter(batch=>!hidePaidBatches||getBatchStatus(batch)!=='Paid').map(batch=>{
            const batchInvoices=batch.invoiceIds.map(id=>invoices.find(i=>i.id===id)).filter(Boolean);
            const totalPaid=getBatchTotalPaid(batch);
            const batchOnlyPaid=(batch.paymentRecords||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
            const preExisting=batch.preExistingPaid||0;
            const outstanding=getBatchOutstanding(batch);
            const status=getBatchStatus(batch);
            const proj=projects.find(p=>p.id===batch.projectId);
            const icons={paynow:'📱',bank_transfer:'🏦',cash:'💵'};
            return (
              <div key={batch.id} style={{background:T.card,border:`2px solid ${status==='Paid'?T.success+'30':status==='Partial'?'#7c3aed30':T.borderLight}`,borderRadius:16,overflow:'hidden',boxShadow:T.shadow}}>
                {/* Batch header */}
                <div style={{padding:'14px 18px',background:status==='Paid'?T.successLight:status==='Partial'?'rgba(124,58,237,0.05)':T.bg,
                  borderBottom:`1px solid ${T.borderLight}`,display:'flex',alignItems:'center',justifyContent:'space-between',flexWrap:'wrap',gap:10}}>
                  <div>
                    <div style={{fontSize:14,fontWeight:700,color:T.text,display:'flex',alignItems:'center',gap:8}}>
                      🔗 {batch.label}
                      <Badge color={ST_CLR[status]||T.muted}>{status}</Badge>
                    </div>
                    <div style={{fontSize:11,color:T.muted,marginTop:3}}>{proj?.name||'—'} · {batch.invoiceIds.length} invoices grouped</div>
                  </div>
                  <div style={{display:'flex',alignItems:'center',gap:12,flexWrap:'wrap'}}>
                    <div style={{textAlign:'right'}}>
                      <div style={{fontSize:11,color:T.dim}}>Combined Total</div>
                      <div style={{fontSize:16,fontWeight:800,color:T.text}}>{fmtSGD(batch.totalAmount)}</div>
                    </div>
                    {preExisting>0&&(
                      <div style={{textAlign:'right'}}>
                        <div style={{fontSize:11,color:T.dim}}>Deposit Paid</div>
                        <div style={{fontSize:13,fontWeight:700,color:T.success}}>{fmtSGD(preExisting)}</div>
                      </div>
                    )}
                    {batchOnlyPaid>0&&(
                      <div style={{textAlign:'right'}}>
                        <div style={{fontSize:11,color:T.dim}}>Batch Paid</div>
                        <div style={{fontSize:13,fontWeight:700,color:T.success}}>{fmtSGD(batchOnlyPaid)}</div>
                      </div>
                    )}
                    <div style={{textAlign:'right'}}>
                      <div style={{fontSize:11,color:T.dim}}>Balance Due</div>
                      <div style={{fontSize:16,fontWeight:800,color:outstanding>0?T.danger:T.success}}>{fmtSGD(outstanding)}</div>
                    </div>
                    {isAdmin&&status!=='Paid'&&(
                      <button type="button" onClick={()=>{
                        const hist=getSupplierHistory(batch.supplier);
                        setBatchPayModal(batch.id);
                        setBatchPayForm({
                          amount:String(Math.round(outstanding*100)/100),
                          date:new Date().toISOString().slice(0,10),
                          method:hist?.method||'paynow',
                          reference:hist?.reference||'',
                          notes:preExisting>0?'Balance payment':'50% deposit',
                          proofImage:null,
                          _histFound:!!hist,
                        });
                      }}
                        style={{display:'flex',alignItems:'center',gap:6,padding:'8px 14px',borderRadius:9,border:'none',cursor:'pointer',fontFamily:'inherit',fontSize:12,fontWeight:700,background:T.accent,color:'#fff'}}>
                        <CreditCard size={13}/>{preExisting>0?`Pay Balance ${fmtSGD(outstanding)}`:'Record Payment'}
                      </button>
                    )}
                    {isAdmin&&(
                      <button type="button" onClick={()=>ungroupBatch(batch.id)}
                        title="Ungroup — dissolve batch, keep invoices"
                        style={{display:'flex',alignItems:'center',gap:5,padding:'7px 12px',borderRadius:9,border:`1px solid ${T.borderLight}`,cursor:'pointer',fontFamily:'inherit',fontSize:11,fontWeight:600,background:T.bg,color:T.muted}}>
                        🔓 Ungroup
                      </button>
                    )}
                  </div>
                </div>

                {/* Progress bar */}
                <div style={{height:4,background:T.borderLight}}>
                  <div style={{height:'100%',width:`${Math.min(100,batch.totalAmount>0?totalPaid/batch.totalAmount*100:0)}%`,
                    background:status==='Paid'?T.success:status==='Partial'?'#7c3aed':T.warning,transition:'width 0.3s'}}/>
                </div>

                {/* Invoices in batch — shown in payment order */}
                <div style={{padding:'12px 18px',display:'flex',flexDirection:'column',gap:6}}>
                  <div style={{fontSize:10,color:T.dim,fontWeight:600,textTransform:'uppercase',letterSpacing:'0.06em',marginBottom:2}}>
                    Payment Order — Invoice 1 cleared first, then VO
                  </div>
                  {(()=>{
                    const batchChannelPaid=(batch.paymentRecords||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
                    let batchRem=batchChannelPaid;
                    const bStatus=getBatchStatus(batch);
                    return batch.invoiceIds.map((id,idx)=>{
                      const inv=invoices.find(i=>i.id===id);
                      if(!inv) return null;
                      const invTotal=parseFloat(inv.total)||0;
                      const ownPaid=(inv.paymentRecords||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
                      const stillNeeds=Math.max(0,invTotal-ownPaid);
                      const batchAlloc=Math.min(batchRem,stillNeeds);
                      batchRem=Math.max(0,batchRem-batchAlloc);
                      const totalCov=ownPaid+batchAlloc;
                      const invStatus=bStatus==='Paid'?'Paid':totalCov<=0?'Pending':totalCov>=invTotal-0.001?'Paid':'Partial';
                      return (
                        <div key={id} style={{display:'flex',alignItems:'center',justifyContent:'space-between',fontSize:12,padding:'8px 10px',background:T.bg,borderRadius:8,border:`1px solid ${invStatus==='Paid'?T.success+'30':invStatus==='Partial'?'#7c3aed30':T.borderLight}`}}>
                          <div style={{display:'flex',alignItems:'center',gap:8}}>
                            <span style={{fontSize:11,fontWeight:700,color:T.dim,minWidth:20}}>#{idx+1}</span>
                            <Badge color={CAT_CLR[inv.category]||T.dim} sm>{inv.category}</Badge>
                            <span style={{color:T.text,fontWeight:600}}>{inv.supplier}</span>
                            <span style={{color:T.dim,fontFamily:'monospace',fontSize:11}}>{inv.invoiceNo}</span>
                          </div>
                          <div style={{display:'flex',alignItems:'center',gap:10}}>
                            {totalCov>0&&totalCov<invTotal&&(
                              <span style={{fontSize:10,color:'#7c3aed',fontWeight:600}}>{fmtSGD(totalCov)} / {fmtSGD(invTotal)}</span>
                            )}
                            <Badge color={ST_CLR[invStatus]||T.muted}>{invStatus}</Badge>
                            <span style={{color:T.text,fontWeight:700}}>{fmtSGD(inv.total)}</span>
                          </div>
                        </div>
                      );
                    });
                  })()}
                </div>

                {/* Payment records */}
                {(batch.paymentRecords||[]).length>0&&(
                  <div style={{padding:'0 18px 12px',borderTop:`1px solid ${T.borderLight}`,marginTop:4}}>
                    <div style={{fontSize:11,fontWeight:600,color:T.dim,textTransform:'uppercase',letterSpacing:'0.06em',margin:'10px 0 6px'}}>Payment History</div>
                    {(batch.paymentRecords||[]).map(r=>(
                      <div key={r.id} style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'7px 10px',background:T.bg,borderRadius:8,marginBottom:4}}>
                        <div style={{display:'flex',alignItems:'center',gap:8}}>
                          <span style={{fontSize:16}}>{icons[r.method]||'💰'}</span>
                          <div>
                            <div style={{fontSize:12,fontWeight:600,color:T.success}}>{fmtSGD(r.amount)}</div>
                            <div style={{fontSize:10,color:T.dim}}>{r.method==='paynow'?'PayNow':r.method==='bank_transfer'?'Bank Transfer':'Cash'}{r.reference&&` · ${r.reference}`}</div>
                          </div>
                        </div>
                        <div style={{display:'flex',alignItems:'center',gap:8}}>
                          <span style={{fontSize:11,color:T.dim}}>{fmtDate(r.date)}{r.notes&&` · ${r.notes}`}</span>
                          {isAdmin&&<button type="button" onClick={()=>deleteBatchPayRecord(batch.id,r.id)}
                            style={{background:'none',border:'none',cursor:'pointer',color:T.dim,padding:2}}><Trash2 size={11}/></button>}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {Object.entries(activeByProject).map(([projId,invList])=>{
        const proj=projects.find(p=>p.id===projId);
        const projTotal=invList.reduce((s,i)=>s+i.total,0);
        return(
          <div key={projId} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,overflow:'hidden',boxShadow:T.shadow}}>
            <div style={{padding:'14px 20px',borderBottom:`1px solid ${T.borderLight}`,display:'flex',alignItems:'center',justifyContent:'space-between',background:T.bg}}>
              <div style={{display:'flex',alignItems:'center',gap:10}}>
                <FolderOpen size={14} style={{color:T.accent}}/>
                <span style={{fontSize:14,fontWeight:600,color:T.text}}>{proj?.name||'Unknown Project'}</span>
                <span style={{fontSize:12,color:T.muted}}>{proj?.client}</span>
              </div>
              <div style={{display:'flex',alignItems:'center',gap:12}}>
                <span style={{fontSize:12,color:T.dim}}>{invList.length} invoice{invList.length!==1?'s':''}</span>
                <span style={{fontSize:13,fontWeight:700,color:T.text}}>{fmtSGD(projTotal)}</span>
              </div>
            </div>
            <InvTable invList={invList}/>
          </div>
        );
      })}

      {archivedInv.length>0&&(
        <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,overflow:'hidden',boxShadow:T.shadow}}>
          <button onClick={()=>setShowArchived(s=>!s)}
            style={{width:'100%',padding:'14px 20px',background:T.bg,border:'none',cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'space-between',fontFamily:'inherit'}}>
            <div style={{display:'flex',alignItems:'center',gap:10}}>
              <CheckCircle size={14} style={{color:T.success}}/>
              <span style={{fontSize:14,fontWeight:600,color:T.muted}}>Paid / Archived Invoices</span>
              <Badge color={T.success} sm>{archivedInv.length}</Badge>
            </div>
            <span style={{fontSize:12,color:T.dim}}>{showArchived?'▲ Hide':'▼ Show'}</span>
          </button>
          {showArchived&&<InvTable invList={archivedInv} isArchived/>}
        </div>
      )}

      {editTarget&&(
        <Modal title="Edit Invoice" onClose={()=>setEditTarget(null)}>
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <Field label="Supplier Name" value={editForm.supplier||''} onChange={ef('supplier')} placeholder="Supplier"/>
            <Field label="Invoice Number" value={editForm.invoiceNo||''} onChange={ef('invoiceNo')} placeholder="Invoice #"/>
            <Field label="Invoice Date" type="date" value={editForm.invoiceDate||''} onChange={ef('invoiceDate')}/>
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:10}}>
              <div>
                <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Subtotal (S$)</label>
                <input type="number" value={editForm.subtotal||''} onChange={e=>{const s=parseFloat(e.target.value)||0;const g=parseFloat(editForm.gst)||0;setEditForm(p=>({...p,subtotal:e.target.value,total:String(s+g)}))} } style={iStyle}/>
              </div>
              <div>
                <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>GST (S$) — 0 = no GST</label>
                <input type="number" value={editForm.gst||''} onChange={e=>{const g=parseFloat(e.target.value)||0;const s=parseFloat(editForm.subtotal)||0;setEditForm(p=>({...p,gst:e.target.value,total:String(s+g)}))} } style={iStyle}/>
              </div>
              <Field label="Total (S$)" type="number" value={editForm.total||''} onChange={ef('total')}/>
            </div>
            <Field label="Category" value={editForm.category||'Carpentry'} onChange={ef('category')} as="select" options={CATS.map(c=>({v:c,l:c}))}/>
            <Field label="Status" value={editForm.status||'Pending'} onChange={ef('status')} as="select" options={INV_STATUSES.map(s=>({v:s,l:s}))}/>

            {(editForm.status==='Paid')&&(
              <div style={{background:T.successLight,border:`1px solid ${T.success}20`,borderRadius:12,padding:'14px 16px',display:'flex',flexDirection:'column',gap:12}}>
                <div style={{fontSize:12,fontWeight:700,color:T.success,display:'flex',alignItems:'center',gap:6}}>
                  <CheckCircle size={13}/>Payment Details
                </div>
                <div>
                  <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:8}}>Payment Method</label>
                  <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:8}}>
                    {[
                      {id:'paynow',label:'PayNow',icon:'📱'},
                      {id:'bank_transfer',label:'Bank Transfer',icon:'🏦'},
                      {id:'cash',label:'Cash',icon:'💵'},
                    ].map(m=>(
                      <button key={m.id} type="button" onClick={()=>ef('paymentMethod')(m.id)}
                        style={{padding:'10px 8px',borderRadius:10,
                          border:`2px solid ${editForm.paymentMethod===m.id?T.success:T.borderLight}`,
                          background:editForm.paymentMethod===m.id?'rgba(29,131,72,0.08)':'transparent',
                          cursor:'pointer',fontFamily:'inherit',display:'flex',flexDirection:'column',
                          alignItems:'center',gap:4,transition:'all 0.12s'}}>
                        <span style={{fontSize:20}}>{m.icon}</span>
                        <span style={{fontSize:11,fontWeight:600,color:editForm.paymentMethod===m.id?T.success:T.muted}}>{m.label}</span>
                      </button>
                    ))}
                  </div>
                </div>
                <Field label="Date Paid" type="date" value={editForm.paidDate||editForm.invoiceDate||''} onChange={ef('paidDate')}/>
              </div>
            )}
          </div>
          <div style={{display:'flex',justifyContent:'flex-end',gap:10,marginTop:22}}>
            <Btn variant="secondary" onClick={()=>setEditTarget(null)}>Cancel</Btn>
            <Btn onClick={saveEdit}>Save Changes</Btn>
          </div>
        </Modal>
      )}

      {/* Hidden payment proof upload input */}
      <input ref={proofRef} type="file" accept="image/*,application/pdf" style={{display:'none'}}
        onChange={e=>{
          const f=e.target.files?.[0];
          if(f&&proofTarget){
            const r=new FileReader();
            r.onload=ev=>{addPayProof(proofTarget,ev.target.result);setProofTarget(null);};
            r.readAsDataURL(f);
          }
          e.target.value='';
        }}/>

      {modal&&(
        <Modal title="Capture Invoice — AI OCR" onClose={()=>setModal(false)} wide>
          <div style={{display:'flex',flexDirection:'column',gap:16}}>

            {/* STEP 1 — Upload invoice */}
            <div>
              <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:8,display:'flex',alignItems:'center',gap:6}}>
                <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>1</span>
                Upload Invoice {(acctSettings?.anthropicApiKey||'').trim()&&<span style={{fontSize:10,color:T.accent,fontWeight:400}}>— AI fills all fields below</span>}
              </div>
              <DropZone accept="image/*,application/pdf" onDrop={f=>doOCR(f)}>
              <div onClick={()=>fileRef.current?.click()}
                style={{border:`2px dashed ${preview?T.success+'60':T.borderLight}`,borderRadius:12,
                  cursor:'pointer',background:T.bg,overflow:'hidden',transition:'border-color .2s',
                  minHeight:preview?'auto':120,display:'flex',alignItems:'center',justifyContent:'center',flexDirection:'column',gap:10,padding:preview?0:20}}>
                {preview?(
                  preview.startsWith('data:application/pdf')?(
                    <div style={{padding:'16px 20px',display:'flex',alignItems:'center',gap:12}}>
                      <span style={{fontSize:36}}>📄</span>
                      <div>
                        <div style={{fontSize:13,fontWeight:600,color:T.text}}>{ocr.loading?'AI reading invoice...':ocr.done?'Invoice read ✓':'PDF uploaded'}</div>
                        <div style={{fontSize:11,color:T.muted,marginTop:2}}>{ocr.done?'Fields filled below — review and confirm':'Ready to process'}</div>
                      </div>
                    </div>
                  ):(
                    <img src={preview} alt="Invoice" style={{width:'100%',maxHeight:200,objectFit:'contain',display:'block'}}/>
                  )
                ):(
                  <>
                    <div style={{width:48,height:48,background:T.accentLight,borderRadius:12,display:'flex',alignItems:'center',justifyContent:'center'}}>
                      <Upload size={20} style={{color:T.accent}}/>
                    </div>
                    <div style={{textAlign:'center'}}>
                      <div style={{fontSize:13,fontWeight:700,color:T.text}}>Drop invoice here or tap to upload</div>
                      <div style={{fontSize:11,color:T.dim,marginTop:3}}>JPG, PNG or PDF — AI extracts all fields</div>
                    </div>
                  </>
                )}
              </div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:8,marginTop:8}}>
                <Btn variant="secondary" size="sm" onClick={()=>fileRef.current?.click()} full><Upload size={12}/>Upload File</Btn>
                <Btn variant="secondary" size="sm" onClick={()=>camRef.current?.click()} full><Camera size={12}/>Take Photo</Btn>
              </div>
              <input ref={fileRef} type="file" accept="image/*,application/pdf" style={{display:'none'}}
                onChange={e=>{if(e.target.files?.[0])doOCR(e.target.files[0]);e.target.value='';}}/>
              <input ref={camRef} type="file" accept="image/*" capture="environment" style={{display:'none'}}
                onChange={e=>{if(e.target.files?.[0])doOCR(e.target.files[0]);e.target.value='';}}/>
              </DropZone>
              {ocr.loading&&<div style={{marginTop:8,display:'flex',alignItems:'center',gap:8,fontSize:12,color:T.accent,background:T.accentLight,borderRadius:10,padding:'9px 12px'}}>
                <Loader2 size={13} style={{animation:'spin 1s linear infinite'}}/>AI is extracting invoice details...
              </div>}
              {ocr.done&&<div style={{marginTop:8,display:'flex',alignItems:'center',gap:8,fontSize:12,color:T.success,background:T.successLight,borderRadius:10,padding:'9px 12px'}}>
                <CheckCircle size={13}/>Details extracted — review and fill in anything missing below
              </div>}
              {preview&&preview.length>MAX_IMAGE_B64&&!ocr.loading&&(
                <div style={{marginTop:8,fontSize:12,color:'#92400e',background:'#fef3c7',border:'1px solid #fcd34d',borderRadius:10,padding:'9px 12px',display:'flex',alignItems:'flex-start',gap:8}}>
                  <AlertCircle size={13} style={{flexShrink:0,marginTop:1}}/>
                  <span><b>PDF kept locally only</b> — this file is {Math.round(preview.length/1333)}KB which exceeds the 100KB cloud limit. The invoice data (supplier, amount, date) will sync to all devices, but the PDF attachment will only be viewable on this device during this session.</span>
                </div>
              )}
              {ocr.err&&<div style={{marginTop:8,fontSize:12,color:T.danger,background:T.dangerLight,borderRadius:10,padding:'9px 12px'}}>{ocr.err}</div>}
              {dup&&<div style={{marginTop:8,fontSize:12,color:T.warning,background:'rgba(245,158,11,0.08)',borderRadius:10,padding:'9px 12px'}}>⚠ {dup}</div>}
            </div>

            {/* STEP 2 — Review and confirm fields */}
            <div>
              <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:12,display:'flex',alignItems:'center',gap:6}}>
                <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>2</span>
                Review & Confirm Details
                {ocr.loading&&<span style={{fontSize:10,color:T.accent,fontWeight:400}}>— filling from invoice...</span>}
              </div>
              <div style={{display:'flex',flexDirection:'column',gap:12}}>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:12}}>
                  <Field label="Project *" value={form.projectId} onChange={ff('projectId')} as="select"
                    options={[{v:'',l:'Select project...'}, ...projects.map(p=>({v:p.id,l:p.name}))]}/>
                  <Field label="Category" value={form.category} onChange={ff('category')} as="select" options={CATS.map(c=>({v:c,l:c}))}/>
                </div>
                <Field label="Supplier Name *" value={form.supplier} onChange={ff('supplier')} placeholder="e.g. Premium Carpentry Works"/>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:12}}>
                  <Field label="Invoice Number *" value={form.invoiceNo} onChange={ff('invoiceNo')} placeholder="e.g. PCW-2025-0234"/>
                  <Field label="Invoice Date" type="date" value={form.invoiceDate} onChange={ff('invoiceDate')}/>
                </div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:8}}>
                  <Field label="Subtotal (S$)" type="number" value={form.subtotal} onChange={ff('subtotal')} placeholder="0"/>
                  <div>
                    <label style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.07em',display:'block',marginBottom:5}}>GST <span style={{fontWeight:400,fontSize:10}}>(0=none)</span></label>
                    <input type="number" value={form.gst} onChange={e=>ff('gst')(e.target.value)} placeholder="0" style={iStyle}/>
                  </div>
                  <Field label="Total (S$)" type="number" value={form.total} onChange={ff('total')} placeholder="0"/>
                </div>
                <Field label="Status" value={form.status} onChange={ff('status')} as="select" options={INV_STATUSES.map(s=>({v:s,l:s}))}/>
                {form.status==='Paid'&&(
                  <div style={{background:T.successLight,border:`1px solid ${T.success}20`,borderRadius:12,padding:'14px 16px',display:'flex',flexDirection:'column',gap:12}}>
                    <div style={{fontSize:12,fontWeight:700,color:T.success,display:'flex',alignItems:'center',gap:6}}><CheckCircle size={13}/>Payment Details</div>
                    <div>
                      <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:8}}>Payment Method</label>
                      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:8}}>
                        {[{id:'paynow',label:'PayNow',icon:'📱'},{id:'bank_transfer',label:'Bank Transfer',icon:'🏦'},{id:'cash',label:'Cash',icon:'💵'}].map(m=>(
                          <button key={m.id} type="button" onClick={()=>ff('paymentMethod')(m.id)}
                            style={{padding:'8px',borderRadius:10,border:`2px solid ${form.paymentMethod===m.id?T.success:T.borderLight}`,
                              background:form.paymentMethod===m.id?'rgba(29,131,72,0.08)':'transparent',cursor:'pointer',fontFamily:'inherit',
                              display:'flex',flexDirection:'column',alignItems:'center',gap:3}}>
                            <span style={{fontSize:18}}>{m.icon}</span>
                            <span style={{fontSize:11,fontWeight:600,color:form.paymentMethod===m.id?T.success:T.muted}}>{m.label}</span>
                          </button>
                        ))}
                      </div>
                    </div>
                    <Field label="Date Paid" type="date" value={form.paidDate||form.invoiceDate||''} onChange={ff('paidDate')}/>
                  </div>
                )}
              </div>
            </div>

            <div style={{display:'flex',justifyContent:'flex-end',gap:10,borderTop:`1px solid ${T.borderLight}`,paddingTop:14,alignItems:'center'}}>
              {err&&<div style={{fontSize:12,color:T.danger,flex:1,fontWeight:500}}>{err}</div>}
              <Btn variant="secondary" onClick={()=>{setModal(false);setErr('');}}>Cancel</Btn>
              <Btn onClick={save_} disabled={!form.projectId||!form.supplier||!form.invoiceNo}>Save Invoice</Btn>
            </div>
          </div>
        </Modal>
      )}
      {deleteTarget&&(
        <ConfirmDelete
          matchValue={deleteTarget.invoiceNo}
          typeLabel="invoice"
          impact={`This will permanently remove the invoice from project records. Ensure you have physical copies before proceeding.`}
          onConfirm={()=>confirmSoftDelete(deleteTarget)}
          onClose={()=>setDeleteTarget(null)}
        />
      )}
      {/* Create Batch Modal */}
      {batchModal&&(
        <Modal title="Create Payment Batch" onClose={()=>{setBatchModal(false);setBatchName('');}}>
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <div style={{background:T.bg,borderRadius:12,padding:'12px 14px'}}>
              <div style={{fontSize:12,color:T.muted,marginBottom:8,fontWeight:600}}>Invoices being grouped:</div>
              {[...selected].map(id=>{
                const inv=invoices.find(i=>i.id===id);
                if(!inv) return null;
                const alreadyPaid=(inv.paymentRecords||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
                const balance=inv.total-alreadyPaid;
                return (
                  <div key={id} style={{padding:'7px 0',borderBottom:`1px solid ${T.borderLight}`}}>
                    <div style={{display:'flex',justifyContent:'space-between',fontSize:12}}>
                      <span style={{color:T.text,fontWeight:500}}>{inv.supplier} <span style={{color:T.dim,fontFamily:'monospace',fontSize:11}}>{inv.invoiceNo}</span></span>
                      <span style={{fontWeight:700,color:T.text}}>{fmtSGD(inv.total)}</span>
                    </div>
                    {alreadyPaid>0&&(
                      <div style={{display:'flex',justifyContent:'space-between',fontSize:11,marginTop:3}}>
                        <span style={{color:T.success}}>✓ Deposit paid</span>
                        <span style={{color:T.success,fontWeight:600}}>−{fmtSGD(alreadyPaid)}</span>
                      </div>
                    )}
                    {alreadyPaid>0&&(
                      <div style={{display:'flex',justifyContent:'space-between',fontSize:11,marginTop:1}}>
                        <span style={{color:T.dim}}>Balance remaining</span>
                        <span style={{color:balance>0?T.danger:T.success,fontWeight:600}}>{fmtSGD(Math.max(0,balance))}</span>
                      </div>
                    )}
                  </div>
                );
              })}
              {(()=>{
                const combined=[...selected].reduce((s,id)=>s+(invoices.find(i=>i.id===id)?.total||0),0);
                const deposits=[...selected].reduce((s,id)=>{
                  const inv=invoices.find(i=>i.id===id);
                  return s+(inv?.paymentRecords||[]).reduce((ps,r)=>ps+(parseFloat(r.amount)||0),0);
                },0);
                const balance=combined-deposits;
                return (
                  <div style={{paddingTop:8}}>
                    <div style={{display:'flex',justifyContent:'space-between',fontSize:12,color:T.muted,marginBottom:3}}>
                      <span>Combined invoice total</span><span>{fmtSGD(combined)}</span>
                    </div>
                    {deposits>0&&(
                      <div style={{display:'flex',justifyContent:'space-between',fontSize:12,color:T.success,marginBottom:3}}>
                        <span>Total deposits already paid</span><span>−{fmtSGD(deposits)}</span>
                      </div>
                    )}
                    <div style={{display:'flex',justifyContent:'space-between',fontSize:14,fontWeight:800,color:T.accent,borderTop:`1px solid ${T.borderLight}`,paddingTop:6,marginTop:3}}>
                      <span>You need to pay</span><span>{fmtSGD(Math.max(0,balance))}</span>
                    </div>
                  </div>
                );
              })()}
            </div>
            <div>
              <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Batch Label *</label>
              <input value={batchName} onChange={e=>setBatchName(e.target.value)}
                placeholder="e.g. Carpentry Works + VO" style={iStyle}/>
              <div style={{fontSize:11,color:T.dim,marginTop:4}}>Give it a name that describes what these invoices cover together</div>
            </div>
            <div style={{display:'flex',justifyContent:'flex-end',gap:10}}>
              <Btn variant="secondary" onClick={()=>{setBatchModal(false);setBatchName('');}}>Cancel</Btn>
              <Btn onClick={createBatch} disabled={!batchName.trim()}>
                <Receipt size={13}/>Create Batch
              </Btn>
            </div>
          </div>
        </Modal>
      )}

      {/* Record Batch Payment Modal */}
      {batchPayModal&&(()=>{
        const batch=(invoiceBatches||[]).find(b=>b.id===batchPayModal);
        if(!batch) return null;
        const totalPaid=(batch.paymentRecords||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
        const outstanding=batch.totalAmount-totalPaid;
        const pf=k=>v=>setBatchPayForm(p=>({...p,[k]:v}));
        const methodLabels={paynow:'PayNow',bank_transfer:'Bank Transfer',cash:'Cash'};
        const methodIcons={paynow:'📱',bank_transfer:'🏦',cash:'💵'};
        return (
          <Modal title={`Record Payment — ${batch.label}`} onClose={()=>setBatchPayModal(null)}>
            <input ref={batchProofRef} type="file" accept="image/*,application/pdf" capture="environment"
              style={{display:'none'}} onChange={e=>{
                const f=e.target.files?.[0];if(!f)return;
                extractProof(f,setBatchPayForm);e.target.value='';
              }}/>
            <div style={{display:'flex',flexDirection:'column',gap:14}}>

              {/* STEP 1 — Upload proof first */}
              <div>
                <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:8,display:'flex',alignItems:'center',gap:6}}>
                  <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>1</span>
                  Upload Payment Screenshot {(acctSettings?.anthropicApiKey||'').trim()&&<span style={{fontSize:10,color:T.accent,fontWeight:400}}>— AI fills details below</span>}
                </div>
                <DropZone accept="image/*,application/pdf" onDrop={f=>extractProof(f,setBatchPayForm)}>
                {batchPayForm.proofImage?(
                  <div style={{position:'relative',background:T.bg,borderRadius:12,border:`1px solid ${T.borderLight}`,overflow:'hidden'}}>
                    {batchPayForm.proofImage.startsWith('data:application/pdf')
                      ?<div style={{padding:16,display:'flex',alignItems:'center',gap:10}}>
                          <span style={{fontSize:32}}>📄</span>
                          <div><div style={{fontSize:13,fontWeight:600,color:T.text}}>{proofOcrLoading?'AI reading...':'PDF uploaded'}</div></div>
                        </div>
                      :<img src={batchPayForm.proofImage} alt="Proof" style={{width:'100%',maxHeight:180,objectFit:'contain',display:'block'}}/>
                    }
                    {proofOcrLoading&&<div style={{position:'absolute',bottom:0,left:0,right:0,background:'rgba(0,113,227,0.9)',padding:'8px 14px',display:'flex',alignItems:'center',gap:8}}>
                      <Loader2 size={13} style={{color:'#fff',animation:'spin 1s linear infinite'}}/><span style={{fontSize:12,color:'#fff',fontWeight:600}}>Reading payment details...</span>
                    </div>}
                    <button type="button" onClick={()=>setBatchPayForm(p=>({...p,proofImage:null}))}
                      style={{position:'absolute',top:8,right:8,background:T.danger,border:'none',borderRadius:'50%',width:26,height:26,cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'center',color:'#fff'}}>
                      <X size={13}/>
                    </button>
                  </div>
                ):(
                  <div onClick={()=>batchProofRef.current?.click()}
                    style={{border:`2px dashed ${T.borderLight}`,borderRadius:12,padding:'20px',textAlign:'center',cursor:'pointer',background:T.bg,display:'flex',flexDirection:'column',alignItems:'center',gap:8}}>
                    <div style={{width:44,height:44,background:T.accentLight,borderRadius:12,display:'flex',alignItems:'center',justifyContent:'center'}}>
                      <Upload size={18} style={{color:T.accent}}/>
                    </div>
                    <div style={{fontSize:13,fontWeight:600,color:T.text}}>Drop screenshot or tap to upload</div>
                    <div style={{fontSize:11,color:T.dim}}>PayNow / Bank Transfer receipt — JPG, PNG or PDF</div>
                  </div>
                )}
                </DropZone>
              </div>

              {/* STEP 2 — Summary */}
              <div>
                <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:8,display:'flex',alignItems:'center',gap:6}}>
                  <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>2</span>
                  Verify Payment Details
                </div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:8,marginBottom:10}}>
                  {[{l:'Batch Total',v:fmtSGD(batch.totalAmount),c:T.text},{l:'Already Paid',v:fmtSGD(totalPaid),c:T.success},{l:'Outstanding',v:fmtSGD(Math.max(0,outstanding)),c:outstanding>0?T.danger:T.success}].map(({l,v,c})=>(
                    <div key={l} style={{background:T.bg,borderRadius:10,padding:'10px 12px'}}>
                      <div style={{fontSize:10,color:T.dim,marginBottom:3}}>{l}</div>
                      <div style={{fontSize:15,fontWeight:700,color:c}}>{v}</div>
                    </div>
                  ))}
                </div>
                <div style={{height:5,background:T.borderLight,borderRadius:3,overflow:'hidden',marginBottom:12}}>
                  <div style={{height:'100%',borderRadius:3,width:`${Math.min(100,batch.totalAmount>0?totalPaid/batch.totalAmount*100:0)}%`,background:outstanding>0?T.warning:T.success}}/>
                </div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:10,marginBottom:10}}>
                  <div>
                    <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Amount (S$) *</label>
                    <input type="number" value={batchPayForm.amount} onChange={e=>pf('amount')(e.target.value)} placeholder="0.00" style={iStyle}/>
                    <div style={{display:'flex',gap:8,marginTop:6}}>
                      {[50,100].map(pct=>(
                        <button key={pct} type="button" onClick={()=>pf('amount')(String(Math.round(outstanding*pct/100*100)/100))}
                          style={{fontSize:11,padding:'3px 9px',borderRadius:7,border:`1px solid ${T.borderLight}`,background:T.bg,cursor:'pointer',fontFamily:'inherit',color:T.muted,fontWeight:600}}>{pct}%</button>
                      ))}
                      <button type="button" onClick={()=>pf('amount')(String(Math.round(Math.max(0,outstanding)*100)/100))}
                        style={{fontSize:11,padding:'3px 9px',borderRadius:7,border:`1px solid ${T.borderLight}`,background:T.bg,cursor:'pointer',fontFamily:'inherit',color:T.muted,fontWeight:600}}>Full</button>
                    </div>
                  </div>
                  <Field label="Date Paid" type="date" value={batchPayForm.date} onChange={pf('date')}/>
                </div>
                <div style={{marginBottom:10}}>
                  <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:8}}>Payment Method</label>
                  <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:8}}>
                    {[{id:'paynow',label:'PayNow',icon:'📱'},{id:'bank_transfer',label:'Bank Transfer',icon:'🏦'},{id:'cash',label:'Cash',icon:'💵'}].map(m=>(
                      <button key={m.id} type="button" onClick={()=>pf('method')(m.id)}
                        style={{padding:'8px',borderRadius:10,border:`2px solid ${batchPayForm.method===m.id?T.accent:T.borderLight}`,background:batchPayForm.method===m.id?T.accentLight:'transparent',cursor:'pointer',fontFamily:'inherit',display:'flex',flexDirection:'column',alignItems:'center',gap:3}}>
                        <span style={{fontSize:18}}>{m.icon}</span>
                        <span style={{fontSize:11,fontWeight:600,color:batchPayForm.method===m.id?T.accent:T.muted}}>{m.label}</span>
                      </button>
                    ))}
                  </div>
                </div>
                <Field label={batchPayForm.method==='paynow'?'PayNow / UEN':batchPayForm.method==='bank_transfer'?'Bank Account':'Receipt / Reference'}
                  value={batchPayForm.reference} onChange={pf('reference')}
                  placeholder={batchPayForm.method==='paynow'?'e.g. 202320231N':batchPayForm.method==='bank_transfer'?'e.g. DBS 0721-0976-05':'e.g. Receipt #001'}/>
                <div style={{marginTop:10}}>
                  <Field label="Notes (optional)" value={batchPayForm.notes} onChange={pf('notes')} placeholder="e.g. 50% deposit"/>
                </div>
              </div>

              <div style={{display:'flex',justifyContent:'flex-end',gap:10,borderTop:`1px solid ${T.borderLight}`,paddingTop:14}}>
                <Btn variant="secondary" onClick={()=>setBatchPayModal(null)}>Cancel</Btn>
                <Btn onClick={recordBatchPayment} disabled={!batchPayForm.amount||parseFloat(batchPayForm.amount)<=0}>
                  <CheckCircle size={13}/>Record Payment
                </Btn>
              </div>
            </div>
          </Modal>
        );
      })()}

      {lightbox?.type==='img'&&<Lightbox src={lightbox.src} title={lightbox.title} onClose={()=>setLightbox(null)}/>}

      {/* Payment History Modal */}
      {lightbox?.type==='payments'&&(()=>{
        const inv=lightbox.inv;
        const records=inv.paymentRecords||[];
        const totalPaid=records.reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
        const outstanding=inv.total-totalPaid;
        const icons={paynow:'📱',bank_transfer:'🏦',cash:'💵'};
        const labels={paynow:'PayNow',bank_transfer:'Bank Transfer',cash:'Cash'};
        return (
          <Modal title={`Payment History — ${inv.invoiceNo}`} onClose={()=>setLightbox(null)}>
            <div style={{display:'flex',flexDirection:'column',gap:14}}>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:8}}>
                {[
                  {l:'Invoice Total',v:fmtSGD(inv.total),c:T.text},
                  {l:'Total Paid',v:fmtSGD(totalPaid),c:T.success},
                  {l:'Outstanding',v:fmtSGD(Math.max(0,outstanding)),c:outstanding>0?T.danger:T.success},
                ].map(({l,v,c})=>(
                  <div key={l} style={{background:T.bg,borderRadius:10,padding:'10px 12px'}}>
                    <div style={{fontSize:10,color:T.dim,marginBottom:3}}>{l}</div>
                    <div style={{fontSize:15,fontWeight:700,color:c}}>{v}</div>
                  </div>
                ))}
              </div>
              <div style={{height:6,background:T.borderLight,borderRadius:3,overflow:'hidden'}}>
                <div style={{height:'100%',width:`${Math.min(100,inv.total>0?totalPaid/inv.total*100:0)}%`,
                  background:outstanding>0?T.warning:T.success,borderRadius:3}}/>
              </div>
              <div style={{display:'flex',flexDirection:'column',gap:8}}>
                {records.length===0&&<div style={{fontSize:13,color:T.dim,textAlign:'center',padding:20}}>No payments recorded yet</div>}
                {records.map(r=>(
                  <div key={r.id} style={{background:T.bg,borderRadius:12,padding:'12px 14px',display:'flex',alignItems:'flex-start',gap:12}}>
                    <div style={{fontSize:22,flexShrink:0}}>{icons[r.method]||'💰'}</div>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:3}}>
                        <span style={{fontSize:14,fontWeight:700,color:T.success}}>{fmtSGD(r.amount)}</span>
                        <span style={{fontSize:11,color:T.dim}}>{fmtDate(r.date)}</span>
                      </div>
                      <div style={{fontSize:12,color:T.muted,fontWeight:600}}>{labels[r.method]||r.method}</div>
                      {r.reference&&<div style={{fontSize:11,color:T.dim,marginTop:2,fontFamily:'monospace'}}>Ref: {r.reference}</div>}
                      {r.notes&&<div style={{fontSize:11,color:T.dim,marginTop:2}}>{r.notes}</div>}
                      {r.proofImage&&<button type="button" onClick={()=>setLightbox({type:'img',src:r.proofImage,title:'Payment Screenshot'})}
                        style={{marginTop:5,background:T.successLight,border:'none',borderRadius:6,padding:'2px 9px',cursor:'pointer',fontSize:11,color:T.success,fontFamily:'inherit',fontWeight:600,display:'inline-flex',alignItems:'center',gap:4}}>
                        <ZoomIn size={10}/>View Screenshot</button>}
                    </div>
                    {isAdmin&&<button type="button" onClick={()=>deletePayRecord(inv.id,r.id)}
                      style={{background:'none',border:'none',cursor:'pointer',color:T.dim,padding:4,borderRadius:6,flexShrink:0}}>
                      <Trash2 size={12}/></button>}
                  </div>
                ))}
              </div>
              {isAdmin&&inv.status!=='Paid'&&(
                <Btn onClick={()=>{setLightbox(null);const rem=inv.total-(inv.paymentRecords||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);setPayModal(inv.id);setPayForm({amount:String(Math.round(Math.max(0,rem)*100)/100),date:new Date().toISOString().slice(0,10),method:'paynow',reference:'',notes:''});}}>
                  <CreditCard size={13}/>Record Another Payment
                </Btn>
              )}
            </div>
          </Modal>
        );
      })()}

      {/* Record Payment Modal */}
      {payModal&&(()=>{
        const inv=invoices.find(i=>i.id===payModal);
        if(!inv) return null;
        const totalPaid=(inv.paymentRecords||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
        const outstanding=inv.total-totalPaid;
        const pf=k=>v=>setPayForm(p=>({...p,[k]:v}));
        const methodLabels={paynow:'PayNow',bank_transfer:'Bank Transfer',cash:'Cash'};
        const methodIcons={paynow:'📱',bank_transfer:'🏦',cash:'💵'};
        return (
          <Modal title={`Record Payment — ${inv.invoiceNo}`} onClose={()=>setPayModal(null)}>
            <input ref={payProofRef} type="file" accept="image/*,application/pdf" capture="environment"
              style={{display:'none'}} onChange={e=>{
                const f=e.target.files?.[0];if(!f)return;
                extractProof(f,setPayForm);e.target.value='';
              }}/>
            <div style={{display:'flex',flexDirection:'column',gap:14}}>

              {/* STEP 1 — Upload proof first */}
              <div>
                <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:8,display:'flex',alignItems:'center',gap:6}}>
                  <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>1</span>
                  Upload Payment Screenshot {(acctSettings?.anthropicApiKey||'').trim()&&<span style={{fontSize:10,color:T.accent,fontWeight:400}}>— AI fills details below</span>}
                </div>
                <DropZone accept="image/*,application/pdf" onDrop={f=>extractProof(f,setPayForm)}>
                {payForm.proofImage?(
                  <div style={{position:'relative',background:T.bg,borderRadius:12,border:`1px solid ${T.borderLight}`,overflow:'hidden'}}>
                    {payForm.proofImage.startsWith('data:application/pdf')
                      ?<div style={{padding:16,display:'flex',alignItems:'center',gap:10}}>
                          <span style={{fontSize:32}}>📄</span>
                          <div><div style={{fontSize:13,fontWeight:600,color:T.text}}>{proofOcrLoading?'AI reading...':'PDF uploaded'}</div></div>
                        </div>
                      :<img src={payForm.proofImage} alt="Proof" style={{width:'100%',maxHeight:180,objectFit:'contain',display:'block'}}/>
                    }
                    {proofOcrLoading&&<div style={{position:'absolute',bottom:0,left:0,right:0,background:'rgba(0,113,227,0.9)',padding:'8px 14px',display:'flex',alignItems:'center',gap:8}}>
                      <Loader2 size={13} style={{color:'#fff',animation:'spin 1s linear infinite'}}/><span style={{fontSize:12,color:'#fff',fontWeight:600}}>Reading payment details...</span>
                    </div>}
                    <button type="button" onClick={()=>setPayForm(p=>({...p,proofImage:null}))}
                      style={{position:'absolute',top:8,right:8,background:T.danger,border:'none',borderRadius:'50%',width:26,height:26,cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'center',color:'#fff'}}>
                      <X size={13}/>
                    </button>
                  </div>
                ):(
                  <div onClick={()=>payProofRef.current?.click()}
                    style={{border:`2px dashed ${T.borderLight}`,borderRadius:12,padding:'20px',textAlign:'center',cursor:'pointer',background:T.bg,display:'flex',flexDirection:'column',alignItems:'center',gap:8}}>
                    <div style={{width:44,height:44,background:T.accentLight,borderRadius:12,display:'flex',alignItems:'center',justifyContent:'center'}}>
                      <Upload size={18} style={{color:T.accent}}/>
                    </div>
                    <div style={{fontSize:13,fontWeight:600,color:T.text}}>Drop screenshot or tap to upload</div>
                    <div style={{fontSize:11,color:T.dim}}>PayNow / Bank Transfer receipt — JPG, PNG or PDF</div>
                  </div>
                )}
                </DropZone>
              </div>

              {/* STEP 2 — Verify details */}
              <div>
                <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:8,display:'flex',alignItems:'center',gap:6}}>
                  <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>2</span>
                  Verify Payment Details
                </div>
                <div style={{background:T.bg,borderRadius:12,padding:'12px 14px',marginBottom:10}}>
                  <div style={{fontSize:12,color:T.muted,marginBottom:2}}>{inv.supplier}</div>
                  <div style={{display:'flex',justifyContent:'space-between',alignItems:'baseline'}}>
                    <span style={{fontSize:13,color:T.muted}}>Outstanding</span>
                    <span style={{fontSize:18,fontWeight:800,color:outstanding>0?T.danger:T.success}}>{fmtSGD(Math.max(0,outstanding))}</span>
                  </div>
                </div>
                <Field label="Amount Paying (S$) *" type="number" value={payForm.amount} onChange={pf('amount')} placeholder="0.00"/>
                <div style={{marginTop:10}}>
                  <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:8}}>Payment Method</label>
                  <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:8}}>
                    {[{id:'paynow',label:'PayNow',icon:'📱'},{id:'bank_transfer',label:'Bank Transfer',icon:'🏦'},{id:'cash',label:'Cash',icon:'💵'}].map(m=>(
                      <button key={m.id} type="button" onClick={()=>pf('method')(m.id)}
                        style={{padding:'8px',borderRadius:10,border:`2px solid ${payForm.method===m.id?T.accent:T.borderLight}`,background:payForm.method===m.id?T.accentLight:'transparent',cursor:'pointer',fontFamily:'inherit',display:'flex',flexDirection:'column',alignItems:'center',gap:3,transition:'all 0.12s'}}>
                        <span style={{fontSize:18}}>{m.icon}</span>
                        <span style={{fontSize:11,fontWeight:600,color:payForm.method===m.id?T.accent:T.muted}}>{m.label}</span>
                      </button>
                    ))}
                  </div>
                </div>
                <div style={{marginTop:10}}>
                  <Field label={payForm.method==='paynow'?'PayNow / UEN':payForm.method==='bank_transfer'?'Bank Account':'Receipt / Reference'}
                    value={payForm.reference} onChange={pf('reference')}
                    placeholder={payForm.method==='paynow'?'e.g. 202320231N':payForm.method==='bank_transfer'?'e.g. DBS 0721-0976-05':'e.g. Receipt #001'}/>
                </div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:12,marginTop:10}}>
                  <Field label="Date Paid" type="date" value={payForm.date} onChange={pf('date')}/>
                  <Field label="Notes (optional)" value={payForm.notes} onChange={pf('notes')} placeholder="e.g. Balance payment"/>
                </div>
              </div>

              <div style={{display:'flex',justifyContent:'flex-end',gap:10,borderTop:`1px solid ${T.borderLight}`,paddingTop:14}}>
                <Btn variant="secondary" onClick={()=>setPayModal(null)}>Cancel</Btn>
                <Btn onClick={recordPayment} disabled={!payForm.amount||parseFloat(payForm.amount)<=0}>
                  <CheckCircle size={13}/>Record Payment
                </Btn>
              </div>
            </div>
          </Modal>
        );
      })()}
    </div>
  );
}
function Payments({payments,setPayments,projects,invoices,isAdmin,onSoftDelete,onShowToast,acctSettings,logAction=()=>{}}){
  const [modal,setModal]=useState(false);
  const [form,setForm]=useState({projectId:'',type:'Deposit',amount:'',date:'',status:'Received',paymentMethod:'',reference:''});
  const [deleteTarget,setDeleteTarget]=useState(null);
  const [receipt,setReceipt]=useState(null);
  const [ocrLoading,setOcrLoading]=useState(false);
  const payFileRef=useRef();
  const payCamRef=useRef();
  const ff=k=>v=>setForm(p=>({...p,[k]:v}));

  const extractPayment=async(file)=>{
    // Store image immediately, compress in background for sync
    const reader=new FileReader();
    reader.onload=ev=>setReceipt(ev.target.result);
    reader.readAsDataURL(file);
    compressForSync(file).then(c=>{if(c)setReceipt(c);});

    const apiKey=(acctSettings?.anthropicApiKey||'').trim();
    if(!apiKey) return; // no key — just store image

    setOcrLoading(true);
    try{
      const b64=await toB64(file);
      const isPDF=file.type==='application/pdf';
      const part=isPDF
        ?{type:'document',source:{type:'base64',media_type:'application/pdf',data:b64}}
        :{type:'image',source:{type:'base64',media_type:file.type,data:b64}};
      const res=await fetch('https://api.anthropic.com/v1/messages',{
        method:'POST',
        headers:{'Content-Type':'application/json','x-api-key':apiKey,
          'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},
        body:JSON.stringify({
          model:'claude-opus-4-5',max_tokens:400,
          messages:[{role:'user',content:[part,{type:'text',
            text:'Extract from this payment receipt/screenshot. Return ONLY valid JSON:\n{"amount":0,"date":"YYYY-MM-DD","method":"paynow|bank_transfer|cash","reference":"","notes":""}\nRules: amount=total paid (numeric). date=YYYY-MM-DD. method=paynow if PayNow/UEN/QR, bank_transfer if bank/FAST/GIRO, cash if cash. reference=PayNow number/UEN/bank account/transaction ref. notes=any other useful info. Use empty string if not found.'}]}]
        })
      });
      const data=await res.json();
      if(data.error) throw new Error(data.error.message);
      const txt=(data.content||[]).map(c=>c.text||'').join('').replace(/```json|```/g,'').trim();
      const parsed=JSON.parse(txt);
      setForm(prev=>({
        ...prev,
        amount:parsed.amount?String(parsed.amount):prev.amount,
        date:parsed.date||prev.date,
        paymentMethod:parsed.method||prev.paymentMethod,
        reference:parsed.reference||prev.reference,
      }));
    }catch(e){console.warn('Payment OCR failed:',e);}
    finally{setOcrLoading(false);}
  };

  const save_=()=>{
    if(!form.projectId||!form.amount)return;
    const pay={...form,id:uid(),amount:parseFloat(form.amount)||0,
      receiptNo:`RCP-${new Date().getFullYear()}-${String(Date.now()).slice(-5)}`};
    const upd=[...payments,pay];
    setPayments(upd);saveS('payments',upd);setModal(false);setReceipt(null);
    const projName=projects.find(p=>p.id===pay.projectId)?.name||'Unknown project';
    logAction('CREATE_PAYMENT',`Recorded ${pay.type} payment $${pay.amount.toLocaleString('en-SG',{minimumFractionDigits:2})} — ${projName} — Status: ${pay.status}`);
    if(pay.status==='Received'){
      const proj=projects.find(p=>p.id===pay.projectId);
      setTimeout(()=>{
        if(window.confirm('Payment recorded! Print client receipt now?')){
          printDoc(buildReceiptHTML(pay,proj,pay.receiptNo,getCo(acctSettings)),`Receipt ${pay.receiptNo}`);
        }
      },300);
    }
  };
  const del_=(id)=>{
    if(!isAdmin){return;}
    const pay=payments.find(p=>p.id===id);
    if(pay)setDeleteTarget(pay);
  };
  const confirmSoftDelete=(pay)=>{
    onSoftDelete({...pay,_trashType:'payment',_deletedAt:new Date().toISOString()});
    const upd=payments.filter(p=>p.id!==pay.id);setPayments(upd);saveS('payments',upd);
    setDeleteTarget(null);
    onShowToast(`${pay.type} payment of ${fmtSGD(pay.amount)} moved to Trash`,()=>{setPayments(prev=>[...prev,pay]);saveS('payments',[...payments,pay]);});
  };

  const totalOut=projects.reduce((s,p)=>{
    const rev=p.contractAmount+(p.variationOrders||0);
    const recv=payments.filter(py=>py.projectId===p.id&&py.status==='Received').reduce((s2,py)=>s2+py.amount,0);
    return s+Math.max(0,rev-recv);
  },0);

  return (
    <div style={{display:'flex',flexDirection:'column',gap:16}}>
      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',flexWrap:'wrap',gap:10}}>
        <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:12,padding:'10px 18px'}}>
          <span style={{fontSize:12,color:T.dim}}>Total Outstanding Receivables: </span>
          <span style={{fontSize:15,fontWeight:700,color:totalOut>0?T.accent:T.success}}>{fmtSGD(totalOut)}</span>
        </div>
        <Btn onClick={()=>{setForm({projectId:'',type:'Deposit',amount:'',date:'',status:'Received',paymentMethod:'',reference:''});setReceipt(null);setModal(true);}}>
          <Plus size={13}/>Record Payment
        </Btn>
      </div>

      {/* Outstanding Summary — projects with balance due */}
      {(()=>{
        const outstanding=projects.filter(p=>{
          if(p.archived||p.status==='Cancelled') return false;
          const rev=p.contractAmount+(p.variationOrders||0);
          const recv=payments.filter(py=>py.projectId===p.id&&py.status==='Received').reduce((s,py)=>s+py.amount,0);
          return rev-recv>0;
        }).map(p=>{
          const rev=p.contractAmount+(p.variationOrders||0);
          const recv=payments.filter(py=>py.projectId===p.id&&py.status==='Received').reduce((s,py)=>s+py.amount,0);
          const out=rev-recv;
          const pct=rev>0?Math.min(recv/rev*100,100):0;
          const isOverdue=p.endDate&&new Date(p.endDate)<new Date();
          return {p,rev,recv,out,pct,isOverdue};
        }).sort((a,b)=>b.out-a.out);
        if(!outstanding.length) return null;
        return (
          <div style={{background:T.card,border:`2px solid ${T.accent}25`,borderRadius:18,overflow:'hidden',boxShadow:T.shadow}}>
            <div style={{padding:'14px 20px',borderBottom:`1px solid ${T.borderLight}`,display:'flex',alignItems:'center',justifyContent:'space-between'}}>
              <div style={{display:'flex',alignItems:'center',gap:8}}>
                <Bell size={14} style={{color:T.accent}}/>
                <span style={{fontSize:14,fontWeight:700,color:T.text}}>Outstanding Balances</span>
                <Badge color={T.accent}>{outstanding.length} project{outstanding.length!==1?'s':''}</Badge>
              </div>
              <span style={{fontSize:15,fontWeight:800,color:T.accent}}>{fmtSGD(outstanding.reduce((s,o)=>s+o.out,0))}</span>
            </div>
            {outstanding.map(({p,rev,recv,out,pct,isOverdue})=>(
              <div key={p.id} style={{padding:'12px 20px',borderBottom:`1px solid ${T.borderLight}`,display:'flex',alignItems:'center',gap:12}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:3,flexWrap:'wrap'}}>
                    <span style={{fontSize:13,fontWeight:700,color:T.text}}>{p.name}</span>
                    <span style={{fontSize:12,color:T.muted}}>{p.client}</span>
                    {isOverdue&&<Badge color={T.danger} sm>Overdue</Badge>}
                  </div>
                  <div style={{display:'flex',alignItems:'center',gap:8}}>
                    <div style={{flex:1,height:5,background:T.bg,borderRadius:3,maxWidth:160,overflow:'hidden'}}>
                      <div style={{height:'100%',width:`${pct}%`,background:pct>=100?T.success:T.accent,borderRadius:3,transition:'width .3s'}}/>
                    </div>
                    <span style={{fontSize:11,color:T.dim,whiteSpace:'nowrap'}}>{pct.toFixed(0)}% collected</span>
                  </div>
                </div>
                <div style={{textAlign:'right',flexShrink:0}}>
                  <div style={{fontSize:15,fontWeight:800,color:isOverdue?T.danger:T.accent}}>{fmtSGD(out)}</div>
                  <div style={{fontSize:10,color:T.dim}}>outstanding</div>
                </div>
                <button type="button"
                  onClick={()=>{
                    const co=getCo(acctSettings);
                    const payNow=co.payNowUen?`\n\n📱 *PayNow UEN:* ${co.payNowUen}`:'';
                    const bank=co.bankName?`\n🏦 *Bank Transfer:* ${co.bankName} ${co.bankAccount||''}`:'';
                    const msg=`Dear ${p.client},\n\nThank you for choosing ${co.companyName}.\n\nThis is a friendly reminder that your renovation project *${p.name}* has an outstanding balance of *${fmtSGD(out)}*.\n\nContract Value: ${fmtSGD(rev)}\nAmount Paid: ${fmtSGD(recv)}\nOutstanding: *${fmtSGD(out)}*\n\nPlease arrange payment at your earliest convenience.${payNow}${bank}\n\nThank you for your prompt attention.\n\nBest regards,\n${co.companyName}\n${co.companyPhone||''}`;
                    navigator.clipboard.writeText(msg).then(()=>alert('✓ Reminder copied! Paste into WhatsApp.')).catch(()=>{const ta=document.createElement('textarea');ta.value=msg;document.body.appendChild(ta);ta.select();document.execCommand('copy');document.body.removeChild(ta);alert('✓ Reminder copied! Paste into WhatsApp.');});
                  }}
                  style={{display:'flex',alignItems:'center',gap:5,padding:'6px 11px',borderRadius:9,border:`1px solid #25D366`,background:'rgba(37,211,102,0.08)',cursor:'pointer',fontFamily:'inherit',fontSize:12,fontWeight:600,color:'#128C7E',whiteSpace:'nowrap',flexShrink:0}}>
                  <span style={{fontSize:13}}>📲</span>WhatsApp
                </button>
              </div>
            ))}
          </div>
        );
      })()}

      {projects.map(proj=>{
        const pp=payments.filter(p=>p.projectId===proj.id);
        if(pp.length===0)return null;
        const projInv=invoices.filter(i=>i.projectId===proj.id);
        const rev=proj.contractAmount+(proj.variationOrders||0);
        const recv=pp.filter(p=>p.status==='Received').reduce((s,p)=>s+p.amount,0);
        const out=rev-recv;
        const pct=rev>0?Math.min(recv/rev*100,100):0;
        return (
          <div key={proj.id} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow}}>
            <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:14}}>
              <div>
                <div style={{fontSize:14,fontWeight:700,color:T.text}}>{proj.name}</div>
                <div style={{fontSize:12,color:T.muted}}>{proj.client}</div>
              </div>
              <div style={{display:'flex',alignItems:'center',gap:8,flexWrap:'wrap',justifyContent:'flex-end'}}>
                {out>0&&(
                  <button type="button"
                    onClick={()=>{
                      const co=getCo(acctSettings);
                      const payNow=co.payNowUen?`\n\n📱 *PayNow UEN:* ${co.payNowUen}`:'';
                      const bank=co.bankName?`\n🏦 *Bank Transfer:* ${co.bankName} ${co.bankAccount||''}`:'';
                      const msg=`Dear ${proj.client},\n\nThank you for choosing ${co.companyName}.\n\nThis is a friendly reminder that your renovation project *${proj.name}* has an outstanding balance of *${fmtSGD(out)}*.\n\nContract Value: ${fmtSGD(proj.contractAmount+(proj.variationOrders||0))}\nAmount Paid: ${fmtSGD(recv)}\nOutstanding: *${fmtSGD(out)}*\n\nPlease arrange payment at your earliest convenience.${payNow}${bank}\n\nThank you for your prompt attention.\n\nBest regards,\n${co.companyName}\n${co.companyPhone||''}`;
                      navigator.clipboard.writeText(msg).then(()=>alert('✓ Reminder copied! Paste into WhatsApp.')).catch(()=>{
                        const ta=document.createElement('textarea');ta.value=msg;document.body.appendChild(ta);ta.select();document.execCommand('copy');document.body.removeChild(ta);
                        alert('✓ Reminder copied! Paste into WhatsApp.');
                      });
                    }}
                    style={{display:'flex',alignItems:'center',gap:6,padding:'6px 12px',borderRadius:9,border:`1px solid #25D366`,background:'rgba(37,211,102,0.08)',cursor:'pointer',fontFamily:'inherit',fontSize:12,fontWeight:600,color:'#128C7E',whiteSpace:'nowrap'}}>
                    <span style={{fontSize:14}}>📲</span>Copy WhatsApp Reminder
                  </button>
                )}
                <Btn variant="secondary" size="sm"
                  onClick={()=>printDoc(buildSOAHTML(proj,projInv,pp,getCo(acctSettings)),`SOA — ${proj.name}`)}>
                  Statement of Account
                </Btn>
                <div style={{textAlign:'right'}}>
                  <div style={{fontSize:11,color:T.dim}}>Outstanding</div>
                  <div style={{fontSize:16,fontWeight:700,color:out>0?T.accent:T.success}}>{fmtSGD(out)}</div>
                </div>
              </div>
            </div>
            <div style={{display:'flex',flexDirection:'column',gap:8,marginBottom:14}}>
              {pp.map(pay=>(
                <div key={pay.id} style={{display:'flex',alignItems:'center',justifyContent:'space-between',
                  background:T.bg,borderRadius:12,padding:'10px 14px'}}>
                  <div style={{display:'flex',alignItems:'center',gap:10}}>
                    <div style={{width:8,height:8,borderRadius:'50%',flexShrink:0,
                      background:pay.status==='Received'?T.success:T.warning}}/>
                    <span style={{fontSize:13,fontWeight:600,color:T.text}}>{pay.type} Payment</span>
                    <span style={{fontSize:12,color:T.dim}}>{fmtDate(pay.date)}</span>
                  </div>
                  <div style={{display:'flex',alignItems:'center',gap:10}}>
                    <span style={{fontSize:14,fontWeight:700,color:T.text}}>{fmtSGD(pay.amount)}</span>
                    <Badge color={ST_CLR[pay.status]||T.dim}>{pay.status}</Badge>
                    {pay.status==='Received'&&pay.receiptNo&&(
                      <button onClick={()=>printDoc(buildReceiptHTML(pay,proj,pay.receiptNo,getCo(acctSettings)),`Receipt ${pay.receiptNo}`)}
                        style={{background:T.successLight,border:'none',cursor:'pointer',color:T.success,
                          display:'flex',padding:'3px 8px',borderRadius:5,fontSize:11,fontWeight:600,
                          fontFamily:'inherit',alignItems:'center',gap:4}}>
                        Receipt
                      </button>
                    )}
                    {pay.proofImage&&(
                      <button onClick={()=>setLightbox({type:'img',src:pay.proofImage,title:`Payment Screenshot — ${pay.type}`})}
                        style={{background:T.infoLight,border:'none',cursor:'pointer',color:T.info,
                          display:'flex',padding:'3px 8px',borderRadius:5,fontSize:11,fontWeight:600,
                          fontFamily:'inherit',alignItems:'center',gap:4}}>
                        <Eye size={11}/>View
                      </button>
                    )}
                    <button onClick={()=>del_(pay.id)}
                      style={{background:'none',border:'none',cursor:'pointer',color:T.dim,display:'flex'}}><Trash2 size={13}/></button>
                  </div>
                </div>
              ))}
            </div>
            <div>
              <div style={{display:'flex',justifyContent:'space-between',fontSize:12,color:T.dim,marginBottom:6}}>
                <span>Collection Progress</span>
                <span>{fmtSGD(recv)} of {fmtSGD(rev)} ({pct.toFixed(0)}%)</span>
              </div>
              <div style={{height:6,background:T.bg,borderRadius:3,overflow:'hidden'}}>
                <div style={{height:'100%',width:`${pct}%`,background:T.success,borderRadius:3,transition:'width .5s'}}/>
              </div>
            </div>
          </div>
        );
      })}

      {projects.every(p=>!payments.find(py=>py.projectId===p.id))&&(
        <div style={{color:T.dim,fontSize:13,textAlign:'center',padding:36}}>No payments recorded yet</div>
      )}

      {modal&&(
        <Modal title="Record Client Payment" onClose={()=>{setModal(false);setReceipt(null);}} wide>
          <div style={{display:'flex',flexDirection:'column',gap:16}}>

            {/* STEP 1 — Upload screenshot */}
            <div>
              <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:8,display:'flex',alignItems:'center',gap:6}}>
                <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>1</span>
                Upload Payment Screenshot {acctSettings?.anthropicApiKey&&<span style={{fontSize:10,color:T.accent,fontWeight:400}}>— AI fills details below</span>}
              </div>
              <DropZone accept="image/*,application/pdf" onDrop={f=>extractPayment(f)}>
              <div onClick={()=>payFileRef.current?.click()}
                style={{border:`2px dashed ${receipt?T.success+'60':T.borderLight}`,borderRadius:12,
                  cursor:'pointer',background:T.bg,overflow:'hidden',
                  minHeight:receipt?'auto':120,display:'flex',alignItems:'center',justifyContent:'center',flexDirection:'column',gap:10,padding:receipt?0:20}}>
                {receipt?(
                  receipt.startsWith('data:application/pdf')?(
                    <div style={{padding:'16px 20px',display:'flex',alignItems:'center',gap:12}}>
                      <span style={{fontSize:36}}>📄</span>
                      <div><div style={{fontSize:13,fontWeight:600,color:T.text}}>{ocrLoading?'AI reading...':'PDF uploaded'}</div></div>
                    </div>
                  ):(
                    <img src={receipt} alt="Payment proof" style={{width:'100%',maxHeight:200,objectFit:'contain',display:'block'}}/>
                  )
                ):(
                  <>
                    <div style={{width:48,height:48,background:T.accentLight,borderRadius:12,display:'flex',alignItems:'center',justifyContent:'center'}}>
                      <Camera size={20} style={{color:T.accent}}/>
                    </div>
                    <div style={{textAlign:'center'}}>
                      <div style={{fontSize:13,fontWeight:700,color:T.text}}>Drop screenshot or tap to upload</div>
                      <div style={{fontSize:11,color:T.dim,marginTop:3}}>PayNow / Bank transfer / Cash receipt</div>
                    </div>
                  </>
                )}
              </div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:8,marginTop:8}}>
                <Btn variant="secondary" size="sm" onClick={()=>payFileRef.current?.click()} full><Upload size={12}/>Upload</Btn>
                <Btn variant="secondary" size="sm" onClick={()=>payCamRef.current?.click()} full><Camera size={12}/>Camera</Btn>
              </div>
              <input ref={payFileRef} type="file" accept="image/*,application/pdf" style={{display:'none'}}
                onChange={e=>{const f=e.target.files?.[0];if(f)extractPayment(f);e.target.value='';}}/>
              <input ref={payCamRef} type="file" accept="image/*" capture="environment" style={{display:'none'}}
                onChange={e=>{const f=e.target.files?.[0];if(f)extractPayment(f);e.target.value='';}}/>
              </DropZone>
              {receipt&&<button type="button" onClick={()=>setReceipt(null)}
                style={{marginTop:6,background:'none',border:'none',cursor:'pointer',color:T.dim,fontSize:12,display:'flex',alignItems:'center',gap:4}}>
                <X size={11}/>Remove
              </button>}
              {ocrLoading&&<div style={{marginTop:8,display:'flex',alignItems:'center',gap:8,fontSize:12,color:T.accent,background:T.accentLight,borderRadius:10,padding:'8px 12px'}}>
                <Loader2 size={13} style={{animation:'spin 1s linear infinite'}}/>AI extracting payment details...
              </div>}
            </div>

            {/* STEP 2 — Review and confirm fields */}
            <div>
              <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:12,display:'flex',alignItems:'center',gap:6}}>
                <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>2</span>
                Review & Confirm Details
                {ocrLoading&&<span style={{fontSize:10,color:T.accent,fontWeight:400}}>— filling from screenshot...</span>}
              </div>
              <div style={{display:'flex',flexDirection:'column',gap:12}}>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:12}}>
                  <Field label="Project *" value={form.projectId} onChange={ff('projectId')} as="select"
                    options={[{v:'',l:'Select project...'}, ...projects.map(p=>({v:p.id,l:p.name}))]}/>
                  <Field label="Payment Type" value={form.type} onChange={ff('type')} as="select" options={PAY_TYPES.map(t=>({v:t,l:t}))}/>
                </div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:12}}>
                  <Field label="Amount Received (S$) *" type="number" value={form.amount} onChange={ff('amount')} placeholder="0.00"/>
                  <Field label="Payment Date" type="date" value={form.date} onChange={ff('date')}/>
                </div>
                <div>
                  <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:8}}>Payment Method</label>
                  <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:8}}>
                    {[{id:'paynow',label:'PayNow',icon:'📱'},{id:'bank_transfer',label:'Bank Transfer',icon:'🏦'},{id:'cash',label:'Cash',icon:'💵'}].map(m=>(
                      <button key={m.id} type="button" onClick={()=>ff('paymentMethod')(m.id)}
                        style={{padding:'8px 6px',borderRadius:10,border:`2px solid ${form.paymentMethod===m.id?T.accent:T.borderLight}`,
                          background:form.paymentMethod===m.id?T.accentLight:'transparent',cursor:'pointer',fontFamily:'inherit',
                          display:'flex',flexDirection:'column',alignItems:'center',gap:3}}>
                        <span style={{fontSize:18}}>{m.icon}</span>
                        <span style={{fontSize:10,fontWeight:600,color:form.paymentMethod===m.id?T.accent:T.muted}}>{m.label}</span>
                      </button>
                    ))}
                  </div>
                </div>
                {form.paymentMethod&&(
                  <Field label={form.paymentMethod==='paynow'?'PayNow Number / UEN':form.paymentMethod==='bank_transfer'?'Bank Account Number':'Receipt / Reference'}
                    value={form.reference||''} onChange={ff('reference')}
                    placeholder={form.paymentMethod==='paynow'?'e.g. 202320231N':form.paymentMethod==='bank_transfer'?'e.g. DBS 0721-0976-05':'e.g. Receipt #001'}/>
                )}
                <Field label="Status" value={form.status} onChange={ff('status')} as="select" options={['Received','Pending'].map(s=>({v:s,l:s}))}/>
              </div>
            </div>

            <div style={{display:'flex',justifyContent:'flex-end',gap:10,borderTop:`1px solid ${T.borderLight}`,paddingTop:14}}>
              <Btn variant="secondary" onClick={()=>{setModal(false);setReceipt(null);}}>Cancel</Btn>
              <Btn onClick={save_} disabled={!form.projectId||!form.amount}>Record Payment</Btn>
            </div>
          </div>
        </Modal>
      )}
      {deleteTarget&&(
        <ConfirmDelete
          matchValue={`${deleteTarget.type}-${fmtSGD(deleteTarget.amount)}`}
          typeLabel="payment record"
          impact="Removing a payment record will affect the outstanding balance calculation for this project."
          onConfirm={()=>confirmSoftDelete(deleteTarget)}
          onClose={()=>setDeleteTarget(null)}
        />
      )}
    </div>
  );
}

function exportToExcel({projects,invoices,payments,dateFrom,dateTo}){
  const filterDate=d=>{
    if(!d)return true;
    if(dateFrom&&new Date(d)<new Date(dateFrom))return false;
    if(dateTo&&new Date(d)>new Date(dateTo+'T23:59:59'))return false;
    return true;
  };
  const escCSV=v=>{if(v==null)return '';const s=String(v);if(s.includes(',')||s.includes('"')||s.includes('\n'))return '"'+s.replace(/"/g,'""')+'"';return s;};
  const row=cols=>cols.map(escCSV).join(',');

  const projRows=[
    row(['Project Name','Client','Status','Contract (S$)','Variation Orders (S$)','Total Revenue (S$)','Total Expenses (S$)','Gross Profit (S$)','Net Margin %','Payments Received (S$)','Outstanding (S$)','Designer','PM','Start Date','End Date','Archived']),
    ...projects.filter(p=>filterDate(p.startDate)).map(p=>{
      const exp=invoices.filter(i=>i.projectId===p.id).reduce((s,i)=>s+i.total,0);
      const rev=p.contractAmount+(p.variationOrders||0);
      const recv=payments.filter(py=>py.projectId===p.id&&py.status==='Received').reduce((s,py)=>s+py.amount,0);
      const gross=rev-exp; const margin=rev>0?gross/rev*100:0;
      return row([p.name,p.client,p.status,p.contractAmount,p.variationOrders||0,rev,exp.toFixed(2),gross.toFixed(2),margin.toFixed(1),recv.toFixed(2),Math.max(0,rev-recv).toFixed(2),p.designer,p.pm,p.startDate,p.endDate,p.archived?'Yes':'No']);
    })
  ].join('\n');

  const invRows=[
    row(['Project','Supplier','Invoice #','Date','Category','Subtotal (S$)','GST (S$)','Total (S$)','Status','Proof Uploaded','Payment Proof']),
    ...invoices.filter(i=>filterDate(i.invoiceDate)).map(i=>{
      const proj=projects.find(p=>p.id===i.projectId);
      return row([proj?.name||'',i.supplier,i.invoiceNo,i.invoiceDate,i.category,i.subtotal,i.gst,i.total,i.status,i.proofImage?'Yes':'No',i.paymentProof?'Yes':'No']);
    })
  ].join('\n');

  const payRows=[
    row(['Project','Client','Payment Type','Amount (S$)','Date','Status']),
    ...payments.filter(py=>filterDate(py.date)).map(py=>{
      const proj=projects.find(p=>p.id===py.projectId);
      return row([proj?.name||'',proj?.client||'',py.type,py.amount,py.date,py.status]);
    })
  ].join('\n');

  const catMap={};
  invoices.filter(i=>filterDate(i.invoiceDate)).forEach(i=>{
    if(!catMap[i.category])catMap[i.category]={subtotal:0,gst:0,total:0,count:0};
    catMap[i.category].subtotal+=i.subtotal;catMap[i.category].gst+=i.gst;catMap[i.category].total+=i.total;catMap[i.category].count++;
  });
  const catRows=[
    row(['Category','Invoice Count','Subtotal (S$)','GST (S$)','Total (S$)']),
    ...Object.entries(catMap).sort((a,b)=>b[1].total-a[1].total).map(([cat,d])=>
      row([cat,d.count,d.subtotal.toFixed(2),d.gst.toFixed(2),d.total.toFixed(2)])
    )
  ].join('\n');

  const period=dateFrom||dateTo?`_${dateFrom||'start'}_to_${dateTo||'end'}`:'';
  const dateStr=new Date().toISOString().slice(0,10);
  const full=[`LEDGERSPACE EXPORT — Generated ${dateStr}`,''
    ,'=== PROJECTS SUMMARY ===',projRows,''
    ,'=== INVOICES ===',invRows,''
    ,'=== CLIENT PAYMENTS ===',payRows,''
    ,'=== EXPENSE CATEGORY BREAKDOWN ===',catRows
  ].join('\n');
  const blob=new Blob([full],{type:'text/csv;charset=utf-8;'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;a.download=`LedgerSpace_Export${period}_${dateStr}.csv`;
  document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
}

function Reports({projects,invoices,payments,acctSettings}){
  const [selProj,setSelProj]=useState(projects[0]?.id||'');
  const [generating,setGenerating]=useState(false);
  const [pnl,setPnl]=useState(null);
  const [exportFrom,setExportFrom]=useState('');
  const [exportTo,setExportTo]=useState('');

  const proj=projects.find(p=>p.id===selProj);
  const projInv=invoices.filter(i=>i.projectId===selProj);
  const co=getCo(acctSettings);
  const projPay=payments.filter(p=>p.projectId===selProj);

  const rev=(proj?.contractAmount||0)+(proj?.variationOrders||0);
  const totExp=projInv.reduce((s,i)=>s+i.total,0);
  const totExpSub=projInv.reduce((s,i)=>s+i.subtotal,0);
  const totGST=projInv.reduce((s,i)=>s+i.gst,0);
  const grossProfit=rev-totExp;
  const {dComm:designerComm,pmComm}=proj?calcComm(proj,invoices):{dComm:0,pmComm:0};
  const netProfit=grossProfit-designerComm-pmComm;
  const margin=rev>0?netProfit/rev*100:0;
  const recv=projPay.filter(p=>p.status==='Received').reduce((s,p)=>s+p.amount,0);

  const catBreak=useMemo(()=>{
    const m={};
    projInv.forEach(i=>{m[i.category]=(m[i.category]||0)+i.total;});
    return Object.entries(m).sort((a,b)=>b[1]-a[1]);
  },[projInv]);

  const generateAI=async()=>{
    if(!proj)return;
    const apiKey=(acctSettings?.anthropicApiKey||'').trim();
    if(!apiKey){
      setPnl('AI Analysis not configured. Go to System → set your Anthropic API key to enable this feature.');
      return;
    }
    setGenerating(true);setPnl(null);
    try{
      const res=await fetch('https://api.anthropic.com/v1/messages',{
        method:'POST',
        headers:{
          'Content-Type':'application/json',
          'x-api-key':apiKey,
          'anthropic-version':'2023-06-01',
          'anthropic-dangerous-direct-browser-access':'true',
        },
        body:JSON.stringify({
          model:'claude-opus-4-5',max_tokens:1000,
          messages:[{role:'user',content:`You are a financial analyst for an interior design firm in Singapore.
Analyze this project and provide a concise 3-paragraph P&L narrative (max 200 words total):
Project: ${proj.name}
Client: ${proj.client}
Status: ${proj.status}
Contract Value: S$${rev.toLocaleString()}
Total Expenses: S$${totExp.toLocaleString()}
Gross Profit: S$${grossProfit.toLocaleString()}
Net Profit (after commissions): S$${netProfit.toLocaleString()}
Margin: ${margin.toFixed(1)}%
Payments Collected: S$${recv.toLocaleString()}
Category breakdown: ${catBreak.map(([c,a])=>`${c}: S$${a.toLocaleString()}`).join(', ')}

Paragraph 1: Overall financial health and key metrics.
Paragraph 2: Expense analysis and category observations.
Paragraph 3: Recommendations and risk flags.
Be specific, use the numbers, and keep it professional.`}]
        })
      });
      const d=await res.json();
      if(d.error) throw new Error(d.error.message);
      setPnl(d.content?.[0]?.text||'No analysis returned.');
    }catch(e){
      setPnl(`Unable to generate analysis: ${e.message||'Please try again.'}`);
    }
    setGenerating(false);
  };

  if(!proj)return <div style={{color:T.dim,fontSize:13,textAlign:'center',padding:40}}>No projects available</div>;

  return (
    <div style={{display:'flex',flexDirection:'column',gap:16}}>

      {/* Export to Excel panel */}
      <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:20,boxShadow:T.shadow,
        display:'flex',alignItems:'center',gap:14,flexWrap:'wrap'}}>
        <div style={{fontSize:13,fontWeight:700,color:T.text,marginRight:4}}>Export to Excel / CSV</div>
        <div style={{display:'flex',alignItems:'center',gap:8,flex:1,flexWrap:'wrap'}}>
          <div style={{display:'flex',alignItems:'center',gap:6}}>
            <label style={{fontSize:11,color:T.dim,fontWeight:600}}>From</label>
            <input type="date" value={exportFrom} onChange={e=>setExportFrom(e.target.value)}
              style={{...iStyle,width:'auto',minWidth:130}}/>
          </div>
          <div style={{display:'flex',alignItems:'center',gap:6}}>
            <label style={{fontSize:11,color:T.dim,fontWeight:600}}>To</label>
            <input type="date" value={exportTo} onChange={e=>setExportTo(e.target.value)}
              style={{...iStyle,width:'auto',minWidth:130}}/>
          </div>
          <span style={{fontSize:11,color:T.dim}}>Leave blank to export all data</span>
        </div>
        <Btn onClick={()=>exportToExcel({projects,invoices,payments,dateFrom:exportFrom,dateTo:exportTo})}>
          <Upload size={13}/>Download CSV
        </Btn>
      </div>

      <div style={{display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
        <select value={selProj} onChange={e=>{setSelProj(e.target.value);setPnl(null);}} style={{...iStyle,width:'auto',minWidth:220}}>
          {projects.map(p=><option key={p.id} value={p.id} style={{}}>{p.name}{p.archived?' (Archived)':''}</option>)}
        </select>
        <Btn onClick={generateAI} loading={generating}><BarChart3 size={13}/>Generate AI Analysis</Btn>
        <Btn variant="secondary" onClick={()=>{
          const html=`
<div style="max-width:760px;margin:0 auto;">
  <div style="background:#0a1628;color:#fff;padding:28px 32px;border-radius:8px 8px 0 0;display:flex;justify-content:space-between;align-items:flex-start;">
    <div><div style="font-size:22px;font-weight:800;">${co.name}</div>
    <div style="font-size:11px;color:#94a3b8;margin-top:4px;">${co.name}</div></div>
    <div style="text-align:right;"><div style="font-size:16px;font-weight:700;color:#e9aa26;">PROJECT P&L REPORT</div>
    <div style="font-size:11px;color:#94a3b8;margin-top:4px;">Generated ${new Date().toLocaleDateString('en-SG',{day:'2-digit',month:'long',year:'numeric'})}</div></div>
  </div>
  <div style="border:1px solid #e2e8f0;border-top:none;border-radius:0 0 8px 8px;padding:28px 32px;">
    <div style="margin-bottom:20px;padding-bottom:16px;border-bottom:2px solid #e2e8f0;">
      <div style="font-size:18px;font-weight:800;color:#1a1a2e;">${proj.name}</div>
      <div style="font-size:13px;color:#64748b;margin-top:4px;">${proj.client} . ${proj.projectType||'Residential'} . ${proj.status}</div>
      <div style="font-size:12px;color:#94a3b8;margin-top:2px;">${proj.startDate?new Date(proj.startDate).toLocaleDateString('en-SG',{day:'2-digit',month:'short',year:'numeric'}):''} – ${proj.endDate?new Date(proj.endDate).toLocaleDateString('en-SG',{day:'2-digit',month:'short',year:'numeric'}):''}</div>
    </div>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:24px;">
      ${[
        {l:'Contract Value',v:`S$${rev.toLocaleString('en-SG')}`,c:'#1a1a2e'},
        {l:'Total Expenses',v:`S$${totExp.toLocaleString('en-SG')}`,c:'#dc2626'},
        {l:'Gross Profit',v:`S$${grossProfit.toLocaleString('en-SG')}`,c:grossProfit>0?'#059669':'#dc2626'},
        {l:'Net Margin',v:`${margin.toFixed(1)}%`,c:margin>15?'#059669':margin>0?'#d97706':'#dc2626'},
      ].map(({l,v,c})=>`<div style="background:#f8fafc;border-radius:8px;padding:12px 14px;">
        <div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:4px;">${l}</div>
        <div style="font-size:15px;font-weight:800;color:${c};">${v}</div>
      </div>`).join('')}
    </div>
    <div style="margin-bottom:24px;">
      <div style="font-size:12px;font-weight:700;color:#1a1a2e;margin-bottom:10px;text-transform:uppercase;letter-spacing:0.06em;">Expense Category Breakdown</div>
      <table style="width:100%;border-collapse:collapse;">
        <thead><tr style="background:#f1f5f9;">
          <th style="padding:8px 10px;text-align:left;font-size:11px;color:#64748b;font-weight:700;text-transform:uppercase;">Category</th>
          <th style="padding:8px 10px;text-align:right;font-size:11px;color:#64748b;font-weight:700;text-transform:uppercase;">Amount</th>
          <th style="padding:8px 10px;text-align:right;font-size:11px;color:#64748b;font-weight:700;text-transform:uppercase;">% of Expenses</th>
        </tr></thead>
        <tbody>${catBreak.map(([cat,amt])=>`<tr style="border-bottom:1px solid #e2e8f0;">
          <td style="padding:8px 10px;font-size:12px;">${cat}</td>
          <td style="padding:8px 10px;text-align:right;font-size:12px;font-weight:700;">S$${amt.toLocaleString('en-SG',{minimumFractionDigits:2})}</td>
          <td style="padding:8px 10px;text-align:right;font-size:12px;color:#64748b;">${totExp>0?(amt/totExp*100).toFixed(1):0}%</td>
        </tr>`).join('')}</tbody>
        <tfoot><tr style="background:#f1f5f9;">
          <td style="padding:8px 10px;font-size:12px;font-weight:700;">Total</td>
          <td style="padding:8px 10px;text-align:right;font-size:13px;font-weight:800;">S$${totExp.toLocaleString('en-SG',{minimumFractionDigits:2})}</td>
          <td></td>
        </tr></tfoot>
      </table>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:24px;">
      <div style="background:#f8fafc;border-radius:8px;padding:16px;">
        <div style="font-size:11px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:10px;">Commission Summary</div>
        ${(()=>{
          const dMethod=proj.designerCommMethod||'profit_pct';
          const pmMethod=proj.pmCommMethod||'profit_pct';
          const dLabel=dMethod==='profit_pct'
            ? `${proj.designer} — ${proj.designerRate||0}% of Gross Profit`
            : dMethod==='project_sum'
              ? `${proj.designer} — ${proj.designerRate||0}% of Project Sum`
              : `${proj.designer} — Fixed Rate`;
          const pmLabel=pmMethod==='profit_pct'
            ? `${proj.pm} — ${proj.pmRate||0}% of Gross Profit`
            : pmMethod==='project_sum'
              ? `${proj.pm} — ${proj.pmRate||0}% of Project Sum`
              : `${proj.pm} — Fixed Rate`;
          return [
            {l:dLabel,v:`S$${designerComm.toLocaleString('en-SG',{minimumFractionDigits:2})}`},
            {l:pmLabel,v:`S$${pmComm.toLocaleString('en-SG',{minimumFractionDigits:2})}`},
            {l:'Net Profit (after commissions)',v:`S$${netProfit.toLocaleString('en-SG',{minimumFractionDigits:2})}`},
          ].map(({l,v})=>`<div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid #e2e8f0;font-size:12px;">
            <span style="color:#475569;flex:1;padding-right:10px;">${l}</span><span style="font-weight:700;flex-shrink:0;">${v}</span>
          </div>`).join('');
        })()}
      </div>
      <div style="background:#f8fafc;border-radius:8px;padding:16px;">
        <div style="font-size:11px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:10px;">GST Summary</div>
        ${[
          {l:'Supplier GST paid (Input)',v:`S$${totGST.toLocaleString('en-SG',{minimumFractionDigits:2})}`},
          {l:'Subtotal ex-GST',v:`S$${totExpSub.toLocaleString('en-SG',{minimumFractionDigits:2})}`},
          {l:'Client charges (Non-GST reg.)',v:'No GST charged'},
        ].map(({l,v})=>`<div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid #e2e8f0;font-size:12px;">
          <span style="color:#475569;">${l}</span><span style="font-weight:700;">${v}</span>
        </div>`).join('')}
      </div>
    </div>
    ${pnl?`<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;padding:16px;margin-bottom:20px;">
      <div style="font-size:11px;font-weight:700;color:#1e40af;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:8px;">AI Financial Analysis</div>
      <div style="font-size:12px;color:#1e3a8a;line-height:1.7;white-space:pre-wrap;">${pnl}</div>
    </div>`:''}
    <div style="text-align:center;color:#94a3b8;font-size:10px;margin-top:16px;">
      Computer-generated report . ${co.name} . No signature required
    </div>
  </div>
</div>`;
          printDoc(html,`P&L Report — ${proj.name}`);
        }}>
          Print Report
        </Btn>
      </div>

      {/* P&L Summary */}
      <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:24,boxShadow:T.shadow}}>
        <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:20}}>
          <div>
            <div style={{fontSize:16,fontWeight:700,color:T.text}}>{proj.name}</div>
            <div style={{fontSize:12,color:T.muted}}>{proj.client} . <Badge color={ST_CLR[proj.status]}>{proj.status}</Badge></div>
          </div>
          <div style={{textAlign:'right'}}>
            <div style={{fontSize:11,color:T.dim}}>Net Profit Margin</div>
            <div style={{fontSize:22,fontWeight:700,color:margin>15?T.success:margin>0?T.accent:T.danger}}>{margin.toFixed(1)}%</div>
          </div>
        </div>

        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(130px,1fr))',gap:12,marginBottom:20}}>
          {[
            {l:'Contract Value',v:fmtSGD(proj.contractAmount),c:T.text},
            {l:'Variation Orders',v:fmtSGD(proj.variationOrders||0),c:T.info},
            {l:'Total Revenue',v:fmtSGD(rev),c:T.text},
            {l:'Total Expenses',v:fmtSGD(totExp),c:T.danger},
            {l:'Gross Profit',v:fmtSGD(grossProfit),c:grossProfit>0?T.success:T.danger},
            {l:'Designer Comm.',v:fmtSGD(designerComm),c:T.muted},
            {l:'PM Commission',v:fmtSGD(pmComm),c:T.muted},
            {l:'Net Profit',v:fmtSGD(netProfit),c:netProfit>0?T.success:T.danger},
          ].map(({l,v,c})=>(
            <div key={l} style={{background:T.bg,borderRadius:12,padding:'12px 14px'}}>
              <div style={{fontSize:10,color:T.dim,textTransform:'uppercase',letterSpacing:'0.06em',marginBottom:4}}>{l}</div>
              <div style={{fontSize:14,fontWeight:700,color:c}}>{v}</div>
            </div>
          ))}
        </div>

        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:16}}>
          <div>
            <div style={{fontSize:12,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.06em',marginBottom:10}}>Expense Breakdown</div>
            {catBreak.map(([cat,amt])=>{
              const pct=totExp>0?amt/totExp*100:0;
              // Match against quoted scope items
              const qItem=(proj?.scopeItems||[]).find(s=>s.category===cat);
              const variance=qItem?amt-qItem.amount:null;
              return (
                <div key={cat} style={{marginBottom:10}}>
                  <div style={{display:'flex',justifyContent:'space-between',fontSize:12,marginBottom:4}}>
                    <span style={{color:T.muted,display:'flex',alignItems:'center',gap:6}}>
                      <span style={{width:7,height:7,borderRadius:'50%',background:CAT_CLR[cat]||T.dim,display:'inline-block'}}/>
                      {cat}
                    </span>
                    <div style={{display:'flex',alignItems:'center',gap:8}}>
                      {variance!==null&&<span style={{fontSize:10,color:variance>0?T.danger:T.success,fontWeight:600}}>
                        {variance>0?'▲':'▼'} {fmtSGD(Math.abs(variance))} vs quoted
                      </span>}
                      <span style={{color:T.text,fontWeight:700}}>{fmtSGD(amt)}</span>
                    </div>
                  </div>
                  <div style={{height:4,background:T.bg,borderRadius:2,overflow:'hidden'}}>
                    <div style={{height:'100%',width:`${pct}%`,background:CAT_CLR[cat]||T.dim,borderRadius:2}}/>
                  </div>
                </div>
              );
            })}
            {catBreak.length===0&&<div style={{color:T.dim,fontSize:12}}>No expenses recorded</div>}
          </div>

          <div>
            <div style={{fontSize:12,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.06em',marginBottom:10}}>GST Summary</div>
            <div style={{display:'flex',flexDirection:'column',gap:8}}>
              {[
                {l:'Input Tax (Supplier GST)',v:fmtSGD(totGST),c:T.danger},
                {l:'Subtotal (ex-GST)',v:fmtSGD(totExpSub),c:T.muted},
              ].map(({l,v,c})=>(
                <div key={l} style={{display:'flex',justifyContent:'space-between',padding:'9px 12px',
                  background:T.bg,borderRadius:8}}>
                  <span style={{fontSize:12,color:T.muted}}>{l}</span>
                  <span style={{fontSize:13,fontWeight:700,color:c}}>{v}</span>
                </div>
              ))}
            </div>

            <div style={{fontSize:12,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.06em',marginBottom:10,marginTop:16}}>Commission Summary</div>
            <div style={{display:'flex',flexDirection:'column',gap:8}}>
              {(()=>{
                const dMethod=proj.designerCommMethod||'profit_pct';
                const pmMethod=proj.pmCommMethod||'profit_pct';
                const dLabel=dMethod==='profit_pct'
                  ? `${proj.designer} — ${proj.designerRate||0}% of Gross Profit`
                  : dMethod==='project_sum'
                    ? `${proj.designer} — ${proj.designerRate||0}% of Project Sum (${fmtSGD(proj.contractAmount+(proj.variationOrders||0))})`
                    : `${proj.designer} — Fixed Rate`;
                const pmLabel=pmMethod==='profit_pct'
                  ? `${proj.pm} — ${proj.pmRate||0}% of Gross Profit`
                  : pmMethod==='project_sum'
                    ? `${proj.pm} — ${proj.pmRate||0}% of Project Sum (${fmtSGD(proj.contractAmount+(proj.variationOrders||0))})`
                    : `${proj.pm} — Fixed Rate`;
                return [
                  {l:dLabel,v:fmtSGD(designerComm),c:T.accent},
                  {l:pmLabel,v:fmtSGD(pmComm),c:T.info},
                ].map(({l,v,c})=>(
                  <div key={l} style={{display:'flex',justifyContent:'space-between',alignItems:'center',padding:'9px 12px',background:T.bg,borderRadius:8}}>
                    <span style={{fontSize:12,color:T.muted,flex:1,paddingRight:12}}>{l}</span>
                    <span style={{fontSize:13,fontWeight:700,color:c,flexShrink:0}}>{v}</span>
                  </div>
                ));
              })()}
            </div>
          </div>
        </div>

      {/* Quoted Scope from Quotation */}
      {(proj?.scopeItems||[]).length>0&&(
        <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:14,padding:'16px 18px',boxShadow:T.shadow}}>
          <div style={{fontSize:12,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.06em',marginBottom:10,display:'flex',alignItems:'center',gap:8}}>
            <FileSpreadsheet size={13} style={{color:T.accent}}/>
            Quoted Scope{proj.refNo&&` — Ref: ${proj.refNo}`}
          </div>
          <div style={{display:'flex',flexDirection:'column',gap:0}}>
            {(proj.scopeItems||[]).map((s,i)=>{
              const actual=projInv.filter(inv=>inv.category===s.category).reduce((sum,inv)=>sum+inv.total,0);
              const variance=actual-s.amount;
              const hasActual=actual>0;
              return (
                <div key={i} style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'9px 0',borderBottom:i<proj.scopeItems.length-1?`1px solid ${T.borderLight}`:'none',gap:10}}>
                  <div style={{display:'flex',alignItems:'center',gap:8,flex:1,minWidth:0}}>
                    <Badge color={CAT_CLR[s.category]||T.dim} sm>{s.category}</Badge>
                    <span style={{fontSize:12,color:T.muted,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{s.description}</span>
                  </div>
                  <div style={{display:'flex',alignItems:'center',gap:14,flexShrink:0}}>
                    <div style={{textAlign:'right'}}>
                      <div style={{fontSize:10,color:T.dim}}>Quoted</div>
                      <div style={{fontSize:13,fontWeight:600,color:T.text}}>{fmtSGD(s.amount)}</div>
                    </div>
                    {hasActual&&(<div style={{textAlign:'right'}}>
                      <div style={{fontSize:10,color:T.dim}}>Actual</div>
                      <div style={{fontSize:13,fontWeight:600,color:T.text}}>{fmtSGD(actual)}</div>
                    </div>)}
                    {hasActual&&(<div style={{textAlign:'right',minWidth:70}}>
                      <div style={{fontSize:10,color:T.dim}}>Variance</div>
                      <div style={{fontSize:12,fontWeight:700,color:variance>0?T.danger:T.success}}>
                        {variance>0?'+':''}{fmtSGD(variance)}
                      </div>
                    </div>)}
                  </div>
                </div>
              );
            })}
            <div style={{display:'flex',justifyContent:'space-between',paddingTop:10,borderTop:`2px solid ${T.borderLight}`,marginTop:2}}>
              <span style={{fontSize:13,fontWeight:700,color:T.text}}>Total Quoted</span>
              <span style={{fontSize:14,fontWeight:800,color:T.accent}}>{fmtSGD((proj.scopeItems||[]).reduce((s,i)=>s+i.amount,0))}</span>
            </div>
          </div>
        </div>
      )}
      </div>

      {(pnl||generating)&&(
        <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:24,boxShadow:T.shadow}}>
          <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:14}}>
            <div style={{width:7,height:7,borderRadius:'50%',background:T.accent}}/>
            <span style={{fontSize:13,fontWeight:700,color:T.text}}>AI Financial Analysis</span>
          </div>
          {generating?(
            <div style={{display:'flex',alignItems:'center',gap:10,color:T.muted,fontSize:13}}>
              <Loader2 size={14} style={{animation:'spin 1s linear infinite',color:T.accent}}/>
              Generating analysis...
            </div>
          ):(
            <div style={{fontSize:13,color:T.muted,lineHeight:1.75,whiteSpace:'pre-wrap'}}>{pnl}</div>
          )}
        </div>
      )}

      <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:24,boxShadow:T.shadow}}>
        <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:14}}>Invoice Register — {proj.name}</div>
        <div style={{overflowX:'auto'}}>
          <table style={{width:'100%',fontSize:12,borderCollapse:'collapse'}}>
            <thead>
              <tr style={{color:T.dim,fontSize:11,fontWeight:700,textTransform:'uppercase',letterSpacing:'0.05em'}}>
                {['Supplier','Invoice #','Date','Category','Subtotal','GST','Total','Status'].map(h=>(
                  <th key={h} style={{textAlign:['Subtotal','GST','Total'].includes(h)?'right':'left',
                    paddingBottom:9,paddingRight:12,whiteSpace:'nowrap'}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {projInv.map(inv=>(
                <tr key={inv.id} style={{borderTop:`1px solid ${T.borderLight}`}}>
                  <td style={{padding:'9px 12px 9px 0',color:T.text}}>{inv.supplier}</td>
                  <td style={{padding:'9px 12px 9px 0',color:T.dim,fontFamily:'monospace',fontSize:11}}>{inv.invoiceNo}</td>
                  <td style={{padding:'9px 12px 9px 0',color:T.dim,whiteSpace:'nowrap'}}>{fmtDate(inv.invoiceDate)}</td>
                  <td style={{padding:'9px 12px 9px 0'}}><Badge color={CAT_CLR[inv.category]||T.dim} sm>{inv.category}</Badge></td>
                  <td style={{padding:'9px 12px 9px 0',textAlign:'right',color:T.muted}}>{fmtSGD(inv.subtotal)}</td>
                  <td style={{padding:'9px 12px 9px 0',textAlign:'right',color:T.dim}}>{fmtSGD(inv.gst)}</td>
                  <td style={{padding:'9px 12px 9px 0',textAlign:'right',color:T.text,fontWeight:700}}>{fmtSGD(inv.total)}</td>
                  <td style={{padding:'9px 0'}}><Badge color={ST_CLR[inv.status]} sm>{inv.status}</Badge></td>
                </tr>
              ))}
              <tr style={{borderTop:`2px solid ${T.border}`}}>
                <td colSpan={4} style={{padding:'10px 0',color:T.dim,fontWeight:700,fontSize:11,textTransform:'uppercase'}}>Totals</td>
                <td style={{padding:'10px 12px 10px 0',textAlign:'right',color:T.muted,fontWeight:700}}>{fmtSGD(totExpSub)}</td>
                <td style={{padding:'10px 12px 10px 0',textAlign:'right',color:T.dim,fontWeight:700}}>{fmtSGD(totGST)}</td>
                <td style={{padding:'10px 12px 10px 0',textAlign:'right',color:T.text,fontWeight:700}}>{fmtSGD(totExp)}</td>
                <td/>
              </tr>
            </tbody>
          </table>
          {projInv.length===0&&<div style={{color:T.dim,fontSize:13,textAlign:'center',padding:24}}>No invoices for this project</div>}
        </div>
      </div>
    </div>
  );
}

// ── Staff Expense Claims ─────────────────────────────────────────────────────
const STAFF_CLAIM_TYPES = [
  {id:'materials',  label:'Materials / Supplies',  icon:'🧱', category:'Miscellaneous'},
  {id:'delivery',   label:'Delivery / Logistics',  icon:'🚚', category:'Preliminaries'},
  {id:'food',       label:'Food & Drinks',          icon:'🍱', category:'Preliminaries'},
  {id:'transport',  label:'Transport (Grab/Taxi)',  icon:'🚕', category:'Preliminaries'},
  {id:'parking',    label:'Parking / ERP',          icon:'🅿️', category:'Preliminaries'},
  {id:'hardware',   label:'Hardware / Tools',       icon:'🔧', category:'Miscellaneous'},
  {id:'other',      label:'Other',                  icon:'📎', category:'Miscellaneous'},
];

function StaffClaims({claims,setClaims,projects,users,activeUser,isAdmin,invoices,setInvoices,acctSettings,trash,setTrash}){
  const [tab,setTab]=useState('my');
  const [form,setForm]=useState(null);
  const [receipt,setReceipt]=useState(null);
  const [rejectTarget,setRejectTarget]=useState(null);
  const [rejectReason,setRejectReason]=useState('');
  const [adminPayTarget,setAdminPayTarget]=useState(null);
  const [adminPayForm,setAdminPayForm]=useState({method:'',reference:'',receiptImage:null,date:new Date().toISOString().slice(0,10)});
  const [adminOcrLoading,setAdminOcrLoading]=useState(false);
  const [claimOcrLoading,setClaimOcrLoading]=useState(false);
  const receiptRef=useRef();
  const companyReceiptRef=useRef();
  const adminPayReceiptRef=useRef();

  const myProjects=projects.filter(p=>!p.archived&&(
    p.designer===activeUser?.name||p.pm===activeUser?.name||isAdmin
  ));

  const myClaims=(claims||[]).filter(c=>c.submittedBy===activeUser?.name)
    .sort((a,b)=>b.submittedAt.localeCompare(a.submittedAt));

  const pendingClaims=(claims||[]).filter(c=>c.status==='Pending')
    .sort((a,b)=>b.submittedAt.localeCompare(a.submittedAt));

  const ff=k=>v=>setForm(p=>({...p,[k]:v}));
  const apf=k=>v=>setAdminPayForm(p=>({...p,[k]:v}));

  // OCR receipt to auto-fill amount, date, description (for staff receipt)
  const extractReceipt=async(file, apiKey)=>{
    const reader=new FileReader();
    reader.onload=ev=>ff('receiptImage')(ev.target.result);
    reader.readAsDataURL(file);
    if(!apiKey) return;
    try{
      const b64=await toB64(file);
      const isPDF=file.type==='application/pdf';
      const part=isPDF
        ?{type:'document',source:{type:'base64',media_type:'application/pdf',data:b64}}
        :{type:'image',source:{type:'base64',media_type:file.type,data:b64}};
      const res=await fetch('https://api.anthropic.com/v1/messages',{
        method:'POST',
        headers:{'Content-Type':'application/json','x-api-key':apiKey,
          'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},
        body:JSON.stringify({
          model:'claude-opus-4-5',max_tokens:300,
          messages:[{role:'user',content:[part,{type:'text',
            text:'Extract from this expense receipt. Return ONLY valid JSON:\n{"amount":0,"date":"YYYY-MM-DD","description":""}\nRules: amount=total paid (numeric). date=YYYY-MM-DD. description=brief item (max 60 chars). Use 0 or empty if not found.'}]}]
        })
      });
      const data=await res.json();
      if(data.error) return;
      const txt=(data.content||[]).map(c=>c.text||'').join('').replace(/```json|```/g,'').trim();
      const parsed=JSON.parse(txt);
      setForm(prev=>({
        ...prev,
        amount:parsed.amount?String(parsed.amount):prev.amount,
        date:parsed.date||prev.date,
        description:parsed.description||prev.description,
      }));
    }catch{}
  };

  // OCR for admin payment receipt — extracts method, reference, date
  const extractAdminPayReceipt=async(file)=>{
    const apiKey=(acctSettings?.anthropicApiKey||'').trim();
    const reader=new FileReader();
    reader.onload=ev=>setAdminPayForm(p=>({...p,receiptImage:ev.target.result}));
    reader.readAsDataURL(file);
    compressForSync(file).then(c=>{if(c)setAdminPayForm(p=>({...p,receiptImage:c}));});
    if(!apiKey) return;
    setAdminOcrLoading(true);
    try{
      const b64=await toB64(file);
      const isPDF=file.type==='application/pdf';
      const part=isPDF
        ?{type:'document',source:{type:'base64',media_type:'application/pdf',data:b64}}
        :{type:'image',source:{type:'base64',media_type:file.type,data:b64}};
      const res=await fetch('https://api.anthropic.com/v1/messages',{
        method:'POST',
        headers:{'Content-Type':'application/json','x-api-key':apiKey,
          'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},
        body:JSON.stringify({
          model:'claude-opus-4-5',max_tokens:300,
          messages:[{role:'user',content:[part,{type:'text',
            text:'Extract from this company payment receipt/screenshot. Return ONLY valid JSON:\n{"date":"YYYY-MM-DD","method":"paynow|bank_transfer|cash","reference":""}\nRules: date=payment date. method=paynow if PayNow/QR/UEN, bank_transfer if bank/FAST/GIRO, cash if cash. reference=PayNow number/UEN/account/transaction ref. Use empty string if not found.'}]}]
        })
      });
      const data=await res.json();
      if(data.error) return;
      const txt=(data.content||[]).map(c=>c.text||'').join('').replace(/```json|```/g,'').trim();
      const parsed=JSON.parse(txt);
      setAdminPayForm(prev=>({
        ...prev,
        date:parsed.date||prev.date,
        method:parsed.method||prev.method,
        reference:parsed.reference||prev.reference,
      }));
    }catch(e){console.warn('Admin pay OCR failed:',e);}
    finally{setAdminOcrLoading(false);}
  };

  // Delete claim AND its linked auto-generated invoice (if any)
  // Receipt upload handler for claim form — OCR auto-fills details
  const handleClaimReceiptUpload=async(file)=>{
    if(!file) return;
    // Store image preview safely
    try{
      const reader=new FileReader();
      reader.onload=ev=>setForm(prev=>prev?{...prev,receiptImage:ev.target.result}:prev);
      reader.readAsDataURL(file);
      compressForSync(file).then(c=>{if(c)setForm(prev=>prev?{...prev,receiptImage:c}:prev);});
    }catch(e){ console.warn('FileReader error:',e); }

    // OCR — only if API key configured, never blocks submit
    const apiKey=(acctSettings?.anthropicApiKey||'').trim();
    if(!apiKey) return;
    setClaimOcrLoading(true);
    try{
      const b64=await toB64(file);
      const isPDF=file.type==='application/pdf';
      const part=isPDF
        ?{type:'document',source:{type:'base64',media_type:'application/pdf',data:b64}}
        :{type:'image',source:{type:'base64',media_type:file.type,data:b64}};
      const res=await fetch('https://api.anthropic.com/v1/messages',{
        method:'POST',
        headers:{'Content-Type':'application/json','x-api-key':apiKey,
          'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},
        body:JSON.stringify({
          model:'claude-opus-4-5',max_tokens:400,
          messages:[{role:'user',content:[part,{type:'text',
            text:'Extract from this expense receipt. Return ONLY valid JSON:\n{"amount":0,"date":"YYYY-MM-DD","description":"","merchant":""}\nRules: amount=total paid (numeric). date=YYYY-MM-DD. description=what was purchased (max 80 chars). merchant=store/vendor name. Use 0 or empty string if not found.'}]}]
        })
      });
      const data=await res.json();
      if(!data.error){
        const txt=(data.content||[]).map(c=>c.text||'').join('').replace(/```json|```/g,'').trim();
        const parsed=JSON.parse(txt);
        setForm(prev=>prev?({
          ...prev,
          amount:parsed.amount?String(parsed.amount):prev.amount,
          date:parsed.date||prev.date,
          description:parsed.description?(parsed.merchant?`${parsed.merchant} — ${parsed.description}`:parsed.description):prev.description,
        }):prev);
      }
    }catch(e){ console.warn('Claim OCR:',e); }
    finally{ setClaimOcrLoading(false); }
  };

  const deleteClaim=(id)=>{
    const claim=(claims||[]).find(c=>c.id===id);
    if(!claim) return;

    // Send linked invoice to trash too (if any)
    const linkedInvs=(invoices||[]).filter(i=>i._fromClaim===id);
    if(linkedInvs.length>0){
      const trashItems=linkedInvs.map(inv=>({...inv,_trashType:'invoice',_deletedAt:new Date().toISOString()}));
      // Remove from invoices
      const updInv=(invoices||[]).filter(i=>i._fromClaim!==id);
      setInvoices(updInv);saveInvoices(updInv);
      // Add to trash
      setTrash(prev=>{const upd=[...prev,...trashItems];saveS('trash',upd);return upd;});
    }

    // Send claim to trash
    const trashClaim={...claim,_trashType:'staffClaim',_deletedAt:new Date().toISOString()};
    setTrash(prev=>{const upd=[...prev,trashClaim];saveS('trash',upd);return upd;});
    const upd=(claims||[]).filter(c=>c.id!==id);
    setClaims(upd);saveS('staffClaims',upd);
  };

  // Save admin payment record on claim
  const saveAdminPayment=()=>{
    const claim=(claims||[]).find(c=>c.id===adminPayTarget);
    if(!claim) return;

    // Personal claim: create project expense entry when payment proof uploaded
    const isPersonal=claim.paidBy==='personal'||!claim.paidBy;
    if(isPersonal&&!claim.reimbursed){
      const ct=STAFF_CLAIM_TYPES.find(t=>t.id===claim.type);
      const newInv={
        id:uid(),projectId:claim.projectId,
        supplier:`Staff Claim — ${claim.submittedBy}`,
        invoiceNo:`CLM-${new Date().getFullYear()}-${String(Date.now()).slice(-5)}`,
        invoiceDate:adminPayForm.date||claim.date,
        subtotal:claim.amount,gst:0,total:claim.amount,
        category:ct?.category||'Miscellaneous',status:'Paid',
        paymentRecords:[{id:uid(),amount:claim.amount,date:adminPayForm.date||claim.date,
          method:adminPayForm.method||'cash',reference:adminPayForm.reference||'',
          notes:`${ct?.label||claim.type}${claim.description?' — '+claim.description:''}`,
          proofImage:adminPayForm.receiptImage||null,recordedAt:new Date().toISOString()}],
        proofImage:adminPayForm.receiptImage||null,
        createdAt:new Date().toISOString(),_fromClaim:claim.id,
      };
      const updInv=[...(invoices||[]),newInv];
      setInvoices(updInv);saveInvoices(updInv);
    }

    const upd=(claims||[]).map(c=>c.id===adminPayTarget?{...c,
      adminPayment:{method:adminPayForm.method,reference:adminPayForm.reference,
        receiptImage:adminPayForm.receiptImage,date:adminPayForm.date,
        recordedBy:activeUser?.name,recordedAt:new Date().toISOString()},
      reimbursed:true,reimbursedAt:new Date().toISOString(),
      reflectedInExpenses:isPersonal?true:c.reflectedInExpenses,
    }:c);
    setClaims(upd);saveS('staffClaims',upd);
    setAdminPayTarget(null);
    setAdminPayForm({method:'',reference:'',receiptImage:null,date:new Date().toISOString().slice(0,10)});
  };

  const submitClaim=()=>{
    if(!form.projectId||!form.type||!form.amount||parseFloat(form.amount)<=0)return;
    const claim={
      id:uid(),
      projectId:form.projectId,
      type:form.type,
      amount:parseFloat(form.amount),
      description:form.description||'',
      receiptImage:form.receiptImage||null,
      paidBy:form.paidBy||'personal',      // 'personal' = staff paid, 'company' = company paid
      companyReceiptImage:form.companyReceiptImage||null,
      submittedBy:activeUser?.name,
      submittedByRole:activeUser?.role,
      submittedAt:new Date().toISOString(),
      date:form.date||new Date().toISOString().slice(0,10),
      status:'Pending',
      reviewedBy:null,reviewedAt:null,notes:'',
      reimbursed:false,reimbursedAt:null,   // for personal claims
      reflectedInExpenses:false,            // for company claims
    };
    const upd=[...(claims||[]),claim];
    setClaims(upd);saveS('staffClaims',upd);
    setForm(null);
  };

  // Approve: create supplier invoice entry — project or company accounts
  const approve=(id)=>{
    const claim=(claims||[]).find(c=>c.id===id);
    if(!claim) return;

    let updInvoices=invoices||[];
    let reflectedInExpenses=false;

    if(claim.paidBy==='company'){
      const ct=STAFF_CLAIM_TYPES.find(t=>t.id===claim.type);
      const isCompanyExp=claim.projectId==='__company__'||!claim.projectId;
      const newInv={
        id:uid(),
        projectId:isCompanyExp?null:claim.projectId,  // null = company account expense
        supplier:`Staff Claim — ${claim.submittedBy}`,
        invoiceNo:`CLM-${new Date().getFullYear()}-${String(Date.now()).slice(-5)}`,
        invoiceDate:claim.date,
        subtotal:claim.amount,
        gst:0,
        total:claim.amount,
        category:ct?.category||'Miscellaneous',
        status:'Paid',
        paymentMethod:'cash',
        _isCompanyExpense:isCompanyExp,
        paymentRecords:[{id:uid(),amount:claim.amount,date:claim.date,method:'cash',
          reference:`Claim ID ${claim.id}`,
          notes:`${ct?.label||claim.type}${claim.description?' — '+claim.description:''}`,
          proofImage:claim.companyReceiptImage||claim.receiptImage||null,
          recordedAt:new Date().toISOString()}],
        proofImage:claim.companyReceiptImage||claim.receiptImage||null,
        createdAt:new Date().toISOString(),
        _fromClaim:claim.id,
      };
      updInvoices=[...(invoices||[]),newInv];
      setInvoices(updInvoices);saveInvoices(updInvoices);
      reflectedInExpenses=true;
    }

    const upd=(claims||[]).map(c=>c.id===id?{...c,
      status:'Approved',
      reviewedBy:activeUser?.name,
      reviewedAt:new Date().toISOString(),
      reflectedInExpenses,
    }:c);
    setClaims(upd);saveS('staffClaims',upd);
  };

  const reject=(id)=>{
    const upd=(claims||[]).map(c=>c.id===id?{...c,
      status:'Rejected',
      reviewedBy:activeUser?.name,
      reviewedAt:new Date().toISOString(),
      notes:rejectReason,
    }:c);
    setClaims(upd);saveS('staffClaims',upd);
    setRejectTarget(null);setRejectReason('');
  };

  const markReimbursed=(id)=>{
    const upd=(claims||[]).map(c=>c.id===id?{...c,reimbursed:true,reimbursedAt:new Date().toISOString()}:c);
    setClaims(upd);saveS('staffClaims',upd);
  };

  const statusClr={Pending:T.warning,Approved:T.success,Rejected:T.danger};
  const statusBg={Pending:T.warningLight,Approved:T.successLight,Rejected:T.dangerLight};

  const totalMy=myClaims.reduce((s,c)=>s+c.amount,0);
  const totalPending=pendingClaims.reduce((s,c)=>s+c.amount,0);
  const approvedPersonal=myClaims.filter(c=>c.status==='Approved'&&c.paidBy==='personal'&&!c.reimbursed).reduce((s,c)=>s+c.amount,0);

  return (
    <div style={{display:'flex',flexDirection:'column',gap:16}}>

      {/* Tabs */}
      <div style={{display:'flex',gap:8}}>
        {[
          {id:'my',label:'My Claims'},
          ...(isAdmin?[{id:'approve',label:`Approve${pendingClaims.length>0?` (${pendingClaims.length})`:''}`,danger:pendingClaims.length>0}]:[]),
        ].map(t=>(
          <button key={t.id} type="button" onClick={()=>setTab(t.id)}
            style={{padding:'8px 18px',borderRadius:10,border:'none',cursor:'pointer',fontFamily:'inherit',
              fontSize:13,fontWeight:tab===t.id?700:500,
              background:tab===t.id?(t.danger?T.dangerLight:T.accentLight):'transparent',
              color:tab===t.id?(t.danger?T.danger:T.accent):T.muted}}>
            {t.label}
          </button>
        ))}
        <div style={{flex:1}}/>
        {tab==='my'&&(
          <button type="button" onClick={()=>setForm({projectId:myProjects[0]?.id||'',type:'',amount:'',description:'',date:new Date().toISOString().slice(0,10),receiptImage:null,paidBy:'personal',companyReceiptImage:null})}
            style={{display:'flex',alignItems:'center',gap:6,padding:'8px 16px',borderRadius:10,border:'none',cursor:'pointer',fontFamily:'inherit',fontSize:13,fontWeight:600,background:T.accent,color:'#fff'}}>
            <Plus size={14}/>Submit Claim
          </button>
        )}
      </div>

      {/* Summary */}
      {tab==='my'&&(
        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(150px,1fr))',gap:10}}>
          {[
            {label:'Total Claimed',val:fmtSGD(totalMy),color:T.text},
            {label:'Approved',val:fmtSGD(myClaims.filter(c=>c.status==='Approved').reduce((s,c)=>s+c.amount,0)),color:T.success},
            {label:'Pending Reimbursement',val:fmtSGD(approvedPersonal),color:approvedPersonal>0?T.warning:T.success,sub:approvedPersonal>0?'Company owes you this':'All settled'},
          ].map(({label,val,color,sub})=>(
            <div key={label} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:14,padding:'14px 16px',boxShadow:T.shadow}}>
              <div style={{fontSize:11,color:T.dim,marginBottom:4}}>{label}</div>
              <div style={{fontSize:18,fontWeight:700,color}}>{val}</div>
              {sub&&<div style={{fontSize:10,color,marginTop:3}}>{sub}</div>}
            </div>
          ))}
        </div>
      )}
      {tab==='approve'&&isAdmin&&(
        <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:14,padding:'14px 16px',boxShadow:T.shadow,display:'flex',alignItems:'center',gap:12}}>
          <div style={{flex:1}}>
            <div style={{fontSize:11,color:T.dim,marginBottom:2}}>Total Pending Value</div>
            <div style={{fontSize:20,fontWeight:700,color:T.warning}}>{fmtSGD(totalPending)}</div>
          </div>
          <div style={{fontSize:13,color:T.muted}}>{pendingClaims.length} claim{pendingClaims.length!==1?'s':''} awaiting review</div>
        </div>
      )}

      {/* MY CLAIMS */}
      {tab==='my'&&(
        <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,overflow:'hidden',boxShadow:T.shadow}}>
          {myClaims.length===0&&(
            <div style={{padding:40,textAlign:'center',color:T.dim,fontSize:13}}>No claims submitted yet. Tap "Submit Claim" to add one.</div>
          )}
          {myClaims.map(c=>{
            const proj=projects.find(p=>p.id===c.projectId);
            const ct=STAFF_CLAIM_TYPES.find(t=>t.id===c.type);
            const isPersonal=c.paidBy==='personal'||!c.paidBy;
            return (
              <div key={c.id} style={{padding:'14px 18px',borderBottom:`1px solid ${T.borderLight}`,display:'flex',alignItems:'flex-start',gap:14}}>
                <div style={{fontSize:24,flexShrink:0,marginTop:2}}>{ct?.icon||'📎'}</div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:4,flexWrap:'wrap'}}>
                    <span style={{fontSize:13,fontWeight:700,color:T.text}}>{ct?.label||c.type}</span>
                    <span style={{fontSize:10,fontWeight:700,padding:'2px 7px',borderRadius:6,
                      background:isPersonal?'rgba(124,58,237,0.1)':T.accentLight,
                      color:isPersonal?'#7c3aed':T.accent}}>
                      {isPersonal?'💳 Paid by me':'🏢 Company paid'}
                    </span>
                    <span style={{fontSize:11,color:T.muted}}>·</span>
                    <span style={{fontSize:12,color:T.muted,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{proj?.name||'Unknown project'}</span>
                  </div>
                  {c.description&&<div style={{fontSize:12,color:T.muted,marginBottom:4}}>{c.description}</div>}
                  <div style={{display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
                    <span style={{fontSize:11,color:T.dim}}>{fmtDate(c.date)}</span>
                    {c.receiptImage&&(
                      <button type="button" onClick={()=>setReceipt(c.receiptImage)}
                        style={{background:T.accentLight,border:'none',borderRadius:6,padding:'2px 8px',cursor:'pointer',fontSize:11,color:T.accent,fontFamily:'inherit',fontWeight:600}}>
                        View Receipt
                      </button>
                    )}
                    {/* Reimbursement status for personal claims */}
                    {isPersonal&&c.status==='Approved'&&(
                      c.reimbursed
                        ?<span style={{fontSize:11,color:T.success,fontWeight:600}}>✓ Reimbursed</span>
                        :<span style={{fontSize:11,color:T.warning,fontWeight:600}}>⏳ Awaiting reimbursement</span>
                    )}
                    {/* Expense reflected badge for company claims */}
                    {!isPersonal&&c.status==='Approved'&&c.reflectedInExpenses&&(
                      <span style={{fontSize:11,color:T.accent,fontWeight:600}}>✓ Recorded as project expense</span>
                    )}
                    {c.status==='Rejected'&&c.notes&&(
                      <span style={{fontSize:11,color:T.danger}}>Reason: {c.notes}</span>
                    )}
                  </div>
                </div>
                <div style={{textAlign:'right',flexShrink:0,display:'flex',flexDirection:'column',alignItems:'flex-end',gap:6}}>
                  <div style={{fontSize:16,fontWeight:800,color:T.text}}>{fmtSGD(c.amount)}</div>
                  <div style={{fontSize:11,fontWeight:700,padding:'2px 9px',borderRadius:8,
                    background:statusBg[c.status],color:statusClr[c.status],display:'inline-block'}}>
                    {c.status}
                  </div>
                  {/* Delete own pending claims */}
                  {c.status==='Pending'&&(
                    <button type="button" onClick={()=>deleteClaim(c.id)}
                      style={{background:'none',border:'none',cursor:'pointer',color:T.dim,padding:4,borderRadius:6,display:'flex',alignItems:'center',gap:3,fontSize:11,fontFamily:'inherit'}}>
                      <Trash2 size={11}/>Delete
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* APPROVE CLAIMS */}
      {tab==='approve'&&isAdmin&&(
        <div style={{display:'flex',flexDirection:'column',gap:12}}>

          {/* ── APPROVED — awaiting payment proof ── */}
          {(()=>{
            const needProof=(claims||[]).filter(c=>c.status==='Approved'&&(c.paidBy==='personal'||!c.paidBy)&&!c.reimbursed);
            if(!needProof.length) return null;
            return (
              <div style={{background:T.card,border:`2px solid ${T.accent}30`,borderRadius:16,overflow:'hidden',boxShadow:T.shadow}}>
                <div style={{background:T.accentLight,padding:'10px 16px',display:'flex',alignItems:'center',gap:8}}>
                  <Upload size={14} style={{color:T.accent,flexShrink:0}}/>
                  <span style={{fontSize:13,fontWeight:700,color:T.accent}}>
                    {needProof.length} approved claim{needProof.length!==1?'s':''} — upload payment proof to complete
                  </span>
                </div>
                {needProof.map((c,idx)=>{
                  const proj=projects.find(p=>p.id===c.projectId);
                  const ct=STAFF_CLAIM_TYPES.find(t=>t.id===c.type);
                  return (
                    <div key={c.id} style={{padding:'14px 16px',borderTop:idx>0?`1px solid ${T.borderLight}`:'none',display:'flex',alignItems:'center',gap:12}}>
                      <span style={{fontSize:22,flexShrink:0}}>{ct?.icon||'📎'}</span>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:13,fontWeight:700,color:T.text}}>{c.submittedBy}</div>
                        <div style={{fontSize:12,color:T.muted}}>{ct?.label||c.type} · {proj?.name||'—'} · {fmtDate(c.date)}</div>
                        {c.description&&<div style={{fontSize:11,color:T.dim,marginTop:2}}>{c.description}</div>}
                        {c.receiptImage&&(
                          <button type="button" onClick={()=>setReceipt(c.receiptImage)}
                            style={{marginTop:5,background:T.bg,border:`1px solid ${T.borderLight}`,borderRadius:6,padding:'2px 9px',cursor:'pointer',fontSize:11,color:T.muted,fontFamily:'inherit',display:'inline-flex',alignItems:'center',gap:4}}>
                            <ZoomIn size={10}/>Staff Receipt
                          </button>
                        )}
                      </div>
                      <div style={{display:'flex',flexDirection:'column',alignItems:'flex-end',gap:6,flexShrink:0}}>
                        <div style={{fontSize:15,fontWeight:800,color:T.text}}>{fmtSGD(c.amount)}</div>
                        <button type="button" onClick={()=>{
                          setAdminPayTarget(c.id);
                          setAdminPayForm({method:'',reference:'',receiptImage:null,date:new Date().toISOString().slice(0,10)});
                        }}
                          style={{display:'flex',alignItems:'center',gap:6,padding:'8px 14px',borderRadius:9,border:'none',cursor:'pointer',fontFamily:'inherit',fontSize:12,fontWeight:700,background:T.accent,color:'#fff',whiteSpace:'nowrap'}}>
                          <Upload size={13}/>Upload Proof
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            );
          })()}

          {/* ── PENDING — awaiting approval ── */}
          {pendingClaims.length===0&&(claims||[]).filter(c=>c.status==='Approved'&&(c.paidBy==='personal'||!c.paidBy)&&!c.reimbursed).length===0&&(
            <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:40,textAlign:'center',color:T.dim,fontSize:13}}>
              No pending claims — all caught up.
            </div>
          )}
          {pendingClaims.length>0&&(
            <div style={{fontSize:12,fontWeight:600,color:T.dim,textTransform:'uppercase',letterSpacing:'0.06em',marginTop:4}}>
              Pending Approval
            </div>
          )}
          {pendingClaims.map(c=>{
            const proj=projects.find(p=>p.id===c.projectId);
            const ct=STAFF_CLAIM_TYPES.find(t=>t.id===c.type);
            const submitter=users.find(u=>u.name===c.submittedBy);
            const isPersonal=c.paidBy==='personal'||!c.paidBy;
            return (
              <div key={c.id} style={{background:T.card,border:`2px solid ${isPersonal?'rgba(124,58,237,0.2)':T.accent+'25'}`,borderRadius:16,padding:'16px 18px',boxShadow:T.shadow}}>
                {/* Paid-by banner */}
                <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:12,padding:'8px 12px',borderRadius:10,
                  background:isPersonal?'rgba(124,58,237,0.06)':T.accentLight}}>
                  <span style={{fontSize:18}}>{isPersonal?'💳':'🏢'}</span>
                  <div>
                    <div style={{fontSize:12,fontWeight:700,color:isPersonal?'#7c3aed':T.accent}}>
                      {isPersonal?`Paid by ${c.submittedBy} personally — needs reimbursement`:'Paid by company — will be recorded as project expense'}
                    </div>
                    {!isPersonal&&<div style={{fontSize:11,color:T.muted}}>Approving this will auto-create a supplier invoice entry in the project</div>}
                    {isPersonal&&<div style={{fontSize:11,color:T.muted}}>After approval, mark as reimbursed once payment is made</div>}
                  </div>
                </div>

                <div style={{display:'flex',alignItems:'flex-start',gap:14}}>
                  <div style={{fontSize:26,flexShrink:0}}>{ct?.icon||'📎'}</div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:6,flexWrap:'wrap'}}>
                      <span style={{fontSize:14,fontWeight:700,color:T.text}}>{ct?.label||c.type}</span>
                      <span style={{fontSize:11,color:T.muted}}>by</span>
                      <span style={{fontSize:13,fontWeight:600,color:ROLE_CLR[submitter?.role]||T.accent}}>{c.submittedBy}</span>
                      <Badge color={ROLE_CLR[submitter?.role]||T.dim}>{ROLE_LABEL[submitter?.role]}</Badge>
                    </div>
                    <div style={{fontSize:12,color:T.muted,marginBottom:4}}>
                      <strong style={{color:T.text}}>{proj?.name||'Unknown project'}</strong>
                      {c.description&&` · ${c.description}`}
                    </div>
                    <div style={{fontSize:11,color:T.dim,marginBottom:10}}>Submitted {fmtDate(c.date)}</div>
                    {c.receiptImage&&(
                      <div style={{marginBottom:10}}>
                        <img src={c.receiptImage} alt="Receipt" onClick={()=>setReceipt(c.receiptImage)}
                          style={{maxHeight:120,maxWidth:220,borderRadius:8,border:`1px solid ${T.borderLight}`,cursor:'pointer',objectFit:'contain'}}/>
                        <div style={{fontSize:10,color:T.dim,marginTop:3}}>Tap to enlarge</div>
                      </div>
                    )}
                    <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
                      <button type="button" onClick={()=>approve(c.id)}
                        style={{display:'flex',alignItems:'center',gap:6,padding:'7px 16px',borderRadius:9,border:'none',cursor:'pointer',fontFamily:'inherit',fontSize:13,fontWeight:600,background:T.success,color:'#fff'}}>
                        <CheckCircle size={14}/>{isPersonal?'Approve Claim':'Approve & Record Expense'}
                      </button>
                      <button type="button" onClick={()=>{setRejectTarget(c.id);setRejectReason('');}}
                        style={{display:'flex',alignItems:'center',gap:6,padding:'7px 16px',borderRadius:9,border:`1px solid ${T.danger}40`,cursor:'pointer',fontFamily:'inherit',fontSize:13,fontWeight:600,background:T.dangerLight,color:T.danger}}>
                        Reject
                      </button>
                      {isAdmin&&(
                        <button type="button" onClick={()=>deleteClaim(c.id)}
                          style={{display:'flex',alignItems:'center',gap:5,padding:'7px 12px',borderRadius:9,border:`1px solid ${T.borderLight}`,cursor:'pointer',fontFamily:'inherit',fontSize:12,background:'none',color:T.dim}}>
                          <Trash2 size={12}/>Delete
                        </button>
                      )}
                    </div>
                  </div>
                  <div style={{textAlign:'right',flexShrink:0}}>
                    <div style={{fontSize:20,fontWeight:800,color:T.text}}>{fmtSGD(c.amount)}</div>
                  </div>
                </div>
              </div>
            );
          })}

          {/* Approved claims — awaiting payment upload */}
          {(()=>{
            const approvedClaims=(claims||[]).filter(c=>c.status==='Approved').sort((a,b)=>b.submittedAt.localeCompare(a.submittedAt));
            const awaitingPayment=approvedClaims.filter(c=>(c.paidBy==='personal'||!c.paidBy)&&!c.reimbursed);
            if(approvedClaims.length===0) return null;
            return (
              <div style={{marginTop:8}}>
                {awaitingPayment.length>0&&(
                  <div style={{background:T.warningLight,border:`1px solid ${T.warning}30`,borderRadius:14,padding:'10px 16px',marginBottom:12,display:'flex',alignItems:'center',gap:8}}>
                    <Bell size={13} style={{color:T.warning,flexShrink:0}}/>
                    <span style={{fontSize:13,color:T.warning,fontWeight:600}}>
                      {awaitingPayment.length} claim{awaitingPayment.length!==1?'s':''} approved — payment proof not yet uploaded
                    </span>
                  </div>
                )}
                <div style={{fontSize:12,fontWeight:600,color:T.dim,textTransform:'uppercase',letterSpacing:'0.06em',marginBottom:8}}>Approved Claims</div>
                <div style={{display:'flex',flexDirection:'column',gap:10}}>
                  {approvedClaims.map(c=>{
                    const proj=projects.find(p=>p.id===c.projectId);
                    const ct=STAFF_CLAIM_TYPES.find(t=>t.id===c.type);
                    const isPersonal=c.paidBy==='personal'||!c.paidBy;
                    const needsPayment=isPersonal&&!c.reimbursed;
                    return (
                      <div key={c.id} style={{background:T.card,border:`2px solid ${needsPayment?T.warning+'40':T.success+'30'}`,borderRadius:14,overflow:'hidden',boxShadow:T.shadow}}>
                        {/* Claim summary row */}
                        <div style={{padding:'14px 16px',display:'flex',alignItems:'flex-start',gap:12}}>
                          <span style={{fontSize:22,flexShrink:0}}>{ct?.icon||'📎'}</span>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:4,flexWrap:'wrap'}}>
                              <span style={{fontSize:13,fontWeight:700,color:T.text}}>{ct?.label||c.type}</span>
                              <span style={{fontSize:11,color:T.muted}}>·</span>
                              <span style={{fontSize:12,color:T.muted}}>{c.submittedBy}</span>
                              <span style={{fontSize:10,fontWeight:700,padding:'1px 6px',borderRadius:5,
                                background:isPersonal?'rgba(124,58,237,0.1)':T.accentLight,
                                color:isPersonal?'#7c3aed':T.accent}}>
                                {isPersonal?'💳 Personal':'🏢 Company'}
                              </span>
                            </div>
                            <div style={{fontSize:11,color:T.dim}}>{proj?.name||'—'} · {fmtDate(c.date)}</div>
                            {c.description&&<div style={{fontSize:11,color:T.muted,marginTop:2}}>{c.description}</div>}
                            {c.receiptImage&&(
                              <button type="button" onClick={()=>setReceipt(c.receiptImage)}
                                style={{marginTop:6,background:T.accentLight,border:'none',borderRadius:6,padding:'2px 9px',cursor:'pointer',fontSize:11,color:T.accent,fontFamily:'inherit',fontWeight:600,display:'inline-flex',alignItems:'center',gap:4}}>
                                <ZoomIn size={10}/>View Staff Receipt
                              </button>
                            )}
                          </div>
                          <div style={{textAlign:'right',flexShrink:0,display:'flex',flexDirection:'column',alignItems:'flex-end',gap:6}}>
                            <div style={{fontSize:17,fontWeight:800,color:T.text}}>{fmtSGD(c.amount)}</div>
                            <button type="button" onClick={()=>deleteClaim(c.id)}
                              style={{background:'none',border:'none',cursor:'pointer',color:T.dim,padding:2,display:'flex'}}>
                              <Trash2 size={12}/>
                            </button>
                          </div>
                        </div>

                        {/* Payment proof section */}
                        {isPersonal&&(
                          <div style={{borderTop:`1px solid ${T.borderLight}`,padding:'12px 16px',background:needsPayment?T.warningLight:'rgba(5,150,105,0.04)'}}>
                            {c.reimbursed&&c.adminPayment?(
                              // Already paid — show proof
                              <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',flexWrap:'wrap',gap:8}}>
                                <div>
                                  <div style={{fontSize:12,fontWeight:700,color:T.success,display:'flex',alignItems:'center',gap:6}}>
                                    <CheckCircle size={13}/>Payment recorded — {fmtDate(c.adminPayment.date)}
                                  </div>
                                  <div style={{fontSize:11,color:T.muted,marginTop:2}}>
                                    {c.adminPayment.method==='paynow'?'📱 PayNow':c.adminPayment.method==='bank_transfer'?'🏦 Bank Transfer':'💵 Cash'}
                                    {c.adminPayment.reference&&` · ${c.adminPayment.reference}`}
                                  </div>
                                  <div style={{fontSize:11,color:T.accent,marginTop:3,fontWeight:600}}>✓ Reflected in project expenses</div>
                                </div>
                                <div style={{display:'flex',gap:8,alignItems:'center'}}>
                                  {c.adminPayment.receiptImage&&(
                                    <button type="button" onClick={()=>setReceipt(c.adminPayment.receiptImage)}
                                      style={{background:T.successLight,border:`1px solid ${T.success}30`,borderRadius:8,padding:'5px 12px',cursor:'pointer',fontSize:12,color:T.success,fontFamily:'inherit',fontWeight:600,display:'flex',alignItems:'center',gap:5}}>
                                      <ZoomIn size={12}/>View Payment Proof
                                    </button>
                                  )}
                                  <button type="button" onClick={()=>{
                                    setAdminPayTarget(c.id);
                                    setAdminPayForm({method:c.adminPayment?.method||'',reference:c.adminPayment?.reference||'',receiptImage:null,date:new Date().toISOString().slice(0,10)});
                                  }}
                                    style={{background:T.bg,border:`1px solid ${T.borderLight}`,borderRadius:8,padding:'5px 12px',cursor:'pointer',fontSize:12,color:T.muted,fontFamily:'inherit',fontWeight:600,display:'flex',alignItems:'center',gap:5}}>
                                    <Edit3 size={11}/>Update
                                  </button>
                                </div>
                              </div>
                            ):(
                              // Not yet paid — show upload button prominently
                              <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:10,flexWrap:'wrap'}}>
                                <div>
                                  <div style={{fontSize:12,fontWeight:700,color:T.warning}}>⏳ Awaiting payment to {c.submittedBy}</div>
                                  <div style={{fontSize:11,color:T.muted,marginTop:2}}>Upload proof to reimburse staff and record as project expense</div>
                                </div>
                                <button type="button" onClick={()=>{
                                  setAdminPayTarget(c.id);
                                  setAdminPayForm({method:'',reference:'',receiptImage:null,date:new Date().toISOString().slice(0,10)});
                                }}
                                  style={{display:'flex',alignItems:'center',gap:7,padding:'9px 16px',borderRadius:10,border:'none',cursor:'pointer',fontFamily:'inherit',fontSize:13,fontWeight:700,background:T.accent,color:'#fff',whiteSpace:'nowrap'}}>
                                  <Upload size={14}/>Upload Payment Proof
                                </button>
                              </div>
                            )}
                          </div>
                        )}
                        {!isPersonal&&c.reflectedInExpenses&&(
                          <div style={{borderTop:`1px solid ${T.borderLight}`,padding:'10px 16px',background:'rgba(5,150,105,0.04)'}}>
                            <span style={{fontSize:12,color:T.success,fontWeight:600,display:'flex',alignItems:'center',gap:6}}>
                              <CheckCircle size={13}/>Recorded as project expense
                            </span>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })()}

          {/* Rejected / other history */}
          {(claims||[]).filter(c=>c.status==='Rejected').length>0&&(
            <div style={{marginTop:8}}>
              <div style={{fontSize:12,fontWeight:600,color:T.dim,textTransform:'uppercase',letterSpacing:'0.06em',marginBottom:8}}>Rejected</div>
              <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:16,overflow:'hidden'}}>
                {(claims||[]).filter(c=>c.status==='Rejected').sort((a,b)=>b.submittedAt.localeCompare(a.submittedAt)).map(c=>{
                  const proj=projects.find(p=>p.id===c.projectId);
                  const ct=STAFF_CLAIM_TYPES.find(t=>t.id===c.type);
                  return (
                    <div key={c.id} style={{padding:'12px 16px',borderBottom:`1px solid ${T.borderLight}`,display:'flex',alignItems:'center',gap:12}}>
                      <span style={{fontSize:18}}>{ct?.icon||'📎'}</span>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:12,fontWeight:600,color:T.text}}>{c.submittedBy} · {ct?.label||c.type}</div>
                        <div style={{fontSize:11,color:T.dim}}>{proj?.name||'—'} · {fmtDate(c.date)}</div>
                        {c.notes&&<div style={{fontSize:11,color:T.danger,marginTop:2}}>Reason: {c.notes}</div>}
                      </div>
                      <div style={{textAlign:'right',flexShrink:0,display:'flex',flexDirection:'column',alignItems:'flex-end',gap:4}}>
                        <div style={{fontSize:13,fontWeight:700,color:T.text}}>{fmtSGD(c.amount)}</div>
                        <button type="button" onClick={()=>deleteClaim(c.id)}
                          style={{background:'none',border:'none',cursor:'pointer',color:T.dim,padding:2,display:'flex'}}>
                          <Trash2 size={11}/>
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      )}

      {/* Reject reason modal */}
      {rejectTarget&&(
        <Modal title="Reject Claim" onClose={()=>{setRejectTarget(null);setRejectReason('');}}>
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <Field label="Reason for rejection (optional)" value={rejectReason} onChange={v=>setRejectReason(v)} placeholder="e.g. Receipt missing, amount incorrect"/>
            <div style={{display:'flex',justifyContent:'flex-end',gap:10}}>
              <Btn variant="secondary" onClick={()=>{setRejectTarget(null);setRejectReason('');}}>Cancel</Btn>
              <Btn variant="danger" onClick={()=>reject(rejectTarget)}>Reject Claim</Btn>
            </div>
          </div>
        </Modal>
      )}

      {/* Submit Expense Claim Modal */}
      {form&&(
        <Modal title="Submit Expense Claim" onClose={()=>setForm(null)} wide>
            <input ref={receiptRef} type="file" accept="image/*,application/pdf" capture="environment"
              style={{display:'none'}} onChange={e=>{const f=e.target.files?.[0];if(f)handleClaimReceiptUpload(f);e.target.value='';}}/>
            <div style={{display:'flex',flexDirection:'column',gap:18}}>

              {/* STEP 1 — Upload receipt */}
              <div>
                <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:8,display:'flex',alignItems:'center',gap:6}}>
                  <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>1</span>
                   Upload Receipt {(acctSettings?.anthropicApiKey||'').trim()&&<span style={{fontSize:10,color:T.accent,fontWeight:400,marginLeft:4}}>— AI will auto-fill details below</span>}
                </div>
              <DropZone accept="image/*,application/pdf" onDrop={f=>handleClaimReceiptUpload(f)}>
                {form.receiptImage?(
                  <div style={{position:'relative',background:T.bg,borderRadius:12,border:`1px solid ${T.borderLight}`,overflow:'hidden'}}>
                    {form.receiptImage.startsWith('data:application/pdf')
                      ?<div style={{padding:'20px',display:'flex',alignItems:'center',gap:12}}>
                          <span style={{fontSize:36}}>📄</span>
                          <div>
                            <div style={{fontSize:13,fontWeight:600,color:T.text}}>PDF uploaded</div>
                            <div style={{fontSize:11,color:T.dim}}>{claimOcrLoading?'AI reading receipt...':'Ready'}</div>
                          </div>
                        </div>
                      :<img src={form.receiptImage} alt="Receipt"
                          style={{width:'100%',maxHeight:200,objectFit:'contain',display:'block',background:T.bg}}/>
                    }
                    {claimOcrLoading&&(
                      <div style={{position:'absolute',bottom:0,left:0,right:0,background:'rgba(0,113,227,0.92)',padding:'8px 14px',display:'flex',alignItems:'center',gap:8}}>
                        <Loader2 size={13} style={{color:'#fff',animation:'spin 1s linear infinite'}}/>
                        <span style={{fontSize:12,color:'#fff',fontWeight:600}}>Reading receipt with AI...</span>
                      </div>
                    )}
                    <button type="button" onClick={()=>ff('receiptImage')(null)}
                      style={{position:'absolute',top:8,right:8,background:T.danger,border:'none',borderRadius:'50%',width:26,height:26,cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'center',color:'#fff',boxShadow:'0 2px 8px rgba(0,0,0,0.2)'}}>
                      <X size={13}/>
                    </button>
                  </div>
                ):(
                  <div onClick={()=>receiptRef.current?.click()}
                    style={{border:`2px dashed ${T.borderLight}`,borderRadius:12,padding:'28px 20px',textAlign:'center',cursor:'pointer',background:T.bg,display:'flex',flexDirection:'column',alignItems:'center',gap:10}}>
                    <div style={{width:52,height:52,background:T.accentLight,borderRadius:14,display:'flex',alignItems:'center',justifyContent:'center'}}>
                      <Camera size={22} style={{color:T.accent}}/>
                    </div>
                    <div style={{fontSize:13,fontWeight:700,color:T.text}}>Take photo or upload receipt</div>
                    <div style={{fontSize:11,color:T.dim}}>JPG, PNG or PDF — drag & drop or tap to upload</div>
                  </div>
                )}
                </DropZone>
              </div>

              {/* STEP 2 — Who paid */}
              <div>
                <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:8,display:'flex',alignItems:'center',gap:6}}>
                  <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>2</span>
                  Who paid?
                </div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:10}}>
                  {[
                    {id:'personal',label:'I paid myself',sub:'Company will reimburse me',icon:'💳',color:'#7c3aed'},
                    {id:'company', label:'Company paid', sub:'Record as project expense',icon:'🏢',color:T.accent},
                  ].map(opt=>(
                    <button key={opt.id} type="button" onClick={()=>ff('paidBy')(opt.id)}
                      style={{padding:'14px',borderRadius:12,textAlign:'left',
                        border:`2px solid ${form.paidBy===opt.id?opt.color:T.borderLight}`,
                        background:form.paidBy===opt.id?`${opt.color}08`:'transparent',
                        cursor:'pointer',fontFamily:'inherit',transition:'all 0.12s'}}>
                      <div style={{fontSize:22,marginBottom:6}}>{opt.icon}</div>
                      <div style={{fontSize:13,fontWeight:700,color:form.paidBy===opt.id?opt.color:T.text}}>{opt.label}</div>
                      <div style={{fontSize:11,color:T.muted,marginTop:2}}>{opt.sub}</div>
                    </button>
                  ))}
                </div>
              </div>

              {/* STEP 3 — Expense type */}
              <div>
                <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:8,display:'flex',alignItems:'center',gap:6}}>
                  <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>3</span>
                  Type of Expense
                </div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:8}}>
                  {STAFF_CLAIM_TYPES.map(t=>(
                    <button key={t.id} type="button" onClick={()=>ff('type')(t.id)}
                      style={{padding:'10px 12px',borderRadius:10,
                        border:`2px solid ${form.type===t.id?T.accent:T.borderLight}`,
                        cursor:'pointer',fontFamily:'inherit',
                        background:form.type===t.id?T.accentLight:'transparent',
                        display:'flex',alignItems:'center',gap:10,textAlign:'left',transition:'all 0.12s'}}>
                      <span style={{fontSize:20}}>{t.icon}</span>
                      <span style={{fontSize:12,fontWeight:600,color:form.type===t.id?T.accent:T.text}}>{t.label}</span>
                    </button>
                  ))}
                </div>
              </div>

              {/* STEP 4 — Amount, date, description, project */}
              <div>
                <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:10,display:'flex',alignItems:'center',gap:6}}>
                  <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>4</span>
                  Claim Details
                  {claimOcrLoading&&<span style={{fontSize:10,color:T.accent,fontWeight:400}}>— filling from receipt...</span>}
                </div>
                <div style={{display:'flex',flexDirection:'column',gap:12}}>
                  <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:12}}>
                    <Field label="Amount (S$) *" type="number" value={form.amount} onChange={ff('amount')} placeholder="0.00"/>
                    <Field label="Date *" type="date" value={form.date} onChange={ff('date')}/>
                  </div>
                  <Field label="Description of items claimed" value={form.description||''} onChange={ff('description')} placeholder="e.g. Nippon paint 10L × 3 tins, brushes"/>
                  <div>
                    <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Project *</label>
                    <select value={form.projectId} onChange={e=>ff('projectId')(e.target.value)}
                      style={{...iStyle,appearance:'none',backgroundImage:`url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8' fill='none'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%23AEAEB2' stroke-width='1.5' stroke-linecap='round'/%3E%3C/svg%3E")`,backgroundRepeat:'no-repeat',backgroundPosition:'right 13px center',paddingRight:36}}>
                      <option value="">— Select project —</option>
                      <option value="__company__">🏢 Company Expense (no project)</option>
                      {myProjects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
                    </select>
                    {form.projectId==='__company__'&&(
                      <div style={{marginTop:6,fontSize:11,color:T.muted,background:T.accentLight,borderRadius:8,padding:'6px 10px'}}>
                        This expense will be recorded under Company Accounts, not linked to a project.
                      </div>
                    )}
                  </div>
                </div>
              </div>

              <div style={{display:'flex',justifyContent:'flex-end',gap:10,paddingTop:4,borderTop:`1px solid ${T.borderLight}`}}>
                {(!form.type||!form.amount||parseFloat(form.amount)<=0||!form.paidBy||!form.projectId)&&(
                  <span style={{fontSize:12,color:T.muted,alignSelf:'center',marginRight:'auto'}}>
                    {!form.projectId?'Select a project':!form.paidBy?'Select who paid':!form.type?'Select expense type':(!form.amount||parseFloat(form.amount)<=0)?'Enter amount':''}
                  </span>
                )}
                <Btn variant="secondary" onClick={()=>setForm(null)}>Cancel</Btn>
                <Btn onClick={submitClaim} disabled={!form.projectId||!form.type||!form.amount||parseFloat(form.amount)<=0||!form.paidBy}>
                  <CheckCircle size={13}/>Submit Claim
                </Btn>
              </div>
            </div>
          </Modal>
      )}

      {/* Admin Payment Modal — record company payment to staff */}
      {adminPayTarget&&(()=>{
        const claim=(claims||[]).find(c=>c.id===adminPayTarget);
        if(!claim) return null;
        const ct=STAFF_CLAIM_TYPES.find(t=>t.id===claim.type);
        return (
          <Modal title="Record Company Payment to Staff" onClose={()=>setAdminPayTarget(null)} wide>
            <input ref={adminPayReceiptRef} type="file" accept="image/*,application/pdf" capture="environment"
              style={{display:'none'}} onChange={e=>{const f=e.target.files?.[0];if(f)extractAdminPayReceipt(f);e.target.value='';}}/>
            <div style={{display:'flex',flexDirection:'column',gap:16}}>

              {/* Claim being paid */}
              <div style={{background:T.bg,borderRadius:12,padding:'12px 14px',display:'flex',alignItems:'center',justifyContent:'space-between',gap:12}}>
                <div>
                  <div style={{fontSize:12,color:T.muted,marginBottom:2}}>Reimbursing</div>
                  <div style={{fontSize:15,fontWeight:700,color:T.text}}>{claim.submittedBy}</div>
                  <div style={{fontSize:11,color:T.dim,marginTop:2}}>
                    {ct?.icon} {ct?.label||claim.type} · {fmtDate(claim.date)}
                    {claim.description&&<span> · {claim.description}</span>}
                  </div>
                  {claim.receiptImage&&(
                    <button type="button" onClick={()=>setReceipt(claim.receiptImage)}
                      style={{marginTop:6,background:T.accentLight,border:'none',borderRadius:6,padding:'2px 9px',cursor:'pointer',fontSize:11,color:T.accent,fontFamily:'inherit',fontWeight:600,display:'inline-flex',alignItems:'center',gap:4}}>
                      <ZoomIn size={10}/>View Staff Receipt
                    </button>
                  )}
                </div>
                <div style={{textAlign:'right',flexShrink:0}}>
                  <div style={{fontSize:20,fontWeight:800,color:T.accent}}>{fmtSGD(claim.amount)}</div>
                  <div style={{fontSize:10,color:T.muted,marginTop:2}}>to reimburse</div>
                </div>
              </div>

              {/* STEP 1 — Upload payment screenshot */}
              <div>
                <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:8,display:'flex',alignItems:'center',gap:6}}>
                  <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>1</span>
                  Upload Payment Screenshot / Proof
                  {acctSettings?.anthropicApiKey&&<span style={{fontSize:10,color:T.accent,fontWeight:400}}>— AI fills details below</span>}
                </div>
                <DropZone accept="image/*,application/pdf" onDrop={f=>extractAdminPayReceipt(f)}>
                {adminPayForm.receiptImage?(
                  <div style={{position:'relative',background:T.bg,borderRadius:12,border:`1px solid ${adminOcrLoading?T.accent:T.success}`,overflow:'hidden'}}>
                    {adminPayForm.receiptImage.startsWith('data:application/pdf')
                      ?<div style={{padding:'16px 20px',display:'flex',alignItems:'center',gap:12}}>
                          <span style={{fontSize:36}}>📄</span>
                          <div>
                            <div style={{fontSize:13,fontWeight:600,color:T.text}}>{adminOcrLoading?'AI reading payment details...':'PDF uploaded ✓'}</div>
                            <div style={{fontSize:11,color:T.dim,marginTop:2}}>{adminOcrLoading?'Extracting method, reference and date...':'Check fields below'}</div>
                          </div>
                        </div>
                      :<img src={adminPayForm.receiptImage} alt="Payment proof"
                          style={{width:'100%',maxHeight:200,objectFit:'contain',display:'block'}}/>
                    }
                    {adminOcrLoading&&(
                      <div style={{position:'absolute',bottom:0,left:0,right:0,background:'rgba(0,113,227,0.9)',padding:'8px 14px',display:'flex',alignItems:'center',gap:8}}>
                        <Loader2 size={13} style={{color:'#fff',animation:'spin 1s linear infinite'}}/>
                        <span style={{fontSize:12,color:'#fff',fontWeight:600}}>Reading payment details with AI...</span>
                      </div>
                    )}
                    <button type="button" onClick={()=>apf('receiptImage')(null)}
                      style={{position:'absolute',top:8,right:8,background:T.danger,border:'none',borderRadius:'50%',width:26,height:26,cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'center',color:'#fff'}}>
                      <X size={13}/>
                    </button>
                  </div>
                ):(
                  <div onClick={()=>adminPayReceiptRef.current?.click()}
                    style={{border:`2px dashed ${T.borderLight}`,borderRadius:12,padding:'24px 20px',textAlign:'center',cursor:'pointer',background:T.bg,display:'flex',flexDirection:'column',alignItems:'center',gap:10}}>
                    <div style={{width:50,height:50,background:T.accentLight,borderRadius:14,display:'flex',alignItems:'center',justifyContent:'center'}}>
                      <Upload size={20} style={{color:T.accent}}/>
                    </div>
                    <div>
                      <div style={{fontSize:13,fontWeight:700,color:T.text}}>Upload or drop payment screenshot</div>
                      <div style={{fontSize:11,color:T.dim,marginTop:3}}>PayNow / Bank Transfer receipt — JPG, PNG or PDF</div>
                    </div>
                  </div>
                )}
                </DropZone>
              </div>

              {/* STEP 2 — Verify payment details */}
              <div>
                <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.08em',marginBottom:10,display:'flex',alignItems:'center',gap:6}}>
                  <span style={{background:T.accent,color:'#fff',borderRadius:'50%',width:18,height:18,display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:10,fontWeight:800,flexShrink:0}}>2</span>
                  Verify Payment Details
                  {adminOcrLoading&&<span style={{fontSize:10,color:T.accent,fontWeight:400}}>— filling from screenshot...</span>}
                </div>
                <div style={{display:'flex',flexDirection:'column',gap:12}}>
                  <div>
                    <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:8}}>Payment Method *</label>
                    <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:8}}>
                      {[{id:'paynow',label:'PayNow',icon:'📱'},{id:'bank_transfer',label:'Bank Transfer',icon:'🏦'},{id:'cash',label:'Cash',icon:'💵'}].map(m=>(
                        <button key={m.id} type="button" onClick={()=>apf('method')(m.id)}
                          style={{padding:'10px 8px',borderRadius:10,border:`2px solid ${adminPayForm.method===m.id?T.accent:T.borderLight}`,
                            background:adminPayForm.method===m.id?T.accentLight:'transparent',cursor:'pointer',fontFamily:'inherit',
                            display:'flex',flexDirection:'column',alignItems:'center',gap:4}}>
                          <span style={{fontSize:20}}>{m.icon}</span>
                          <span style={{fontSize:11,fontWeight:600,color:adminPayForm.method===m.id?T.accent:T.muted}}>{m.label}</span>
                        </button>
                      ))}
                    </div>
                  </div>
                  {adminPayForm.method&&(
                    <Field label={adminPayForm.method==='paynow'?'PayNow Number / UEN':adminPayForm.method==='bank_transfer'?'Bank Account Number':'Receipt / Reference'}
                      value={adminPayForm.reference} onChange={apf('reference')}
                      placeholder={adminPayForm.method==='paynow'?'e.g. 9123 4567 or UEN':adminPayForm.method==='bank_transfer'?'e.g. DBS 0721-0976-05':'e.g. Receipt #001'}/>
                  )}
                  <Field label="Date Paid *" type="date" value={adminPayForm.date} onChange={apf('date')}/>
                </div>
              </div>

              <div style={{display:'flex',justifyContent:'flex-end',gap:10,borderTop:`1px solid ${T.borderLight}`,paddingTop:14}}>
                <Btn variant="secondary" onClick={()=>setAdminPayTarget(null)}>Cancel</Btn>
                <Btn onClick={saveAdminPayment} disabled={!adminPayForm.date}>
                  <CheckCircle size={13}/>Confirm Payment — {fmtSGD(claim.amount)}
                </Btn>
              </div>
            </div>
          </Modal>
        );
      })()}

      {/* Receipt lightbox */}
      {receipt&&(
        <Lightbox src={receipt} title="Receipt / Proof" onClose={()=>setReceipt(null)}/>
      )}
    </div>
  );
}

function Commissions({projects,setProjects,invoices,isAdmin,users=[]}){
  const [payModal,setPayModal]=useState(null);
  const [payForm,setPayForm]=useState({date:new Date().toISOString().slice(0,10),method:'bank_transfer',reference:'',notes:'',dPaid:true,pmPaid:true});
  const pf=k=>v=>setPayForm(p=>({...p,[k]:v}));

  const markPaid=(proj)=>{
    const payout={
      date:payForm.date,
      method:payForm.method,
      reference:payForm.reference,
      notes:payForm.notes,
      dPaid:payForm.dPaid,
      pmPaid:payForm.pmPaid,
      paidBy:users.find(u=>u.role==='admin')?.name||'Admin',
      paidAt:new Date().toISOString(),
    };
    const upd=projects.map(p=>p.id===proj.id?{
      ...p,
      commissionPaid:true,
      commissionPaidAt:new Date().toISOString(),
      commissionPayout:payout,
      // Store individual payout records
      commissionPayoutHistory:[...(p.commissionPayoutHistory||[]),{...payout,dComm:proj.dComm,pmComm:proj.pmComm,id:uid()}],
    }:p);
    setProjects(upd);saveProjects(upd);setPayModal(null);
  };

  const canPayout=(proj)=>{
    const projInv=invoices.filter(i=>i.projectId===proj.id);
    return projInv.length>0&&projInv.every(i=>i.status==='Paid');
  };

  const allRows=projects
    .filter(p=>!p.archived||(p.archived&&!p.commissionPaid))
    .map(p=>({...p,...calcComm(p,invoices)}))
    .filter(r=>r.dComm>0||r.pmComm>0);

  const unpaidRows=allRows.filter(r=>!r.commissionPaid);
  const paidRows=allRows.filter(r=>r.commissionPaid).sort((a,b)=>new Date(b.commissionPaidAt)-new Date(a.commissionPaidAt));

  const byDesigner={},byPM={};
  const yearlyDesigner={},yearlyPM={};

  unpaidRows.forEach(r=>{
    byDesigner[r.designer]=(byDesigner[r.designer]||0)+r.dComm;
    byPM[r.pm]=(byPM[r.pm]||0)+r.pmComm;
  });

  allRows.forEach(r=>{
    const yr=r.startDate?new Date(r.startDate).getFullYear():new Date().getFullYear();
    if(!yearlyDesigner[r.designer]) yearlyDesigner[r.designer]={};
    if(!yearlyPM[r.pm]) yearlyPM[r.pm]={};
    yearlyDesigner[r.designer][yr]=(yearlyDesigner[r.designer][yr]||0)+r.dComm;
    yearlyPM[r.pm][yr]=(yearlyPM[r.pm][yr]||0)+r.pmComm;
  });

  // Per-person paid history
  const personHistory=useMemo(()=>{
    const map={};
    paidRows.forEach(r=>{
      const po=r.commissionPayout||{};
      const addEntry=(name,role,amount)=>{
        if(!map[name]) map[name]={name,role,entries:[],total:0};
        map[name].entries.push({project:r.name,amount,date:po.date||r.commissionPaidAt,method:po.method,reference:po.reference,notes:po.notes});
        map[name].total+=amount;
      };
      if(r.dComm>0) addEntry(r.designer,'designer',r.dComm);
      if(r.pmComm>0) addEntry(r.pm,'pm',r.pmComm);
    });
    return Object.values(map).sort((a,b)=>a.name.localeCompare(b.name));
  },[paidRows]);

  const allYears=[...new Set(allRows.map(r=>r.startDate?new Date(r.startDate).getFullYear():new Date().getFullYear()))].sort((a,b)=>b-a);
  const pending=payModal?allRows.find(r=>r.id===payModal):null;
  const methodLabels={bank_transfer:'🏦 Bank Transfer',paynow:'📱 PayNow',cash:'💵 Cash',cheque:'📝 Cheque'};

  return (
    <div style={{display:'flex',flexDirection:'column',gap:20}}>

      {/* Individual balances */}
      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:16}}>
        {/* Designers */}
        <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow}}>
          <div style={{fontSize:14,fontWeight:600,color:T.text,marginBottom:16,display:'flex',alignItems:'center',gap:8}}>
            <Users size={15} style={{color:T.accent}}/>Designers — Outstanding Commission
          </div>
          {Object.keys(byDesigner).length===0&&(
            <div style={{fontSize:13,color:T.dim,fontStyle:'italic'}}>All designer commissions paid</div>
          )}
          {Object.entries(byDesigner).map(([name,unpaid])=>{
            const yearData=yearlyDesigner[name]||{};
            return (
              <div key={name} style={{padding:'12px 16px',background:T.bg,borderRadius:12,marginBottom:10,border:`1px solid ${T.borderLight}`}}>
                <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:8}}>
                  <div>
                    <div style={{fontSize:14,fontWeight:600,color:T.text}}>{name}</div>
                    <div style={{fontSize:11,color:T.dim}}>{ROLE_LABEL[users.find(u=>u.name===name)?.role]||'Designer'} · Design Commission</div>
                  </div>
                  <div style={{textAlign:'right'}}>
                    <div style={{fontSize:18,fontWeight:700,color:T.accent}}>{fmtSGD(unpaid)}</div>
                    <div style={{fontSize:11,color:T.dim}}>outstanding</div>
                  </div>
                </div>
                {allYears.length>0&&(
                  <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
                    {allYears.map(yr=>(
                      yearData[yr]?<div key={yr} style={{background:'rgba(0,113,227,0.07)',borderRadius:7,padding:'3px 10px',display:'flex',gap:6,alignItems:'center'}}>
                        <span style={{fontSize:10,color:T.accent,fontWeight:600}}>{yr}</span>
                        <span style={{fontSize:11,color:T.text,fontWeight:700}}>{fmtSGD(yearData[yr])}</span>
                      </div>:null
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>

        {/* PMs */}
        <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow}}>
          <div style={{fontSize:14,fontWeight:600,color:T.text,marginBottom:16,display:'flex',alignItems:'center',gap:8}}>
            <Users size={15} style={{color:T.info}}/>Project Managers — Outstanding Commission
          </div>
          {Object.keys(byPM).length===0&&(
            <div style={{fontSize:13,color:T.dim,fontStyle:'italic'}}>All PM commissions paid</div>
          )}
          {Object.entries(byPM).map(([name,unpaid])=>{
            const yearData=yearlyPM[name]||{};
            return (
              <div key={name} style={{padding:'12px 16px',background:T.bg,borderRadius:12,marginBottom:10,border:`1px solid ${T.borderLight}`}}>
                <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:8}}>
                  <div>
                    <div style={{fontSize:14,fontWeight:600,color:T.text}}>{name}</div>
                    <div style={{fontSize:11,color:T.dim}}>{ROLE_LABEL[users.find(u=>u.name===name)?.role]||'Project Manager'} · PM Commission</div>
                  </div>
                  <div style={{textAlign:'right'}}>
                    <div style={{fontSize:18,fontWeight:700,color:T.info}}>{fmtSGD(unpaid)}</div>
                    <div style={{fontSize:11,color:T.dim}}>outstanding</div>
                  </div>
                </div>
                {allYears.length>0&&(
                  <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
                    {allYears.map(yr=>(
                      yearData[yr]?<div key={yr} style={{background:'rgba(8,100,200,0.07)',borderRadius:7,padding:'3px 10px',display:'flex',gap:6,alignItems:'center'}}>
                        <span style={{fontSize:10,color:T.info,fontWeight:600}}>{yr}</span>
                        <span style={{fontSize:11,color:T.text,fontWeight:700}}>{fmtSGD(yearData[yr])}</span>
                      </div>:null
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>

      {/* Unpaid by project */}
      <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,overflow:'hidden',boxShadow:T.shadow}}>
        <div style={{padding:'16px 22px',borderBottom:`1px solid ${T.borderLight}`,display:'flex',alignItems:'center',justifyContent:'space-between'}}>
          <div>
            <div style={{fontSize:14,fontWeight:600,color:T.text}}>Unpaid Commission by Project</div>
            <div style={{fontSize:12,color:T.dim,marginTop:2}}>Only projects with outstanding commission are shown</div>
          </div>
          {unpaidRows.length>0&&<Badge color={T.warning}>{unpaidRows.length} outstanding</Badge>}
        </div>
        {unpaidRows.length===0?(
          <div style={{padding:'32px',textAlign:'center',color:T.dim,fontSize:13}}>All project commissions have been paid out.</div>
        ):(
          <div style={{overflowX:'auto'}}>
            <table style={{width:'100%',fontSize:13,borderCollapse:'collapse'}}>
              <thead>
                <tr style={{color:T.dim,fontSize:11,fontWeight:600,textTransform:'uppercase',letterSpacing:'0.06em',background:T.bg}}>
                  {['Project','Method','Designer','PM','Revenue','Expenses','Gross Profit','D. Commission','PM Commission','Net Profit',''].map(h=>(
                    <th key={h} style={{textAlign:['Revenue','Expenses','Gross Profit','D. Commission','PM Commission','Net Profit'].includes(h)?'right':'left',padding:'10px 14px',whiteSpace:'nowrap'}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {unpaidRows.map(r=>{
                  const net=r.gross-r.dComm-r.pmComm;
                  return (
                    <tr key={r.id} style={{borderTop:`1px solid ${T.borderLight}`}}>
                      <td style={{padding:'10px 14px',color:T.text,fontWeight:600,maxWidth:140,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.name}</td>
                      <td style={{padding:'10px 14px'}}><Badge color={T.info} sm>{COMM_METHOD_LABEL[r.method]||r.method}</Badge></td>
                      <td style={{padding:'10px 14px',color:T.muted,fontSize:12}}>{r.designer}</td>
                      <td style={{padding:'10px 14px',color:T.muted,fontSize:12}}>{r.pm}</td>
                      <td style={{padding:'10px 14px',textAlign:'right',color:T.text}}>{fmtSGD(r.rev)}</td>
                      <td style={{padding:'10px 14px',textAlign:'right',color:T.danger}}>{fmtSGD(r.exp)}</td>
                      <td style={{padding:'10px 14px',textAlign:'right',color:r.gross>0?T.success:T.danger,fontWeight:600}}>{fmtSGD(r.gross)}</td>
                      <td style={{padding:'10px 14px',textAlign:'right',color:T.accent,fontWeight:600}}>{fmtSGD(r.dComm)}</td>
                      <td style={{padding:'10px 14px',textAlign:'right',color:T.info,fontWeight:600}}>{fmtSGD(r.pmComm)}</td>
                      <td style={{padding:'10px 14px',textAlign:'right',color:net>0?T.success:T.danger,fontWeight:600}}>{fmtSGD(net)}</td>
                      <td style={{padding:'10px 14px'}}>
                        {isAdmin&&(
                          canPayout(r)
                            ? <Btn size="sm" onClick={()=>setPayModal(r.id)}><CheckCircle size={11}/>Pay Out</Btn>
                            : <span title="All supplier invoices must be Paid before commission can be released"
                                style={{fontSize:11,color:T.dim,display:'flex',alignItems:'center',gap:4,cursor:'help'}}>
                                <Lock size={11}/>Invoices pending
                              </span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Per-person payout history */}
      {personHistory.length>0&&(
        <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,overflow:'hidden',boxShadow:T.shadow}}>
          <div style={{padding:'16px 22px',borderBottom:`1px solid ${T.borderLight}`}}>
            <div style={{fontSize:14,fontWeight:600,color:T.text}}>Commission Payout History — By Person</div>
            <div style={{fontSize:12,color:T.dim,marginTop:2}}>Full record of all commission payments made</div>
          </div>
          <div style={{padding:'16px 22px',display:'flex',flexDirection:'column',gap:14}}>
            {personHistory.map(person=>(
              <div key={person.name} style={{background:T.bg,borderRadius:14,overflow:'hidden',border:`1px solid ${T.borderLight}`}}>
                {/* Person header */}
                <div style={{padding:'12px 16px',display:'flex',alignItems:'center',justifyContent:'space-between',borderBottom:`1px solid ${T.borderLight}`,background:T.card}}>
                  <div style={{display:'flex',alignItems:'center',gap:10}}>
                    <div style={{width:34,height:34,borderRadius:'50%',background:person.role==='designer'?T.accentLight:'rgba(8,100,200,0.1)',display:'flex',alignItems:'center',justifyContent:'center',fontSize:14}}>
                      {person.name.charAt(0)}
                    </div>
                    <div>
                      <div style={{fontSize:13,fontWeight:700,color:T.text}}>{person.name}</div>
                      <div style={{fontSize:11,color:T.dim}}>{person.role==='designer'?'Designer':'Project Manager'} · {person.entries.length} payout{person.entries.length!==1?'s':''}</div>
                    </div>
                  </div>
                  <div style={{textAlign:'right'}}>
                    <div style={{fontSize:16,fontWeight:800,color:person.role==='designer'?T.accent:T.info}}>{fmtSGD(person.total)}</div>
                    <div style={{fontSize:10,color:T.dim}}>total paid</div>
                  </div>
                </div>
                {/* Entries */}
                {person.entries.map((e,i)=>(
                  <div key={i} style={{padding:'10px 16px',borderBottom:i<person.entries.length-1?`1px solid ${T.borderLight}`:'none',display:'flex',alignItems:'center',gap:12}}>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:12,fontWeight:600,color:T.text,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{e.project}</div>
                      <div style={{display:'flex',gap:8,alignItems:'center',marginTop:3,flexWrap:'wrap'}}>
                        <span style={{fontSize:11,color:T.muted}}>{fmtDate(e.date)}</span>
                        {e.method&&<span style={{fontSize:11,color:T.muted}}>{methodLabels[e.method]||e.method}</span>}
                        {e.reference&&<span style={{fontSize:11,color:T.dim}}>Ref: {e.reference}</span>}
                        {e.notes&&<span style={{fontSize:11,color:T.dim,fontStyle:'italic'}}>{e.notes}</span>}
                      </div>
                    </div>
                    <div style={{fontSize:13,fontWeight:700,color:T.success,display:'flex',alignItems:'center',gap:5,flexShrink:0}}>
                      <CheckCircle size={12}/>{fmtSGD(e.amount)}
                    </div>
                  </div>
                ))}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Paid by project (compact) */}
      {paidRows.length>0&&(
        <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,overflow:'hidden',boxShadow:T.shadow}}>
          <div style={{padding:'16px 22px',borderBottom:`1px solid ${T.borderLight}`}}>
            <div style={{fontSize:14,fontWeight:600,color:T.text}}>Paid Projects</div>
          </div>
          <div style={{overflowX:'auto'}}>
            <table style={{width:'100%',fontSize:13,borderCollapse:'collapse'}}>
              <thead>
                <tr style={{color:T.dim,fontSize:11,fontWeight:600,textTransform:'uppercase',letterSpacing:'0.06em',background:T.bg}}>
                  {['Project','Designer','PM','D. Comm','PM Comm','Method','Ref','Paid On'].map(h=>(
                    <th key={h} style={{textAlign:['D. Comm','PM Comm'].includes(h)?'right':'left',padding:'10px 14px',whiteSpace:'nowrap'}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {paidRows.map(r=>{
                  const po=r.commissionPayout||{};
                  return (
                    <tr key={r.id} style={{borderTop:`1px solid ${T.borderLight}`,opacity:0.75}}>
                      <td style={{padding:'10px 14px',color:T.text,fontWeight:600,maxWidth:140,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.name}</td>
                      <td style={{padding:'10px 14px',color:T.muted,fontSize:12}}>{r.designer}</td>
                      <td style={{padding:'10px 14px',color:T.muted,fontSize:12}}>{r.pm}</td>
                      <td style={{padding:'10px 14px',textAlign:'right',color:T.dim}}>{fmtSGD(r.dComm)}</td>
                      <td style={{padding:'10px 14px',textAlign:'right',color:T.dim}}>{fmtSGD(r.pmComm)}</td>
                      <td style={{padding:'10px 14px',fontSize:12,color:T.muted}}>{methodLabels[po.method]||po.method||'—'}</td>
                      <td style={{padding:'10px 14px',fontSize:12,color:T.dim,maxWidth:100,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{po.reference||'—'}</td>
                      <td style={{padding:'10px 14px',color:T.success,fontSize:12,display:'flex',alignItems:'center',gap:5}}>
                        <CheckCircle size={12}/>{fmtDate(r.commissionPaidAt)}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Pay out confirmation modal — enhanced */}
      {payModal&&pending&&(
        <Modal title="Record Commission Payout" onClose={()=>setPayModal(null)}>
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            {/* Summary */}
            <div style={{background:T.bg,borderRadius:12,padding:'14px 16px'}}>
              <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:8}}>{pending.name}</div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:8}}>
                <div style={{background:T.card,borderRadius:10,padding:'10px 14px',border:`1px solid ${T.accent}20`}}>
                  <div style={{fontSize:11,color:T.dim,marginBottom:2}}>{pending.designer} (Designer)</div>
                  <div style={{fontSize:16,fontWeight:800,color:T.accent}}>{fmtSGD(pending.dComm)}</div>
                </div>
                <div style={{background:T.card,borderRadius:10,padding:'10px 14px',border:`1px solid ${T.info}20`}}>
                  <div style={{fontSize:11,color:T.dim,marginBottom:2}}>{pending.pm} (Project Manager)</div>
                  <div style={{fontSize:16,fontWeight:800,color:T.info}}>{fmtSGD(pending.pmComm)}</div>
                </div>
              </div>
              <div style={{marginTop:10,padding:'8px 12px',background:T.successLight,borderRadius:9,display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                <span style={{fontSize:12,color:T.success,fontWeight:600}}>Total Payout</span>
                <span style={{fontSize:15,fontWeight:800,color:T.success}}>{fmtSGD(pending.dComm+pending.pmComm)}</span>
              </div>
            </div>

            {/* Payment details */}
            <div>
              <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:8}}>Payment Method</label>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:8}}>
                {[{id:'bank_transfer',label:'Bank Transfer',icon:'🏦'},{id:'paynow',label:'PayNow',icon:'📱'},{id:'cash',label:'Cash',icon:'💵'},{id:'cheque',label:'Cheque',icon:'📝'}].map(m=>(
                  <button key={m.id} type="button" onClick={()=>pf('method')(m.id)}
                    style={{padding:'8px 10px',borderRadius:10,border:`2px solid ${payForm.method===m.id?T.accent:T.borderLight}`,
                      background:payForm.method===m.id?T.accentLight:'transparent',cursor:'pointer',fontFamily:'inherit',
                      display:'flex',alignItems:'center',gap:8}}>
                    <span style={{fontSize:18}}>{m.icon}</span>
                    <span style={{fontSize:12,fontWeight:600,color:payForm.method===m.id?T.accent:T.muted}}>{m.label}</span>
                  </button>
                ))}
              </div>
            </div>
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:12}}>
              <Field label="Date Paid" type="date" value={payForm.date} onChange={pf('date')}/>
              <Field label="Reference / Transfer No." value={payForm.reference} onChange={pf('reference')} placeholder="e.g. TRF-2025-0012"/>
            </div>
            <Field label="Notes (optional)" value={payForm.notes} onChange={pf('notes')} placeholder="e.g. Combined with March salary"/>

            <div style={{display:'flex',justifyContent:'flex-end',gap:10,marginTop:4,borderTop:`1px solid ${T.borderLight}`,paddingTop:14}}>
              <Btn variant="secondary" onClick={()=>setPayModal(null)}>Cancel</Btn>
              <Btn onClick={()=>markPaid(pending)}><CheckCircle size={13}/>Confirm Payout</Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

function Admin({users,setUsers,projects,onSoftDelete,onShowToast,actionLog=[],onUndoAction,isSuperAdmin}){
  const [adminTab,setAdminTab]=useState('users'); // 'users' | 'log'
  const [modal,setModal]=useState(null);
  const [permModal,setPermModal]=useState(null);
  const [pwModal,setPwModal]=useState(null);
  const [search,setSearch]=useState('');
  // password is kept in its own state, completely isolated from form
  // so no other setForm call can ever corrupt it
  const [pwInput,setPwInput]=useState('');
  const [showPw,setShowPw]=useState(false);

  const blankUser={name:'',email:'',role:'expense_entry',assignedProjects:[],
    photo:'', tabs:ROLE_DEFAULT_TABS.expense_entry,
    widgets:DASH_WIDGETS.map(w=>w.id), active:true};
  const [form,setForm]=useState(blankUser);
  const ff=k=>v=>setForm(p=>({...p,[k]:v}));

  const openNew=()=>{
    setForm({...blankUser});
    setPwInput('');
    setShowPw(false);
    setModal('new');
  };

  const openEdit=(u)=>{
    setForm({...u});
    setPwInput('');
    setShowPw(false);
    setModal(u.id);
  };

  const closeModal=()=>{
    setModal(null);
    setPwInput('');
    setShowPw(false);
  };

  const editingUser = modal && modal!=='new' ? users.find(u=>u.id===modal) : null;
  const permUser = permModal ? users.find(u=>u.id===permModal) : null;

  const onRoleChange=role=>{
    setForm(p=>({...p,role,tabs:ROLE_DEFAULT_TABS[role]||[],assignedProjects:p.assignedProjects}));
  };

  const toggleProject=(uid,pid)=>{
    setUsers(prev=>prev.map(u=>{
      if(u.id!==uid)return u;
      const has=u.assignedProjects.includes(pid);
      return {...u,assignedProjects:has?u.assignedProjects.filter(x=>x!==pid):[...u.assignedProjects,pid]};
    }));
  };

  const toggleTab=(uid,tabId)=>{
    setUsers(prev=>prev.map(u=>{
      if(u.id!==uid)return u;

      if(u.role==='admin'&&tabId==='admin')return u;
      const has=u.tabs.includes(tabId);
      return {...u,tabs:has?u.tabs.filter(x=>x!==tabId):[...u.tabs,tabId]};
    }));
  };

  const toggleWidget=(uid,wid)=>{
    setUsers(prev=>prev.map(u=>{
      if(u.id!==uid)return u;
      const has=u.widgets.includes(wid);
      return {...u,widgets:has?u.widgets.filter(x=>x!==wid):[...u.widgets,wid]};
    }));
  };

  const toggleActive=(uid)=>{
    const upd=users.map(u=>u.id===uid?{...u,active:!u.active}:u);
    setUsers(upd);saveUsers(upd);
  };

  const saveUser=()=>{
    if(!form.name||!form.email)return;
    const existingPw = modal!=='new' ? (users.find(u=>u.id===modal)?.password||'') : '';
    const finalPw = pwInput.trim() ? pwInput.trim() : existingPw;
    const userRecord = {...form, password:finalPw};
    let upd;
    if(modal==='new'){
      upd=[...users,{...userRecord,id:uid()}];
    } else {
      upd=users.map(u=>u.id===modal?{...userRecord,id:modal}:u);
    }
    setUsers(upd);
    // Save with photos (already compressed to ~5KB each, safe for Firestore)
    saveUsers(upd);
    closeModal();
  };

  const savePerms=()=>{ saveUsers(users); setPermModal(null); };
  const [deleteTarget,setDeleteTarget]=useState(null);

  const delUser=(id)=>{ const u=users.find(x=>x.id===id); if(u)setDeleteTarget(u); };
  const confirmDelUser=(u)=>{
    onSoftDelete({...u,_trashType:'user',_deletedAt:new Date().toISOString()});
    const upd=users.filter(x=>x.id!==u.id);setUsers(upd);saveUsers(upd);
    setDeleteTarget(null);
    onShowToast(`User ${u.name} moved to Trash`,()=>{setUsers(prev=>[...prev,u]);saveUsers([...users,u]);});
  };

  const ALL_TABS=[
    {id:'dashboard',label:'Dashboard'},
    {id:'projects',label:'Projects'},
    {id:'invoices',label:'Invoices'},
    {id:'payments',label:'Payments'},
    {id:'contacts',label:'Contacts'},
    {id:'reports',label:'Reports'},
    {id:'commissions',label:'Commissions'},
    {id:'warranty',label:'Warranty'},
    {id:'workers',label:'Site Workers'},
    {id:'checkin',label:'Check In/Out (Workers)'},
    {id:'accounts',label:'Company Accounts (Admin only)'},
    {id:'trash',label:'Trash (Admin only)'},
    {id:'admin',label:'Admin (restricted)'},
  ];

  const filtered=users.filter(u=>(u.name+u.email).toLowerCase().includes(search.toLowerCase()));

  const Toggle=({on,onClick})=>(
    <button onClick={onClick}
      style={{background:'none',border:'none',cursor:'pointer',color:on?T.success:T.dim,display:'flex',padding:0,transition:'color .15s'}}>
      {on?<ToggleRight size={22}/>:<ToggleLeft size={22}/>}
    </button>
  );

  return (
    <div style={{display:'flex',flexDirection:'column',gap:16}}>
      {/* Tab switcher */}
      <div style={{display:'flex',background:T.bg,borderRadius:10,border:`1px solid ${T.borderLight}`,overflow:'hidden',alignSelf:'flex-start'}}>
        {[{id:'users',label:'Users'},{id:'log',label:`Action Log (${actionLog.length})`}].map(t=>(
          <button key={t.id} onClick={()=>setAdminTab(t.id)}
            style={{padding:'8px 18px',border:'none',cursor:'pointer',fontFamily:'inherit',fontSize:13,fontWeight:600,
              background:adminTab===t.id?T.text:'transparent',
              color:adminTab===t.id?'#F8F6F2':T.muted,transition:'all .15s'}}>
            {t.label}
          </button>
        ))}
      </div>

      {adminTab==='log'&&(()=>{
        const [undoTarget,setUndoTarget]=useState(null);
        const [logSearch,setLogSearch]=useState('');
        const [logUser,setLogUser]=useState('All');

        const ACTION_ICON={
          CREATE_PROJECT:'📁',EDIT_PROJECT:'✏️',CLOSE_PROJECT:'✅',REOPEN_PROJECT:'🔄',
          CREATE_INVOICE:'🧾',PAY_INVOICE:'💳',
          CREATE_PAYMENT:'💰',
          CREATE_WORKER:'👷',EDIT_WORKER:'✏️',
          DELETE_PROJECT:'🗑️',DELETE_INVOICE:'🗑️',DELETE_PAYMENT:'🗑️',DELETE_USER:'🗑️',DELETE_STAFFCLAIM:'🗑️',
          RESTORE_PROJECT:'↩️',RESTORE_INVOICE:'↩️',RESTORE_PAYMENT:'↩️',RESTORE_USER:'↩️',
          PERMANENT_DELETE:'⛔',
        };
        const ACTION_COLOR={
          CREATE_PROJECT:T.success,EDIT_PROJECT:T.info,CLOSE_PROJECT:T.success,REOPEN_PROJECT:T.tan,
          CREATE_INVOICE:'#8A6A3A',PAY_INVOICE:T.success,CREATE_PAYMENT:T.success,
          CREATE_WORKER:T.info,EDIT_WORKER:T.info,
          DELETE_PROJECT:T.danger,DELETE_INVOICE:T.danger,DELETE_PAYMENT:T.danger,DELETE_USER:T.danger,DELETE_STAFFCLAIM:T.danger,
          RESTORE_PROJECT:T.tan,RESTORE_INVOICE:T.tan,RESTORE_PAYMENT:T.tan,RESTORE_USER:T.tan,
          PERMANENT_DELETE:'#6d28d9',
        };
        const canUndo=entry=>entry.snapshot&&entry.action.startsWith('DELETE_')&&!entry.action.includes('PERMANENT');

        // Filter and sort
        const uniqueUsers=['All',...[...new Set(actionLog.map(e=>e.userName))].sort()];
        const filtered=actionLog
          .filter(e=>{
            const matchUser=logUser==='All'||e.userName===logUser;
            const matchSearch=!logSearch||(e.detail+e.userName+e.action).toLowerCase().includes(logSearch.toLowerCase());
            return matchUser&&matchSearch;
          })
          .sort((a,b)=>new Date(b.at)-new Date(a.at));

        // Group by user for the grouped view
        const byUser={};
        filtered.forEach(e=>{
          if(!byUser[e.userName]) byUser[e.userName]={userName:e.userName,userRole:e.userRole,entries:[]};
          byUser[e.userName].entries.push(e);
        });
        const userGroups=Object.values(byUser).sort((a,b)=>new Date(b.entries[0].at)-new Date(a.entries[0].at));

        return (
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            {/* Toolbar */}
            <div style={{display:'flex',gap:10,flexWrap:'wrap'}}>
              <div style={{position:'relative',flex:1,minWidth:180}}>
                <Search size={12} style={{position:'absolute',left:10,top:'50%',transform:'translateY(-50%)',color:T.dim}}/>
                <input value={logSearch} onChange={e=>setLogSearch(e.target.value)} placeholder="Search actions…"
                  style={{...iStyle,paddingLeft:30,fontSize:13}}/>
              </div>
              <select value={logUser} onChange={e=>setLogUser(e.target.value)}
                style={{...iStyle,width:'auto',fontSize:13,cursor:'pointer'}}>
                {uniqueUsers.map(u=><option key={u} value={u}>{u}</option>)}
              </select>
              <div style={{fontSize:12,color:T.dim,display:'flex',alignItems:'center'}}>
                {filtered.length} action{filtered.length!==1?'s':''}
              </div>
            </div>

            {userGroups.length===0&&(
              <div style={{textAlign:'center',padding:'48px 0',color:T.muted,fontSize:13}}>
                No actions recorded yet. Actions are automatically tracked from here on.
              </div>
            )}

            {/* User groups */}
            {userGroups.map(group=>(
              <div key={group.userName} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:16,overflow:'hidden',boxShadow:T.shadow}}>
                {/* User header */}
                <div style={{padding:'12px 18px',background:T.bg,borderBottom:`1px solid ${T.borderLight}`,display:'flex',alignItems:'center',justifyContent:'space-between'}}>
                  <div style={{display:'flex',alignItems:'center',gap:10}}>
                    <div style={{width:30,height:30,borderRadius:10,background:T.text,display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0}}>
                      <span style={{fontSize:12,color:'#F8F6F2',fontWeight:700}}>{group.userName.charAt(0).toUpperCase()}</span>
                    </div>
                    <div>
                      <div style={{fontSize:13,fontWeight:700,color:T.text}}>{group.userName}</div>
                      <div style={{fontSize:10,color:T.dim}}>{ROLE_LABEL[group.userRole]||group.userRole} · {group.entries.length} action{group.entries.length!==1?'s':''}</div>
                    </div>
                  </div>
                  <div style={{fontSize:11,color:T.dim}}>
                    Last: {(()=>{const d=new Date(group.entries[0].at);const age=Math.floor((Date.now()-d)/60000);return age<2?'Just now':age<60?`${age}m ago`:age<1440?`${Math.floor(age/60)}h ago`:d.toLocaleDateString('en-SG',{day:'numeric',month:'short'});})()}
                  </div>
                </div>
                {/* Action rows */}
                <div>
                  {group.entries.map((entry,idx)=>{
                    const color=ACTION_COLOR[entry.action]||T.muted;
                    const icon=ACTION_ICON[entry.action]||'•';
                    const d=new Date(entry.at);
                    const timeStr=d.toLocaleString('en-SG',{day:'numeric',month:'short',year:'2-digit',hour:'2-digit',minute:'2-digit',hour12:false});
                    return (
                      <div key={entry.id} style={{display:'flex',alignItems:'flex-start',gap:12,padding:'10px 18px',
                        borderTop:idx===0?'none':`1px solid ${T.borderLight}`,
                        background:'transparent'}}>
                        <div style={{width:24,height:24,borderRadius:7,background:color+'14',display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0,marginTop:1,fontSize:13}}>
                          {icon}
                        </div>
                        <div style={{flex:1,minWidth:0}}>
                          <div style={{fontSize:12,fontWeight:600,color:T.text,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{entry.detail}</div>
                          <div style={{fontSize:10,color:T.dim,marginTop:2,display:'flex',gap:8,flexWrap:'wrap'}}>
                            <span style={{fontFamily:'monospace',background:T.bg,padding:'1px 5px',borderRadius:4,fontSize:9}}>{entry.action}</span>
                            <span>{timeStr}</span>
                          </div>
                        </div>
                        {isSuperAdmin&&canUndo(entry)&&(
                          <button onClick={()=>setUndoTarget(entry)}
                            style={{background:T.accentLight,border:`1px solid ${T.borderLight}`,borderRadius:7,
                              padding:'3px 10px',cursor:'pointer',fontSize:11,fontWeight:600,
                              color:T.text,fontFamily:'inherit',display:'flex',alignItems:'center',gap:4,flexShrink:0}}>
                            <RotateCcw size={9}/>Undo
                          </button>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}

            {/* Undo confirmation modal */}
            {undoTarget&&(
              <Modal title="Confirm Undo" onClose={()=>setUndoTarget(null)}>
                <div style={{display:'flex',flexDirection:'column',gap:16}}>
                  <div style={{background:T.bg,borderRadius:12,padding:'14px 16px'}}>
                    <div style={{fontSize:11,color:T.muted,marginBottom:4}}>Action to undo:</div>
                    <div style={{fontSize:14,fontWeight:600,color:T.text}}>{undoTarget.detail}</div>
                    <div style={{fontSize:11,color:T.dim,marginTop:4}}>
                      By {undoTarget.userName} · {new Date(undoTarget.at).toLocaleString('en-SG',{day:'numeric',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit',hour12:false})}
                    </div>
                  </div>
                  <div style={{fontSize:13,color:T.muted,lineHeight:1.6}}>
                    This will restore the deleted item back to its original location.
                  </div>
                  <div style={{display:'flex',justifyContent:'flex-end',gap:10}}>
                    <Btn variant="secondary" onClick={()=>setUndoTarget(null)}>Cancel</Btn>
                    <Btn onClick={()=>{if(undoTarget.snapshot)onUndoAction(undoTarget.snapshot);setUndoTarget(null);}}>
                      <RotateCcw size={13}/>Confirm Undo
                    </Btn>
                  </div>
                </div>
              </Modal>
            )}
          </div>
        );
      })()}

      {adminTab==='users'&&(<>
      {/* Stats row */}
      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(150px,1fr))',gap:12}}>
        {[
          {l:'Total Users',v:users.length,c:T.accent},
          {l:'Active',v:users.filter(u=>u.active).length,c:T.success},
          {l:'Inactive',v:users.filter(u=>!u.active).length,c:T.dim},
          {l:'Admins',v:users.filter(u=>u.role==='admin').length,c:T.danger},
        ].map(({l,v,c})=>(
          <div key={l} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:12,padding:'14px 18px'}}>
            <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.07em',marginBottom:6}}>{l}</div>
            <div style={{fontSize:22,fontWeight:700,color:c}}>{v}</div>
          </div>
        ))}
      </div>

      {/* Toolbar */}
      <div style={{display:'flex',gap:10,flexWrap:'wrap'}}>
        <div style={{flex:1,minWidth:180,position:'relative'}}>
          <Search size={13} style={{position:'absolute',left:11,top:'50%',transform:'translateY(-50%)',color:T.dim}}/>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search users..."
            style={{...iStyle,paddingLeft:33}}/>
        </div>
        <Btn onClick={openNew}>
          <UserPlus size={13}/>Create User
        </Btn>
      </div>

      {/* User table */}
      <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,overflow:'hidden',boxShadow:T.shadow}}>
        <table style={{width:'100%',fontSize:13,borderCollapse:'collapse'}}>
          <thead>
            <tr style={{color:T.dim,fontSize:11,fontWeight:700,textTransform:'uppercase',letterSpacing:'0.06em',
              background:T.bg}}>
              {['User','Role','Last Login','Assigned Projects','Status','Actions'].map(h=>(
                <th key={h} style={{textAlign:'left',padding:'11px 16px',whiteSpace:'nowrap'}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map(u=>(
              <tr key={u.id} style={{borderTop:`1px solid ${T.borderLight}`,opacity:u.active?1:0.45}}>
                <td style={{padding:'12px 16px'}}>
                  <div style={{display:'flex',alignItems:'center',gap:10}}>
                    <Avatar
                      photo={u.photo}
                      name={u.name}
                      size={36}
                      color={ROLE_CLR[u.role]}
                    />
                    <div>
                      <div style={{fontSize:13,fontWeight:700,color:T.text}}>{u.name}</div>
                      <div style={{fontSize:11,color:T.dim}}>{u.email}</div>
                    </div>
                  </div>
                </td>
                <td style={{padding:'12px 16px'}}>
                  <Badge color={ROLE_CLR[u.role]}>{ROLE_LABEL[u.role]}</Badge>
                </td>
                <td style={{padding:'12px 16px',minWidth:110}}>
                  {u.lastLoginAt?(()=>{
                    const d=new Date(u.lastLoginAt);
                    const now=new Date();
                    const diffMs=now-d;
                    const diffMins=Math.floor(diffMs/60000);
                    const diffHrs=Math.floor(diffMs/3600000);
                    const diffDays=Math.floor(diffMs/86400000);
                    let when='';
                    if(diffMins<2) when='Just now';
                    else if(diffMins<60) when=`${diffMins}m ago`;
                    else if(diffHrs<24) when=`${diffHrs}h ago`;
                    else if(diffDays<7) when=`${diffDays}d ago`;
                    else when=d.toLocaleDateString('en-SG',{day:'numeric',month:'short',year:'2-digit'});
                    return (
                      <div>
                        <div style={{fontSize:12,fontWeight:500,color:T.text}}>{when}</div>
                        <div style={{fontSize:10,color:T.dim,marginTop:1}}>
                          {d.toLocaleDateString('en-SG',{day:'numeric',month:'short'})} {d.toLocaleTimeString('en-SG',{hour:'2-digit',minute:'2-digit'})}
                        </div>
                      </div>
                    );
                  })():(
                    <span style={{fontSize:11,color:T.dim,fontStyle:'italic'}}>Never</span>
                  )}
                </td>
                <td style={{padding:'12px 16px'}}>
                  {u.role==='admin'||u.role==='accounts'?(
                    <span style={{fontSize:12,color:T.dim,fontStyle:'italic'}}>All projects</span>
                  ):(
                    <div style={{display:'flex',flexWrap:'wrap',gap:4}}>
                      {u.assignedProjects.length===0?(
                        <span style={{fontSize:12,color:T.dim,fontStyle:'italic'}}>None assigned</span>
                      ):u.assignedProjects.map(pid=>{
                        const p=projects.find(x=>x.id===pid);
                        return p?<Badge key={pid} color={T.info} sm>{p.name.split(' ').slice(0,2).join(' ')}</Badge>:null;
                      })}
                    </div>
                  )}
                </td>
                <td style={{padding:'12px 16px'}}>
                  <div style={{display:'flex',alignItems:'center',gap:8}}>
                    <Toggle on={u.active} onClick={()=>toggleActive(u.id)}/>
                    <span style={{fontSize:12,color:u.active?T.success:T.dim}}>{u.active?'Active':'Inactive'}</span>
                  </div>
                </td>
                <td style={{padding:'12px 16px'}}>
                  <div style={{display:'flex',gap:4}}>
                    <button title="Edit user" onClick={()=>openEdit(u)}
                      style={{background:T.bg,border:'none',cursor:'pointer',color:T.muted,
                        display:'flex',padding:'5px 7px',borderRadius:6,alignItems:'center',gap:4,fontSize:12}}>
                      <Edit3 size={12}/>Edit
                    </button>
                    <button title="Permissions & visibility" onClick={()=>setPermModal(u.id)}
                      style={{background:T.infoLight,border:'none',cursor:'pointer',color:T.info,
                        display:'flex',padding:'5px 7px',borderRadius:6,alignItems:'center',gap:4,fontSize:12}}>
                      <Shield size={12}/>Perms
                    </button>
                    {u.role!=='admin'&&(
                      <button title="Delete user" onClick={()=>delUser(u.id)}
                        style={{background:T.dangerLight,border:'none',cursor:'pointer',color:T.danger,
                          display:'flex',padding:'5px 7px',borderRadius:6,alignItems:'center',gap:4,fontSize:12}}>
                        <Trash2 size={12}/>
                      </button>
                    )}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {filtered.length===0&&<div style={{color:T.dim,fontSize:13,textAlign:'center',padding:36}}>No users found</div>}
      </div>

      {/* Create / Edit user modal */}
      {(modal==='new'||editingUser)&&(
        <Modal title={modal==='new'?'Create New User':'Edit User'} onClose={closeModal}>
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            {/* Profile photo */}
            <div style={{display:'flex',alignItems:'center',gap:16,padding:'14px 16px',background:T.bg,borderRadius:12,border:`1px solid ${T.borderLight}`}}>
              <Avatar
                photo={form.photo}
                name={form.name||'?'}
                size={68}
                color={ROLE_CLR[form.role]||T.accent}
                editable
                onUpload={v=>ff('photo')(v)}
              />
              <div style={{flex:1}}>
                <div style={{fontSize:13,fontWeight:600,color:T.text,marginBottom:4}}>{form.name||'New User'}</div>
                {form.photo
                  ? <button type="button" onClick={()=>ff('photo')('')}
                      style={{background:'none',border:'none',cursor:'pointer',fontSize:12,color:T.danger,padding:0,fontFamily:'inherit',fontWeight:500}}>
                      Remove photo
                    </button>
                  : <label style={{fontSize:12,color:T.muted}}>
                      Tap the circle or camera icon to upload a photo
                    </label>
                }
              </div>
            </div>

            <Field label="Full Name *" value={form.name} onChange={ff('name')} placeholder="e.g. Priya Sharma"/>
            <Field label="Email Address *" type="email" value={form.email} onChange={ff('email')} placeholder="priya@company.sg"/>
            <div>
              <label style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.07em',display:'block',marginBottom:5}}>Role *</label>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:8}}>
                {ROLES.map(r=>(
                  <button type="button" key={r} onClick={()=>onRoleChange(r)}
                    style={{padding:'8px 12px',borderRadius:10,border:`1px solid ${form.role===r?ROLE_CLR[r]:'rgba(255,255,255,0.08)'}`,
                      background:form.role===r?`${ROLE_CLR[r]}18`:'rgba(255,255,255,0.03)',
                      cursor:'pointer',fontFamily:'inherit',fontSize:12,fontWeight:600,
                      color:form.role===r?ROLE_CLR[r]:T.muted,textAlign:'left',transition:'all .15s'}}>
                    {ROLE_LABEL[r]}
                  </button>
                ))}
              </div>
              <div style={{marginTop:8,fontSize:11,color:T.dim,background:T.bg,borderRadius:10,padding:'8px 12px'}}>
                {form.role==='admin'&&'Full access to all modules including Admin panel.'}
                {form.role==='accounts'&&'Access to invoices, payments, and reports. No project creation.'}
                {form.role==='designer'&&'Can enter expenses for assigned projects only. Dashboard summary visible.'}
                {form.role==='pm'&&'Can manage assigned projects, enter expenses, and view payments.'}
                {form.role==='expense_entry'&&'Can only capture invoices/expenses for their assigned projects. No financial visibility.'}
              </div>
            </div>

            {(form.role==='designer'||form.role==='pm'||form.role==='expense_entry')&&(
              <div>
                <label style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.07em',display:'block',marginBottom:8}}>Assigned Projects</label>
                <div style={{display:'flex',flexDirection:'column',gap:6,maxHeight:150,overflowY:'auto'}}>
                  {projects.map(p=>{
                    const has=form.assignedProjects.includes(p.id);
                    return (
                      <label key={p.id} style={{display:'flex',alignItems:'center',gap:10,cursor:'pointer',
                        padding:'8px 12px',background:has?'rgba(59,130,246,0.08)':'rgba(255,255,255,0.03)',
                        borderRadius:10,border:`1px solid ${has?T.info+'40':'transparent'}`}}>
                        <input type="checkbox" checked={has}
                          onChange={()=>setForm(prev=>{
                            const a=prev.assignedProjects;
                            return {...prev,assignedProjects:has?a.filter(x=>x!==p.id):[...a,p.id]};
                          })}
                          style={{accentColor:T.info}}/>
                        <div>
                          <div style={{fontSize:13,fontWeight:600,color:T.text}}>{p.name}</div>
                          <div style={{fontSize:11,color:T.dim}}>{p.client}</div>
                        </div>
                      </label>
                    );
                  })}
                </div>
              </div>
            )}

            <div>
              <label style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.07em',display:'block',marginBottom:5}}>
                {modal==='new'?'Set Password':'Reset Password (leave blank to keep current)'}
              </label>
              <div style={{position:'relative'}}>
                <input
                  type={showPw?'text':'password'}
                  value={pwInput}
                  onChange={e=>setPwInput(e.target.value)}
                  placeholder={modal==='new'?'Enter password':'New password (optional)'}
                  autoComplete="new-password"
                  spellCheck="false"
                  style={{...iStyle,paddingRight:44}}/>
                <button type="button"
                  onClick={()=>setShowPw(s=>!s)}
                  style={{position:'absolute',right:10,top:'50%',transform:'translateY(-50%)',
                    background:'none',border:'none',cursor:'pointer',color:T.muted,
                    display:'flex',alignItems:'center',justifyContent:'center',
                    width:32,height:32,borderRadius:8}}>
                  {showPw?<EyeOff size={16}/>:<Eye size={16}/>}
                </button>
              </div>
              {pwInput.length>0&&(
                <div style={{marginTop:6,display:'flex',alignItems:'center',gap:6}}>
                  <div style={{height:4,flex:1,borderRadius:2,background:T.borderLight,overflow:'hidden'}}>
                    <div style={{height:'100%',borderRadius:2,
                      width:pwInput.length<6?'33%':pwInput.length<10?'66%':'100%',
                      background:pwInput.length<6?T.danger:pwInput.length<10?T.warning:T.success,
                      transition:'width 0.2s,background 0.2s'}}/>
                  </div>
                  <span style={{fontSize:10,color:pwInput.length<6?T.danger:pwInput.length<10?T.warning:T.success,fontWeight:600,flexShrink:0}}>
                    {pwInput.length<6?'Too short':pwInput.length<10?'OK':'Strong'}
                  </span>
                </div>
              )}
            </div>
          </div>
          <div style={{display:'flex',justifyContent:'flex-end',gap:10,marginTop:22}}>
            <Btn variant="secondary" onClick={closeModal}>Cancel</Btn>
            <Btn onClick={saveUser} disabled={!form.name||!form.email}>{modal==='new'?'Create User':'Save Changes'}</Btn>
          </div>
        </Modal>
      )}

      {/* Permissions & Visibility modal */}
      {permUser&&(
        <Modal title={`Permissions — ${permUser.name}`} onClose={()=>setPermModal(null)} wide>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:24}}>
            {/* Nav tabs */}
            <div>
              <div style={{fontSize:12,fontWeight:700,color:T.text,marginBottom:4}}>Navigation Access</div>
              <div style={{fontSize:11,color:T.dim,marginBottom:12}}>Which pages this user can see in the sidebar</div>
              <div style={{display:'flex',flexDirection:'column',gap:8}}>
                {ALL_TABS.map(({id,label})=>{
                  const on=permUser.tabs.includes(id);
                  const locked=permUser.role==='admin';
                  return (
                    <div key={id} style={{display:'flex',alignItems:'center',justifyContent:'space-between',
                      padding:'10px 14px',background:on?'rgba(16,185,129,0.07)':'rgba(255,255,255,0.025)',
                      borderRadius:12,border:`1px solid ${on?T.success+'30':T.border}`}}>
                      <div>
                        <span style={{fontSize:13,fontWeight:600,color:on?T.text:T.muted}}>{label}</span>
                        {locked&&<span style={{fontSize:10,color:T.dim,marginLeft:6}}>(locked)</span>}
                      </div>
                      {locked?(
                        <Lock size={14} style={{color:T.dim}}/>
                      ):(
                        <Toggle on={on} onClick={()=>toggleTab(permUser.id,id)}/>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>

            <div style={{display:'flex',flexDirection:'column',gap:20}}>
              {/* Dashboard widgets */}
              <div>
                <div style={{fontSize:12,fontWeight:700,color:T.text,marginBottom:4}}>Dashboard Widgets</div>
                <div style={{fontSize:11,color:T.dim,marginBottom:12}}>What this user sees on their dashboard</div>
                <div style={{display:'flex',flexDirection:'column',gap:8}}>
                  {DASH_WIDGETS.map(({id,label})=>{
                    const on=permUser.widgets.includes(id);
                    return (
                      <div key={id} style={{display:'flex',alignItems:'center',justifyContent:'space-between',
                        padding:'10px 14px',background:on?'rgba(59,130,246,0.07)':'rgba(255,255,255,0.025)',
                        borderRadius:12,border:`1px solid ${on?T.info+'30':T.border}`}}>
                        <span style={{fontSize:13,fontWeight:600,color:on?T.text:T.muted}}>{label}</span>
                        <Toggle on={on} onClick={()=>toggleWidget(permUser.id,id)}/>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Project access */}
              {(permUser.role==='designer'||permUser.role==='pm'||permUser.role==='expense_entry')&&(
                <div>
                  <div style={{fontSize:12,fontWeight:700,color:T.text,marginBottom:4}}>Project Access</div>
                  <div style={{fontSize:11,color:T.dim,marginBottom:12}}>Projects this user can submit expenses for</div>
                  <div style={{display:'flex',flexDirection:'column',gap:8}}>
                    {projects.map(p=>{
                      const on=permUser.assignedProjects.includes(p.id);
                      return (
                        <div key={p.id} style={{display:'flex',alignItems:'center',justifyContent:'space-between',
                          padding:'10px 14px',background:on?'rgba(233,170,38,0.07)':'rgba(255,255,255,0.025)',
                          borderRadius:12,border:`1px solid ${on?T.accent+'30':T.border}`}}>
                          <div>
                            <div style={{fontSize:13,fontWeight:600,color:on?T.text:T.muted}}>{p.name}</div>
                            <div style={{fontSize:11,color:T.dim}}>{p.client}</div>
                          </div>
                          <Toggle on={on} onClick={()=>toggleProject(permUser.id,p.id)}/>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          </div>
          <div style={{display:'flex',justifyContent:'flex-end',gap:10,marginTop:22}}>
            <Btn variant="secondary" onClick={()=>setPermModal(null)}>Close</Btn>
            <Btn onClick={savePerms}>Save Permissions</Btn>
          </div>
        </Modal>
      )}

      {deleteTarget&&(
        <ConfirmDelete
          matchValue={deleteTarget.email}
          typeLabel="user"
          impact="The user will lose access immediately. Their data and records will remain intact in Trash."
          onConfirm={()=>confirmDelUser(deleteTarget)}
          onClose={()=>setDeleteTarget(null)}
        />
      )}
      </>)}
    </div>
  );
}

function Contacts({projects,invoices,payments}){
  const [view,setView]=useState('clients');
  const [search,setSearch]=useState('');
  const [yearFilt,setYearFilt]=useState('All');
  const [typeFilt,setTypeFilt]=useState('All');
  const [lightbox,setLightbox]=useState(null);


  const clientMap=useMemo(()=>{
    const m={};
    projects.forEach(p=>{
      const key=p.client;
      if(!m[key]) m[key]={
        name:p.client,
        email:p.clientEmail||'',
        phone:p.clientPhone||'',
        projects:[],
      };
      const year=p.startDate?new Date(p.startDate).getFullYear():'—';
      const recv=payments.filter(py=>py.projectId===p.id&&py.status==='Received').reduce((s,py)=>s+py.amount,0);
      const exp=invoices.filter(i=>i.projectId===p.id).reduce((s,i)=>s+i.total,0);
      const rev=p.contractAmount+(p.variationOrders||0);
      m[key].projects.push({
        id:p.id, name:p.name, year, type:p.projectType||'Residential',
        status:p.status, contractAmount:rev, collected:recv,
        grossProfit:rev-exp, archived:p.archived
      });
    });
    return Object.values(m).sort((a,b)=>a.name.localeCompare(b.name));
  },[projects,invoices,payments]);

  const years=useMemo(()=>{
    const s=new Set();
    projects.forEach(p=>{if(p.startDate)s.add(new Date(p.startDate).getFullYear());});
    return ['All',...Array.from(s).sort((a,b)=>b-a)];
  },[projects]);

  const filteredClients=useMemo(()=>clientMap.filter(c=>{
    const matchSearch=(c.name+c.email).toLowerCase().includes(search.toLowerCase());
    const matchYear=yearFilt==='All'||c.projects.some(p=>String(p.year)===String(yearFilt));
    const matchType=typeFilt==='All'||c.projects.some(p=>p.type===typeFilt);
    return matchSearch&&matchYear&&matchType;
  }),[clientMap,search,yearFilt,typeFilt]);


  const supplierMap=useMemo(()=>{
    const m={};
    invoices.forEach(inv=>{
      const key=inv.supplier;
      if(!m[key]) m[key]={name:inv.supplier,categories:new Set(),invoices:[],yearlyPaid:{}};
      m[key].categories.add(inv.category);
      const year=inv.invoiceDate?new Date(inv.invoiceDate).getFullYear():'—';
      if(inv.status==='Paid'){
        m[key].yearlyPaid[year]=(m[key].yearlyPaid[year]||0)+inv.total;
      }
      const proj=projects.find(p=>p.id===inv.projectId);
      m[key].invoices.push({
        ...inv,
        projectName:proj?.name||'—',
        year,
      });
    });
    return Object.values(m)
      .map(s=>({...s,categories:Array.from(s.categories),
        totalPaid:s.invoices.filter(i=>i.status==='Paid').reduce((a,i)=>a+i.total,0),
        totalInvoices:s.invoices.length,
      }))
      .sort((a,b)=>b.totalPaid-a.totalPaid);
  },[invoices,projects]);

  const filteredSuppliers=useMemo(()=>supplierMap.filter(s=>
    s.name.toLowerCase().includes(search.toLowerCase())||
    s.categories.some(c=>c.toLowerCase().includes(search.toLowerCase()))
  ),[supplierMap,search]);


  const invYears=useMemo(()=>{
    const s=new Set();
    invoices.forEach(i=>{if(i.invoiceDate)s.add(new Date(i.invoiceDate).getFullYear());});
    return Array.from(s).sort((a,b)=>b-a);
  },[invoices]);

  return (
    <div style={{display:'flex',flexDirection:'column',gap:16}}>
      {/* Tab switcher */}
      <div style={{display:'flex',gap:10,flexWrap:'wrap',alignItems:'center'}}>
        <div style={{display:'flex',background:T.bg,borderRadius:12,
          border:`1px solid ${T.borderLight}`,overflow:'hidden'}}>
          {[
            {v:'clients',l:'Clients',Icon:Home},
            {v:'suppliers',l:'Suppliers & Subcons',Icon:Building2},
          ].map(({v,l,Icon})=>(
            <button key={v} onClick={()=>{setView(v);setSearch('');setYearFilt('All');setTypeFilt('All');}}
              style={{padding:'8px 18px',border:'none',cursor:'pointer',fontFamily:'inherit',
                fontSize:13,fontWeight:600,display:'flex',alignItems:'center',gap:7,
                background:view===v?T.text:'transparent',
                color:view===v?'#F8F6F2':T.muted,transition:'all .15s'}}>
              <Icon size={13}/>{l}
            </button>
          ))}
        </div>

        <div style={{flex:1,minWidth:180,position:'relative'}}>
          <Search size={13} style={{position:'absolute',left:11,top:'50%',transform:'translateY(-50%)',color:T.dim}}/>
          <input value={search} onChange={e=>setSearch(e.target.value)}
            placeholder={view==='clients'?'Search clients…':'Search suppliers or categories…'}
            style={{...iStyle,paddingLeft:33}}/>
        </div>

        {view==='clients'&&<>
          <select value={yearFilt} onChange={e=>setYearFilt(e.target.value)} style={{...iStyle,width:'auto',minWidth:100}}>
            {years.map(y=><option key={y} value={y} style={{}}>{y}</option>)}
          </select>
          <select value={typeFilt} onChange={e=>setTypeFilt(e.target.value)} style={{...iStyle,width:'auto',minWidth:120}}>
            {['All',...PROJ_TYPES].map(t=><option key={t} value={t} style={{}}>{t}</option>)}
          </select>
        </>}
      </div>

      {/* -- CLIENT LIST ------------------------------------------------- */}
      {view==='clients'&&(
        <div style={{display:'flex',flexDirection:'column',gap:12}}>
          {filteredClients.length===0&&(
            <div style={{color:T.dim,fontSize:13,textAlign:'center',padding:40}}>No clients found</div>
          )}
          {filteredClients.map(c=>{
            const totalRev=c.projects.reduce((s,p)=>s+p.contractAmount,0);
            const totalCollected=c.projects.reduce((s,p)=>s+p.collected,0);
            const totalProfit=c.projects.reduce((s,p)=>s+p.grossProfit,0);
            const shownProjects=c.projects.filter(p=>
              (yearFilt==='All'||String(p.year)===String(yearFilt))&&
              (typeFilt==='All'||p.type===typeFilt)
            );
            return (
              <div key={c.name} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow}}>
                <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:14,flexWrap:'wrap',gap:10}}>
                  <div style={{display:'flex',alignItems:'center',gap:12}}>
                    <div style={{width:42,height:42,borderRadius:'50%',background:'rgba(233,170,38,0.12)',
                      display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0,
                      fontSize:16,fontWeight:700,color:T.accent}}>
                      {c.name.charAt(0)}
                    </div>
                    <div>
                      <div style={{fontSize:14,fontWeight:700,color:T.text}}>{c.name}</div>
                      <div style={{display:'flex',gap:12,marginTop:3,flexWrap:'wrap'}}>
                        {c.email
                          ? <span style={{fontSize:12,color:T.dim,display:'flex',alignItems:'center',gap:4}}><Mail size={11}/>{c.email}</span>
                          : <span style={{fontSize:12,color:T.dim,fontStyle:'italic'}}>No email recorded</span>}
                        {c.phone&&<span style={{fontSize:12,color:T.dim,display:'flex',alignItems:'center',gap:4}}><Phone size={11}/>{c.phone}</span>}
                      </div>
                    </div>
                  </div>
                  <div style={{display:'flex',gap:10,flexWrap:'wrap'}}>
                    {[
                      {l:'Total Revenue',v:fmtSGD(totalRev),c:T.text},
                      {l:'Collected',v:fmtSGD(totalCollected),c:T.success},
                      {l:'Gross Profit',v:fmtSGD(totalProfit),c:totalProfit>0?T.accent:T.danger},
                    ].map(({l,v,c:col})=>(
                      <div key={l} style={{background:T.bg,borderRadius:12,padding:'8px 14px',textAlign:'right'}}>
                        <div style={{fontSize:10,color:T.dim,textTransform:'uppercase',letterSpacing:'0.06em',marginBottom:2}}>{l}</div>
                        <div style={{fontSize:14,fontWeight:700,color:col}}>{v}</div>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Projects list */}
                <div style={{display:'flex',flexDirection:'column',gap:7}}>
                  {shownProjects.map(p=>(
                    <div key={p.id} style={{display:'flex',alignItems:'center',gap:10,flexWrap:'wrap',
                      background:T.bg,borderRadius:12,padding:'9px 14px'}}>
                      <div style={{display:'flex',alignItems:'center',gap:6,flex:1,minWidth:160}}>
                        {p.type==='Residential'?<Home size={12} style={{color:T.muted,flexShrink:0}}/>:<Building2 size={12} style={{color:T.info,flexShrink:0}}/>}
                        <span style={{fontSize:13,fontWeight:600,color:T.text}}>{p.name}</span>
                        {p.archived&&<Badge color={T.dim} sm>Archived</Badge>}
                      </div>
                      <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
                        <Badge color={p.type==='Residential'?T.accent:T.info} sm>{p.type}</Badge>
                        <Badge color={ST_CLR[p.status]||T.dim} sm>{p.status}</Badge>
                        <span style={{fontSize:12,color:T.dim}}>{p.year}</span>
                        <span style={{fontSize:12,fontWeight:700,color:T.text}}>{fmtSGD(p.contractAmount)}</span>
                      </div>
                    </div>
                  ))}
                  {shownProjects.length===0&&(
                    <div style={{fontSize:12,color:T.dim,padding:'6px 0',fontStyle:'italic'}}>No projects match current filters</div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* -- SUPPLIER LIST ----------------------------------------------- */}
      {view==='suppliers'&&(
        <div style={{display:'flex',flexDirection:'column',gap:12}}>
          {/* Year summary header */}
          {invYears.length>0&&(
            <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow}}>
              <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:14}}>Annual Supplier Payments</div>
              <div style={{display:'flex',gap:12,flexWrap:'wrap'}}>
                {invYears.map(yr=>{
                  const ytdPaid=invoices.filter(i=>i.status==='Paid'&&i.invoiceDate&&new Date(i.invoiceDate).getFullYear()===yr).reduce((s,i)=>s+i.total,0);
                  const ytdCount=invoices.filter(i=>i.status==='Paid'&&i.invoiceDate&&new Date(i.invoiceDate).getFullYear()===yr).length;
                  return (
                    <div key={yr} style={{background:T.bg,borderRadius:12,padding:'12px 18px',minWidth:140}}>
                      <div style={{fontSize:11,color:T.dim,fontWeight:700,marginBottom:4}}>{yr}</div>
                      <div style={{fontSize:18,fontWeight:700,color:T.accent}}>{fmtSGD(ytdPaid)}</div>
                      <div style={{fontSize:11,color:T.dim,marginTop:2}}>{ytdCount} invoices paid</div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {filteredSuppliers.length===0&&(
            <div style={{color:T.dim,fontSize:13,textAlign:'center',padding:40}}>No suppliers found</div>
          )}
          {filteredSuppliers.map(s=>(
            <div key={s.name} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow}}>
              <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:14,flexWrap:'wrap',gap:10}}>
                <div style={{display:'flex',alignItems:'center',gap:12}}>
                  <div style={{width:42,height:42,borderRadius:12,background:T.infoLight,
                    display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0,
                    fontSize:16,fontWeight:700,color:T.info}}>
                    {s.name.charAt(0)}
                  </div>
                  <div>
                    <div style={{fontSize:14,fontWeight:700,color:T.text}}>{s.name}</div>
                    <div style={{display:'flex',gap:5,marginTop:4,flexWrap:'wrap'}}>
                      {s.categories.map(cat=>(
                        <Badge key={cat} color={CAT_CLR[cat]||T.dim} sm>{cat}</Badge>
                      ))}
                    </div>
                  </div>
                </div>
                <div style={{display:'flex',gap:10,flexWrap:'wrap'}}>
                  {[
                    {l:'Total Invoices',v:s.totalInvoices,c:T.text},
                    {l:'Total Paid',v:fmtSGD(s.totalPaid),c:T.success},
                  ].map(({l,v,c})=>(
                    <div key={l} style={{background:T.bg,borderRadius:12,padding:'8px 14px',textAlign:'right'}}>
                      <div style={{fontSize:10,color:T.dim,textTransform:'uppercase',letterSpacing:'0.06em',marginBottom:2}}>{l}</div>
                      <div style={{fontSize:14,fontWeight:700,color:c}}>{v}</div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Yearly paid breakdown */}
              {Object.keys(s.yearlyPaid).length>0&&(
                <div style={{display:'flex',gap:8,marginBottom:12,flexWrap:'wrap'}}>
                  {Object.entries(s.yearlyPaid).sort((a,b)=>b[0]-a[0]).map(([yr,amt])=>(
                    <div key={yr} style={{background:T.successLight,border:'1px solid rgba(16,185,129,0.2)',
                      borderRadius:7,padding:'5px 12px',display:'flex',gap:8,alignItems:'center'}}>
                      <span style={{fontSize:11,color:T.success,fontWeight:700}}>{yr}</span>
                      <span style={{fontSize:12,color:T.text,fontWeight:700}}>{fmtSGD(amt)}</span>
                    </div>
                  ))}
                </div>
              )}

              {/* Invoice list */}
              <div style={{overflowX:'auto'}}>
                <table style={{width:'100%',fontSize:12,borderCollapse:'collapse'}}>
                  <thead>
                    <tr style={{color:T.dim,fontSize:11,fontWeight:700,textTransform:'uppercase',letterSpacing:'0.05em'}}>
                      {['Invoice #','Project','Date','Category','Total','Status','Invoice','Payment'].map(h=>(
                        <th key={h} style={{textAlign:'left',paddingBottom:8,paddingRight:12,whiteSpace:'nowrap'}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {s.invoices.map(inv=>(
                      <tr key={inv.id} style={{borderTop:`1px solid ${T.borderLight}`}}>
                        <td style={{padding:'7px 12px 7px 0',color:T.dim,fontFamily:'monospace',fontSize:11}}>{inv.invoiceNo}</td>
                        <td style={{padding:'7px 12px 7px 0',color:T.muted,maxWidth:130,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{inv.projectName}</td>
                        <td style={{padding:'7px 12px 7px 0',color:T.dim,whiteSpace:'nowrap'}}>{fmtDate(inv.invoiceDate)}</td>
                        <td style={{padding:'7px 12px 7px 0'}}><Badge color={CAT_CLR[inv.category]||T.dim} sm>{inv.category}</Badge></td>
                        <td style={{padding:'7px 12px 7px 0',color:T.text,fontWeight:700}}>{fmtSGD(inv.total)}</td>
                        <td style={{padding:'7px 12px 7px 0'}}><Badge color={ST_CLR[inv.status]} sm>{inv.status}</Badge></td>
                        <td style={{padding:'7px 12px 7px 0'}}>
                          {inv.proofImage
                            ? <button onClick={()=>setLightbox({src:inv.proofImage,title:`Invoice — ${inv.invoiceNo}`})}
                                style={{background:T.infoLight,border:'none',cursor:'pointer',color:T.info,
                                  display:'flex',padding:'3px 7px',borderRadius:5,alignItems:'center',gap:3,fontSize:11,fontFamily:'inherit'}}>
                                <ZoomIn size={10}/>View
                              </button>
                            : <span style={{fontSize:11,color:T.dim,fontStyle:'italic'}}>—</span>}
                        </td>
                        <td style={{padding:'7px 0'}}>
                          {inv.paymentProof
                            ? <button onClick={()=>setLightbox({src:inv.paymentProof,title:`Payment Proof — ${inv.invoiceNo}`})}
                                style={{background:T.successLight,border:'none',cursor:'pointer',color:T.success,
                                  display:'flex',padding:'3px 7px',borderRadius:5,alignItems:'center',gap:3,fontSize:11,fontFamily:'inherit'}}>
                                <ZoomIn size={10}/>View
                              </button>
                            : <span style={{fontSize:11,color:T.dim,fontStyle:'italic'}}>—</span>}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ))}
        </div>
      )}

      {lightbox&&<Lightbox src={lightbox.src} title={lightbox.title} onClose={()=>setLightbox(null)}/>}
    </div>
  );
}

function Warranty({warranties,setWarranties,projects,isAdmin,acctSettings}){
  const [search,setSearch]=useState('');
  const [issueModal,setIssueModal]=useState(null);
  const [claimModal,setClaimModal]=useState(null);
  const [viewClaim,setViewClaim]=useState(null);
  const [lightbox,setLightbox]=useState(null);
  const claimFileRef=useRef();
  const [claimPhotoTarget,setClaimPhotoTarget]=useState(null);
  const claimCamRef=useRef();

  const blankClaim={issue:'',attendedBy:'',cost:'',date:new Date().toISOString().slice(0,10),photos:[],notes:''};
  const [claimForm,setClaimForm]=useState(blankClaim);
  const cf=k=>v=>setClaimForm(p=>({...p,[k]:v}));

  const issueWarranty=(projectId)=>{
    const serial=genSerial();
    const now=new Date().toISOString();
    const expiry=new Date();expiry.setFullYear(expiry.getFullYear()+1);
    const w={id:uid(),projectId,serial,issuedAt:now,expiresAt:expiry.toISOString(),claims:[]};
    const upd=[...warranties,w];
    setWarranties(upd);saveS('warranties',upd);
    setIssueModal(null);
    setTimeout(()=>{
      const proj=projects.find(p=>p.id===projectId);
      if(window.confirm(`Warranty issued! Serial: ${serial}\nPrint warranty card now?`)){
        printDoc(buildWarrantyHTML(proj,serial,getCo(acctSettings)),`Warranty — ${serial}`);
      }
    },300);
  };

  const saveClaim=(warrantyId)=>{
    if(!claimForm.issue||!claimForm.attendedBy)return;
    const claim={...claimForm,id:uid(),cost:parseFloat(claimForm.cost)||0,loggedAt:new Date().toISOString()};
    const upd=warranties.map(w=>w.id===warrantyId?{...w,claims:[...(w.claims||[]),claim]}:w);
    setWarranties(upd);saveS('warranties',upd);
    setClaimModal(null);setClaimForm(blankClaim);
  };

  const addClaimPhoto=(dataUrl)=>{
    setClaimForm(p=>({...p,photos:[...(p.photos||[]),dataUrl]}));
  };

  const filtered=warranties.filter(w=>{
    const proj=projects.find(p=>p.id===w.projectId);
    return (w.serial+proj?.name+proj?.client).toLowerCase().includes(search.toLowerCase());
  });

  const now=new Date();


  const eligibleProjects=projects.filter(p=>
    p.archived&&!warranties.find(w=>w.projectId===p.id)
  );

  return (
    <div style={{display:'flex',flexDirection:'column',gap:16}}>
      {/* Search + Issue warranty */}
      <div style={{display:'flex',gap:10,flexWrap:'wrap',alignItems:'center'}}>
        <div style={{flex:1,minWidth:200,position:'relative'}}>
          <Search size={13} style={{position:'absolute',left:11,top:'50%',transform:'translateY(-50%)',color:T.dim}}/>
          <input value={search} onChange={e=>setSearch(e.target.value)}
            placeholder="Search by serial number, project or client…"
            style={{...iStyle,paddingLeft:33}}/>
        </div>
        {isAdmin&&eligibleProjects.length>0&&(
          <Btn onClick={()=>setIssueModal('pick')}><Plus size={13}/>Issue Warranty</Btn>
        )}
      </div>

      {/* Stats */}
      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(150px,1fr))',gap:12}}>
        {[
          {l:'Total Issued',v:warranties.length,c:T.accent},
          {l:'Active',v:warranties.filter(w=>new Date(w.expiresAt)>now).length,c:T.success},
          {l:'Expired',v:warranties.filter(w=>new Date(w.expiresAt)<=now).length,c:T.dim},
          {l:'Claims Logged',v:warranties.reduce((s,w)=>s+(w.claims?.length||0),0),c:T.danger},
        ].map(({l,v,c})=>(
          <div key={l} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:12,padding:'14px 18px'}}>
            <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',letterSpacing:'0.07em',marginBottom:6}}>{l}</div>
            <div style={{fontSize:22,fontWeight:700,color:c}}>{v}</div>
          </div>
        ))}
      </div>

      {/* Warranty cards */}
      {filtered.length===0&&(
        <div style={{color:T.dim,fontSize:13,textAlign:'center',padding:48,background:T.card,
          borderRadius:14,border:`1px solid ${T.borderLight}`}}>
          {warranties.length===0
            ?'No warranties issued yet. Close a project first, then issue a warranty.'
            :'No warranties match your search.'}
        </div>
      )}

      {filtered.map(w=>{
        const proj=projects.find(p=>p.id===w.projectId);
        const isActive=new Date(w.expiresAt)>now;
        const daysLeft=Math.ceil((new Date(w.expiresAt)-now)/(1000*60*60*24));
        const claimsTotal=(w.claims||[]).reduce((s,c)=>s+(c.cost||0),0);
        return (
          <div key={w.id} style={{background:T.card,border:`1px solid ${isActive?T.border:'rgba(100,116,139,0.2)'}`,
            borderRadius:14,padding:20,opacity:isActive?1:0.75}}>

            {/* Header */}
            <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:16,flexWrap:'wrap',gap:10}}>
              <div style={{display:'flex',alignItems:'center',gap:14}}>
                <div style={{background:isActive?'rgba(16,185,129,0.1)':'rgba(100,116,139,0.1)',
                  borderRadius:12,padding:'10px 14px',border:`1px solid ${isActive?T.success+'30':T.dim+'30'}`}}>
                  <div style={{fontSize:10,color:T.dim,fontWeight:700,marginBottom:3,textTransform:'uppercase',letterSpacing:'0.07em'}}>Serial No.</div>
                  <div style={{fontSize:14,fontWeight:700,color:isActive?T.success:T.dim,fontFamily:'monospace'}}>{w.serial}</div>
                </div>
                <div>
                  <div style={{fontSize:14,fontWeight:700,color:T.text}}>{proj?.name||'—'}</div>
                  <div style={{fontSize:12,color:T.muted}}>{proj?.client}</div>
                  <div style={{fontSize:11,color:T.dim,marginTop:2}}>
                    Issued {fmtDate(w.issuedAt)} . Expires {fmtDate(w.expiresAt)}
                  </div>
                </div>
              </div>
              <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
                <Badge color={isActive?T.success:T.dim}>{isActive?`${daysLeft}d left`:'Expired'}</Badge>
                <Btn variant="secondary" size="sm"
                  onClick={()=>{if(proj)printDoc(buildWarrantyHTML(proj,w.serial,getCo(acctSettings)),`Warranty — ${w.serial}`);}}>
                  Print Card
                </Btn>
                {isAdmin&&isActive&&(
                  <Btn size="sm" onClick={()=>{setClaimModal(w.id);setClaimForm(blankClaim);}}>
                    <Plus size={12}/>Log Claim
                  </Btn>
                )}
              </div>
            </div>

            {/* Claims list */}
            {(w.claims||[]).length>0&&(
              <div>
                <div style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',
                  letterSpacing:'0.07em',marginBottom:10,display:'flex',justifyContent:'space-between'}}>
                  <span>Warranty Claims ({w.claims.length})</span>
                  <span style={{color:T.danger}}>Total cost: {fmtSGD(claimsTotal)}</span>
                </div>
                <div style={{display:'flex',flexDirection:'column',gap:8}}>
                  {(w.claims||[]).map(claim=>(
                    <div key={claim.id} style={{background:T.dangerLight,border:'1px solid rgba(239,68,68,0.15)',
                      borderRadius:12,padding:'12px 16px'}}>
                      <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',flexWrap:'wrap',gap:8}}>
                        <div style={{flex:1,minWidth:0}}>
                          <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:4}}>{claim.issue}</div>
                          <div style={{display:'flex',gap:14,fontSize:12,color:T.muted,flexWrap:'wrap'}}>
                            <span>Date: {fmtDate(claim.date)}</span>
                            <span>Attended by: <strong style={{color:T.text}}>{claim.attendedBy}</strong></span>
                            <span>Cost: <strong style={{color:T.danger}}>{fmtSGD(claim.cost)}</strong></span>
                          </div>
                          {claim.notes&&<div style={{fontSize:12,color:T.muted,marginTop:6,fontStyle:'italic'}}>{claim.notes}</div>}
                        </div>
                        {(claim.photos||[]).length>0&&(
                          <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
                            {claim.photos.map((ph,idx)=>(
                              <img key={idx} src={ph} alt="Claim photo"
                                onClick={()=>setLightbox({src:ph,title:`Claim Photo — ${claim.issue}`})}
                                style={{width:52,height:52,objectFit:'cover',borderRadius:6,cursor:'zoom-in',
                                  border:'1px solid rgba(239,68,68,0.3)'}}/>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {(w.claims||[]).length===0&&(
              <div style={{fontSize:12,color:T.dim,fontStyle:'italic',padding:'6px 0'}}>No claims logged</div>
            )}
          </div>
        );
      })}

      {/* Issue warranty picker */}
      {issueModal==='pick'&&(
        <Modal title="Issue Warranty Certificate" onClose={()=>setIssueModal(null)}>
          <div style={{display:'flex',flexDirection:'column',gap:12}}>
            <div style={{fontSize:13,color:T.muted,marginBottom:4}}>
              Select a completed & archived project to issue a warranty for.
              A unique serial number will be generated automatically.
            </div>
            {eligibleProjects.map(p=>(
              <button key={p.id} onClick={()=>issueWarranty(p.id)}
                style={{background:T.bg,border:`1px solid ${T.borderLight}`,borderRadius:12,
                  padding:'14px 16px',cursor:'pointer',textAlign:'left',fontFamily:'inherit',
                  display:'flex',alignItems:'center',justifyContent:'space-between',transition:'border-color .15s'}}>
                <div>
                  <div style={{fontSize:13,fontWeight:700,color:T.text}}>{p.name}</div>
                  <div style={{fontSize:12,color:T.muted}}>{p.client} . Archived {fmtDate(p.archivedAt)}</div>
                </div>
                <div style={{fontSize:12,color:T.accent,fontWeight:600}}>Issue →</div>
              </button>
            ))}
            {eligibleProjects.length===0&&(
              <div style={{color:T.dim,fontSize:13,textAlign:'center',padding:20}}>
                All archived projects already have a warranty issued.
              </div>
            )}
          </div>
        </Modal>
      )}

      {/* Log claim modal */}
      {claimModal&&(
        <Modal title="Log Warranty Claim" onClose={()=>{setClaimModal(null);setClaimForm(blankClaim);}} wide>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:16}}>
            <div style={{gridColumn:'1/-1'}}>
              <Field label="Issue Description *" value={claimForm.issue} onChange={cf('issue')}
                placeholder="e.g. Cabinet hinge broken, paint peeling on wall"/>
            </div>
            <Field label="Attended By *" value={claimForm.attendedBy} onChange={cf('attendedBy')}
              placeholder="Name of staff / subcon who attended"/>
            <Field label="Date of Attendance" type="date" value={claimForm.date} onChange={cf('date')}/>
            <Field label="Cost Incurred (S$)" type="number" value={claimForm.cost} onChange={cf('cost')} placeholder="0"/>
            <div style={{gridColumn:'1/-1'}}>
              <label style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',
                letterSpacing:'0.07em',display:'block',marginBottom:5}}>Notes / Write-up</label>
              <textarea value={claimForm.notes} onChange={e=>cf('notes')(e.target.value)}
                placeholder="Describe the defect, cause, and remedial action taken…"
                rows={3} style={{...iStyle,resize:'vertical'}}/>
            </div>
            <div style={{gridColumn:'1/-1'}}>
              <label style={{fontSize:11,fontWeight:700,color:T.dim,textTransform:'uppercase',
                letterSpacing:'0.07em',display:'block',marginBottom:8}}>Photos</label>
              <div style={{display:'flex',gap:8,flexWrap:'wrap',alignItems:'center'}}>
                {(claimForm.photos||[]).map((ph,idx)=>(
                  <img key={idx} src={ph} alt="claim"
                    style={{width:64,height:64,objectFit:'cover',borderRadius:10,
                      border:`1px solid ${T.borderLight}`,cursor:'pointer'}}
                    onClick={()=>setLightbox({src:ph,title:'Claim Photo'})}/>
                ))}
                <button onClick={()=>claimFileRef.current?.click()}
                  style={{width:64,height:64,background:T.bg,
                    border:`2px dashed ${T.border}`,borderRadius:10,cursor:'pointer',
                    display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',gap:4,color:T.dim}}>
                  <Upload size={14}/><span style={{fontSize:10}}>Upload</span>
                </button>
                <button onClick={()=>claimCamRef.current?.click()}
                  style={{width:64,height:64,background:T.bg,
                    border:`2px dashed ${T.border}`,borderRadius:10,cursor:'pointer',
                    display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',gap:4,color:T.dim}}>
                  <Camera size={14}/><span style={{fontSize:10}}>Camera</span>
                </button>
                <input ref={claimFileRef} type="file" accept="image/*" style={{display:'none'}}
                  onChange={e=>{const f=e.target.files?.[0];if(f){const r=new FileReader();r.onload=ev=>addClaimPhoto(ev.target.result);r.readAsDataURL(f);}e.target.value='';}}/>
                <input ref={claimCamRef} type="file" accept="image/*" capture="environment" style={{display:'none'}}
                  onChange={e=>{const f=e.target.files?.[0];if(f){const r=new FileReader();r.onload=ev=>addClaimPhoto(ev.target.result);r.readAsDataURL(f);}e.target.value='';}}/>
              </div>
            </div>
          </div>
          <div style={{display:'flex',justifyContent:'flex-end',gap:10,marginTop:22}}>
            <Btn variant="secondary" onClick={()=>{setClaimModal(null);setClaimForm(blankClaim);}}>Cancel</Btn>
            <Btn onClick={()=>saveClaim(claimModal)} disabled={!claimForm.issue||!claimForm.attendedBy}>Save Claim</Btn>
          </div>
        </Modal>
      )}

      {lightbox&&<Lightbox src={lightbox.src} title={lightbox.title} onClose={()=>setLightbox(null)}/>}
    </div>
  );
}

function WorkerPortal({worker, onLogout, attendance, setAttendance, projects, claims, setClaims}){
  const [selProjectId,setSelProjectId]=useState('');
  const [locating,setLocating]=useState(false);
  const [locErr,setLocErr]=useState('');
  const [view,setView]=useState('main'); // 'main' | 'history' | 'claims' | 'newclaim'
  const claimTypes=['Food / Meals','Transport','Petrol','ERP / Parking','Material Purchase','Other'];
  const [claimForm,setClaimForm]=useState({type:'Food / Meals',amount:'',notes:'',photo:null,date:new Date().toISOString().slice(0,10)});
  const claimPhotoRef=useRef();
  const claimCamRef=useRef();
  const cf=k=>v=>setClaimForm(p=>({...p,[k]:v}));

  const workerClaims=(claims||[]).filter(c=>c.workerId===worker.id)
    .sort((a,b)=>b.date.localeCompare(a.date));

  const submitClaim=()=>{
    if(!claimForm.amount||parseFloat(claimForm.amount)<=0)return;
    const c={id:uid(),workerId:worker.id,
      projectId:selProjectId||worker.projectId||'',
      type:claimForm.type,amount:parseFloat(claimForm.amount),
      notes:claimForm.notes,photo:claimForm.photo,
      date:claimForm.date,status:'Pending',submittedAt:new Date().toISOString()};
    const upd=[...(claims||[]),c];
    setClaims(upd);saveS('workerClaims',upd);
    setClaimForm({type:'Food / Meals',amount:'',notes:'',photo:null,date:new Date().toISOString().slice(0,10)});
    setView('claims');
  };

  const assignedProjects=projects.filter(p=>
    (worker.assignedProjects||[worker.projectId]).includes(p.id)&&p.status!=='Cancelled'
  );
  const todayStr=new Date().toISOString().slice(0,10);
  const ym=todayStr.slice(0,7);
  const todayRec=selProjectId?attendance.find(a=>a.workerId===worker.id&&a.date===todayStr&&a.projectId===selProjectId):null;
  const isCheckedIn=!!(todayRec&&todayRec.checkIn&&!todayRec.checkOut);
  const selProj=projects.find(p=>p.id===selProjectId);
  const monthHours=attendance.filter(a=>a.workerId===worker.id&&a.date.startsWith(ym)&&a.totalHours).reduce((s,a)=>s+a.totalHours,0);
  const todayTotal=attendance.filter(a=>a.workerId===worker.id&&a.date===todayStr&&a.totalHours).reduce((s,a)=>s+a.totalHours,0);

  const getGeo=()=>new Promise((res,rej)=>{
    if(!navigator.geolocation){rej('Geolocation not supported');return;}
    navigator.geolocation.getCurrentPosition(p=>res({lat:p.coords.latitude,lng:p.coords.longitude}),()=>rej('Location access denied. Please allow location in browser settings.'),{enableHighAccuracy:true,timeout:15000});
  });

  const reverseGeo=async(lat,lng)=>{
    try{const r=await fetch(`https://nominatim.openstreetmap.org/reverse?lat=${lat}&lon=${lng}&format=json`);const d=await r.json();return d.display_name?.split(',').slice(0,3).join(', ')||`${lat.toFixed(4)}, ${lng.toFixed(4)}`;}catch{return `${lat.toFixed(4)}, ${lng.toFixed(4)}`;}
  };

  const handleCheckIn=async()=>{
    if(!selProjectId){setLocErr('Please select a site first.');return;}
    // Guard: already checked in at this project today?
    const existing=attendance.find(a=>a.workerId===worker.id&&a.date===todayStr&&a.projectId===selProjectId&&a.checkIn);
    if(existing&&!existing.checkOut){setLocErr('You are already checked in at this site today.');return;}
    if(existing&&existing.checkOut){setLocErr('You have already completed a session at this site today.');return;}
    // Guard: checked in at a DIFFERENT project and not yet checked out?
    const openSession=attendance.find(a=>a.workerId===worker.id&&a.date===todayStr&&a.checkIn&&!a.checkOut);
    if(openSession){
      const openProj=projects.find(p=>p.id===openSession.projectId);
      setLocErr(`Please check out from ${openProj?.name||'your current site'} first before checking in here.`);
      return;
    }
    setLocating(true);setLocErr('');
    try{
      const{lat,lng}=await getGeo();
      const addr=await reverseGeo(lat,lng);
      const now=new Date().toISOString();
      const rec={id:uid(),workerId:worker.id,projectId:selProjectId,date:todayStr,checkIn:{time:now,lat,lng,addr},checkOut:null,totalHours:null,isVO:false,notes:''};
      const upd=[...attendance,rec];setAttendance(upd);saveS('attendance',upd);
    }catch(e){setLocErr(String(e));}
    setLocating(false);
  };

  const handleCheckOut=async()=>{
    if(!todayRec)return;
    setLocating(true);setLocErr('');
    try{
      const{lat,lng}=await getGeo();
      const addr=await reverseGeo(lat,lng);
      const now=new Date().toISOString();
      const hrs=(new Date(now)-new Date(todayRec.checkIn.time))/3600000;
      if(hrs>24){setLocErr('Session exceeds 24 hours. Please contact admin to fix this record.');setLocating(false);return;}
      const upd=attendance.map(a=>a.id===todayRec.id?{...a,checkOut:{time:now,lat,lng,addr},totalHours:parseFloat(hrs.toFixed(2))}:a);
      setAttendance(upd);saveS('attendance',upd);
    }catch(e){setLocErr(String(e));}
    setLocating(false);
  };

  // Forgot-to-check-out: any open session from YESTERDAY or earlier
  const forgotCheckout=attendance.find(a=>
    a.workerId===worker.id && a.checkIn && !a.checkOut && a.date<todayStr
  );

  const recent=useMemo(()=>
    attendance.filter(a=>a.workerId===worker.id&&a.checkIn).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,10),
    [attendance,worker.id]
  );
  const timeStr=new Date().toLocaleTimeString('en-SG',{hour:'2-digit',minute:'2-digit',hour12:true});
  const dateStr2=new Date().toLocaleDateString('en-SG',{weekday:'long',day:'2-digit',month:'long'});

  return (
    <div style={{minHeight:'100vh',background:'#F5F5F7',display:'flex',flexDirection:'column',alignItems:'center',padding:'16px 16px 48px',maxWidth:440,margin:'0 auto',gap:14}}>
      <div style={{width:'100%',display:'flex',alignItems:'center',justifyContent:'space-between',paddingTop:8}}>
        <div style={{display:'flex',alignItems:'center',gap:10}}>
          <Avatar photo={worker.photo} name={worker.name} size={38} color={T.accent}/>
          <div>
            <div style={{fontSize:16,fontWeight:700,color:T.text}}>{worker.name}</div>
            <div style={{fontSize:12,color:T.muted}}>{timeStr} . {dateStr2}</div>
          </div>
        </div>
        <div style={{display:'flex',gap:6}}>
          <button onClick={()=>setView(v=>v==='history'?'main':'history')}
            style={{background:view==='history'?T.accentLight:T.card,border:`1px solid ${view==='history'?T.accent+'40':T.borderLight}`,borderRadius:10,padding:'6px 10px',cursor:'pointer',fontFamily:'inherit',fontSize:11,fontWeight:500,color:view==='history'?T.accent:T.muted}}>
            History
          </button>
          <button onClick={()=>setView(v=>v==='claims'||v==='newclaim'?'main':'claims')}
            style={{background:view==='claims'||view==='newclaim'?T.accentLight:T.card,border:`1px solid ${view==='claims'||view==='newclaim'?T.accent+'40':T.borderLight}`,borderRadius:10,padding:'6px 10px',cursor:'pointer',fontFamily:'inherit',fontSize:11,fontWeight:500,color:view==='claims'||view==='newclaim'?T.accent:T.muted}}>
            Claims{workerClaims.filter(c=>c.status==='Pending').length>0?` (${workerClaims.filter(c=>c.status==='Pending').length})`:''}
          </button>
          <button onClick={onLogout} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:10,padding:'6px 10px',cursor:'pointer',fontFamily:'inherit',fontSize:11,color:T.muted}}>
            Sign Out
          </button>
        </div>
      </div>

      {view==='history'?(
        <div style={{width:'100%',background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,overflow:'hidden',boxShadow:T.shadow}}>
          <div style={{padding:'14px 18px',borderBottom:`1px solid ${T.borderLight}`,fontSize:14,fontWeight:600,color:T.text}}>Attendance History</div>
          {recent.length===0&&<div style={{padding:28,textAlign:'center',color:T.dim,fontSize:13}}>No records yet</div>}
          {recent.map(a=>{const p=projects.find(x=>x.id===a.projectId);return(
            <div key={a.id} style={{padding:'11px 18px',borderBottom:`1px solid ${T.borderLight}`,display:'flex',justifyContent:'space-between',alignItems:'center'}}>
              <div>
                <div style={{fontSize:13,fontWeight:600,color:T.text}}>{new Date(a.date+'T00:00:00').toLocaleDateString('en-SG',{weekday:'short',day:'2-digit',month:'short'})}</div>
                <div style={{fontSize:11,color:T.muted,marginTop:1}}>{p?.name||'—'}</div>
                <div style={{fontSize:11,color:T.dim}}>{fmtTime(a.checkIn?.time)} → {a.checkOut?fmtTime(a.checkOut.time):'Not checked out'}</div>
              </div>
              <div style={{fontSize:14,fontWeight:700,color:a.totalHours?T.accent:T.warning}}>{a.totalHours?fmtHours(a.totalHours):'Active'}</div>
            </div>
          );})}
        </div>
      ):(
        <>
          {/* Forgot-to-check-out warning */}
          {forgotCheckout&&(
            <div style={{width:'100%',background:'#FEF9EC',border:'1px solid #F5DFA0',borderRadius:14,padding:'12px 16px',display:'flex',alignItems:'center',gap:10}}>
              <AlertTriangle size={16} style={{color:'#B7860A',flexShrink:0}}/>
              <div>
                <div style={{fontSize:13,fontWeight:700,color:'#78350F'}}>You forgot to check out</div>
                <div style={{fontSize:12,color:'#92400E',marginTop:2}}>Check-in on {forgotCheckout.date} at {fmtTime(forgotCheckout.checkIn?.time)} was never closed. Contact admin to fix this.</div>
              </div>
            </div>
          )}

          <div style={{width:'100%',background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:'16px 22px',boxShadow:T.shadow}}>
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:0}}>
              <div style={{textAlign:'center',padding:'6px 0'}}>
                <div style={{fontSize:11,color:T.dim,marginBottom:4,textTransform:'uppercase',letterSpacing:'0.06em'}}>Today</div>
                <div style={{fontSize:22,fontWeight:700,color:T.accent}}>{todayTotal>0?fmtHours(todayTotal):isCheckedIn?'⏱':'—'}</div>
              </div>
              <div style={{textAlign:'center',padding:'6px 0',borderLeft:`1px solid ${T.borderLight}`}}>
                <div style={{fontSize:11,color:T.dim,marginBottom:4,textTransform:'uppercase',letterSpacing:'0.06em'}}>This Month</div>
                <div style={{fontSize:22,fontWeight:700,color:T.text}}>{fmtHours(monthHours)}</div>
              </div>
            </div>
          </div>

          <div style={{width:'100%'}}>
            <div style={{fontSize:13,fontWeight:600,color:T.muted,marginBottom:8,textAlign:'center'}}>
              {isCheckedIn?'Currently working at:':'Select site to check in:'}
            </div>
            <div style={{display:'flex',flexDirection:'column',gap:8}}>
              {assignedProjects.map(p=>{
                const pr=attendance.find(a=>a.workerId===worker.id&&a.date===todayStr&&a.projectId===p.id);
                const active=!!(pr&&pr.checkIn&&!pr.checkOut);
                const done=!!(pr&&pr.checkOut);
                const sel=selProjectId===p.id;
                return(
                  <button key={p.id} onClick={()=>{if(!isCheckedIn)setSelProjectId(p.id);else if(active)setSelProjectId(p.id);}}
                    disabled={isCheckedIn&&!active}
                    style={{width:'100%',background:T.card,border:`2px solid ${active?T.success:sel?T.accent:T.borderLight}`,borderRadius:16,padding:'14px 18px',cursor:(isCheckedIn&&!active)?'not-allowed':'pointer',fontFamily:'inherit',textAlign:'left',transition:'all 0.15s',opacity:(isCheckedIn&&!active&&!done)?0.38:1,boxShadow:(sel||active)?T.shadowMd:T.shadow}}>
                    <div style={{display:'flex',alignItems:'center',justifyContent:'space-between'}}>
                      <div>
                        <div style={{fontSize:15,fontWeight:700,color:T.text}}>{p.name}</div>
                        <div style={{fontSize:12,color:T.muted,marginTop:2}}>{p.client}</div>
                      </div>
                      <div style={{fontSize:12,fontWeight:700}}>
                        {active&&<span style={{color:T.success,display:'flex',alignItems:'center',gap:5}}><span style={{width:8,height:8,borderRadius:'50%',background:T.success,display:'inline-block'}}/>Checked In</span>}
                        {done&&<span style={{color:T.dim}}>✓ Done</span>}
                        {!active&&!done&&sel&&<span style={{color:T.accent}}>Selected ✓</span>}
                      </div>
                    </div>
                  </button>
                );
              })}
              {assignedProjects.length===0&&<div style={{padding:24,textAlign:'center',color:T.dim,fontSize:13,background:T.card,borderRadius:14,border:`1px solid ${T.borderLight}`}}>No sites assigned. Contact admin.</div>}
            </div>
          </div>

          {locErr&&<div style={{width:'100%',background:T.dangerLight,border:`1px solid ${T.danger}30`,borderRadius:12,padding:'10px 16px',fontSize:13,color:T.danger,textAlign:'center'}}>{locErr}</div>}

          {selProjectId&&(
            <div style={{width:'100%'}}>
              {!todayRec&&(
                <button onClick={handleCheckIn} disabled={locating} style={{width:'100%',background:T.success,color:'#fff',border:'none',borderRadius:20,padding:'22px',fontSize:20,fontWeight:700,cursor:locating?'wait':'pointer',fontFamily:'inherit',display:'flex',alignItems:'center',justifyContent:'center',gap:12,boxShadow:'0 6px 24px rgba(29,131,72,0.35)',opacity:locating?0.7:1}}>
                  {locating?<Loader2 size={24} style={{animation:'spin 1s linear infinite'}}/>:<LogIn size={24}/>}
                  {locating?'Getting your location…':'CHECK IN'}
                </button>
              )}
              {isCheckedIn&&(
                <div style={{display:'flex',flexDirection:'column',gap:10}}>
                  <div style={{background:T.successLight,border:`1px solid ${T.success}30`,borderRadius:14,padding:'12px 18px',display:'flex',alignItems:'center',gap:12}}>
                    <div style={{width:10,height:10,borderRadius:'50%',background:T.success,flexShrink:0}}/>
                    <div><div style={{fontSize:13,fontWeight:700,color:T.success}}>Working at {selProj?.name}</div><div style={{fontSize:12,color:T.muted,marginTop:2}}>Since {fmtTime(todayRec.checkIn.time)} . {todayRec.checkIn.addr}</div></div>
                  </div>
                  <button onClick={handleCheckOut} disabled={locating} style={{width:'100%',background:T.danger,color:'#fff',border:'none',borderRadius:20,padding:'22px',fontSize:20,fontWeight:700,cursor:locating?'wait':'pointer',fontFamily:'inherit',display:'flex',alignItems:'center',justifyContent:'center',gap:12,boxShadow:'0 6px 24px rgba(192,57,43,0.35)',opacity:locating?0.7:1}}>
                    {locating?<Loader2 size={24} style={{animation:'spin 1s linear infinite'}}/>:<LogOut size={24}/>}
                    {locating?'Getting your location…':'CHECK OUT'}
                  </button>
                </div>
              )}
              {todayRec&&!isCheckedIn&&(
                <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:16,padding:'20px 24px',textAlign:'center',boxShadow:T.shadow}}>
                  <CheckCircle size={28} style={{color:T.success,margin:'0 auto 10px'}}/>
                  <div style={{fontSize:16,fontWeight:700,color:T.text}}>Done for today!</div>
                  <div style={{fontSize:13,color:T.muted,marginTop:4}}>{fmtTime(todayRec.checkIn.time)} → {fmtTime(todayRec.checkOut?.time)}</div>
                  <div style={{fontSize:16,fontWeight:700,color:T.accent,marginTop:4}}>{fmtHours(todayRec.totalHours)} at {selProj?.name}</div>
                </div>
              )}
            </div>
          )}
        </>
      )}

      {/* Claims view */}
      {view==='claims'&&(
        <div style={{width:'100%',display:'flex',flexDirection:'column',gap:12}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
            <div style={{fontSize:14,fontWeight:600,color:T.text}}>My Claims</div>
            <Btn size="sm" onClick={()=>setView('newclaim')}><Plus size={12}/>New Claim</Btn>
          </div>
          {workerClaims.length===0&&(
            <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:14,padding:28,textAlign:'center',color:T.dim,fontSize:13,boxShadow:T.shadow}}>
              No claims submitted yet
            </div>
          )}
          {workerClaims.map(c=>(
            <div key={c.id} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:14,padding:'14px 16px',boxShadow:T.shadow}}>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:6}}>
                <div>
                  <div style={{fontSize:14,fontWeight:700,color:T.text}}>{c.type}</div>
                  <div style={{fontSize:12,color:T.muted,marginTop:2}}>{c.date}</div>
                  {c.notes&&<div style={{fontSize:12,color:T.dim,marginTop:3,fontStyle:'italic'}}>{c.notes}</div>}
                </div>
                <div style={{textAlign:'right'}}>
                  <div style={{fontSize:16,fontWeight:700,color:T.text}}>{fmtSGD(parseFloat(c.amount))}</div>
                  <Badge color={c.status==='Approved'?T.success:c.status==='Rejected'?T.danger:T.warning} sm>
                    {c.status}
                  </Badge>
                </div>
              </div>
              {c.photo&&(
                <img src={c.photo} alt="receipt" style={{width:'100%',maxHeight:140,objectFit:'cover',borderRadius:8,marginTop:6,border:`1px solid ${T.borderLight}`}}/>
              )}
            </div>
          ))}
        </div>
      )}

      {/* New claim form */}
      {view==='newclaim'&&(
        <div style={{width:'100%',background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:20,boxShadow:T.shadow,display:'flex',flexDirection:'column',gap:14}}>
          <div style={{fontSize:15,fontWeight:700,color:T.text}}>Submit a Claim</div>

          <div>
            <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Claim Type</label>
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:6}}>
              {claimTypes.map(t=>(
                <button key={t} onClick={()=>cf('type')(t)}
                  style={{padding:'10px 8px',borderRadius:10,textAlign:'center',fontSize:12,fontWeight:600,
                    border:`1px solid ${claimForm.type===t?T.accent:T.borderLight}`,
                    background:claimForm.type===t?T.accentLight:'transparent',
                    color:claimForm.type===t?T.accent:T.muted,cursor:'pointer',fontFamily:'inherit'}}>
                  {t}
                </button>
              ))}
            </div>
          </div>

          <div>
            <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Amount (S$) *</label>
            <input type="number" value={claimForm.amount} onChange={e=>cf('amount')(e.target.value)}
              placeholder="0.00" style={{...iStyle,fontSize:18,fontWeight:700,textAlign:'center'}}/>
          </div>

          <div>
            <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Date</label>
            <input type="date" value={claimForm.date} onChange={e=>cf('date')(e.target.value)} style={iStyle}/>
          </div>

          <div>
            <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Notes (optional)</label>
            <input value={claimForm.notes} onChange={e=>cf('notes')(e.target.value)}
              placeholder="e.g. Lunch at site, Grab from MRT to site..." style={iStyle}/>
          </div>

          <div>
            <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:8}}>Receipt / Photo</label>
            {claimForm.photo?(
              <div style={{position:'relative'}}>
                <img src={claimForm.photo} alt="receipt" style={{width:'100%',maxHeight:160,objectFit:'cover',borderRadius:10,border:`1px solid ${T.borderLight}`}}/>
                <button onClick={()=>cf('photo')(null)}
                  style={{position:'absolute',top:6,right:6,background:'rgba(0,0,0,0.5)',border:'none',borderRadius:6,padding:'4px 8px',cursor:'pointer',color:'#fff',fontSize:11}}>
                  Remove
                </button>
              </div>
            ):(
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:8}}>
                <button onClick={()=>claimPhotoRef.current?.click()}
                  style={{background:T.bg,border:`2px dashed ${T.borderLight}`,borderRadius:12,padding:'16px 8px',cursor:'pointer',fontFamily:'inherit',color:T.muted,fontSize:12,display:'flex',flexDirection:'column',alignItems:'center',gap:6}}>
                  <Upload size={20}/><span>Upload</span><span style={{fontSize:10,color:T.dim}}>from gallery</span>
                </button>
                <button onClick={()=>claimCamRef.current?.click()}
                  style={{background:T.bg,border:`2px dashed ${T.borderLight}`,borderRadius:12,padding:'16px 8px',cursor:'pointer',fontFamily:'inherit',color:T.muted,fontSize:12,display:'flex',flexDirection:'column',alignItems:'center',gap:6}}>
                  <Camera size={20}/><span>Camera</span><span style={{fontSize:10,color:T.dim}}>take photo</span>
                </button>
                <input ref={claimPhotoRef} type="file" accept="image/*" style={{display:'none'}}
                  onChange={async e=>{const f=e.target.files?.[0];if(f){const r=new FileReader();r.onload=async ev=>{cf('photo')(await compressPhoto(ev.target.result));};r.readAsDataURL(f);}e.target.value='';}}/>
                <input ref={claimCamRef} type="file" accept="image/*" capture="environment" style={{display:'none'}}
                  onChange={async e=>{const f=e.target.files?.[0];if(f){const r=new FileReader();r.onload=async ev=>{cf('photo')(await compressPhoto(ev.target.result));};r.readAsDataURL(f);}e.target.value='';}}/>
              </div>
            )}
          </div>

          <div style={{display:'flex',gap:10,marginTop:4}}>
            <Btn variant="secondary" full onClick={()=>setView('claims')}>Cancel</Btn>
            <Btn full onClick={submitClaim} disabled={!claimForm.amount||parseFloat(claimForm.amount)<=0}>
              <CheckCircle size={14}/>Submit Claim
            </Btn>
          </div>
        </div>
      )}
    </div>
  );
}

function WorkerAdmin({siteWorkers,setSiteWorkers,attendance,setAttendance,projects,invoices,setInvoices,claims=[],setClaims,acctSettings,logAction=()=>{}}){
  const [view,setView]=useState('workers');
  const [selWorker,setSelWorker]=useState(null);
  const [modal,setModal]=useState(null);
  const [month,setMonth]=useState(new Date().toISOString().slice(0,7));
  const [deleteTarget,setDeleteTarget]=useState(null);
  const certFileRef=useRef();

  const blank={id:'',name:'',phone:'',nric:'',nationality:'',projectId:'',assignedProjects:[],
    pin:'',costRate:'12',billingRate:'18',monthlySalary:'1800',
    workPassNo:'',workPassExpiry:'',certificates:[],status:'Active',createdAt:new Date().toISOString()};
  const [form,setForm]=useState(blank);
  const ff=k=>v=>setForm(p=>({...p,[k]:v}));

  const saveWorker=()=>{
    if(!form.name)return;
    const d={...form,costRate:parseFloat(form.costRate)||0,
      billingRate:parseFloat(form.billingRate)||0,
      monthlySalary:parseFloat(form.monthlySalary)||0,
      assignedProjects:form.assignedProjects||[],
      pin:form.pin||'1234'};
    const isNew=modal==='new';
    const upd=isNew?[...siteWorkers,{...d,id:uid()}]:siteWorkers.map(w=>w.id===d.id?d:w);
    setSiteWorkers(upd);saveS('siteWorkers',upd);setModal(null);
    logAction(isNew?'CREATE_WORKER':'EDIT_WORKER',`${isNew?'Added':'Edited'} site worker: ${d.name} — ${d.trade||'General'} (${d.status})`);
  };

  const addCert=()=>setForm(p=>({...p,certificates:[...p.certificates,
    {id:uid(),name:'',issueDate:'',expiryDate:''}]}));
  const updateCert=(i,k,v)=>setForm(p=>({...p,certificates:p.certificates.map((c,idx)=>idx===i?{...c,[k]:v}:c)}));
  const removeCert=(i)=>setForm(p=>({...p,certificates:p.certificates.filter((_,idx)=>idx!==i)}));


  const alerts=[];
  siteWorkers.filter(w=>w.status==='Active').forEach(w=>{
    const wpd=daysUntil(w.workPassExpiry);
    if(wpd<=60) alerts.push({worker:w.name,doc:'Work Pass',days:wpd,expiry:w.workPassExpiry});
    (w.certificates||[]).forEach(c=>{
      if(!c.expiryDate)return;
      const cd=daysUntil(c.expiryDate);
      if(cd<=60) alerts.push({worker:w.name,doc:c.name,days:cd,expiry:c.expiryDate});
    });
  });



  const monthRecs=attendance.filter(a=>a.date.startsWith(month));

  // Total hours for the month per worker
  const workerMonth=(wid)=>{
    const recs=monthRecs.filter(a=>a.workerId===wid&&a.totalHours);
    const hrs=recs.reduce((s,a)=>s+a.totalHours,0);
    const days=[...new Set(recs.map(a=>a.date))].length;
    return {hrs,days,recs};
  };

  // Per-project breakdown — only VO-marked records for billing, all records for payslip
  const workerMonthByProject=(wid,voOnly=false)=>{
    const recs=monthRecs.filter(a=>a.workerId===wid&&a.totalHours&&(voOnly?a.isVO:true));
    const byProj={};
    recs.forEach(a=>{
      if(!byProj[a.projectId])byProj[a.projectId]={hrs:0,days:0,dates:[]};
      byProj[a.projectId].hrs+=a.totalHours;
      if(!byProj[a.projectId].dates.includes(a.date)){byProj[a.projectId].dates.push(a.date);byProj[a.projectId].days++;}
    });
    return Object.entries(byProj).map(([projId,v])=>({
      projId,proj:projects.find(p=>p.id===projId),hrs:v.hrs,days:v.days,
      billing:v.hrs*(siteWorkers.find(w=>w.id===wid)?.billingRate||0)
    }));
  };

  // Generate VO invoice entry in a project for worker billing
  const generateWorkerVO=(worker,projId,hrs,billingAmt,monthLabel,setInvoices,invoices)=>{
    const voRef=`VO-WK-${worker.name.split(' ')[0].toUpperCase()}-${month.replace('-','')}`;
    // Check if VO already exists for this worker+project+month
    const exists=invoices.find(i=>i.invoiceNo===voRef&&i.projectId===projId);
    if(exists){console.warn('VO already posted:', voRef);return;}
    const voInv={
      id:uid(), projectId:projId, invoiceNo:voRef,
      supplier:`Site Worker — ${worker.name}`,
      category:'Labour (VO)', invoiceDate:new Date(month+'-28').toISOString().slice(0,10),
      subtotal:parseFloat(billingAmt.toFixed(2)), gst:0,
      total:parseFloat(billingAmt.toFixed(2)),
      status:'Approved', isVO:true,
      notes:`${worker.name} — ${hrs.toFixed(1)}h @ S$${worker.billingRate}/hr — ${monthLabel}`,
      createdAt:new Date().toISOString(),
    };
    const upd=[...invoices,voInv];
    setInvoices(upd);saveInvoices(upd);
    return voRef;
  };

  const generatePayslip=(worker,invoices,setInvoices,acctSettings)=>{
    const co=getCo(acctSettings);
    const projBreakdown=workerMonthByProject(worker.id,false);   // all hours for payslip
    const projVORows=workerMonthByProject(worker.id,true);        // VO-only for billing
    const{hrs,days}=workerMonth(worker.id);
    const hrPay=hrs*(worker.costRate||0);
    const total=hrPay+(worker.monthlySalary||0);
    const monthLabel=new Date(month+'-01').toLocaleDateString('en-SG',{month:'long',year:'numeric'});

    // Build per-project payslip rows (all hours, no billing amounts shown)
    const projRows=projBreakdown.map(({proj,hrs:ph,days:pd})=>{
      const amt=ph*(worker.costRate||0);
      return {projName:proj?.name||'Unknown Site',client:proj?.client||'',hrs:ph,days:pd,amt,projId:proj?.id};
    });

    const html=`
<div style="max-width:660px;margin:0 auto;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;">
  <!-- Header -->
  <div style="background:#0a1628;color:#fff;padding:26px 32px;border-radius:10px 10px 0 0;display:flex;justify-content:space-between;align-items:flex-start;">
    <div>
      <div style="font-size:22px;font-weight:800;">${co.name}</div>
      <div style="font-size:11px;color:#94a3b8;margin-top:3px;">${co.name}</div>
    </div>
    <div style="text-align:right;">
      <div style="font-size:18px;font-weight:800;color:#0071E3;">PAYSLIP</div>
      <div style="font-size:11px;color:#94a3b8;margin-top:3px;">${monthLabel}</div>
    </div>
  </div>

  <!-- Body -->
  <div style="border:1px solid #e2e8f0;border-top:none;border-radius:0 0 10px 10px;padding:26px 32px;">

    <!-- Employee only -->
    <div style="margin-bottom:24px;padding-bottom:20px;border-bottom:2px solid #e2e8f0;">
      <div style="font-size:10px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:0.1em;margin-bottom:10px;">Employee Details</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">
        <div>
          <div style="font-size:17px;font-weight:800;color:#1a1a2e;">${worker.name}</div>
          <div style="font-size:12px;color:#64748b;margin-top:4px;">${worker.nationality}</div>
        </div>
        <div style="font-size:12px;color:#64748b;line-height:1.9;">
          <div><strong style="color:#1a1a2e;">NRIC / FIN:</strong> ${worker.nric||'—'}</div>
          <div><strong style="color:#1a1a2e;">Phone:</strong> ${worker.phone||'—'}</div>
          <div><strong style="color:#1a1a2e;">Work Pass:</strong> ${worker.workPassNo||'—'}</div>
          <div><strong style="color:#1a1a2e;">Expires:</strong> ${worker.workPassExpiry?new Date(worker.workPassExpiry).toLocaleDateString('en-SG',{day:'2-digit',month:'short',year:'numeric'}):'—'}</div>
        </div>
      </div>
    </div>

    <!-- Earnings table with per-project breakdown -->
    <div style="margin-bottom:24px;">
      <div style="font-size:11px;font-weight:700;color:#1a1a2e;margin-bottom:10px;text-transform:uppercase;letter-spacing:0.07em;">Earnings Breakdown</div>
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:#f8fafc;">
            <th style="padding:9px 12px;text-align:left;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.06em;">Description</th>
            <th style="padding:9px 12px;text-align:center;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;">Days</th>
            <th style="padding:9px 12px;text-align:right;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;">Hours</th>
            <th style="padding:9px 12px;text-align:right;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;">Rate</th>
            <th style="padding:9px 12px;text-align:right;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;">Amount</th>
          </tr>
        </thead>
        <tbody>
          ${worker.monthlySalary>0?`
          <tr style="border-bottom:1px solid #e2e8f0;">
            <td style="padding:10px 12px;font-size:13px;color:#1a1a2e;">Basic Monthly Salary</td>
            <td style="padding:10px 12px;text-align:center;font-size:12px;color:#64748b;">${days}d</td>
            <td style="padding:10px 12px;text-align:right;font-size:12px;color:#64748b;">—</td>
            <td style="padding:10px 12px;text-align:right;font-size:12px;color:#64748b;">S$${worker.monthlySalary.toFixed(2)}/mth</td>
            <td style="padding:10px 12px;text-align:right;font-size:13px;font-weight:700;color:#1a1a2e;">S$${worker.monthlySalary.toFixed(2)}</td>
          </tr>`:''}
          ${projRows.length>0?`
          <tr style="background:#f0f9ff;">
            <td colspan="5" style="padding:7px 12px;font-size:10px;font-weight:700;color:#0369a1;text-transform:uppercase;letter-spacing:0.06em;">Hourly Pay — Site Breakdown @ S$${worker.costRate}/hr</td>
          </tr>
          ${projRows.map(r=>`
          <tr style="border-bottom:1px solid #e2e8f0;">
            <td style="padding:10px 12px 10px 20px;font-size:13px;color:#1a1a2e;">
              ${r.projName}
              <div style="font-size:10px;color:#94a3b8;margin-top:2px;">${r.client}</div>
            </td>
            <td style="padding:10px 12px;text-align:center;font-size:12px;color:#64748b;">${r.days}d</td>
            <td style="padding:10px 12px;text-align:right;font-size:12px;color:#64748b;">${r.hrs.toFixed(1)}h</td>
            <td style="padding:10px 12px;text-align:right;font-size:12px;color:#64748b;">S$${worker.costRate}/hr</td>
            <td style="padding:10px 12px;text-align:right;font-size:13px;font-weight:700;color:#1a1a2e;">S$${r.amt.toFixed(2)}</td>
          </tr>`).join('')}
          <tr style="background:#f8fafc;">
            <td colspan="2" style="padding:8px 12px;font-size:12px;color:#64748b;">Total Hourly Pay</td>
            <td style="padding:8px 12px;text-align:right;font-size:12px;font-weight:700;color:#0369a1;">${hrs.toFixed(1)}h</td>
            <td></td>
            <td style="padding:8px 12px;text-align:right;font-size:13px;font-weight:700;color:#1a1a2e;">S$${hrPay.toFixed(2)}</td>
          </tr>`:''}
        </tbody>
        <tfoot>
          <tr style="border-top:2px solid #e2e8f0;">
            <td colspan="4" style="padding:12px 12px;font-size:15px;font-weight:800;color:#1a1a2e;">TOTAL TAKE-HOME PAY</td>
            <td style="padding:12px 12px;text-align:right;font-size:20px;font-weight:800;color:#0071E3;">S$${total.toFixed(2)}</td>
          </tr>
        </tfoot>
      </table>
    </div>

    <!-- Footer -->
    <div style="text-align:center;color:#94a3b8;font-size:10px;line-height:1.8;border-top:1px solid #e2e8f0;padding-top:14px;">
      Computer-generated payslip . No signature required . ${co.name} . ${monthLabel}
    </div>
  </div>
</div>`;

    printDoc(html,`Payslip — ${worker.name} — ${monthLabel}`);

    // Auto-post VO to each project — only for records marked "Bill as VO"
    let voPosted=0;
    projVORows.forEach(r=>{
      if(r.projId&&r.billing>0){
        const ref=generateWorkerVO(worker,r.projId,r.hrs,r.billing,monthLabel,setInvoices,invoices);
        if(ref)voPosted++;
      }
    });
    // VO posting complete — log result (no blocking alert)
    console.log(`Payslip printed. VO posted: ${voPosted} project(s).`);
  };

  return (
    <div style={{display:'flex',flexDirection:'column',gap:20}}>

      {/* Expiry Alerts */}
      {alerts.length>0&&(
        <div style={{background:T.card,border:`1px solid ${T.danger}30`,borderRadius:14,overflow:'hidden',boxShadow:T.shadow}}>
          <div style={{padding:'12px 18px',background:T.dangerLight,display:'flex',alignItems:'center',gap:10,borderBottom:`1px solid ${T.danger}20`}}>
            <AlertTriangle size={16} style={{color:T.danger}}/>
            <span style={{fontSize:13,fontWeight:700,color:T.danger}}>{alerts.length} Document{alerts.length>1?'s':''} Expiring Soon</span>
          </div>
          <div style={{display:'flex',flexDirection:'column',gap:0}}>
            {alerts.sort((a,b)=>a.days-b.days).map((al,i)=>(
              <div key={i} style={{display:'flex',alignItems:'center',justifyContent:'space-between',
                padding:'10px 18px',borderBottom:i<alerts.length-1?`1px solid ${T.borderLight}`:'none'}}>
                <div style={{display:'flex',alignItems:'center',gap:10}}>
                  <CreditCard size={14} style={{color:expiryColor(al.days),flexShrink:0}}/>
                  <div>
                    <span style={{fontSize:13,fontWeight:600,color:T.text}}>{al.worker}</span>
                    <span style={{fontSize:12,color:T.muted,marginLeft:8}}>{al.doc}</span>
                  </div>
                </div>
                <div style={{display:'flex',alignItems:'center',gap:8}}>
                  <span style={{fontSize:12,color:T.muted}}>{fmtDate(al.expiry)}</span>
                  <Badge color={expiryColor(al.days)}>{expiryLabel(al.days)}</Badge>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Sub-tabs */}
      <div style={{display:'flex',background:'rgba(0,0,0,0.04)',borderRadius:12,padding:4,gap:2,width:'fit-content'}}>
        {[{id:'workers',l:'Workers'},{id:'attendance',l:'Attendance'},{id:'payroll',l:'Payroll & Billing'},{id:'claims',l:`Claims (${claims.filter(c=>c.status==='Pending').length})`}].map(({id,l})=>(
          <button key={id} onClick={()=>setView(id)}
            style={{padding:'7px 18px',borderRadius:10,border:'none',cursor:'pointer',fontFamily:'inherit',
              fontSize:13,fontWeight:view===id?600:400,
              background:view===id?T.card:'transparent',
              color:view===id?T.text:T.muted,
              boxShadow:view===id?T.shadow:'none',transition:'all 0.15s'}}>
            {l}
          </button>
        ))}
      </div>

      {/* Workers list */}
      {view==='workers'&&(
        <div style={{display:'flex',flexDirection:'column',gap:12}}>
          <div style={{display:'flex',justifyContent:'flex-end'}}>
            <Btn onClick={()=>{setForm(blank);setModal('new');}}>
              <Plus size={13}/>Add Worker
            </Btn>
          </div>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(min(340px,100%),1fr))',gap:14}}>
            {siteWorkers.map(w=>{
              const wpDays=daysUntil(w.workPassExpiry);
              const certAlerts=(w.certificates||[]).filter(c=>daysUntil(c.expiryDate)<=60);
              const {hrs,days}=workerMonth(w.id);
              const proj=projects.find(p=>p.id===w.projectId);
              return (
                <div key={w.id} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:20,boxShadow:T.shadow,opacity:w.status==='Active'?1:0.55}}>
                  <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:14}}>
                    <div style={{display:'flex',alignItems:'center',gap:12}}>
                      <Avatar
                        photo={w.photo}
                        name={w.name}
                        size={52}
                        color={T.accent}
                        editable
                        onUpload={photo=>{
                          const upd=siteWorkers.map(x=>x.id===w.id?{...x,photo}:x);
                          setSiteWorkers(upd);saveS('siteWorkers',upd);
                        }}
                      />
                      <div>
                        <div style={{fontSize:14,fontWeight:700,color:T.text}}>{w.name}</div>
                        <div style={{fontSize:12,color:T.muted}}>{w.nationality} . {w.phone}</div>
                        <div style={{fontSize:11,color:T.dim,marginTop:2}}>{proj?.name||'No project'}</div>
                      </div>
                    </div>
                    <div style={{display:'flex',gap:6}}>
                      <button onClick={()=>{setForm({...w,costRate:String(w.costRate),billingRate:String(w.billingRate),monthlySalary:String(w.monthlySalary),assignedProjects:w.assignedProjects||[w.projectId].filter(Boolean),pin:w.pin||''});setModal('edit');}}
                        style={{background:T.bg,border:`1px solid ${T.borderLight}`,borderRadius:8,padding:'5px 8px',cursor:'pointer',color:T.muted,display:'flex'}}>
                        <Edit3 size={13}/>
                      </button>
                    </div>
                  </div>

                  {/* Rates */}
                  <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:8,marginBottom:12}}>
                    {[{l:'Cost/hr',v:`S$${w.costRate}`},{l:'Bill/hr',v:`S$${w.billingRate}`},{l:'Salary',v:fmtSGD(w.monthlySalary)}].map(({l,v})=>(
                      <div key={l} style={{background:T.bg,borderRadius:8,padding:'8px 10px'}}>
                        <div style={{fontSize:10,color:T.dim,marginBottom:2}}>{l}</div>
                        <div style={{fontSize:13,fontWeight:700,color:T.text}}>{v}</div>
                      </div>
                    ))}
                  </div>

                  {/* This month */}
                  <div style={{display:'flex',gap:10,marginBottom:12}}>
                    <div style={{background:T.accentLight,borderRadius:8,padding:'7px 12px',flex:1,textAlign:'center'}}>
                      <div style={{fontSize:10,color:T.accent}}>This Month</div>
                      <div style={{fontSize:14,fontWeight:700,color:T.accent}}>{fmtHours(hrs)}</div>
                      <div style={{fontSize:10,color:T.muted}}>{days} days</div>
                    </div>
                    <div style={{background:T.infoLight,borderRadius:8,padding:'7px 12px',flex:1,textAlign:'center'}}>
                      <div style={{fontSize:10,color:T.info}}>Billing this month</div>
                      <div style={{fontSize:14,fontWeight:700,color:T.info}}>{fmtSGD(hrs*(w.billingRate||0))}</div>
                    </div>
                  </div>

                  {/* Work Pass */}
                  <div style={{padding:'8px 12px',background:wpDays<=60?T.dangerLight:T.bg,
                    border:`1px solid ${wpDays<=60?T.danger+'30':T.borderLight}`,
                    borderRadius:8,marginBottom:8,display:'flex',alignItems:'center',justifyContent:'space-between'}}>
                    <div style={{display:'flex',alignItems:'center',gap:6}}>
                      <CreditCard size={12} style={{color:expiryColor(wpDays)}}/>
                      <span style={{fontSize:12,color:T.text}}>Work Pass {w.workPassNo}</span>
                    </div>
                    <Badge color={expiryColor(wpDays)} sm>{expiryLabel(wpDays)}</Badge>
                  </div>

                  {/* Certificates */}
                  {(w.certificates||[]).map(c=>{
                    const cd=daysUntil(c.expiryDate);
                    return (
                      <div key={c.id} style={{padding:'6px 12px',background:cd<=60?T.dangerLight:T.bg,
                        border:`1px solid ${cd<=60?T.danger+'30':T.borderLight}`,
                        borderRadius:8,marginBottom:6,display:'flex',alignItems:'center',justifyContent:'space-between'}}>
                        <div style={{display:'flex',alignItems:'center',gap:6}}>
                          <Star size={11} style={{color:expiryColor(cd)}}/>
                          <span style={{fontSize:12,color:T.text}}>{c.name}</span>
                        </div>
                        <Badge color={expiryColor(cd)} sm>{expiryLabel(cd)}</Badge>
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Attendance */}
      {view==='attendance'&&(
        <div style={{display:'flex',flexDirection:'column',gap:14}}>
          <div style={{display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
            <input type="month" value={month} onChange={e=>setMonth(e.target.value)}
              style={{...iStyle,width:'auto',minWidth:160}}/>
            <select value={selWorker||''} onChange={e=>setSelWorker(e.target.value||null)}
              style={{...iStyle,width:'auto',minWidth:160}}>
              <option value="">All Workers</option>
              {siteWorkers.map(w=><option key={w.id} value={w.id}>{w.name}</option>)}
            </select>
            <div style={{fontSize:13,color:T.muted}}>{monthRecs.filter(a=>!selWorker||a.workerId===selWorker).length} records</div>
            <Btn variant="secondary" size="sm" onClick={()=>{
              const monthLabel=new Date(month+'-01').toLocaleDateString('en-SG',{month:'long',year:'numeric'});
              const wb=XLSX.utils.book_new();
              const rows=[['Worker','Nationality','NRIC','Project','Date','Check In','Check Out','Hours','Billing Type','Cost Rate','Cost','Bill Rate','Billing']];
              const recs=monthRecs.filter(a=>!selWorker||a.workerId===selWorker);
              recs.sort((a,b)=>a.date.localeCompare(b.date)||(a.workerId.localeCompare(b.workerId))).forEach(a=>{
                const w=siteWorkers.find(x=>x.id===a.workerId);
                const p=projects.find(x=>x.id===a.projectId);
                const hrs=a.totalHours||0;
                rows.push([w?.name||'',w?.nationality||'',w?.nric||'',p?.name||'',a.date,
                  a.checkIn?.time?new Date(a.checkIn.time).toLocaleTimeString('en-SG',{hour:'2-digit',minute:'2-digit',hour12:true}):'',
                  a.checkOut?.time?new Date(a.checkOut.time).toLocaleTimeString('en-SG',{hour:'2-digit',minute:'2-digit',hour12:true}):'In progress',
                  hrs>0?parseFloat(hrs.toFixed(2)):'',
                  a.isVO?'Bill as VO':'Included in Contract',
                  w?.costRate||0, hrs>0?parseFloat((hrs*(w?.costRate||0)).toFixed(2)):'',
                  w?.billingRate||0, (a.isVO&&hrs>0)?parseFloat((hrs*(w?.billingRate||0)).toFixed(2)):'']);
              });
              const ws=XLSX.utils.aoa_to_sheet(rows);
              ws['!cols']=[{wch:22},{wch:14},{wch:14},{wch:28},{wch:12},{wch:10},{wch:12},{wch:8},{wch:20},{wch:10},{wch:10},{wch:10},{wch:10}];
              XLSX.utils.book_append_sheet(wb,ws,`Attendance ${month}`);
              XLSX.writeFile(wb,`Attendance_${month}_${new Date().toISOString().slice(0,10)}.xlsx`);
            }}>
              <Download size={12}/>Export Excel
            </Btn>
          </div>

          {siteWorkers.filter(w=>!selWorker||w.id===selWorker).map(w=>{
            const recs=monthRecs.filter(a=>a.workerId===w.id).sort((a,b)=>a.date.localeCompare(b.date));
            if(recs.length===0&&selWorker)return null;
            if(recs.length===0&&!selWorker)return null;
            const totHrs=recs.filter(a=>a.totalHours).reduce((s,a)=>s+a.totalHours,0);
            return (
              <div key={w.id} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,overflow:'hidden',boxShadow:T.shadow}}>
                <div style={{padding:'14px 20px',background:T.bg,borderBottom:`1px solid ${T.borderLight}`,display:'flex',alignItems:'center',justifyContent:'space-between'}}>
                  <div style={{display:'flex',alignItems:'center',gap:10}}>
                    <Users size={14} style={{color:T.accent}}/>
                    <span style={{fontSize:14,fontWeight:600,color:T.text}}>{w.name}</span>
                    <span style={{fontSize:12,color:T.muted}}>{recs.length} days</span>
                  </div>
                  <div style={{display:'flex',gap:12,fontSize:12}}>
                    <span style={{color:T.accent,fontWeight:700}}>{fmtHours(totHrs)} total</span>
                    <span style={{color:T.info}}>{fmtSGD(totHrs*(w.billingRate||0))} billed</span>
                  </div>
                </div>
                <div style={{overflowX:'auto'}}>
                  <table style={{width:'100%',fontSize:13,borderCollapse:'collapse'}}>
                    <thead>
                      <tr style={{color:T.dim,fontSize:11,fontWeight:600,textTransform:'uppercase',letterSpacing:'0.05em',background:T.bg}}>
                        {['Date','Check In','Location In','Check Out','Location Out','Hours','Billing',''].map(h=>(
                          <th key={h} style={{textAlign:'left',padding:'9px 14px',whiteSpace:'nowrap'}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {recs.map(a=>(
                        <tr key={a.id} style={{borderTop:`1px solid ${T.borderLight}`}}>
                          <td style={{padding:'9px 14px',color:T.text,fontWeight:600,whiteSpace:'nowrap'}}>
                            {new Date(a.date+'T00:00:00').toLocaleDateString('en-SG',{weekday:'short',day:'2-digit',month:'short'})}
                          </td>
                          <td style={{padding:'9px 14px',color:T.success,fontWeight:600}}>{fmtTime(a.checkIn?.time)}</td>
                          <td style={{padding:'9px 14px',color:T.dim,fontSize:11,maxWidth:120,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
                            {a.checkIn?.addr||'—'}
                          </td>
                          <td style={{padding:'9px 14px',color:T.danger,fontWeight:600}}>{a.checkOut?fmtTime(a.checkOut.time):'—'}</td>
                          <td style={{padding:'9px 14px',color:T.dim,fontSize:11,maxWidth:120,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
                            {a.checkOut?.addr||'—'}
                          </td>
                          <td style={{padding:'9px 14px',color:T.accent,fontWeight:700}}>{a.totalHours?fmtHours(a.totalHours):'In progress'}</td>
                          <td style={{padding:'9px 14px',color:T.info,fontWeight:600}}>{a.totalHours?fmtSGD(a.totalHours*(w.billingRate||0)):'—'}</td>
                          <td style={{padding:'9px 14px'}}>
                            <button onClick={()=>{
                                const upd=attendance.map(x=>x.id===a.id?{...x,isVO:!x.isVO}:x);
                                setAttendance(upd);saveS('attendance',upd);
                              }}
                              title={a.isVO?'Click to mark as Included in Contract':'Click to mark as Bill to Client (VO)'}
                              style={{background:a.isVO?T.warningLight:T.successLight,border:`1px solid ${a.isVO?T.warning+'40':T.success+'40'}`,borderRadius:6,padding:'3px 9px',cursor:'pointer',fontFamily:'inherit',fontSize:11,fontWeight:600,color:a.isVO?T.warning:T.success}}>
                              {a.isVO?'Bill as VO':'In Contract'}
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            );
          })}
          {monthRecs.length===0&&(
            <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:'40px',textAlign:'center',color:T.dim,boxShadow:T.shadow}}>No attendance records for this month</div>
          )}
        </div>
      )}

      {/* Payroll */}
      {view==='payroll'&&(
        <div style={{display:'flex',flexDirection:'column',gap:14}}>
          <div style={{display:'flex',gap:10,alignItems:'center'}}>
            <input type="month" value={month} onChange={e=>setMonth(e.target.value)}
              style={{...iStyle,width:'auto',minWidth:160}}/>
            <div style={{fontSize:13,color:T.muted}}>
              {new Date(month+'-01').toLocaleDateString('en-SG',{month:'long',year:'numeric'})} payroll
            </div>
          </div>

          <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,overflow:'hidden',boxShadow:T.shadow}}>
            <table style={{width:'100%',fontSize:13,borderCollapse:'collapse'}}>
              <thead>
                <tr style={{color:T.dim,fontSize:11,fontWeight:600,textTransform:'uppercase',letterSpacing:'0.05em',background:T.bg}}>
                  {['Worker','Project','Days','Hours','Salary','Hour Pay','Total Pay','Billing to Client',''].map(h=>(
                    <th key={h} style={{textAlign:['Days','Hours','Salary','Hour Pay','Total Pay','Billing to Client'].includes(h)?'right':'left',padding:'10px 14px',whiteSpace:'nowrap'}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {siteWorkers.map(w=>{
                  const{hrs,days}=workerMonth(w.id);
                  const hrPay=hrs*(w.costRate||0);
                  const totalPay=(w.monthlySalary||0)+hrPay;
                  const billing=hrs*(w.billingRate||0);
                  const proj=projects.find(p=>p.id===w.projectId);
                  return (
                    <tr key={w.id} style={{borderTop:`1px solid ${T.borderLight}`}}>
                      <td style={{padding:'11px 14px',color:T.text,fontWeight:600}}>{w.name}</td>
                      <td style={{padding:'11px 14px',color:T.muted,fontSize:12,maxWidth:130,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{proj?.name||'—'}</td>
                      <td style={{padding:'11px 14px',textAlign:'right',color:T.text}}>{days}</td>
                      <td style={{padding:'11px 14px',textAlign:'right',color:T.text}}>{hrs.toFixed(1)}h</td>
                      <td style={{padding:'11px 14px',textAlign:'right',color:T.muted}}>{fmtSGD(w.monthlySalary||0)}</td>
                      <td style={{padding:'11px 14px',textAlign:'right',color:T.muted}}>{fmtSGD(hrPay)}</td>
                      <td style={{padding:'11px 14px',textAlign:'right',color:T.accent,fontWeight:700}}>{fmtSGD(totalPay)}</td>
                      <td style={{padding:'11px 14px',textAlign:'right',color:T.info,fontWeight:700}}>{fmtSGD(billing)}</td>
                      <td style={{padding:'11px 14px'}}>
                        <Btn size="sm" onClick={()=>generatePayslip(w,invoices,setInvoices,acctSettings)}>Payslip</Btn>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr style={{borderTop:`2px solid ${T.borderLight}`,background:T.bg}}>
                  <td colSpan={4} style={{padding:'11px 14px',fontWeight:700,color:T.text}}>Totals</td>
                  <td style={{padding:'11px 14px',textAlign:'right',fontWeight:700,color:T.muted}}>
                    {fmtSGD(siteWorkers.reduce((s,w)=>s+(w.monthlySalary||0),0))}
                  </td>
                  <td style={{padding:'11px 14px',textAlign:'right',fontWeight:700,color:T.muted}}>
                    {fmtSGD(siteWorkers.reduce((s,w)=>{const{hrs}=workerMonth(w.id);return s+hrs*(w.costRate||0);},0))}
                  </td>
                  <td style={{padding:'11px 14px',textAlign:'right',fontWeight:700,color:T.accent}}>
                    {fmtSGD(siteWorkers.reduce((s,w)=>{const{hrs}=workerMonth(w.id);return s+(w.monthlySalary||0)+hrs*(w.costRate||0);},0))}
                  </td>
                  <td style={{padding:'11px 14px',textAlign:'right',fontWeight:700,color:T.info}}>
                    {fmtSGD(siteWorkers.reduce((s,w)=>{const{hrs}=workerMonth(w.id);return s+hrs*(w.billingRate||0);},0))}
                  </td>
                  <td/>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}

      {/* Claims review */}
      {view==='claims'&&(
        <div style={{display:'flex',flexDirection:'column',gap:12}}>
          {claims.length===0&&<div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:'36px',textAlign:'center',color:T.dim,boxShadow:T.shadow,fontSize:13}}>No claims submitted yet</div>}
          {[...claims].sort((a,b)=>new Date(b.submittedAt)-new Date(a.submittedAt)).map(c=>{
            const worker=siteWorkers.find(w=>w.id===c.workerId);
            const proj=projects.find(p=>p.id===c.projectId);
            return (
              <div key={c.id} style={{background:T.card,border:`1px solid ${c.status==='Pending'?T.warning+'40':T.borderLight}`,borderRadius:14,padding:'16px 18px',boxShadow:T.shadow}}>
                <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:8}}>
                  <div>
                    <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:4}}>
                      <span style={{fontSize:14,fontWeight:700,color:T.text}}>{c.type}</span>
                      <Badge color={c.status==='Approved'?T.success:c.status==='Rejected'?T.danger:T.warning} sm>{c.status}</Badge>
                    </div>
                    <div style={{fontSize:12,color:T.muted}}>{worker?.name} . {c.date}</div>
                    {proj&&<div style={{fontSize:12,color:T.dim}}>{proj.name}</div>}
                    {c.notes&&<div style={{fontSize:12,color:T.muted,marginTop:3,fontStyle:'italic'}}>{c.notes}</div>}
                  </div>
                  <div style={{textAlign:'right'}}>
                    <div style={{fontSize:18,fontWeight:700,color:T.text}}>S${parseFloat(c.amount).toFixed(2)}</div>
                  </div>
                </div>
                {c.photo&&<img src={c.photo} alt="receipt" style={{width:'100%',maxHeight:120,objectFit:'cover',borderRadius:8,marginBottom:10,border:`1px solid ${T.borderLight}`}}/>}
                {c.status==='Pending'&&(
                  <div style={{display:'flex',gap:8}}>
                    <Btn size="sm" onClick={()=>{const u=claims.map(x=>x.id===c.id?{...x,status:'Approved'}:x);setClaims(u);saveS('workerClaims',u);}}>
                      <CheckCircle size={11}/>Approve
                    </Btn>
                    <Btn variant="danger" size="sm" onClick={()=>{const u=claims.map(x=>x.id===c.id?{...x,status:'Rejected'}:x);setClaims(u);saveS('workerClaims',u);}}>
                      Reject
                    </Btn>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {/* Add/Edit Worker Modal */}
      {(modal==='new'||modal==='edit')&&(
        <Modal title={modal==='new'?'Add Site Worker':'Edit Worker'} onClose={()=>setModal(null)} wide>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:14}}>
            <div style={{gridColumn:'1/-1'}}><Field label="Full Name *" value={form.name} onChange={ff('name')} placeholder="e.g. Ahmad Bin Salleh"/></div>
            <Field label="Phone" value={form.phone} onChange={ff('phone')} placeholder="+65 9111 2222"/>
            <Field label="NRIC / FIN / Pass No." value={form.nric} onChange={ff('nric')} placeholder="FT-WP-001"/>
            <Field label="Nationality" value={form.nationality} onChange={ff('nationality')} placeholder="Malaysian"/>
            <div>
              <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>4-Digit PIN *</label>
              <input type="number" value={form.pin} onChange={e=>ff('pin')(e.target.value.slice(0,4))} placeholder="e.g. 1234" maxLength={4} style={{...iStyle}} />
              <div style={{fontSize:11,color:T.dim,marginTop:4}}>Worker uses this PIN to log in at Check In/Out page</div>
            </div>

            <div style={{gridColumn:'1/-1',background:T.bg,borderRadius:12,padding:14,border:`1px solid ${T.borderLight}`}}>
              <div style={{fontSize:12,fontWeight:600,color:T.muted,marginBottom:10}}>Assigned Sites (Projects)</div>
              <div style={{display:'flex',flexDirection:'column',gap:6,maxHeight:160,overflowY:'auto'}}>
                {projects.filter(p=>p.status!=='Cancelled'&&!p.archived).map(p=>{
                  const has=(form.assignedProjects||[]).includes(p.id);
                  return(
                    <label key={p.id} style={{display:'flex',alignItems:'center',gap:10,cursor:'pointer',padding:'8px 12px',background:has?T.accentLight:'transparent',borderRadius:8,border:`1px solid ${has?T.accent+'30':'transparent'}`,transition:'all 0.12s'}}>
                      <input type="checkbox" checked={has} onChange={()=>{
                        const cur=form.assignedProjects||[];
                        const next=has?cur.filter(x=>x!==p.id):[...cur,p.id];
                        setForm(prev=>({...prev,assignedProjects:next,projectId:next[0]||''}));
                      }} style={{accentColor:T.accent}}/>
                      <div>
                        <div style={{fontSize:13,fontWeight:600,color:T.text}}>{p.name}</div>
                        <div style={{fontSize:11,color:T.muted}}>{p.client}</div>
                      </div>
                    </label>
                  );
                })}
              </div>
            </div>            <div style={{gridColumn:'1/-1',background:T.bg,borderRadius:12,padding:14,border:`1px solid ${T.borderLight}`}}>
              <div style={{fontSize:12,fontWeight:600,color:T.muted,marginBottom:10}}>Pay & Billing Rates</div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:10}}>
                <Field label="Cost Rate (S$/hr)" type="number" value={form.costRate} onChange={ff('costRate')} placeholder="12"/>
                <Field label="Billing Rate (S$/hr)" type="number" value={form.billingRate} onChange={ff('billingRate')} placeholder="18"/>
                <Field label="Monthly Salary (S$)" type="number" value={form.monthlySalary} onChange={ff('monthlySalary')} placeholder="1800"/>
              </div>
            </div>

            <div style={{gridColumn:'1/-1',background:T.bg,borderRadius:12,padding:14,border:`1px solid ${T.borderLight}`}}>
              <div style={{fontSize:12,fontWeight:600,color:T.muted,marginBottom:10}}>Work Pass</div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:10}}>
                <Field label="Work Pass No." value={form.workPassNo} onChange={ff('workPassNo')} placeholder="WP2024001"/>
                <Field label="Expiry Date" type="date" value={form.workPassExpiry} onChange={ff('workPassExpiry')}/>
              </div>
            </div>

            <div style={{gridColumn:'1/-1',background:T.bg,borderRadius:12,padding:14,border:`1px solid ${T.borderLight}`}}>
              <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:10}}>
                <div style={{fontSize:12,fontWeight:600,color:T.muted}}>Certificates & Licences</div>
                <Btn variant="secondary" size="sm" onClick={addCert}><Plus size={11}/>Add</Btn>
              </div>
              {form.certificates.length===0&&<div style={{fontSize:12,color:T.dim,fontStyle:'italic'}}>No certificates added</div>}
              {form.certificates.map((c,i)=>(
                <div key={c.id} style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(130px,1fr))',gap:8,marginBottom:8,alignItems:'end'}}>
                  <Field label={i===0?'Certificate Name':''} value={c.name} onChange={v=>updateCert(i,'name',v)} placeholder="CSOC Card"/>
                  <Field label={i===0?'Issue Date':''} type="date" value={c.issueDate} onChange={v=>updateCert(i,'issueDate',v)}/>
                  <Field label={i===0?'Expiry Date':''} type="date" value={c.expiryDate} onChange={v=>updateCert(i,'expiryDate',v)}/>
                  <button onClick={()=>removeCert(i)} style={{background:T.dangerLight,border:'none',borderRadius:8,padding:'8px 10px',cursor:'pointer',color:T.danger,display:'flex',alignItems:'center',marginTop:i===0?22:0}}>
                    <X size={13}/>
                  </button>
                </div>
              ))}
            </div>

            <Field label="Status" value={form.status} onChange={ff('status')} as="select" options={[{v:'Active',l:'Active'},{v:'Inactive',l:'Inactive'}]}/>
          </div>
          <div style={{display:'flex',justifyContent:'flex-end',gap:10,marginTop:22}}>
            <Btn variant="secondary" onClick={()=>setModal(null)}>Cancel</Btn>
            <Btn onClick={saveWorker} disabled={!form.name}>{modal==='new'?'Add Worker':'Save Changes'}</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

function CompanyAccounts({projects,invoices,payments,acctSettings,setAcctSettings}){
  const [tab,setTab]=useState('overview');
  const [settings,setSettings]=useState(acctSettings);
  const [generating,setGenerating]=useState(false);
  const [uploadMonth,setUploadMonth]=useState(null);
  const [reconcileMonth,setReconcileMonth]=useState(null);
  const [reconChecked,setReconChecked]=useState({});
  const bankFileRef=useRef();
  const sf=k=>v=>setSettings(p=>{const n={...p,[k]:v};setAcctSettings(n);saveS('acctSettings',n);return n;});

  const today=new Date();
  const fyEnd=new Date(today.getFullYear(),settings.fyEndMonth-1,settings.fyEndDay);
  if(fyEnd<today) fyEnd.setFullYear(today.getFullYear()+1);
  const fyStart=new Date(fyEnd);
  fyStart.setFullYear(fyEnd.getFullYear()-1);
  fyStart.setDate(fyStart.getDate()+1);

  const fyLabel=`${fyStart.getFullYear()}/${String(fyEnd.getFullYear()).slice(2)}`;
  const yaYear=fyEnd.getFullYear()+1;


  const eciDeadline=new Date(fyEnd.getFullYear(),5,30);
  const formDeadline=new Date(fyEnd.getFullYear(),10,30);
  const daysToECI=Math.ceil((eciDeadline-today)/864e5);
  const daysToForm=Math.ceil((formDeadline-today)/864e5);


  const inFY=d=>{if(!d)return false;const dt=new Date(d);return dt>=fyStart&&dt<=fyEnd;};

  const fyRevenue=payments.filter(p=>p.status==='Received'&&inFY(p.date)).reduce((s,p)=>s+p.amount,0);
  const fyCOS=invoices.filter(i=>i.status==='Paid'&&inFY(i.invoiceDate)).reduce((s,i)=>s+i.total,0);
  const fyCommissions=projects
    .filter(p=>p.commissionPaid&&inFY(p.commissionPaidAt))
    .reduce((p2,p)=>{const{dComm,pmComm}=calcComm(p,invoices);return p2+dComm+pmComm;},0);
  const fyStaff=parseFloat(settings.staffCosts)||0;
  const fyAdmin=parseFloat(settings.adminFees)||0;
  const fyBank=parseFloat(settings.bankCharges)||0;
  const fyOther=parseFloat(settings.otherExpenses)||0;
  const fyOtherExp=fyCommissions+fyStaff+fyAdmin+fyBank+fyOther;
  const fyGrossProfit=fyRevenue-fyCOS;
  const fyNetProfit=fyGrossProfit-fyOtherExp;
  const priorLoss=parseFloat(settings.priorYearLoss)||0;
  const chargeableIncome=Math.max(0,fyNetProfit-priorLoss);
  const lossAbsorbed=Math.min(priorLoss,Math.max(0,fyNetProfit));
  const lossCF=priorLoss-lossAbsorbed;
  const taxPayable=chargeableIncome>0?chargeableIncome*0.17:0;


  const months=[];
  for(let m=0;m<12;m++){
    const d=new Date(fyStart.getFullYear(),fyStart.getMonth()+m,1);
    const key=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    const label=d.toLocaleDateString('en-SG',{month:'short',year:'numeric'});
    const recv=payments.filter(p=>p.status==='Received'&&p.date?.startsWith(key)).reduce((s,p)=>s+p.amount,0);
    const paid=invoices.filter(i=>i.status==='Paid'&&i.invoiceDate?.startsWith(key)).reduce((s,i)=>s+i.total,0);
    const hasStatement=!!(settings.bankStatements||{})[key];
    months.push({key,label,recv,paid,net:recv-paid,hasStatement});
  }


  const revItems = payments.filter(p=>p.status==='Received'&&inFY(p.date)).map(p=>{
    const proj=projects.find(x=>x.id===p.projectId);
    return {desc:`${proj?.client||'Client'} – ${proj?.name||''} (${p.type} Payment)`,amt:p.amount};
  });
  const cosItems = invoices.filter(i=>i.status==='Paid'&&inFY(i.invoiceDate)).map(i=>{
    const proj=projects.find(x=>x.id===i.projectId);
    return {desc:`${i.supplier} – ${proj?.name||''} (${i.category})`,amt:i.total};
  });
  const commItems = projects.filter(p=>p.commissionPaid&&inFY(p.commissionPaidAt)).map(p=>{
    const{dComm,pmComm}=calcComm(p,invoices);
    return {desc:`Commission – ${p.name} (${p.designer}/${p.pm})`,amt:dComm+pmComm};
  });


  const generateExcel=()=>{
    setGenerating(true);
    try{
      const wb=XLSX.utils.book_new();
      const co=settings.companyName||'TDI WORKSPACE PTE. LTD.';
      const uen=settings.uen||'196800306E';
      const dir=settings.director||'Ng Zhi Wei Kelvin';
      const fyS=fyStart.toLocaleDateString('en-SG',{day:'2-digit',month:'long',year:'numeric'});
      const fyE=fyEnd.toLocaleDateString('en-SG',{day:'2-digit',month:'long',year:'numeric'});
      const fyFull=`${fyS} – ${fyE}`;
      const fmt2=(n)=>n?.toFixed(2)||'0.00';


      const s1=[
        [co],['IRAS CORPORATE INCOME TAX FILING PACKAGE'],
        [`Year of Assessment ${yaYear}  |  Financial Year Ended ${fyE}`],
        ['Form C-S Lite  (Revenue ≤ SGD 200,000  |  Simple Tax Adjustments Only)'],[''],
        ['SECTION A — COMPANY PARTICULARS'],[''],
        ['','Company Name',co],
        ['','UEN / Reg No.',uen],
        ['','Director',dir],
        ['','Bank Account',settings.bankAccount||''],
        ['','Financial Year End',`${settings.fyEndDay} ${new Date(2000,settings.fyEndMonth-1,1).toLocaleDateString('en-SG',{month:'long'})}`],
        ['','Year of Assessment',`${yaYear}`],
        ['','Form Type','Form C-S Lite   (Revenue < SGD 200,000)'],
        ['','ECI Filing Deadline',`30 June ${fyEnd.getFullYear()}`],
        ['','Form C-S Lite Deadline',`30 November ${fyEnd.getFullYear()}`],
        ['','Filing Method','myTax Portal — https://mytax.iras.gov.sg'],[''],
        ['SECTION B — FORM C-S LITE ELIGIBILITY CHECK'],[''],
        ['✓','Incorporated in Singapore','YES'],
        ['✓',`Revenue ≤ SGD 200,000 for the YA`,`YES  (Revenue = SGD ${fmt2(fyRevenue)})`],
        ['✓','Derives income taxable at 17% only','YES'],
        ['✓','No capital allowance claims','YES'],
        ['✓','ELIGIBLE FOR FORM C-S LITE','CONFIRMED'],[''],
        ['SECTION C — KEY FINANCIALS FOR IRAS SUBMISSION (4-Line Statement)'],
        ['','Item',`FY ${fyLabel}  (SGD)`,'Prior FY  (SGD)'],
        ['','Line 1: Revenue (Turnover)',fyRevenue,''],
        ['','Line 2: Cost of Sales',fyCOS,''],
        ['','Gross Profit / (Loss)',fyGrossProfit,''],
        ['','Line 3: Other Allowable Expenses',fyOtherExp,''],
        ['','Line 4: Net Profit / (Loss)',fyNetProfit,''],[''],
        ['SECTION D — FILING DEADLINES'],
        ['','ECI Deadline',`30 June ${fyEnd.getFullYear()}`,'File even if ECI = NIL'],
        ['','Form C-S Lite Deadline',`30 November ${fyEnd.getFullYear()}`,'DO NOT miss'],
      ];
      const ws1=XLSX.utils.aoa_to_sheet(s1);
      ws1['!cols']=[{wch:4},{wch:45},{wch:25},{wch:20}];
      XLSX.utils.book_append_sheet(wb,ws1,'1. Filing Checklist');


      const s2=[
        [co],['TAX COMPUTATION — YEAR OF ASSESSMENT '+yaYear],
        [`Financial Year: ${fyFull}  |  Form C-S Lite  |  Corporate Tax Rate: 17%`],[''],
        ['Description','Reference',`FY ${fyLabel}  (SGD)`,'Notes'],[''],
        ['PART 1 — NET PROFIT / (LOSS) PER FINANCIAL STATEMENTS'],
        ['Revenue (Turnover)','Line 1',fyRevenue,''],
        ['Less: Cost of Sales','Line 2',-fyCOS,''],
        ['Gross Profit / (Loss)','',fyGrossProfit,''],
        ['Less: Commissions Paid','L3(a)',-fyCommissions,''],
        ['Less: Staff Costs / Wages','L3(b)',-fyStaff,''],
        ['Less: Admin & Professional Fees','L3(c)',-fyAdmin,''],
        ['Less: Bank Charges & Fees','L3(d)',-fyBank,''],
        ['Less: Other Expenses','L3(e)',-fyOther,''],
        ['Net Profit / (Loss) per Financial Statements','Line 4',fyNetProfit,''],[''],
        ['PART 2 — TAX ADJUSTMENTS'],
        ['Add: Private / non-business expenses','T1',0,'Confirm if any personal expenses included'],
        ['Add: Capital expenditure (non-deductible)','T2',0,''],
        ['Add: Penalties / fines','T3',0,''],
        ['Less: Capital allowances (S19/19A ITA)','T4',0,'No qualifying plant & machinery'],
        ['TOTAL TAX ADJUSTMENTS','',0,''],[''],
        ['PART 3 — ADJUSTED PROFIT / (LOSS)'],
        ['Net Profit / (Loss) per Accounts','',fyNetProfit,''],
        ['Total Tax Adjustments','',0,''],
        ['ADJUSTED PROFIT / (LOSS) FOR THE YEAR','',fyNetProfit,''],[''],
        ['PART 4 — UNABSORBED LOSSES SET-OFF'],
        ['Accumulated unabsorbed loss b/f','',-priorLoss,'Per prior year tax computation'],
        ['Current year adjusted profit / (loss)','',fyNetProfit,''],
        ['Losses set off against current year profit','',-lossAbsorbed,'Loss absorbed up to current year profit only'],
        ['CHARGEABLE INCOME AFTER LOSS SET-OFF','',chargeableIncome,''],[''],
        ['PART 5 — PARTIAL TAX EXEMPTION  (Only if Chargeable Income > 0)'],
        ['Chargeable Income (from Part 4)','',chargeableIncome,''],
        ['First SGD 10,000 @ 75% exempt (taxable 25%)','',chargeableIncome>0?Math.min(chargeableIncome,10000)*0.25:0,''],
        ['Next SGD 190,000 @ 50% exempt (taxable 50%)','',chargeableIncome>10000?Math.min(chargeableIncome-10000,190000)*0.5:0,''],
        ['TAXABLE CHARGEABLE INCOME','',chargeableIncome>0?Math.min(chargeableIncome,10000)*0.25+Math.max(0,Math.min(chargeableIncome-10000,190000))*0.5:0,''],[''],
        ['PART 6 — INCOME TAX PAYABLE'],
        ['Corporate Tax Rate','',0.17,'Singapore headline rate'],
        ['ESTIMATED INCOME TAX PAYABLE','',taxPayable,''],[''],
        ['PART 7 — UNABSORBED LOSSES CARRIED FORWARD'],
        ['Accumulated unabsorbed loss b/f','',-priorLoss,''],
        ['Current year adjusted profit / (loss)','',fyNetProfit,''],
        ['Less: Losses absorbed by current year income','',-lossAbsorbed,''],
        ['TOTAL UNABSORBED LOSS CARRIED FORWARD','',-lossCF,''],[''],
        ['PART 8 — TAX SUMMARY'],
        ['Item','Line',`FY ${fyLabel}`,''],
        ['Revenue (Turnover)','Line 1',fyRevenue,''],
        ['Cost of Sales','Line 2',fyCOS,''],
        ['Other Operating Expenses','Line 3',fyOtherExp,''],
        ['Net Profit / (Loss)','Line 4',fyNetProfit,''],
        ['Total Tax Adjustments','',0,''],
        ['Adjusted Profit / (Loss)','',fyNetProfit,''],
        ['Less: Unabsorbed Loss Set-Off','',-lossAbsorbed,''],
        ['Chargeable Income','',chargeableIncome,''],
        ['Tax Payable','',taxPayable,''],
        ['Unabsorbed Loss c/f','',-lossCF,''],
      ];
      const ws2=XLSX.utils.aoa_to_sheet(s2);
      ws2['!cols']=[{wch:50},{wch:12},{wch:22},{wch:45}];
      XLSX.utils.book_append_sheet(wb,ws2,'2. Tax Computation');


      const s3=[
        [co],['4-LINE STATEMENT OF ACCOUNTS'],
        [`Year of Assessment ${yaYear}  |  Financial Year ${fyFull}`],
        ['Prepared for IRAS Income Tax Filing (Form C-S Lite)  |  All amounts in SGD'],[''],
        ['Description / Transaction',`FY ${fyLabel} (SGD)  [YA ${yaYear}]`,'Reference'],[''],
        ['LINE 1 — REVENUE (TURNOVER)'],
        ...revItems.map(r=>[`  ${r.desc}`,r.amt,'']),
        [''],['TOTAL REVENUE',fyRevenue,'Note 1'],[''],
        ['LINE 2 — COST OF SALES  (DIRECT PROJECT COSTS)'],
        ...cosItems.map(c=>[`  ${c.desc}`,c.amt,'']),
        [''],['TOTAL COST OF SALES',fyCOS,'Note 2'],
        ['GROSS PROFIT / (LOSS)  [Line 1 – Line 2]',fyGrossProfit,''],[''],
        ['LINE 3 — OTHER ALLOWABLE BUSINESS EXPENSES'],
        ...commItems.map(c=>[`  ${c.desc}`,c.amt,'']),
        fyStaff>0?[`  Staff Costs / Wages`,fyStaff,'']:null,
        fyAdmin>0?[`  Admin & Professional Fees`,fyAdmin,'']:null,
        fyBank>0?[`  Bank Service Charges`,fyBank,'']:null,
        fyOther>0?[`  Other Operating Expenses`,fyOther,'']:null,
        [''],['TOTAL OTHER EXPENSES',fyOtherExp,'Note 3'],[''],
        ['LINE 4 — NET PROFIT / (LOSS)'],
        ['NET PROFIT / (LOSS)  [Gross Profit – Line 3]',fyNetProfit,''],[''],[''],
        ['I certify that the above 4-Line Statement is prepared from the accounting records of'],
        [`${co} and is, to the best of my knowledge, true, correct and complete.`],[''],
        ['Signature:','___________________________'],
        ['Name:  '+dir+' (Director)'],
        ['UEN:   '+uen],
        ['Date:  ___________________________'],
        ['Company: '+co],
      ].filter(Boolean);
      const ws3=XLSX.utils.aoa_to_sheet(s3);
      ws3['!cols']=[{wch:60},{wch:28},{wch:12}];
      XLSX.utils.book_append_sheet(wb,ws3,'3. 4-Line Statement');


      const s4=[
        [co],['UNAUDITED PROFIT & LOSS STATEMENT'],
        [`Financial Year ${fyFull}  |  Cash basis  |  Unaudited`],
        ['Prepared from bank records. Unaudited — subject to accountant review before IRAS submission.'],[''],
        ['Description','Note',`FY ${fyLabel}  SGD`,'Source / Remarks'],[''],
        ['REVENUE (TURNOVER)'],
        ...revItems.map(r=>[`    ${r.desc}`,'',r.amt,'']),
        ['TOTAL REVENUE','1',fyRevenue,''],[''],
        ['COST OF SALES  (DIRECT PROJECT COSTS)'],
        ...cosItems.map(c=>[`    ${c.desc}`,'',c.amt,'']),
        ['TOTAL COST OF SALES','2',fyCOS,''],
        ['GROSS PROFIT / (LOSS)  [Revenue – Cost of Sales]','',fyGrossProfit,''],[''],
        ['OTHER OPERATING EXPENSES'],
        ...commItems.map(c=>[`    ${c.desc}`,'',c.amt,'']),
        fyStaff>0?[`    Staff Costs / Wages`,'',fyStaff,'']:null,
        fyAdmin>0?[`    Admin & Professional Fees`,'',fyAdmin,'']:null,
        fyBank>0?[`    Bank Service Charges`,'',fyBank,'']:null,
        fyOther>0?[`    Other Operating Expenses`,'',fyOther,'']:null,
        ['TOTAL OTHER OPERATING EXPENSES','3',fyOtherExp,''],
        ['NET PROFIT / (LOSS) BEFORE TAX','',fyNetProfit,''],[''],
        ['RETAINED EARNINGS / (ACCUMULATED DEFICIT) ROLL-FORWARD'],
        ['Accumulated Deficit – Opening','',-priorLoss,'Per prior year computation'],
        ['Net Profit / (Loss) for the Year','',fyNetProfit,''],
        ['Accumulated Deficit – Closing','',-lossCF,'Carried forward to next FY'],[''],
        ['NOTES TO FINANCIAL STATEMENTS'],
        ['Note 1:',`Revenue represents cash receipts from clients (cash basis). FY ${fyLabel}.`],
        ['Note 2:','Direct project costs include payments to subcontractors, suppliers, and trade payables.'],
        ['Note 3:','Commissions, staff costs, admin fees and bank charges per records.'],
        ['Note 4:','Prepared on cash basis. Unaudited. Qualified accountant should review before IRAS submission.'],
      ].filter(Boolean);
      const ws4=XLSX.utils.aoa_to_sheet(s4);
      ws4['!cols']=[{wch:60},{wch:8},{wch:22},{wch:45}];
      XLSX.utils.book_append_sheet(wb,ws4,'4. Profit & Loss');


      const cashBalance=0;
      const tradePayables=0;
      const otherPayables=0;
      const paidUpCapital=1;
      const accDeficit=-(lossCF);
      const totalAssets=cashBalance;
      const totalLiabilities=tradePayables+otherPayables;
      const totalEquity=paidUpCapital+accDeficit+fyNetProfit;
      const s5=[
        [co],['UNAUDITED STATEMENT OF FINANCIAL POSITION'],
        [`As at ${fyEnd.toLocaleDateString('en-SG',{day:'2-digit',month:'long',year:'numeric'})}  |  Unaudited`],
        ['Yellow = estimated — confirm actual amounts before submission'],[''],
        ['Description','Note',`${fyEnd.toLocaleDateString('en-SG',{day:'2-digit',month:'short',year:'numeric'})}  SGD`,'Source / Notes'],[''],
        ['CURRENT ASSETS'],
        ['Cash & Cash Equivalents – DBS Account','5','[ENTER CLOSING BANK BALANCE]','Confirmed: per DBS statements'],
        ['Trade Debtors / Receivables','5',0,'Confirm with client files'],
        ['TOTAL CURRENT ASSETS','5','[ENTER TOTAL]',''],[''],
        ['NON-CURRENT ASSETS'],
        ['Property, Plant & Equipment (net)','',0,'Confirm if any equipment held'],
        ['TOTAL NON-CURRENT ASSETS','',0,''],
        ['TOTAL ASSETS','','[ENTER TOTAL]',''],[''],
        ['CURRENT LIABILITIES'],
        ['Trade Payables (subcontractors & suppliers)','6','[ENTER OUTSTANDING PAYABLES]','Confirm from outstanding supplier invoices'],
        ['Other Payables & Accrued Expenses','6','[ENTER ACCRUALS]','Confirm actual accruals at year end'],
        ['Income Tax Payable','',(taxPayable>0?taxPayable:0),taxPayable>0?'Per Tax Computation':'TAX PAYABLE = NIL'],
        ['TOTAL CURRENT LIABILITIES','6','[ENTER TOTAL]',''],[''],
        ['TOTAL LIABILITIES','','[ENTER TOTAL]',''],[''],
        ["SHAREHOLDERS' EQUITY / (DEFICIT)"],
        ['Issued & Paid-Up Capital (1 ordinary share)','4',paidUpCapital,'SGD 1.00 — 1 ordinary share. Director: '+dir],
        ['Retained Earnings / (Accumulated Deficit) – b/f','',-priorLoss,'Per prior year P&L'],
        ['Net Profit / (Loss) for the Year','',fyNetProfit,'Per P&L Statement'],
        ["TOTAL SHAREHOLDERS' EQUITY / (DEFICIT)",'7',-lossCF+1,''],[''],
        ['TOTAL LIABILITIES & SHAREHOLDERS EQUITY','','[ENTER TOTAL]','Should equal Total Assets'],[''],
        ['Balance Check: must = zero','','[ENTER FORMULA: Total Assets - Total Liabilities - Total Equity]','If non-zero: update payables / receivables'],[''],
        ['NOTES TO THE FINANCIAL STATEMENTS'],
        ['Note 1:',`Prepared on cash basis from bank account records for ${fyFull}. Unaudited.`],
        ['Note 4:','Issued and paid-up capital: SGD 1.00. Director: '+dir+'. UEN: '+uen],
        ['Note 5:','Cash balance per DBS bank statement — confirm closing balance at year end.'],
        ['Note 6:','Trade and other payables — confirm from outstanding supplier invoices at year end.'],
        ['Note 7:','Accumulated deficit position. Directors to assess going-concern position.'],
      ];
      const ws5=XLSX.utils.aoa_to_sheet(s5);
      ws5['!cols']=[{wch:55},{wch:6},{wch:25},{wch:55}];
      XLSX.utils.book_append_sheet(wb,ws5,'5. Balance Sheet');


      const s6=[
        [co],['(Incorporated in the Republic of Singapore)'],[''],
        ["DIRECTOR'S DECLARATION"],
        [`Financial Year: ${fyFull}`],[''],[''],
        [`I, ${dir}, being a director of ${co} (the "Company"), do hereby declare as follows:`],[''],
        ['1.','In my opinion, the accompanying financial statements are drawn up so as to give a true and fair'],
        ['','view of the state of affairs of the Company as at '+fyEnd.toLocaleDateString('en-SG',{day:'2-digit',month:'long',year:'numeric'})],
        ['','and of the results of the Company for the financial year ended on that date.'],[''],
        ['2.','At the date of this declaration, there are reasonable grounds to believe that the Company'],
        ['','will be able to pay its debts as and when they fall due.'],[''],
        ['3.',`These financial statements have been prepared on a cash basis from DBS Bank account records`],
        ['','and are unaudited. Submitted for income tax filing with IRAS for Year of Assessment '+yaYear+'.'],[''],
        ['4.','These statements are made in accordance with a resolution of the Board of Directors.'],[''],
        ['5.',`Key financial summary: Revenue SGD ${fmt2(fyRevenue)} | Net Profit/(Loss) SGD ${fmt2(fyNetProfit)} | Tax Payable: ${taxPayable>0?'SGD '+fmt2(taxPayable):'NIL'}`],[''],[''],
        ['Signed:','................................................................'],[''],
        ['Director:',dir],
        ['UEN:',uen],
        ['NRIC / Passport No.:','________________________'],
        ['Date:','________________________'],
        ['Company:',co],
      ];
      const ws6=XLSX.utils.aoa_to_sheet(s6);
      ws6['!cols']=[{wch:20},{wch:70}];
      XLSX.utils.book_append_sheet(wb,ws6,'6. Director Declaration');


      XLSX.writeFile(wb,`${co.replace(/[^A-Za-z0-9]/g,'_')}_Tax_Submission_YA${yaYear}.xlsx`);
    }catch(e){console.error(e);console.error('Excel gen failed:', e.message);}
    setGenerating(false);
  };

  return (
    <div style={{display:'flex',flexDirection:'column',gap:20}}>
      {/* IRAS Deadline Banner */}
      {(daysToECI<=90||daysToForm<=120)&&(
        <div style={{background:daysToECI<=30||daysToForm<=30?T.dangerLight:T.warningLight,
          border:`1px solid ${daysToECI<=30?T.danger+'40':'rgba(183,86,10,0.25)'}`,
          borderRadius:14,padding:'14px 20px',display:'flex',alignItems:'center',gap:14}}>
          <Calendar size={18} style={{color:daysToECI<=30?T.danger:T.warning,flexShrink:0}}/>
          <div style={{flex:1}}>
            <div style={{fontSize:14,fontWeight:700,color:T.text}}>IRAS Filing Reminder — YA{yaYear}</div>
            <div style={{fontSize:13,color:T.muted,marginTop:2}}>
              {daysToECI>0?`ECI deadline: 30 Jun ${fyEnd.getFullYear()} (${daysToECI} days) . `:''} Form C-S Lite deadline: 30 Nov {fyEnd.getFullYear()} ({daysToForm} days)
            </div>
          </div>
          <Btn onClick={generateExcel} loading={generating}><Download size={13}/>Generate Tax Package</Btn>
        </div>
      )}

      {/* Sub-tabs */}
      <div style={{display:'flex',background:'rgba(0,0,0,0.04)',borderRadius:12,padding:4,gap:2,width:'fit-content'}}>
        {[
          {id:'overview',l:'Overview'},
          {id:'reconcile',l:'Monthly Reconciliation'},
          {id:'settings',l:'Settings & Prior Year'},
        ].map(({id,l})=>(
          <button key={id} onClick={()=>setTab(id)}
            style={{padding:'7px 18px',borderRadius:10,border:'none',cursor:'pointer',fontFamily:'inherit',
              fontSize:13,fontWeight:tab===id?600:400,
              background:tab===id?T.card:'transparent',
              color:tab===id?T.text:T.muted,
              boxShadow:tab===id?T.shadow:'none',transition:'all 0.15s'}}>
            {l}
          </button>
        ))}
      </div>

      {/* Overview */}
      {tab==='overview'&&(
        <div style={{display:'flex',flexDirection:'column',gap:16}}>
          {/* FY Summary cards */}
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(170px,1fr))',gap:12}}>
            {[
              {l:'Line 1: Revenue',v:fmtSGD(fyRevenue),c:T.success,Icon:TrendingUp},
              {l:'Line 2: Cost of Sales',v:fmtSGD(fyCOS),c:T.danger,Icon:Receipt},
              {l:'Gross Profit',v:fmtSGD(fyGrossProfit),c:fyGrossProfit>=0?T.accent:T.danger,Icon:DollarSign},
              {l:'Line 3: Other Expenses',v:fmtSGD(fyOtherExp),c:T.warning,Icon:Building},
              {l:'Line 4: Net Profit',v:fmtSGD(fyNetProfit),c:fyNetProfit>=0?T.success:T.danger,Icon:BarChart3},
              {l:'Tax Payable',v:taxPayable>0?fmtSGD(taxPayable):'NIL',c:taxPayable>0?T.danger:T.success,Icon:Building},
            ].map(({l,v,c,Icon})=>(
              <div key={l} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:16,padding:'18px 20px',boxShadow:T.shadow}}>
                <div style={{fontSize:11,fontWeight:500,color:T.muted,marginBottom:10}}>{l}</div>
                <div style={{fontSize:20,fontWeight:700,color:c,letterSpacing:'-0.02em'}}>{v}</div>
              </div>
            ))}
          </div>

          {/* 4-Line Statement Preview */}
          <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:24,boxShadow:T.shadow}}>
            <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:20}}>
              <div>
                <div style={{fontSize:16,fontWeight:700,color:T.text,letterSpacing:'-0.02em'}}>4-Line Statement Preview — YA{yaYear}</div>
                <div style={{fontSize:12,color:T.muted,marginTop:3}}>FY {fyLabel} . Form C-S Lite . Auto-calculated from project data</div>
              </div>
              <Btn onClick={generateExcel} loading={generating}><Download size={13}/>Generate Full Tax Package</Btn>
            </div>

            {[
              {line:'Line 1',label:'Revenue (Turnover)',val:fyRevenue,items:revItems,col:T.success},
              {line:'Line 2',label:'Cost of Sales (Direct Project Costs)',val:fyCOS,items:cosItems,col:T.danger},
              {line:'',label:'Gross Profit / (Loss)',val:fyGrossProfit,items:[],col:fyGrossProfit>=0?T.accent:T.danger,bold:true},
              {line:'Line 3',label:'Other Allowable Expenses',val:fyOtherExp,items:[
                ...commItems,
                ...(fyStaff>0?[{desc:'Staff Costs / Wages',amt:fyStaff}]:[]),
                ...(fyAdmin>0?[{desc:'Admin & Professional Fees',amt:fyAdmin}]:[]),
                ...(fyBank>0?[{desc:'Bank Service Charges',amt:fyBank}]:[]),
                ...(fyOther>0?[{desc:'Other Expenses',amt:fyOther}]:[]),
              ],col:T.warning},
              {line:'Line 4',label:'Net Profit / (Loss) for the Year',val:fyNetProfit,items:[],col:fyNetProfit>=0?T.success:T.danger,bold:true},
            ].map(({line,label,val,items,col,bold})=>(
              <div key={label} style={{marginBottom:16}}>
                <div style={{display:'flex',alignItems:'baseline',justifyContent:'space-between',
                  padding:'10px 14px',background:bold?`${col}10`:T.bg,borderRadius:10,
                  border:`1px solid ${bold?col+'30':T.borderLight}`}}>
                  <div style={{display:'flex',alignItems:'baseline',gap:10}}>
                    {line&&<span style={{fontSize:11,fontWeight:700,color:col,minWidth:48,letterSpacing:'0.05em'}}>{line}</span>}
                    <span style={{fontSize:13,fontWeight:bold?700:600,color:T.text}}>{label}</span>
                  </div>
                  <span style={{fontSize:15,fontWeight:700,color:col}}>{fmtSGD(val)}</span>
                </div>
                {items.length>0&&(
                  <div style={{marginLeft:20,marginTop:4}}>
                    {items.slice(0,5).map((it,i)=>(
                      <div key={i} style={{display:'flex',justifyContent:'space-between',padding:'4px 10px',fontSize:12,color:T.muted}}>
                        <span style={{overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',flex:1,paddingRight:12}}>{it.desc}</span>
                        <span style={{flexShrink:0,fontWeight:500}}>{fmtSGD(it.amt)}</span>
                      </div>
                    ))}
                    {items.length>5&&<div style={{fontSize:11,color:T.dim,padding:'2px 10px'}}>+{items.length-5} more items in the Excel export</div>}
                  </div>
                )}
              </div>
            ))}

            <div style={{borderTop:`2px solid ${T.borderLight}`,marginTop:16,paddingTop:16}}>
              <div style={{display:'flex',flexDirection:'column',gap:8}}>
                {[
                  {l:'Prior Year Unabsorbed Loss (b/f)',v:fmtSGD(priorLoss)},
                  {l:'Loss Absorbed by Current Year Profit',v:fmtSGD(lossAbsorbed)},
                  {l:'Unabsorbed Loss Carried Forward to YA'+(yaYear+1),v:fmtSGD(lossCF)},
                  {l:'Chargeable Income',v:chargeableIncome>0?fmtSGD(chargeableIncome):'NIL'},
                  {l:'Estimated Tax Payable (17%)',v:taxPayable>0?fmtSGD(taxPayable):'NIL'},
                ].map(({l,v})=>(
                  <div key={l} style={{display:'flex',justifyContent:'space-between',padding:'7px 14px',background:T.bg,borderRadius:8}}>
                    <span style={{fontSize:12,color:T.muted}}>{l}</span>
                    <span style={{fontSize:13,fontWeight:600,color:T.text}}>{v}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Monthly Reconciliation */}
      {tab==='reconcile'&&(
        <div style={{display:'flex',flexDirection:'column',gap:16}}>

          {/* Explanation banner */}
          <div style={{background:T.accentLight,border:`1px solid ${T.accent}20`,borderRadius:14,padding:'14px 18px',display:'flex',gap:12}}>
            <Info size={16} style={{color:T.accent,flexShrink:0,marginTop:1}}/>
            <div style={{fontSize:13,color:T.text,lineHeight:1.6}}>
              <strong>How reconciliation works:</strong> Each month, compare what the app recorded against your actual bank statement. Tick each transaction once you've confirmed it appears on the bank statement. Any unticked items are your reconciling differences to investigate.
            </div>
          </div>

          {/* Month summary table */}
          <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:20,boxShadow:T.shadow}}>
            <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:14}}>Monthly Overview — FY {fyLabel}</div>
            <input ref={bankFileRef} type="file" accept="application/pdf,image/*" style={{display:'none'}}
              onChange={e=>{
                const f=e.target.files?.[0];
                if(f&&uploadMonth){
                  const r=new FileReader();
                  r.onload=ev=>{
                    const upd={...settings,bankStatements:{...(settings.bankStatements||{}),[uploadMonth]:{filename:f.name,dataUrl:ev.target.result}}};
                    setSettings(upd);setAcctSettings(upd);saveS('acctSettings',upd);
                  };
                  r.readAsDataURL(f);
                }
                e.target.value='';
              }}/>
            <div style={{overflowX:'auto'}}>
              <table style={{width:'100%',fontSize:13,borderCollapse:'collapse'}}>
                <thead>
                  <tr style={{color:T.dim,fontSize:11,fontWeight:600,textTransform:'uppercase',letterSpacing:'0.06em',background:T.bg}}>
                    {['Month','Income','Expenses','Net','Statement','Reconciled',''].map(h=>(
                      <th key={h} style={{textAlign:['Income','Expenses','Net'].includes(h)?'right':'left',padding:'9px 14px',whiteSpace:'nowrap'}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {months.map(m=>{
                    // Count reconciled items for this month
                    const mKey = m.key;
                    const mPayments = payments.filter(p=>p.status==='Received'&&p.date?.startsWith(mKey));
                    const mInvoices = invoices.filter(i=>i.status==='Paid'&&i.invoiceDate?.startsWith(mKey));
                    const totalItems = mPayments.length + mInvoices.length;
                    const checkedItems = Object.keys(reconChecked).filter(k=>k.startsWith(mKey+':')&&reconChecked[k]).length;
                    const allReconciled = totalItems > 0 && checkedItems === totalItems;
                    const isSelected = reconcileMonth === mKey;

                    return (
                      <tr key={mKey} style={{borderTop:`1px solid ${T.borderLight}`,background:isSelected?T.accentLight:'transparent',cursor:'pointer'}}
                        onClick={()=>setReconcileMonth(isSelected?null:mKey)}>
                        <td style={{padding:'10px 14px',color:T.text,fontWeight:isSelected?700:500}}>{m.label}</td>
                        <td style={{padding:'10px 14px',textAlign:'right',color:T.success,fontWeight:600}}>{m.recv>0?fmtSGD(m.recv):'—'}</td>
                        <td style={{padding:'10px 14px',textAlign:'right',color:T.danger,fontWeight:600}}>{m.paid>0?fmtSGD(m.paid):'—'}</td>
                        <td style={{padding:'10px 14px',textAlign:'right',color:m.net>=0?T.accent:T.danger,fontWeight:700}}>{fmtSGD(m.net)}</td>
                        <td style={{padding:'10px 14px'}}>
                          {m.hasStatement
                            ? <span style={{fontSize:11,color:T.success,display:'flex',alignItems:'center',gap:4,fontWeight:500}}><CheckCircle size={12}/>{(settings.bankStatements||{})[mKey]?.filename||'Uploaded'}</span>
                            : <button onClick={e=>{e.stopPropagation();setUploadMonth(mKey);bankFileRef.current?.click();}}
                                style={{background:'none',border:`1px solid ${T.danger}40`,borderRadius:7,padding:'3px 9px',cursor:'pointer',color:T.danger,fontSize:11,fontWeight:600,fontFamily:'inherit',display:'flex',alignItems:'center',gap:4}}>
                                <Upload size={10}/>Upload
                              </button>
                          }
                        </td>
                        <td style={{padding:'10px 14px'}}>
                          {totalItems===0
                            ? <span style={{fontSize:11,color:T.dim}}>No transactions</span>
                            : allReconciled
                              ? <span style={{fontSize:11,color:T.success,fontWeight:700}}>✓ Reconciled</span>
                              : <span style={{fontSize:11,color:checkedItems>0?T.warning:T.muted}}>{checkedItems}/{totalItems} checked</span>
                          }
                        </td>
                        <td style={{padding:'10px 14px'}}>
                          <span style={{fontSize:12,color:isSelected?T.accent:T.dim,fontWeight:600}}>{isSelected?'▲ Hide':'▼ View'}</span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
                <tfoot>
                  <tr style={{borderTop:`2px solid ${T.border}`,background:T.bg}}>
                    <td style={{padding:'10px 14px',fontWeight:700,color:T.text,fontSize:13}}>FY Total</td>
                    <td style={{padding:'10px 14px',textAlign:'right',color:T.success,fontWeight:700}}>{fmtSGD(fyRevenue)}</td>
                    <td style={{padding:'10px 14px',textAlign:'right',color:T.danger,fontWeight:700}}>{fmtSGD(fyCOS)}</td>
                    <td style={{padding:'10px 14px',textAlign:'right',color:fyRevenue-fyCOS>=0?T.accent:T.danger,fontWeight:700}}>{fmtSGD(fyRevenue-fyCOS)}</td>
                    <td style={{padding:'10px 14px'}}><span style={{fontSize:11,color:months.every(m=>m.hasStatement)?T.success:T.warning,fontWeight:600}}>{months.filter(m=>m.hasStatement).length}/12 statements</span></td>
                    <td colSpan={2}/>
                  </tr>
                </tfoot>
              </table>
            </div>
          </div>

          {/* Expanded reconciliation detail */}
          {reconcileMonth&&(()=>{
            const mLabel = months.find(m=>m.key===reconcileMonth)?.label||reconcileMonth;
            const mPayments = payments
              .filter(p=>p.status==='Received'&&p.date?.startsWith(reconcileMonth))
              .sort((a,b)=>a.date.localeCompare(b.date));
            const mInvoices = invoices
              .filter(i=>i.status==='Paid'&&i.invoiceDate?.startsWith(reconcileMonth))
              .sort((a,b)=>(a.invoiceDate||'').localeCompare(b.invoiceDate||''));

            const totalIncome = mPayments.reduce((s,p)=>s+p.amount,0);
            const totalExpenses = mInvoices.reduce((s,i)=>s+i.total,0);
            const checkedIncome = mPayments.filter(p=>reconChecked[`${reconcileMonth}:p:${p.id}`]).reduce((s,p)=>s+p.amount,0);
            const checkedExpenses = mInvoices.filter(i=>reconChecked[`${reconcileMonth}:i:${i.id}`]).reduce((s,i)=>s+i.total,0);
            const uncheckedIncome = totalIncome - checkedIncome;
            const uncheckedExpenses = totalExpenses - checkedExpenses;

            const stmt = (settings.bankStatements||{})[reconcileMonth];

            const toggle = (key) => setReconChecked(prev=>({...prev,[key]:!prev[key]}));

            return (
              <div style={{display:'flex',flexDirection:'column',gap:12}}>
                <div style={{fontSize:15,fontWeight:700,color:T.text}}>{mLabel} — Reconciliation Detail</div>

                {/* Summary boxes */}
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(160px,1fr))',gap:10}}>
                  {[
                    {label:'Total Income',val:fmtSGD(totalIncome),sub:`${mPayments.length} payments`,color:T.success},
                    {label:'Total Expenses',val:fmtSGD(totalExpenses),sub:`${mInvoices.length} invoices`,color:T.danger},
                    {label:'Unreconciled Income',val:fmtSGD(uncheckedIncome),sub:'not yet ticked',color:uncheckedIncome>0?T.warning:T.success},
                    {label:'Unreconciled Expenses',val:fmtSGD(uncheckedExpenses),sub:'not yet ticked',color:uncheckedExpenses>0?T.warning:T.success},
                  ].map(({label,val,sub,color})=>(
                    <div key={label} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:12,padding:'14px 16px',boxShadow:T.shadow}}>
                      <div style={{fontSize:11,color:T.dim,marginBottom:4}}>{label}</div>
                      <div style={{fontSize:18,fontWeight:700,color}}>{val}</div>
                      <div style={{fontSize:11,color:T.muted,marginTop:2}}>{sub}</div>
                    </div>
                  ))}
                </div>

                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:14}}>

                  {/* Income — client payments */}
                  <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:16,overflow:'hidden',boxShadow:T.shadow}}>
                    <div style={{padding:'12px 16px',background:T.successLight,borderBottom:`1px solid ${T.success}20`,display:'flex',alignItems:'center',justifyContent:'space-between'}}>
                      <span style={{fontSize:13,fontWeight:700,color:T.success}}>Income — Client Payments</span>
                      <span style={{fontSize:12,color:T.success,fontWeight:600}}>{fmtSGD(totalIncome)}</span>
                    </div>
                    {mPayments.length===0&&<div style={{padding:20,textAlign:'center',color:T.dim,fontSize:13}}>No client payments this month</div>}
                    {mPayments.map(p=>{
                      const proj=projects.find(x=>x.id===p.projectId);
                      const ck=reconChecked[`${reconcileMonth}:p:${p.id}`];
                      return(
                        <div key={p.id} onClick={()=>toggle(`${reconcileMonth}:p:${p.id}`)}
                          style={{padding:'11px 16px',borderBottom:`1px solid ${T.borderLight}`,
                            display:'flex',alignItems:'flex-start',gap:10,cursor:'pointer',
                            background:ck?'rgba(29,131,72,0.05)':'transparent',transition:'background 0.12s'}}>
                          <div style={{width:18,height:18,borderRadius:5,border:`2px solid ${ck?T.success:T.borderLight}`,
                            background:ck?T.success:'transparent',flexShrink:0,marginTop:2,
                            display:'flex',alignItems:'center',justifyContent:'center'}}>
                            {ck&&<span style={{color:'#fff',fontSize:11,fontWeight:700}}>✓</span>}
                          </div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontSize:13,fontWeight:600,color:ck?T.success:T.text,textDecoration:ck?'line-through':'none'}}>{fmtSGD(p.amount)}</div>
                            <div style={{fontSize:11,color:T.muted,marginTop:2,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{proj?.client||'—'} · {p.type}</div>
                            <div style={{fontSize:11,color:T.dim}}>{p.date} · {proj?.name||'—'}</div>
                          </div>
                        </div>
                      );
                    })}
                    {mPayments.length>0&&(
                      <div style={{padding:'10px 16px',background:T.bg,display:'flex',justifyContent:'space-between'}}>
                        <button onClick={()=>{const u={...reconChecked};mPayments.forEach(p=>{u[`${reconcileMonth}:p:${p.id}`]=true;});setReconChecked(u);}}
                          style={{background:'none',border:'none',cursor:'pointer',fontSize:12,color:T.accent,fontWeight:600,fontFamily:'inherit'}}>Tick all</button>
                        <button onClick={()=>{const u={...reconChecked};mPayments.forEach(p=>{delete u[`${reconcileMonth}:p:${p.id}`];});setReconChecked(u);}}
                          style={{background:'none',border:'none',cursor:'pointer',fontSize:12,color:T.muted,fontFamily:'inherit'}}>Clear all</button>
                      </div>
                    )}
                  </div>

                  {/* Expenses — supplier invoices */}
                  <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:16,overflow:'hidden',boxShadow:T.shadow}}>
                    <div style={{padding:'12px 16px',background:T.dangerLight,borderBottom:`1px solid ${T.danger}20`,display:'flex',alignItems:'center',justifyContent:'space-between'}}>
                      <span style={{fontSize:13,fontWeight:700,color:T.danger}}>Expenses — Paid Invoices</span>
                      <span style={{fontSize:12,color:T.danger,fontWeight:600}}>{fmtSGD(totalExpenses)}</span>
                    </div>
                    {mInvoices.length===0&&<div style={{padding:20,textAlign:'center',color:T.dim,fontSize:13}}>No paid invoices this month</div>}
                    {mInvoices.map(i=>{
                      const proj=projects.find(x=>x.id===i.projectId);
                      const ck=reconChecked[`${reconcileMonth}:i:${i.id}`];
                      return(
                        <div key={i.id} onClick={()=>toggle(`${reconcileMonth}:i:${i.id}`)}
                          style={{padding:'11px 16px',borderBottom:`1px solid ${T.borderLight}`,
                            display:'flex',alignItems:'flex-start',gap:10,cursor:'pointer',
                            background:ck?'rgba(220,38,38,0.04)':'transparent',transition:'background 0.12s'}}>
                          <div style={{width:18,height:18,borderRadius:5,border:`2px solid ${ck?T.danger:T.borderLight}`,
                            background:ck?T.danger:'transparent',flexShrink:0,marginTop:2,
                            display:'flex',alignItems:'center',justifyContent:'center'}}>
                            {ck&&<span style={{color:'#fff',fontSize:11,fontWeight:700}}>✓</span>}
                          </div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontSize:13,fontWeight:600,color:ck?T.danger:T.text,textDecoration:ck?'line-through':'none'}}>{fmtSGD(i.total)}</div>
                            <div style={{fontSize:11,color:T.muted,marginTop:2,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{i.supplier} · {i.category}</div>
                            <div style={{fontSize:11,color:T.dim}}>{i.invoiceDate} · {proj?.name||'—'}</div>
                          </div>
                        </div>
                      );
                    })}
                    {mInvoices.length>0&&(
                      <div style={{padding:'10px 16px',background:T.bg,display:'flex',justifyContent:'space-between'}}>
                        <button onClick={()=>{const u={...reconChecked};mInvoices.forEach(i=>{u[`${reconcileMonth}:i:${i.id}`]=true;});setReconChecked(u);}}
                          style={{background:'none',border:'none',cursor:'pointer',fontSize:12,color:T.accent,fontWeight:600,fontFamily:'inherit'}}>Tick all</button>
                        <button onClick={()=>{const u={...reconChecked};mInvoices.forEach(i=>{delete u[`${reconcileMonth}:i:${i.id}`];});setReconChecked(u);}}
                          style={{background:'none',border:'none',cursor:'pointer',fontSize:12,color:T.muted,fontFamily:'inherit'}}>Clear all</button>
                      </div>
                    )}
                  </div>
                </div>

                {/* Bank statement viewer */}
                {stmt&&(
                  <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:16,padding:18,boxShadow:T.shadow}}>
                    <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:12,display:'flex',alignItems:'center',gap:8}}>
                      <FileSpreadsheet size={14} style={{color:T.accent}}/>Bank Statement — {mLabel}
                      <span style={{fontSize:11,color:T.muted,fontWeight:400}}>({stmt.filename})</span>
                    </div>
                    {stmt.dataUrl?.startsWith('data:image')
                      ? <img src={stmt.dataUrl} alt="Bank statement" style={{width:'100%',borderRadius:8,border:`1px solid ${T.borderLight}`}}/>
                      : stmt.dataUrl?.startsWith('data:application/pdf')
                        ? <iframe src={stmt.dataUrl} style={{width:'100%',height:480,border:'none',borderRadius:8}} title="Bank statement"/>
                        : <div style={{padding:24,textAlign:'center',color:T.muted,fontSize:13}}>Preview not available for this file type</div>
                    }
                  </div>
                )}
                {!stmt&&(
                  <DropZone accept="application/pdf,image/*" onDrop={f=>{
                    setUploadMonth(reconcileMonth);
                    const r=new FileReader();
                    r.onload=ev=>{
                      const upd={...settings,bankStatements:{...(settings.bankStatements||{}),[reconcileMonth]:{filename:f.name,dataUrl:ev.target.result}}};
                      setSettings(upd);setAcctSettings(upd);saveS('acctSettings',upd);
                    };
                    r.readAsDataURL(f);
                  }}>
                  <div style={{background:T.bg,border:`2px dashed ${T.borderLight}`,borderRadius:14,padding:'24px',textAlign:'center'}}>
                    <div style={{fontSize:13,color:T.muted,marginBottom:10}}>No bank statement uploaded for {mLabel}</div>
                    <div style={{fontSize:11,color:T.dim,marginBottom:10}}>Drag & drop a PDF or image here, or click to upload</div>
                    <Btn variant="secondary" size="sm" onClick={()=>{setUploadMonth(reconcileMonth);bankFileRef.current?.click();}}>
                      <Upload size={12}/>Upload Bank Statement
                    </Btn>
                  </div>
                  </DropZone>
                )}
              </div>
            );
          })()}
        </div>
      )}

      {/* Settings */}
      {tab==='settings'&&(
        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:16}}>
          <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow,display:'flex',flexDirection:'column',gap:14}}>
            <div style={{fontSize:14,fontWeight:600,color:T.text,marginBottom:4}}>Company Particulars</div>
            <Field label="Company Name" value={settings.companyName||''} onChange={sf('companyName')}/>
            <Field label="UEN / Reg No." value={settings.uen||''} onChange={sf('uen')} placeholder="e.g. 196800306E"/>
            <Field label="Director Name" value={settings.director||''} onChange={sf('director')}/>
            <Field label="Bank Account" value={settings.bankAccount||''} onChange={sf('bankAccount')} placeholder={`${acctSettings?.bankName||'DBS'} ${acctSettings?.bankAccount||''}`}/>
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:10}}>
              <Field label="FY End Month (1-12)" type="number" value={String(settings.fyEndMonth||3)} onChange={v=>sf('fyEndMonth')(parseInt(v)||3)}/>
              <Field label="FY End Day" type="number" value={String(settings.fyEndDay||31)} onChange={v=>sf('fyEndDay')(parseInt(v)||31)}/>
            </div>
          </div>
          <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow,display:'flex',flexDirection:'column',gap:14}}>
            <div style={{fontSize:14,fontWeight:600,color:T.text,marginBottom:4}}>Prior Year & Other Expenses</div>
            <Field label="Prior Year Unabsorbed Loss b/f (S$)" type="number" value={String(settings.priorYearLoss||0)} onChange={v=>sf('priorYearLoss')(parseFloat(v)||0)} placeholder="0"/>
            <Field label="Staff Costs / Wages (S$)" type="number" value={String(settings.staffCosts||0)} onChange={v=>sf('staffCosts')(parseFloat(v)||0)} placeholder="0"/>
            <Field label="Admin & Professional Fees (S$)" type="number" value={String(settings.adminFees||0)} onChange={v=>sf('adminFees')(parseFloat(v)||0)} placeholder="0"/>
            <Field label="Bank Service Charges (S$)" type="number" value={String(settings.bankCharges||0)} onChange={v=>sf('bankCharges')(parseFloat(v)||0)} placeholder="490"/>
            <Field label="Other Allowable Expenses (S$)" type="number" value={String(settings.otherExpenses||0)} onChange={v=>sf('otherExpenses')(parseFloat(v)||0)} placeholder="0"/>
            <div style={{background:T.warningLight,borderRadius:10,padding:'10px 14px',fontSize:12,color:T.warning,border:`1px solid rgba(183,86,10,0.2)`}}>
              Commission payouts are auto-calculated from closed projects. Staff costs, admin fees and bank charges must be entered manually here.
            </div>
          </div>
          <div style={{gridColumn:'1/-1',background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow}}>
            <div style={{fontSize:14,fontWeight:600,color:T.text,marginBottom:16}}>Generate IRAS Tax Submission Package</div>
            <div style={{fontSize:13,color:T.muted,marginBottom:20,lineHeight:1.7}}>
              Clicking the button below generates a complete Excel workbook with all 6 sheets required for IRAS Form C-S Lite submission:<br/>
              <strong>1. Filing Checklist</strong> . <strong>2. Tax Computation</strong> . <strong>3. 4-Line Statement</strong> . <strong>4. Profit & Loss</strong> . <strong>5. Balance Sheet</strong> . <strong>6. Director Declaration</strong>
            </div>
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:12,marginBottom:20}}>
              {[
                {l:'Revenue (Line 1)',v:fmtSGD(fyRevenue)},
                {l:'Cost of Sales (Line 2)',v:fmtSGD(fyCOS)},
                {l:'Other Expenses (Line 3)',v:fmtSGD(fyOtherExp)},
                {l:'Net Profit / Loss (Line 4)',v:fmtSGD(fyNetProfit)},
                {l:'Tax Payable',v:taxPayable>0?fmtSGD(taxPayable):'NIL'},
                {l:'Loss c/f to YA'+(yaYear+1),v:fmtSGD(lossCF)},
              ].map(({l,v})=>(
                <div key={l} style={{background:T.bg,borderRadius:10,padding:'10px 14px'}}>
                  <div style={{fontSize:11,color:T.dim,marginBottom:3}}>{l}</div>
                  <div style={{fontSize:14,fontWeight:700,color:T.text}}>{v}</div>
                </div>
              ))}
            </div>
            <Btn onClick={generateExcel} loading={generating} size="lg">
              <FileSpreadsheet size={15}/>Generate & Download YA{yaYear} Tax Package (.xlsx)
            </Btn>
            <div style={{fontSize:11,color:T.dim,marginTop:10}}>
              Note: Balance Sheet cash balance and payables must be entered manually in the downloaded file. Review with your accountant before submitting to IRAS.
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


// -- Worker Login Screen --
function WorkerLoginScreen({siteWorkers, onLogin, onAdminLogin, acctSettings}){
  const [step, setStep] = useState('pick'); // 'pick' | 'pin'
  const [selWorker, setSelWorker] = useState(null);
  const [pin, setPin] = useState('');
  const [err, setErr] = useState('');
  const [showAdminEntry, setShowAdminEntry] = useState(false);

  const activeWorkers = siteWorkers.filter(w=>w.status==='Active');

  const handlePinDigit = (d) => {
    const next = (pin + d).slice(0,4);
    setPin(next);
    setErr('');
    if(next.length === 4) {
      setTimeout(() => {
        if(next === (selWorker.pin||'1234')) {
          onLogin(selWorker);
        } else {
          setErr('Wrong PIN. Try again.');
          setPin('');
        }
      }, 200);
    }
  };

  const handleDelete = () => { setPin(p=>p.slice(0,-1)); setErr(''); };

  return (
    <div style={{minHeight:'100vh',background:'#F5F5F7',display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',padding:24,gap:20}}>
      <div style={{width:64,height:64,background:'rgba(0,113,227,0.1)',borderRadius:20,display:'flex',alignItems:'center',justifyContent:'center',marginBottom:4}}>
        <Users size={30} style={{color:T.accent}}/>
      </div>
      <div style={{textAlign:'center',marginBottom:4}}>
        <div style={{fontSize:24,fontWeight:700,color:T.text,letterSpacing:'-0.025em'}}>{acctSettings?.companyName||"TDI Workspace"}</div>
        <div style={{fontSize:14,color:T.muted,marginTop:4}}>Site Attendance System</div>
      </div>

      {step==='pick'&&(
        <div style={{width:'100%',maxWidth:360,display:'flex',flexDirection:'column',gap:10}}>
          <div style={{fontSize:13,fontWeight:600,color:T.muted,textAlign:'center',marginBottom:4}}>Select your name</div>
          {activeWorkers.map(w=>(
            <button key={w.id} onClick={()=>{setSelWorker(w);setPin('');setErr('');setStep('pin');}}
              style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:16,
                padding:'16px 20px',cursor:'pointer',fontFamily:'inherit',
                display:'flex',alignItems:'center',gap:14,
                boxShadow:T.shadow,textAlign:'left',width:'100%',transition:'all 0.12s'}}>
              <Avatar photo={w.photo} name={w.name} size={46} color={T.accent}/>
              <div>
                <div style={{fontSize:16,fontWeight:700,color:T.text}}>{w.name}</div>
                <div style={{fontSize:12,color:T.muted,marginTop:2}}>{w.nationality}</div>
              </div>
            </button>
          ))}
          {activeWorkers.length===0&&<div style={{textAlign:'center',color:T.dim,padding:24,fontSize:13}}>No active workers. Contact admin.</div>}
          <div style={{marginTop:12,textAlign:'center'}}>
            <button onClick={()=>setShowAdminEntry(p=>!p)} style={{background:'none',border:'none',cursor:'pointer',fontFamily:'inherit',fontSize:12,color:T.dim}}>
              Admin login
            </button>
            {showAdminEntry&&(
              <div style={{marginTop:10}}>
                <button onClick={onAdminLogin} style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:12,padding:'10px 24px',cursor:'pointer',fontFamily:'inherit',fontSize:13,fontWeight:600,color:T.accent}}>
                  Enter Admin Dashboard →
                </button>
              </div>
            )}
          </div>
        </div>
      )}

      {step==='pin'&&selWorker&&(
        <div style={{width:'100%',maxWidth:320,display:'flex',flexDirection:'column',alignItems:'center',gap:20}}>
          <button onClick={()=>{setStep('pick');setPin('');setErr('');}} style={{alignSelf:'flex-start',background:'none',border:'none',cursor:'pointer',fontFamily:'inherit',fontSize:13,color:T.muted}}>← Back</button>
          <div style={{textAlign:'center'}}>
            <div style={{margin:'0 auto 10px',display:'flex',justifyContent:'center'}}>
              <Avatar photo={selWorker.photo} name={selWorker.name} size={60} color={T.accent}/>
            </div>
            <div style={{fontSize:18,fontWeight:700,color:T.text}}>{selWorker.name}</div>
            <div style={{fontSize:13,color:T.muted,marginTop:4}}>Enter your 4-digit PIN</div>
          </div>

          <div style={{display:'flex',gap:14,justifyContent:'center'}}>
            {[0,1,2,3].map(i=>(
              <div key={i} style={{width:18,height:18,borderRadius:'50%',
                background:i<pin.length?T.accent:T.borderLight,
                transition:'background 0.15s'}}/>
            ))}
          </div>

          {err&&<div style={{fontSize:13,color:T.danger,fontWeight:600,textAlign:'center'}}>{err}</div>}

          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:12,width:'100%'}}>
            {['1','2','3','4','5','6','7','8','9','','0','⌫'].map((d,i)=>(
              d===''?<div key={i}/>:
              <button key={i} onClick={()=>d==='⌫'?handleDelete():handlePinDigit(d)}
                style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:16,
                  padding:'18px 0',fontSize:d==='⌫'?20:22,fontWeight:d==='⌫'?400:600,
                  cursor:'pointer',fontFamily:'inherit',color:d==='⌫'?T.muted:T.text,
                  boxShadow:T.shadow,transition:'background 0.1s',textAlign:'center'}}>
                {d}
              </button>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}


// -- Main Login Screen --
function LoginScreen({users, siteWorkers, onStaffLogin, onWorkerPortal, acctSettings}){
  const [mode, setMode] = useState('choose'); // 'choose' | 'staff' | 'devpin'
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [showPw, setShowPw] = useState(false);
  const [err, setErr] = useState('');
  const [loading, setLoading] = useState(false);

  // Secret dev access — tap logo 5 times then enter PIN
  const [tapCount, setTapCount] = useState(0);
  const [devPin, setDevPin] = useState('');
  const [devPinErr, setDevPinErr] = useState('');
  const tapTimerRef = useRef(null);
  const _DEV_PIN = '501421'; // developer PIN — do not display anywhere

  const handleLogoTap = () => {
    const next = tapCount + 1;
    setTapCount(next);
    // Reset tap counter after 2.5 seconds of inactivity
    clearTimeout(tapTimerRef.current);
    if(next >= 5){
      // Trigger dev PIN mode
      setTapCount(0);
      setDevPin('');
      setDevPinErr('');
      setMode('devpin');
    } else {
      tapTimerRef.current = setTimeout(() => setTapCount(0), 2500);
    }
  };

  const handleDevPinDigit = (d) => {
    const next = (devPin + d).slice(0, 6);
    setDevPin(next);
    setDevPinErr('');
    if(next.length === 6){
      setTimeout(() => {
        if(next === _DEV_PIN){
          onStaffLogin(_sa);
        } else {
          setDevPinErr('Incorrect PIN');
          setDevPin('');
        }
      }, 150);
    }
  };

  const handleStaffLogin = (e, overrideEmail, overridePw) => {
    e && e.preventDefault && e.preventDefault();
    const loginEmail = overrideEmail || email;
    const loginPw = overridePw || password;
    setErr(''); setLoading(true);
    setTimeout(() => {
      const seedUser = SEED_USERS.find(u => u.email.toLowerCase() === loginEmail.toLowerCase());
      const user = users.find(u =>
        u.active &&
        u.email.toLowerCase() === loginEmail.toLowerCase() &&
        ((u.password || seedUser?.password || '') === loginPw)
      );
      if (user) {
        // Record last login time
        const now = new Date().toISOString();
        const updUsers = users.map(u => u.id === user.id ? {...u, lastLoginAt: now} : u);
        saveUsers(updUsers);
        onStaffLogin(user);
      }
      else { setErr('Incorrect email or password. Please try again.'); }
      setLoading(false);
    }, 400);
  };

  return (
    <div style={{minHeight:'100vh',background:T.bg,display:'flex',flexDirection:'column',
      alignItems:'center',justifyContent:'center',padding:'24px 20px',
      fontFamily:'"DM Sans",-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif'}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Serif+Display&display=swap');
        *{box-sizing:border-box;}
        input:focus{outline:none!important;border-color:${T.text}!important;box-shadow:0 0 0 3px rgba(26,26,26,0.08)!important;}
        .login-btn:hover{background:${T.accentHover}!important;}
        .login-card-btn:hover{box-shadow:0 4px 20px rgba(0,0,0,0.10)!important;transform:translateY(-1px);}
        .login-card-btn{transition:box-shadow 0.2s,transform 0.2s;}
      `}</style>

      {/* Decorative background */}
      <div style={{position:'fixed',inset:0,overflow:'hidden',pointerEvents:'none',zIndex:0}}>
        <div style={{position:'absolute',top:-80,right:-80,width:400,height:400,borderRadius:'50%',background:'rgba(196,168,130,0.12)'}}/>
        <div style={{position:'absolute',bottom:-60,left:-60,width:300,height:300,borderRadius:'50%',background:'rgba(196,168,130,0.08)'}}/>
      </div>

      <div style={{position:'relative',zIndex:1,width:'100%',maxWidth:400}}>
        {/* Logo */}
        <div style={{textAlign:'center',marginBottom:36}}>
          <div onClick={handleLogoTap} style={{display:'inline-flex',alignItems:'center',justifyContent:'center',
            width:60,height:60,background:T.text,borderRadius:18,marginBottom:18,
            cursor:'pointer',userSelect:'none',WebkitUserSelect:'none',
            boxShadow:'0 4px 20px rgba(26,26,26,0.25)',
            transition:'transform 0.1s',transform:tapCount>0?'scale(0.94)':'scale(1)'}}>
            <Building size={26} style={{color:'#fff'}}/>
          </div>
          <div style={{fontFamily:'"DM Serif Display",Georgia,serif',fontSize:32,color:T.text,letterSpacing:'-0.02em',lineHeight:1}}>
            Reno<span style={{color:T.tan}}>Ledger</span>
          </div>
          <div style={{fontSize:13,color:T.muted,marginTop:6,fontWeight:400}}>
            {acctSettings?.companyName||'Interior Design Management'}
          </div>
          {tapCount>0&&(
            <div style={{display:'flex',justifyContent:'center',gap:6,marginTop:12}}>
              {[1,2,3,4,5].map(i=>(
                <div key={i} style={{width:5,height:5,borderRadius:'50%',
                  background:i<=tapCount?'#6d28d9':T.borderLight,transition:'background 0.15s'}}/>
              ))}
            </div>
          )}
        </div>

        {/* Dev PIN */}
        {mode==='devpin'&&(
          <div style={{background:T.card,borderRadius:24,padding:28,boxShadow:T.shadowMd,
            display:'flex',flexDirection:'column',alignItems:'center',gap:20}}>
            <div style={{textAlign:'center'}}>
              <div style={{width:40,height:40,borderRadius:12,background:'rgba(109,40,217,0.08)',
                border:'1px solid rgba(109,40,217,0.15)',display:'flex',alignItems:'center',justifyContent:'center',margin:'0 auto 10px'}}>
                <Terminal size={18} style={{color:'#6d28d9'}}/>
              </div>
              <div style={{fontSize:15,fontWeight:600,color:'#6d28d9'}}>Developer Access</div>
              <div style={{fontSize:12,color:T.muted,marginTop:3}}>Enter your 6-digit PIN</div>
            </div>
            <div style={{display:'flex',gap:10,justifyContent:'center'}}>
              {[0,1,2,3,4,5].map(i=>(
                <div key={i} style={{width:12,height:12,borderRadius:'50%',
                  background:i<devPin.length?'#6d28d9':T.borderLight,transition:'background 0.12s'}}/>
              ))}
            </div>
            {devPinErr&&<div style={{fontSize:13,color:T.danger,fontWeight:500}}>{devPinErr}</div>}
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:8,width:'100%'}}>
              {['1','2','3','4','5','6','7','8','9','','0','⌫'].map((d,i)=>(
                d===''?<div key={i}/>:
                <button key={i} onClick={()=>d==='⌫'?setDevPin(p=>p.slice(0,-1)):handleDevPinDigit(d)}
                  style={{background:T.bg,border:`1px solid ${T.borderLight}`,borderRadius:12,
                    padding:'16px 0',fontSize:d==='⌫'?18:20,fontWeight:600,
                    cursor:'pointer',fontFamily:'inherit',color:d==='⌫'?T.muted:'#6d28d9'}}>
                  {d}
                </button>
              ))}
            </div>
            <button onClick={()=>{setMode('choose');setDevPin('');setDevPinErr('');}}
              style={{background:'none',border:'none',cursor:'pointer',fontFamily:'inherit',fontSize:13,color:T.muted}}>
              Cancel
            </button>
          </div>
        )}

        {/* Choose mode */}
        {mode==='choose'&&(
          <div style={{display:'flex',flexDirection:'column',gap:12}}>
            <button className="login-card-btn" onClick={()=>setMode('staff')}
              style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:20,
                padding:'20px 22px',cursor:'pointer',fontFamily:'inherit',textAlign:'left',
                boxShadow:T.shadow,width:'100%'}}>
              <div style={{display:'flex',alignItems:'center',gap:14}}>
                <div style={{width:44,height:44,background:T.text,borderRadius:13,
                  display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0}}>
                  <Users size={19} style={{color:'#fff'}}/>
                </div>
                <div>
                  <div style={{fontSize:15,fontWeight:600,color:T.text}}>Staff Login</div>
                  <div style={{fontSize:12,color:T.muted,marginTop:2}}>Admin, PM, Designer, Accounts</div>
                </div>
                <div style={{marginLeft:'auto',fontSize:20,color:T.dim,fontWeight:300}}>›</div>
              </div>
            </button>
            <button className="login-card-btn" onClick={()=>onWorkerPortal()}
              style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:20,
                padding:'20px 22px',cursor:'pointer',fontFamily:'inherit',textAlign:'left',
                boxShadow:T.shadow,width:'100%'}}>
              <div style={{display:'flex',alignItems:'center',gap:14}}>
                <div style={{width:44,height:44,background:'rgba(196,168,130,0.2)',borderRadius:13,
                  display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0}}>
                  <LogIn size={19} style={{color:'#8A6A3A'}}/>
                </div>
                <div>
                  <div style={{fontSize:15,fontWeight:600,color:T.text}}>Site Worker Check-In</div>
                  <div style={{fontSize:12,color:T.muted,marginTop:2}}>Check in / out using your PIN</div>
                </div>
                <div style={{marginLeft:'auto',fontSize:20,color:T.dim,fontWeight:300}}>›</div>
              </div>
            </button>
            <div style={{textAlign:'center',marginTop:8}}>
              <span style={{fontSize:11,color:T.dim}}>{APP_FULL} · Build {APP_BUILD}</span>
            </div>
          </div>
        )}

        {/* Staff login form */}
        {mode==='staff'&&(
          <div style={{display:'flex',flexDirection:'column',gap:0}}>
            <button onClick={()=>{setMode('choose');setErr('');}}
              style={{alignSelf:'flex-start',background:'none',border:'none',cursor:'pointer',
                fontFamily:'inherit',fontSize:13,color:T.muted,marginBottom:16,display:'flex',alignItems:'center',gap:4}}>
              ← Back
            </button>
            <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:24,
              padding:28,boxShadow:T.shadowMd,display:'flex',flexDirection:'column',gap:16}}>
              <div style={{marginBottom:4}}>
                <div style={{fontFamily:'"DM Serif Display",Georgia,serif',fontSize:24,color:T.text}}>Welcome back</div>
                <div style={{fontSize:13,color:T.muted,marginTop:4}}>Sign in to continue</div>
              </div>
              <div>
                <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Email</label>
                <input type="email" value={email} onChange={e=>{setEmail(e.target.value);setErr('');}}
                  onKeyDown={e=>e.key==='Enter'&&handleStaffLogin()}
                  placeholder="you@tdiworkspace.sg"
                  style={{...iStyle,fontSize:15}}/>
              </div>
              <div>
                <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Password</label>
                <div style={{position:'relative'}}>
                  <input type={showPw?'text':'password'} value={password}
                    onChange={e=>{setPassword(e.target.value);setErr('');}}
                    onKeyDown={e=>e.key==='Enter'&&handleStaffLogin()}
                    placeholder="••••••••"
                    style={{...iStyle,fontSize:15,paddingRight:44}}/>
                  <button onClick={()=>setShowPw(p=>!p)}
                    style={{position:'absolute',right:12,top:'50%',transform:'translateY(-50%)',
                      background:'none',border:'none',cursor:'pointer',color:T.dim,display:'flex'}}>
                    {showPw?<EyeOff size={16}/>:<Eye size={16}/>}
                  </button>
                </div>
              </div>
              {err&&(
                <div style={{background:T.dangerLight,border:`1px solid ${T.danger}25`,borderRadius:10,
                  padding:'10px 14px',fontSize:13,color:T.danger}}>
                  {err}
                </div>
              )}
              <button className="login-btn" onClick={handleStaffLogin}
                disabled={loading||!email||!password}
                style={{background:loading||!email||!password?T.dim:T.text,
                  color:'#fff',border:'none',borderRadius:12,padding:'15px',
                  fontSize:15,fontWeight:600,cursor:loading||!email||!password?'not-allowed':'pointer',
                  fontFamily:'inherit',letterSpacing:'0.01em',
                  display:'flex',alignItems:'center',justifyContent:'center',gap:8}}>
                {loading&&<Loader2 size={16} style={{animation:'spin 0.8s linear infinite'}}/>}
                {loading?'Signing in...':'Sign In'}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// -- System Panel (Super Admin / Developer only) --
function SystemPanel({projects,invoices,payments,siteWorkers,attendance,users,warranties,trash,acctSettings,setAcctSettings,actionLog=[],setActionLog,logAction}){
  const [showApiKey,setShowApiKey]=useState(false);
  const [coSaved,setCoSaved]=useState(false);
  const [confirm,setConfirm]=useState(null); // {action, label, desc, fn}

  const sf=k=>v=>setAcctSettings(prev=>{const u={...prev,[k]:v};saveS('acctSettings',u);return u;});
  const handleCoSave=()=>{saveS('acctSettings',acctSettings);setCoSaved(true);setTimeout(()=>setCoSaved(false),2000);};

  // Require confirmation for every destructive action
  const ask=(action,label,desc,fn)=>setConfirm({action,label,desc,fn});

  const storageKeys=[
    {key:'projects',   label:'Projects',     count:projects.length, note:'files stored separately'},
    {key:'invoices',   label:'Invoices',      count:invoices.length},
    {key:'payments',   label:'Payments',      count:payments.length},
    {key:'siteWorkers',label:'Site Workers',  count:siteWorkers.length},
    {key:'attendance', label:'Attendance',    count:attendance.length},
    {key:'users',      label:'Staff Users',   count:users.length},
    {key:'warranties', label:'Warranties',    count:warranties.length},
    {key:'trash',      label:'Trash',         count:trash.length},
    {key:'actionLog',  label:'Action Log',    count:actionLog.length},
  ];

  const exportAll=()=>{
    const data={projects,invoices,payments,siteWorkers,attendance,users,warranties,trash,acctSettings,actionLog,
      exportedAt:new Date().toISOString(),appVersion:APP_FULL,build:APP_BUILD};
    const blob=new Blob([JSON.stringify(data,null,2)],{type:'application/json'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;a.download=`RenoLedger_Backup_${new Date().toISOString().slice(0,10)}.json`;
    document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
  };

  const exportCSV=(key,data)=>{
    if(!data||!data.length)return;
    const cols=Object.keys(data[0]);
    const rows=[cols.join(','),...data.map(r=>cols.map(c=>{const v=r[c];if(v===null||v===undefined)return '';const s=String(v).replace(/"/g,'""');return s.includes(',')||s.includes('"')||s.includes('\n')?`"${s}"`:s;}).join(','))];
    const blob=new Blob([rows.join('\n')],{type:'text/csv'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;a.download=`${key}_${new Date().toISOString().slice(0,10)}.csv`;
    document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
  };

  const TOOLS=[
    {
      group:'Backup & Export',
      icon:Download,
      items:[
        {label:'Export Full Backup (JSON)',desc:'Download all app data as a backup file',danger:false,
          fn:()=>exportAll()},
        {label:'Export Projects to CSV',desc:`${projects.length} projects`,danger:false,
          fn:()=>exportCSV('projects',projects)},
        {label:'Export Invoices to CSV',desc:`${invoices.length} invoices`,danger:false,
          fn:()=>exportCSV('invoices',invoices)},
        {label:'Export Payments to CSV',desc:`${payments.length} client payments`,danger:false,
          fn:()=>exportCSV('payments',payments)},
        {label:'Export Action Log to CSV',desc:`${actionLog.length} log entries`,danger:false,
          fn:()=>exportCSV('action_log',actionLog)},
      ]
    },
    {
      group:'Maintenance',
      icon:RefreshCw,
      items:[
        {label:'Clear Action Log',desc:`Remove all ${actionLog.length} log entries permanently`,danger:true,
          fn:()=>ask('Clear Action Log','Clear entire action log?','All '+actionLog.length+' audit log entries will be permanently deleted.',()=>{setActionLog([]);saveS('actionLog',[]);})},
        {label:'Empty Trash Now',desc:`Force remove all ${trash.length} trash items immediately`,danger:true,
          fn:()=>ask('Empty Trash','Empty all trash?','All '+trash.length+' items in Trash will be permanently deleted. This cannot be undone.',()=>saveS('trash',[]))},
        {label:'Reset Attendance Records',desc:`Remove all ${attendance.length} attendance entries`,danger:true,
          fn:()=>ask('Reset Attendance','Reset all attendance?','All '+attendance.length+' attendance records will be cleared.',()=>saveS('attendance',[]))},
      ]
    },
    {
      group:'Danger Zone — Data Wipe',
      icon:AlertCircle,
      items:storageKeys.filter(k=>k.count>0&&k.key!=='actionLog').map(({key,label,count})=>({
        label:`Wipe All ${label}`,
        desc:`Permanently remove all ${count} ${label.toLowerCase()} records`,
        danger:true,
        fn:()=>ask(`Wipe ${label}`,`Permanently wipe all ${label}?`,`This will delete all ${count} ${label.toLowerCase()} records from Firebase. This CANNOT be undone. Make sure you have a backup first.`,()=>clearS(key)),
      })),
    },
  ];

  return (
    <div style={{display:'flex',flexDirection:'column',gap:20}}>

      {/* Developer banner */}
      <div style={{background:'rgba(109,40,217,0.06)',border:'1px solid rgba(109,40,217,0.22)',borderRadius:14,padding:'14px 20px',display:'flex',alignItems:'center',gap:14}}>
        <Terminal size={18} style={{color:'#6d28d9',flexShrink:0}}/>
        <div>
          <div style={{fontSize:14,fontWeight:700,color:'#6d28d9'}}>Developer Mode — {APP_FULL} · Build {APP_BUILD}</div>
          <div style={{fontSize:12,color:'#7c3aed',marginTop:2}}>Super Admin access only. All destructive actions require confirmation.</div>
        </div>
      </div>

      {/* Company Settings */}
      <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow}}>
        <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:18,display:'flex',alignItems:'center',gap:8}}>
          <Building size={15}/>Company Settings
        </div>
        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:12}}>
          <div style={{gridColumn:'1/-1'}}>
            <label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Company Name *</label>
            <input value={acctSettings.companyName||''} onChange={e=>sf('companyName')(e.target.value)} placeholder="e.g. TDI Workspace Pte. Ltd." style={{...iStyle,fontSize:16,fontWeight:600}}/>
          </div>
          <div><label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>UEN</label><input value={acctSettings.uen||''} onChange={e=>sf('uen')(e.target.value)} placeholder="202312345N" style={iStyle}/></div>
          <div><label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>PayNow UEN</label><input value={acctSettings.payNowUen||''} onChange={e=>sf('payNowUen')(e.target.value)} placeholder="202312345N" style={iStyle}/></div>
          <div><label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Bank Name</label><input value={acctSettings.bankName||''} onChange={e=>sf('bankName')(e.target.value)} placeholder="DBS" style={iStyle}/></div>
          <div><label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Bank Account No.</label><input value={acctSettings.bankAccount||''} onChange={e=>sf('bankAccount')(e.target.value)} placeholder="0721-0976-05" style={iStyle}/></div>
          <div><label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Director / Signatory</label><input value={acctSettings.director||''} onChange={e=>sf('director')(e.target.value)} placeholder="Full name" style={iStyle}/></div>
          <div><label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Phone</label><input value={acctSettings.companyPhone||''} onChange={e=>sf('companyPhone')(e.target.value)} placeholder="+65 6123 4567" style={iStyle}/></div>
          <div><label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Email</label><input value={acctSettings.companyEmail||''} onChange={e=>sf('companyEmail')(e.target.value)} placeholder="hello@company.sg" style={iStyle}/></div>
          <div style={{gridColumn:'1/-1'}}><label style={{fontSize:12,fontWeight:500,color:T.muted,display:'block',marginBottom:6}}>Address</label><input value={acctSettings.companyAddress||''} onChange={e=>sf('companyAddress')(e.target.value)} placeholder="1 Business St, Singapore 123456" style={iStyle}/></div>
        </div>
        <div style={{display:'flex',justifyContent:'flex-end',marginTop:14,alignItems:'center',gap:12}}>
          {coSaved&&<span style={{fontSize:13,color:T.success,fontWeight:600}}>✓ Saved</span>}
          <Btn onClick={handleCoSave}><CheckCircle size={13}/>Save Settings</Btn>
        </div>
      </div>

      {/* AI API Key */}
      <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow}}>
        <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:6,display:'flex',alignItems:'center',gap:8}}>
          <Terminal size={15}/>AI OCR — Anthropic API Key
        </div>
        <div style={{fontSize:12,color:T.muted,marginBottom:14,lineHeight:1.6}}>Required for AI extraction. Get at <strong>console.anthropic.com</strong> → API Keys.</div>
        <div style={{position:'relative'}}>
          <input type={showApiKey?'text':'password'} value={acctSettings.anthropicApiKey||''}
            onChange={e=>{const v=e.target.value;setAcctSettings(prev=>{const u={...prev,anthropicApiKey:v};saveS('acctSettings',u);return u;});}}
            placeholder="sk-ant-api03-..." style={{...iStyle,paddingRight:48,fontFamily:'monospace',fontSize:13}}/>
          <button type="button" onClick={()=>setShowApiKey(s=>!s)}
            style={{position:'absolute',right:10,top:'50%',transform:'translateY(-50%)',background:'none',border:'none',cursor:'pointer',color:T.muted,display:'flex',padding:6}}>
            {showApiKey?<EyeOff size={15}/>:<Eye size={15}/>}
          </button>
        </div>
        {(acctSettings.anthropicApiKey||'').startsWith('sk-ant-')&&<div style={{marginTop:8,fontSize:12,color:T.success,fontWeight:600}}>✓ API key configured — AI OCR enabled</div>}
      </div>

      {/* Data overview */}
      <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow}}>
        <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:14,display:'flex',alignItems:'center',gap:8}}>
          <Database size={15}/>Data Storage Overview
        </div>
        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(130px,1fr))',gap:10}}>
          {storageKeys.map(({key,label,count})=>(
            <div key={key} style={{background:T.bg,borderRadius:10,padding:'12px 14px',border:`1px solid ${T.borderLight}`}}>
              <div style={{fontSize:11,color:T.dim,marginBottom:4}}>{label}</div>
              <div style={{fontSize:20,fontWeight:700,color:T.text}}>{count}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Developer Tools — grouped with confirmations */}
      {TOOLS.map(group=>(
        <div key={group.group} style={{background:T.card,border:`1px solid ${group.group.includes('Danger')?T.danger+'40':T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow}}>
          <div style={{fontSize:14,fontWeight:700,color:group.group.includes('Danger')?T.danger:T.text,marginBottom:14,display:'flex',alignItems:'center',gap:8}}>
            <group.icon size={15} style={{color:group.group.includes('Danger')?T.danger:T.muted}}/>{group.group}
            {group.group.includes('Danger')&&<Badge color={T.danger} sm>Requires confirmation</Badge>}
          </div>
          <div style={{display:'flex',flexDirection:'column',gap:8}}>
            {group.items.map(item=>(
              <div key={item.label} style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:12,
                padding:'11px 14px',borderRadius:10,
                background:item.danger?T.dangerLight:T.bg,
                border:`1px solid ${item.danger?T.danger+'20':T.borderLight}`}}>
                <div style={{minWidth:0}}>
                  <div style={{fontSize:13,fontWeight:600,color:T.text}}>{item.label}</div>
                  <div style={{fontSize:11,color:T.muted,marginTop:2}}>{item.desc}</div>
                </div>
                <Btn variant={item.danger?'danger':'secondary'} size="sm" onClick={item.danger
                  ?()=>item.fn()
                  :item.fn}>
                  {item.danger?<Trash2 size={12}/>:<Download size={12}/>}
                  {item.danger?'Clear':'Export'}
                </Btn>
              </div>
            ))}
          </div>
        </div>
      ))}

      {/* App info */}
      <div style={{background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,padding:22,boxShadow:T.shadow}}>
        <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:14}}>Application Info</div>
        <div style={{display:'flex',flexDirection:'column',gap:8}}>
          {[
            {l:'Version',v:APP_FULL},{l:'Build',v:APP_BUILD},
            {l:'Firebase Project',v:FIREBASE.projectId},
            {l:'Trash retention',v:'12 months'},
            {l:'Action log retention',v:'6 months'},
            {l:'Date',v:new Date().toLocaleDateString('en-SG',{day:'2-digit',month:'long',year:'numeric'})},
          ].map(({l,v})=>(
            <div key={l} style={{display:'flex',justifyContent:'space-between',padding:'8px 12px',background:T.bg,borderRadius:8,alignItems:'center',gap:12}}>
              <span style={{fontSize:12,color:T.muted}}>{l}</span>
              <span style={{fontSize:12,fontWeight:600,color:T.text,fontFamily:'monospace',textAlign:'right'}}>{v}</span>
            </div>
          ))}
        </div>
      </div>

      {/* Confirmation modal */}
      {confirm&&(
        <Modal title="Confirm Action" onClose={()=>setConfirm(null)}>
          <div style={{display:'flex',flexDirection:'column',gap:16}}>
            <div style={{background:T.dangerLight,border:`1px solid ${T.danger}20`,borderRadius:12,padding:'14px 16px',display:'flex',gap:12,alignItems:'flex-start'}}>
              <AlertCircle size={18} style={{color:T.danger,flexShrink:0,marginTop:1}}/>
              <div>
                <div style={{fontSize:14,fontWeight:700,color:T.danger}}>{confirm.label}</div>
                <div style={{fontSize:13,color:T.muted,marginTop:4,lineHeight:1.6}}>{confirm.desc}</div>
              </div>
            </div>
            <div style={{fontSize:13,color:T.muted}}>Type <strong style={{color:T.text}}>CONFIRM</strong> to proceed:</div>
            {(()=>{
              const [input,setInput]=useState('');
              return (<>
                <input value={input} onChange={e=>setInput(e.target.value)} placeholder="Type CONFIRM"
                  style={{...iStyle,fontFamily:'monospace'}}/>
                <div style={{display:'flex',justifyContent:'flex-end',gap:10}}>
                  <Btn variant="secondary" onClick={()=>setConfirm(null)}>Cancel</Btn>
                  <Btn variant="danger" disabled={input!=='CONFIRM'} onClick={()=>{confirm.fn();setConfirm(null);}}>
                    <AlertCircle size={13}/>Execute
                  </Btn>
                </div>
              </>);
            })()}
          </div>
        </Modal>
      )}
    </div>
  );
}

const ALL_NAV=[
  // Group 1 — Project (teal accent)
  {id:'dashboard',   label:'Dashboard',        Icon:LayoutDashboard, group:'project'},
  {id:'projects',    label:'Projects',         Icon:FolderOpen,      group:'project'},
  {id:'payments',    label:'Client Payments',  Icon:CreditCard,      group:'project'},
  {id:'reports',     label:'Reports',          Icon:BarChart3,       group:'project'},
  {id:'warranty',    label:'Warranty',         Icon:CheckCircle,     group:'project'},
  // Group 2 — Expenses (amber accent)
  {id:'invoices',    label:'Supplier Invoices',Icon:Receipt,         group:'expenses'},
  {id:'claims',      label:'Expense Claims',   Icon:DollarSign,      group:'expenses'},
  {id:'commissions', label:'Commissions',      Icon:Users,           group:'expenses'},
  // Group 3 — Admin (slate accent)
  {id:'admin',       label:'Admin',            Icon:Shield,          group:'admin'},
  {id:'workers',     label:'Site Workers',     Icon:Users,           group:'admin'},
  {id:'checkin',     label:'Check In/Out',     Icon:LogIn,           group:'admin'},
  {id:'accounts',    label:'Company Accounts', Icon:Building,        group:'admin'},
  {id:'contacts',    label:'Contacts',         Icon:BookUser,        group:'admin'},
  {id:'trash',       label:'Trash',            Icon:Trash,           group:'admin'},
  {id:'system',      label:'System',           Icon:Terminal,        group:'admin'},
];

// Group metadata for sidebar visual styling
const NAV_GROUPS={
  project: {label:'Project',  color:'#0891b2', bg:'rgba(8,145,178,0.08)',  dot:'#0891b2'},
  expenses:{label:'Expenses', color:'#d97706', bg:'rgba(217,119,6,0.08)',  dot:'#d97706'},
  admin:   {label:'Admin',    color:'#64748b', bg:'rgba(100,116,139,0.08)',dot:'#64748b'},
};

export default function App(){
  // ── Dark mode ─────────────────────────────────────────────────────────────
  const [darkMode,setDarkMode]=useState(()=>{
    try{ return localStorage.getItem('rl_dark')==='1'; }catch{ return false; }
  });
  // Keep T and iStyle in sync with dark mode — called before every render
  T = THEMES[darkMode?'dark':'light'];
  iStyle = makeIStyle();
  useEffect(()=>{
    try{ localStorage.setItem('rl_dark',darkMode?'1':'0'); }catch{}
    document.body.style.background = T.bg;
  },[darkMode]);

  // ── Offline detection ─────────────────────────────────────────────────────
  const [isOnline,setIsOnline]=useState(()=>navigator.onLine);
  const [offlineQueue,setOfflineQueue]=useState([]);
  useEffect(()=>{
    const goOn =()=>setIsOnline(true);
    const goOff=()=>setIsOnline(false);
    window.addEventListener('online', goOn);
    window.addEventListener('offline',goOff);
    return ()=>{ window.removeEventListener('online',goOn); window.removeEventListener('offline',goOff); };
  },[]);

  const [tab,setTab]=useState('dashboard');
  const [projects,setProjects]=useState(()=>{
    // Seed from localStorage cache for instant load
    try{ const c=localStorage.getItem('rl_cache_projects'); return c?JSON.parse(c):SEED_PROJ; }catch{ return SEED_PROJ; }
  });
  const [invoices,setInvoices]=useState(()=>{
    try{ const c=localStorage.getItem('rl_cache_invoices'); return c?JSON.parse(c):SEED_INV; }catch{ return SEED_INV; }
  });
  const [payments,setPayments]=useState(()=>{
    try{ const c=localStorage.getItem('rl_cache_payments'); return c?JSON.parse(c):SEED_PAY; }catch{ return SEED_PAY; }
  });
  const [users,setUsers]=useState(()=>{
    try{ const c=localStorage.getItem('rl_cache_users'); return c?JSON.parse(c):SEED_USERS; }catch{ return SEED_USERS; }
  });
  const [warranties,setWarranties]=useState(SEED_WARRANTIES);
  const [trash,setTrash]=useState([]);
  const [actionLog,setActionLog]=useState([]); // audit log — 6 months of user actions
  const [toast,setToast]=useState(null);
  const [syncStatus,setSyncStatus]=useState(null); // null|'saving'|'saved'|'error'
  const [syncError,setSyncError]=useState('');

  // Wire global sync reporter to this component's state
  useEffect(()=>{
    _onSyncStatus=(status,err)=>{
      setSyncStatus(status);
      setSyncError(err||'');
      if(status==='saved') setTimeout(()=>setSyncStatus(null),3000);
    };
    return ()=>{ _onSyncStatus=null; };
  },[]);
  const [acctSettings,setAcctSettings]=useState(SEED_ACCT_SETTINGS);
  const [siteWorkers,setSiteWorkers]=useState(SEED_WORKERS);
  const [attendance,setAttendance]=useState(SEED_ATTENDANCE);
  const [workerClaims,setWorkerClaims]=useState([]);
  const [staffClaims,setStaffClaims]=useState([]);
  const [invoiceBatches,setInvoiceBatches]=useState([]);
  const [notices,setNotices]=useState([]);

  const [activeUserId,setActiveUserId]=useState(null); // null = not logged in
  const [workerSession,setWorkerSession]=useState(null);
  const [showWorkerLogin,setShowWorkerLogin]=useState(false);
  const [ready,setReady]=useState(false);
  const [isMobile,setIsMobile]=useState(()=>typeof window!=='undefined'&&window.innerWidth<768);
  const [notifOpen,setNotifOpen]=useState(false);
  const [notifSeen,setNotifSeen]=useState(()=>{
    try{ return new Set(JSON.parse(localStorage.getItem('rl_notif_seen')||'[]')); }catch{ return new Set(); }
  });

  const notifications = useMemo(()=>{
    const items=[];
    const now=new Date();
    const _isAdmin = activeUserId==='__sa__' || (users.find(u=>u.id===activeUserId)?.role==='admin');

    // 1. Overdue client payments — projects past end date with outstanding balance
    projects.filter(p=>!p.archived&&p.status!=='Cancelled'&&p.endDate).forEach(p=>{
      const endD=new Date(p.endDate);
      const overdueDays=Math.floor((now-endD)/864e5);
      if(overdueDays>0){
        const rev=(p.contractAmount||0)+(p.variationOrders||0);
        const recv=payments.filter(py=>py.projectId===p.id&&py.status==='Received').reduce((s,py)=>s+py.amount,0);
        const outstanding=rev-recv;
        if(outstanding>0) items.push({id:`overdue_${p.id}`,type:'danger',title:'Overdue Payment',
          body:`${p.name} — ${fmtSGD(outstanding)} outstanding, ${overdueDays}d past end date`,
          action:()=>{setTab('payments');setNotifOpen(false);},at:endD.toISOString()});
      }
    });

    // 2. Pending supplier invoices >30 days old
    invoices.filter(i=>i.status==='Pending').forEach(i=>{
      const age=Math.floor((now-new Date(i.invoiceDate||now))/864e5);
      if(age>30) items.push({id:`inv_aging_${i.id}`,type:'warning',title:'Aging Invoice',
        body:`${i.supplier} — ${i.invoiceNo||''} — ${fmtSGD(i.total)} unpaid for ${age} days`,
        action:()=>{setTab('invoices');setNotifOpen(false);},at:i.invoiceDate});
    });

    // 3. Worker document expiry — within 60 days
    siteWorkers.filter(w=>w.status==='Active').forEach(w=>{
      const wpDays=daysUntil(w.workPassExpiry);
      if(wpDays!==null&&wpDays<=60) items.push({id:`wp_${w.id}`,type:wpDays<=14?'danger':'warning',
        title:'Work Pass Expiring',body:`${w.name} — Work Pass expires in ${wpDays} day${wpDays!==1?'s':''}`,
        action:()=>{setTab('workers');setNotifOpen(false);},at:w.workPassExpiry});
      (w.certificates||[]).forEach(c=>{
        if(!c.expiryDate)return;
        const cd=daysUntil(c.expiryDate);
        if(cd!==null&&cd<=60) items.push({id:`cert_${w.id}_${c.id}`,type:cd<=14?'danger':'warning',
          title:'Certificate Expiring',body:`${w.name} — ${c.name} expires in ${cd} day${cd!==1?'s':''}`,
          action:()=>{setTab('workers');setNotifOpen(false);},at:c.expiryDate});
      });
    });

    // 4. Projects starting within 7 days
    projects.filter(p=>p.status==='Planning'&&p.startDate).forEach(p=>{
      const startDays=Math.floor((new Date(p.startDate)-now)/864e5);
      if(startDays>=0&&startDays<=7) items.push({id:`starting_${p.id}`,type:'info',title:'Project Starting Soon',
        body:`${p.name} starts in ${startDays===0?'today':`${startDays} day${startDays!==1?'s':''}`}`,
        action:()=>{setTab('projects');setNotifOpen(false);},at:p.startDate});
    });

    // 5. Pending staff expense claims
    const pendingClaims=staffClaims.filter(c=>c.status==='Pending');
    if(pendingClaims.length>0&&_isAdmin) items.push({id:'pending_claims',type:'info',title:'Pending Expense Claims',
      body:`${pendingClaims.length} staff claim${pendingClaims.length>1?'s':''} awaiting approval`,
      action:()=>{setTab('claims');setNotifOpen(false);},at:pendingClaims[0]?.submittedAt||now.toISOString()});

    return items.sort((a,b)=>{
      const order={danger:0,warning:1,info:2};
      return (order[a.type]||3)-(order[b.type]||3);
    });
  },[projects,invoices,payments,siteWorkers,staffClaims,activeUserId,users]);

  const unreadCount=notifications.filter(n=>!notifSeen.has(n.id)).length;

  const markAllSeen=()=>{
    const newSeen=new Set([...notifSeen,...notifications.map(n=>n.id)]);
    setNotifSeen(newSeen);
    try{ localStorage.setItem('rl_notif_seen',JSON.stringify([...newSeen])); }catch{}
  };

  const [moreOpen,setMoreOpen]=useState(false);
  const [onlinePresence,setOnlinePresence]=useState([]); // [{userId,name,role,photo,loginAt,lastSeen,currentTab}]

  useEffect(()=>{
    const onResize=()=>setIsMobile(window.innerWidth<768);
    window.addEventListener('resize',onResize);
    return ()=>window.removeEventListener('resize',onResize);
  },[]);

  // Presence: write/update own record when logged in, delete on logout, heartbeat every 30s
  const presenceKey = activeUserId ? `presence:${activeUserId}` : null;

  // Derive activeUser early so presence useEffects can reference it
  const isSuperAdmin = activeUserId === '__sa__';
  const activeUser = isSuperAdmin ? _sa : (users.find(u=>u.id===activeUserId)||null);

  // Sanitize userId for use as Firestore document ID (strip leading/trailing underscores)
  const safePresenceKey = (uid) => `presence__${uid.replace(/^_+|_+$/g, 'x')}`;

  useEffect(()=>{
    if(!activeUserId||!activeUser||isSuperAdmin) return; // super admin is local — skip Firebase presence
    const record = {
      userId: activeUserId, name: activeUser.name, role: activeUser.role,
      photo: activeUser.photo||'', loginAt: new Date().toISOString(),
      lastSeen: new Date().toISOString(), currentTab: tab,
    };
    // Write presence to Firestore
    saveS(safePresenceKey(activeUserId), record);

    const heartbeat = setInterval(()=>{
      saveS(safePresenceKey(activeUserId), {...record, lastSeen: new Date().toISOString(), currentTab: tab});
    }, 30000);

    return ()=>{ clearInterval(heartbeat); };
  },[activeUserId, activeUser?.name, tab]);

  // Update currentTab in presence on tab change
  useEffect(()=>{
    if(!activeUserId||!activeUser||isSuperAdmin) return;
    loadS(safePresenceKey(activeUserId), null).then(d=>{
      if(d) saveS(safePresenceKey(activeUserId), {...d, lastSeen: new Date().toISOString(), currentTab: tab});
    });
  },[tab, activeUserId]);

  // Poll presence — list all presence__ docs every 20s via Firestore REST
  useEffect(()=>{
    const readPresence = async() => {
      try{
        const r = await fetch(`${FS_BASE}?key=${FIREBASE.apiKey}&pageSize=50`);
        if(!r.ok){ setOnlinePresence([]); return; }
        const data = await r.json();
        const now = Date.now();
        const docs = (data.documents||[])
          .filter(doc=>doc.name?.includes('presence__'))
          .map(doc=>{
            try{ return JSON.parse(doc.fields?.value?.stringValue||'null'); }catch{ return null; }
          })
          .filter(d=>d&&(now-new Date(d.lastSeen).getTime())<120000);
        setOnlinePresence(docs);
      }catch{ setOnlinePresence([]); }
    };
    readPresence();
    const poll = setInterval(readPresence, 20000);
    return ()=>clearInterval(poll);
  },[]);

  // Delete own presence on logout
  const logout = useCallback(()=>{
    if(activeUserId && !isSuperAdmin) clearS(safePresenceKey(activeUserId));
    setActiveUserId(null);
    setTab('dashboard');
  },[activeUserId]);


  // Log a user action — kept 6 months, saved to Firebase
  const logAction = useCallback((action, detail, snapshot=null)=>{
    const entry={
      id: uid(),
      action,           // e.g. 'DELETE_PROJECT', 'RESTORE_INVOICE', 'CREATE_USER'
      detail,           // human-readable description
      userId: activeUserId,
      userName: activeUser?.name||'Unknown',
      userRole: activeUser?.role||'unknown',
      at: new Date().toISOString(),
      snapshot,         // original data for undo
    };
    setActionLog(prev=>{
      const sixMonthsAgo=new Date();
      sixMonthsAgo.setMonth(sixMonthsAgo.getMonth()-6);
      const fresh=[...prev.filter(e=>new Date(e.at)>sixMonthsAgo), entry];
      saveS('actionLog', fresh);
      return fresh;
    });
  },[activeUserId, activeUser]);

  const handleSoftDelete = useCallback((item)=>{
    setTrash(prev=>{const upd=[...prev,item];saveS('trash',upd);return upd;});
    const typeName={project:'Project',invoice:'Invoice',payment:'Payment',user:'User',staffClaim:'Expense Claim'}[item._trashType]||item._trashType;
    const label=item.name||item.invoiceNo||item.email||item.id;
    logAction(`DELETE_${(item._trashType||'').toUpperCase()}`, `Deleted ${typeName}: ${label}`, item);
  },[logAction]);

  const handleShowToast = useCallback((message,undoFn)=>{
    setToast({message,undoFn});
  },[]);

  const handleRestore = useCallback((item)=>{
    const {_trashType,_deletedAt,...orig}=item;
    if(_trashType==='project'){setProjects(p=>{const u=[...p,orig];saveProjects(u);return u;});}
    else if(_trashType==='invoice'){setInvoices(p=>{const u=[...p,orig];saveInvoices(u);return u;});}
    else if(_trashType==='payment'){setPayments(p=>{const u=[...p,orig];saveS('payments',u);return u;});}
    else if(_trashType==='user'){setUsers(p=>{const u=[...p,orig];saveUsers(u);return u;});}
    else if(_trashType==='staffClaim'){
      setStaffClaims(p=>{const u=[...p,orig];saveS('staffClaims',u);return u;});
      setTrash(prev=>{
        const linkedInTrash=prev.filter(t=>t._trashType==='invoice'&&t._fromClaim===orig.id);
        if(linkedInTrash.length>0){
          setInvoices(p=>{
            const restored=linkedInTrash.map(({_trashType,_deletedAt,...inv})=>inv);
            const u=[...p,...restored];saveInvoices(u);return u;
          });
          return prev.filter(t=>!(t._trashType==='invoice'&&t._fromClaim===orig.id));
        }
        return prev;
      });
    }
    setTrash(prev=>{const upd=prev.filter(t=>t.id!==item.id);saveS('trash',upd);return upd;});
    const typeName={project:'Project',invoice:'Invoice',payment:'Payment',user:'User',staffClaim:'Expense Claim'}[_trashType]||_trashType;
    const label=orig.name||orig.invoiceNo||orig.email||orig.id;
    logAction(`RESTORE_${(_trashType||'').toUpperCase()}`, `Restored ${typeName}: ${label}`, orig);
  },[logAction]);

  const handlePermanentDelete = useCallback((id)=>{
    setTrash(prev=>{
      const item=prev.find(t=>t.id===id);
      if(item){
        const typeName={project:'Project',invoice:'Invoice',payment:'Payment',user:'User',staffClaim:'Expense Claim'}[item._trashType]||item._trashType;
        logAction('PERMANENT_DELETE', `Permanently deleted ${typeName}: ${item.name||item.invoiceNo||item.email||id}`, item);
      }
      const upd=prev.filter(t=>t.id!==id);saveS('trash',upd);return upd;
    });
  },[logAction]);

  const [lastSync,setLastSync]=useState(null); // timestamp of last successful Firebase load
  const [syncing,setSyncing]=useState(false);

  const loadAllData = useCallback(async()=>{
    setSyncing(true);
    try{
      const [p,i,py,us,ws,tr,as,sw,att,wc,sc,ib,al,no]=await Promise.all([
        loadS('projects',SEED_PROJ),
        loadS('invoices',SEED_INV),
        loadS('payments',SEED_PAY),
        loadS('users',SEED_USERS),
        loadS('warranties',SEED_WARRANTIES),
        loadS('trash',[]),
        loadS('acctSettings',SEED_ACCT_SETTINGS),
        loadS('siteWorkers',SEED_WORKERS),
        loadS('attendance',SEED_ATTENDANCE),
        loadS('workerClaims',[]),
        loadS('staffClaims',[]),
        loadS('invoiceBatches',[]),
        loadS('actionLog',[]),
        loadS('notices',[]),
      ]);
      let finalProjects = Array.isArray(p) ? p : SEED_PROJ;
      // Rehydrate quotation files and VO files from separate per-project keys
      if(finalProjects.length > 0){
        const fileResults = await Promise.all(
          finalProjects.map(proj => loadS(`proj_file_${proj.id}`, null))
        );
        finalProjects = finalProjects.map((proj, i) => {
          const fd = fileResults[i];
          if(!fd) return proj;
          const rehydrated = {...proj};
          if(fd.quotationFile) rehydrated.quotationFile = fd.quotationFile;
          if(fd.quotationFilename && !proj.quotationFilename) rehydrated.quotationFilename = fd.quotationFilename;
          if(proj.voList){
            rehydrated.voList = proj.voList.map(vo => {
              const vf = fd[`vo_${vo.id}`];
              return vf ? {...vo, file: vf} : vo;
            });
          }
          return rehydrated;
        });
      }
      setProjects(finalProjects);
      // Cache to localStorage for offline
      try{ localStorage.setItem('rl_cache_projects',JSON.stringify(finalProjects)); }catch{}
      let rawInvoices=Array.isArray(i)?i:SEED_INV;
      // Rehydrate proof images from per-invoice file documents
      if(rawInvoices.length>0){
        const invFileResults=await Promise.all(rawInvoices.map(inv=>loadS(`inv_file_${inv.id}`,null)));
        rawInvoices=rawInvoices.map((inv,idx)=>{
          const fd=invFileResults[idx];
          if(!fd) return inv;
          const r={...inv};
          if(fd.proofImage) r.proofImage=fd.proofImage;
          if(inv.paymentRecords)
            r.paymentRecords=inv.paymentRecords.map(pr=>{
              const rp=fd[`pay_${pr.id}`];
              return rp?{...pr,proofImage:rp}:pr;
            });
          return r;
        });
      }
      setInvoices(rawInvoices);
      // Cache without images to avoid localStorage quota issues
      try{ localStorage.setItem('rl_cache_invoices',JSON.stringify(rawInvoices.map(inv=>({...inv,proofImage:null})))); }catch{}
      const rawPay=Array.isArray(py)?py:SEED_PAY;
      setPayments(rawPay);
      try{ localStorage.setItem('rl_cache_payments',JSON.stringify(rawPay)); }catch{}
      const rawUsers=Array.isArray(us)&&us.length>0?us:SEED_USERS;
      const cleanUsers=rawUsers.map(u=>(!u.photo||u.photo.length<=66666)?u:{...u,photo:''});
      setUsers(cleanUsers);
      try{ localStorage.setItem('rl_cache_users',JSON.stringify(cleanUsers)); }catch{}
      setWarranties(Array.isArray(ws)?ws:SEED_WARRANTIES);
      setAcctSettings({...SEED_ACCT_SETTINGS,...(as&&typeof as==='object'?as:{})});
      setSiteWorkers(Array.isArray(sw)?sw:SEED_WORKERS);
      setAttendance(Array.isArray(att)?att:SEED_ATTENDANCE);
      setWorkerClaims(Array.isArray(wc)?wc:[]);
      setStaffClaims(Array.isArray(sc)?sc:[]);
      setInvoiceBatches(Array.isArray(ib)?ib:[]);
      setNotices(Array.isArray(no)?no:[]);
      // Trash kept for 12 months (previously 30 days)
      const twelveMonthsAgo=Date.now()-365*24*60*60*1000;
      const freshTrash=(Array.isArray(tr)?tr:[]).filter(t=>new Date(t._deletedAt).getTime()>twelveMonthsAgo);
      setTrash(freshTrash);
      if(freshTrash.length!==(Array.isArray(tr)?tr:[]).length) saveS('trash',freshTrash);
      // Action log — keep 6 months
      const sixMonthsAgo=new Date();sixMonthsAgo.setMonth(sixMonthsAgo.getMonth()-6);
      const freshLog=(Array.isArray(al)?al:[]).filter(e=>new Date(e.at)>sixMonthsAgo);
      setActionLog(freshLog);
      setLastSync(new Date());
    }catch(err){
      console.error('Storage load error:',err);
      setUsers(SEED_USERS);setAcctSettings(SEED_ACCT_SETTINGS);
    }finally{setSyncing(false);}
  },[]);

  useEffect(()=>{
    // Safety timeout — Firebase can take a few seconds on first load
    const safetyTimer=setTimeout(()=>setReady(true),8000);
    (async()=>{
      await loadAllData();
      clearTimeout(safetyTimer);
      setReady(true);
    })();
  },[]);

  // Auto-refresh when user returns to the app (switches tabs, unlocks phone, etc.)
  useEffect(()=>{
    const onVisible=()=>{
      if(document.visibilityState==='visible'&&activeUserId){
        loadAllData();
      }
    };
    document.addEventListener('visibilitychange',onVisible);
    return ()=>document.removeEventListener('visibilitychange',onVisible);
  },[activeUserId,loadAllData]);

  // Super admin sees all tabs; regular users see their role-assigned tabs
  const visibleNav = ALL_NAV.filter(n=>(activeUser?.tabs||[]).includes(n.id));

  // Data scope — super admin and admin see everything
  const userProjects = (isSuperAdmin||activeUser?.role==='admin'||activeUser?.role==='accounts')
    ? projects
    : projects.filter(p=>(activeUser?.assignedProjects||[]).includes(p.id));

  const pendCount=invoices.filter(i=>i.status==='Pending').length;
  const trashCount=trash.length;
  const isAdmin = isSuperAdmin || activeUser?.role==='admin';


  useEffect(()=>{
    if(ready && !visibleNav.find(n=>n.id===tab)){
      setTab(visibleNav[0]?.id||'dashboard');
    }
  },[activeUserId,ready]);

  const TAB_SUB={
    dashboard:`${projects.length} projects . ${invoices.length} invoices . ${payments.length} payments`,
    projects:'Manage your renovation and interior design projects',
    invoices:'Capture and manage supplier invoices with AI-assisted OCR extraction',
    payments:'Track payments received from clients and outstanding receivables',
    contacts:'Client directory and supplier & subcontractor records',
    reports:'Project P&L reports with AI financial analysis',
    commissions:'Automatic commission calculations for designers and PMs',
    warranty:'Warranty certificates, serial search and claims management',
    workers:'Site worker management — rates, work pass, certificates, payroll',
    checkin:'Worker check-in & check-out — tap to record your attendance',
    accounts:'Company Accounts — Bank reconciliation, IRAS tax submission, 4-line statement',
    trash:`${trashCount} item${trashCount!==1?'s':''} — items auto-purge after 30 days`,
    admin:'User management, roles, and dashboard visibility controls',
    system:`Developer panel — ${APP_FULL} . Build ${APP_BUILD}`,
  };

  if(!FIREBASE_CONFIGURED) return <FirebaseSetupBanner/>;
  if(!ready) return (
    <div style={{background:T.bg,minHeight:'100vh',display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',
      fontFamily:'"DM Sans",-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',gap:20}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@500&family=DM+Serif+Display&display=swap');@keyframes blink{0%,100%{opacity:0.2}50%{opacity:1}}.ld1{animation:blink 1.2s ease-in-out infinite 0s}.ld2{animation:blink 1.2s ease-in-out infinite 0.2s}.ld3{animation:blink 1.2s ease-in-out infinite 0.4s}`}</style>
      <div style={{display:'flex',alignItems:'center',gap:12}}>
        <div style={{width:44,height:44,background:T.text,borderRadius:13,display:'flex',alignItems:'center',justifyContent:'center',boxShadow:'0 4px 16px rgba(26,26,26,0.20)'}}>
          <Building size={20} style={{color:'#fff'}}/>
        </div>
        <div style={{fontFamily:'"DM Serif Display",Georgia,serif',fontSize:26,color:T.text,lineHeight:1}}>
          RenoLedger
        </div>
      </div>
      <div style={{display:'flex',gap:8,alignItems:'center'}}>
        <div className="ld1" style={{width:7,height:7,borderRadius:'50%',background:T.tan}}/>
        <div className="ld2" style={{width:7,height:7,borderRadius:'50%',background:T.tan}}/>
        <div className="ld3" style={{width:7,height:7,borderRadius:'50%',background:T.tan}}/>
      </div>
    </div>
  );

  // Worker portal session - show standalone fullscreen portal
  if(workerSession) return (
    <div style={{fontFamily:'-apple-system,BlinkMacSystemFont,"SF Pro Display","SF Pro Text",system-ui,sans-serif'}}>
      <style>{`*{box-sizing:border-box;margin:0;padding:0;}@keyframes spin{to{transform:rotate(360deg);}}`}</style>
      <WorkerPortal worker={workerSession} onLogout={()=>setWorkerSession(null)} attendance={attendance} setAttendance={setAttendance} projects={projects} claims={workerClaims} setClaims={setWorkerClaims}/>
    </div>
  );

  // Worker login / check-in portal (standalone, no staff session needed)
  if(showWorkerLogin) return (
    <div style={{fontFamily:'-apple-system,BlinkMacSystemFont,"SF Pro Display","SF Pro Text",system-ui,sans-serif'}}>
      <style>{`*{box-sizing:border-box;margin:0;padding:0;}@keyframes spin{to{transform:rotate(360deg);}}`}</style>
      <WorkerLoginScreen
        siteWorkers={siteWorkers}
        onLogin={(w)=>{setWorkerSession(w);setShowWorkerLogin(false);}}
        onAdminLogin={()=>setShowWorkerLogin(false)}
        acctSettings={acctSettings}
      />
    </div>
  );

  // Not logged in — show main login screen
  if(!activeUserId) return (
    <LoginScreen
      users={users}
      siteWorkers={siteWorkers}
      onStaffLogin={(user)=>setActiveUserId(user.id)}
      onWorkerPortal={()=>setShowWorkerLogin(true)}
        acctSettings={acctSettings}
    />
  );


  return (
    <div style={{background:T.bg,minHeight:'100vh',fontFamily:'"DM Sans",-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',color:T.text}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Serif+Display&display=swap');
        html,body{font-size:15px;background:${T.bg};}
        *{box-sizing:border-box;margin:0;padding:0;}
        ::placeholder{color:${T.dim};}
        ::-webkit-scrollbar{width:4px;height:4px;}
        ::-webkit-scrollbar-track{background:transparent;}
        ::-webkit-scrollbar-thumb{background:${T.borderLight};border-radius:4px;}
        input[type=number]::-webkit-inner-spin-button{opacity:0;}
        input:focus,select:focus{border-color:${T.text} !important;box-shadow:0 0 0 3px ${T.accentLight} !important;outline:none;}
        @keyframes spin{to{transform:rotate(360deg);}}
        @keyframes fadeIn{from{opacity:0;transform:translateY(4px);}to{opacity:1;transform:translateY(0);}}
        @keyframes slideDown{from{opacity:0;transform:translateY(-8px);}to{opacity:1;transform:translateY(0);}}
        .page-content{animation:fadeIn 0.18s ease;}
        button{-webkit-tap-highlight-color:transparent;font-family:"DM Sans",-apple-system,sans-serif;}
        select option{background:${T.card};color:${T.text};}
      `}</style>

      {/* ── Offline banner ── */}
      {!isOnline&&(
        <div style={{position:'fixed',top:0,left:0,right:0,zIndex:999,background:'#9A6A00',
          color:'#fff',textAlign:'center',fontSize:12,fontWeight:600,padding:'8px',
          display:'flex',alignItems:'center',justifyContent:'center',gap:8}}>
          <AlertCircle size={13}/>
          You're offline — showing cached data. Changes will sync when connection returns.
        </div>
      )}

      {/* ── Notification panel ── */}
      {notifOpen&&(
        <div style={{position:'fixed',inset:0,zIndex:200}} onClick={()=>setNotifOpen(false)}>
          <div style={{position:'absolute',top:isMobile?52:60,right:isMobile?8:16,
            width:Math.min(360,window.innerWidth-16),
            background:T.card,border:`1px solid ${T.borderLight}`,borderRadius:18,
            boxShadow:T.shadowLg,overflow:'hidden',animation:'slideDown 0.18s ease'}}
            onClick={e=>e.stopPropagation()}>
            <div style={{padding:'16px 18px 12px',borderBottom:`1px solid ${T.borderLight}`,
              display:'flex',alignItems:'center',justifyContent:'space-between'}}>
              <div style={{fontWeight:700,fontSize:14,color:T.text}}>Notifications {notifications.length>0&&`(${notifications.length})`}</div>
              <div style={{display:'flex',gap:10,alignItems:'center'}}>
                {unreadCount>0&&(
                  <button onClick={markAllSeen} style={{background:'none',border:'none',
                    fontSize:11,color:T.muted,cursor:'pointer',fontFamily:'inherit'}}>Mark all read</button>
                )}
                <button onClick={()=>setNotifOpen(false)} style={{background:'none',border:'none',cursor:'pointer',color:T.dim,display:'flex'}}><X size={14}/></button>
              </div>
            </div>
            <div style={{maxHeight:380,overflowY:'auto',WebkitOverflowScrolling:'touch'}}>
              {notifications.length===0?(
                <div style={{padding:'32px 18px',textAlign:'center',color:T.muted,fontSize:13}}>
                  <div style={{fontSize:28,marginBottom:8}}>✓</div>All clear — no alerts
                </div>
              ):notifications.map(n=>{
                const isNew=!notifSeen.has(n.id);
                const NCLR={danger:T.danger,warning:T.warning,info:T.info};
                const NBKG={danger:T.dangerLight,warning:T.warningLight,info:T.infoLight};
                const NICON={danger:'⚠️',warning:'🕐',info:'ℹ️'};
                return (
                  <div key={n.id} onClick={()=>{n.action&&n.action();markAllSeen();}}
                    style={{padding:'12px 18px',borderBottom:`1px solid ${T.borderLight}`,
                      cursor:'pointer',background:isNew?NBKG[n.type]:'transparent',
                      display:'flex',gap:10,alignItems:'flex-start'}}>
                    <span style={{fontSize:16,flexShrink:0,marginTop:1}}>{NICON[n.type]}</span>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{display:'flex',alignItems:'center',gap:6,marginBottom:2}}>
                        <span style={{fontSize:12,fontWeight:700,color:NCLR[n.type]}}>{n.title}</span>
                        {isNew&&<span style={{width:6,height:6,borderRadius:'50%',background:NCLR[n.type],display:'inline-block'}}/>}
                      </div>
                      <div style={{fontSize:12,color:T.muted,lineHeight:1.5}}>{n.body}</div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}
      {!isMobile&&(
        <div style={{position:'fixed',left:0,top:0,bottom:0,width:240,
          background:T.card,borderRight:`1px solid ${T.borderLight}`,
          display:'flex',flexDirection:'column',zIndex:20,overflowY:'auto'}}>
          {/* Logo area */}
          <div style={{padding:'24px 20px 20px',borderBottom:`1px solid ${T.borderLight}`}}>
            <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:14}}>
              <div style={{width:34,height:34,background:T.text,borderRadius:10,
                display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0}}>
                <Building size={16} style={{color:'#fff'}}/>
              </div>
              <div>
                <div style={{fontFamily:'"DM Serif Display",Georgia,serif',fontSize:17,color:T.text,lineHeight:1}}>
                  RenoLedger
                </div>
                <div style={{fontSize:10,color:T.dim,marginTop:1}}>{acctSettings?.companyName||'TDI Workspace'}</div>
              </div>
            </div>
            {/* Sync status */}
            <div style={{display:'flex',alignItems:'center',justifyContent:'space-between'}}>
              <span style={{fontSize:10,color:syncing?T.tan:syncStatus==='error'?T.danger:T.dim,display:'flex',alignItems:'center',gap:4}}>
                {syncing&&<Loader2 size={9} style={{animation:'spin 1s linear infinite'}}/>}
                {syncStatus==='error'&&<AlertCircle size={9}/>}
                {syncStatus==='saved'&&<CheckCircle size={9} style={{color:T.success}}/>}
                {syncing?'Syncing…':syncStatus==='saved'?'Saved':syncStatus==='error'?
                  <span style={{cursor:'pointer'}} onClick={()=>alert(`Sync failed: ${syncError}`)}>Not synced</span>
                  :lastSync?`${lastSync.getHours().toString().padStart(2,'0')}:${lastSync.getMinutes().toString().padStart(2,'0')}`:''}
              </span>
              <div style={{display:'flex',alignItems:'center',gap:4}}>
                <button onClick={()=>{setNotifOpen(o=>!o);if(!notifOpen)markAllSeen();}}
                  style={{background:'none',border:'none',cursor:'pointer',position:'relative',
                    padding:4,color:T.muted,display:'flex',alignItems:'center'}}>
                  <Bell size={14}/>
                  {unreadCount>0&&(
                    <span style={{position:'absolute',top:0,right:0,width:14,height:14,
                      background:T.danger,borderRadius:'50%',fontSize:8,fontWeight:700,
                      color:'#fff',display:'flex',alignItems:'center',justifyContent:'center'}}>
                      {unreadCount>9?'9+':unreadCount}
                    </span>
                  )}
                </button>
                <button onClick={()=>loadAllData()} disabled={syncing}
                  style={{background:'none',border:`1px solid ${T.borderLight}`,borderRadius:6,padding:'2px 8px',
                    cursor:'pointer',fontSize:10,color:T.dim,fontFamily:'inherit',display:'flex',alignItems:'center',gap:3}}>
                  <RefreshCw size={9} style={syncing?{animation:'spin 1s linear infinite'}:{}}/>Refresh
                </button>
              </div>
            </div>
            {isSuperAdmin&&(
              <div style={{marginTop:8,background:'rgba(109,40,217,0.06)',border:'1px solid rgba(109,40,217,0.18)',borderRadius:8,padding:'5px 10px',display:'flex',alignItems:'center',gap:6}}>
                <Terminal size={10} style={{color:'#6d28d9'}}/><span style={{fontSize:10,fontWeight:700,color:'#6d28d9'}}>Developer Mode</span>
              </div>
            )}
          </div>

          {/* User card */}
          <div style={{padding:'14px 16px',borderBottom:`1px solid ${T.borderLight}`}}>
            <div style={{display:'flex',alignItems:'center',gap:10}}>
              <Avatar photo={activeUser?.photo} name={activeUser?.name||'?'} size={32} color={ROLE_CLR[activeUser?.role]||T.tan}/>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:13,fontWeight:600,color:T.text,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{activeUser?.name}</div>
                <div style={{fontSize:11,color:T.muted}}>{ROLE_LABEL[activeUser?.role]||'Staff'}</div>
              </div>
            </div>
            <button onClick={logout} style={{background:'none',border:'none',cursor:'pointer',
              fontSize:11,color:T.dim,fontFamily:'inherit',padding:'2px 0',whiteSpace:'nowrap'}}>
              Sign out
            </button>
            {/* Dark mode toggle */}
            <button onClick={()=>setDarkMode(d=>!d)}
              style={{background:'none',border:'none',cursor:'pointer',fontSize:16,padding:'2px 0',lineHeight:1}}
              title={darkMode?'Switch to light mode':'Switch to dark mode'}>
              {darkMode?'☀️':'🌙'}
            </button>
          </div>

          {/* Nav groups */}
          <nav style={{flex:1,padding:'8px 10px',display:'flex',flexDirection:'column',gap:1,overflowY:'auto'}}>
            {(()=>{
              const groups=['project','expenses','admin'];
              return groups.map(gKey=>{
                const gItems=visibleNav.filter(n=>n.group===gKey);
                if(!gItems.length) return null;
                const g=NAV_GROUPS[gKey];
                return (
                  <div key={gKey} style={{marginBottom:4}}>
                    <div style={{padding:'8px 12px 4px',fontSize:9,fontWeight:700,
                      color:T.dim,textTransform:'uppercase',letterSpacing:'0.12em'}}>
                      {g.label}
                    </div>
                    {gItems.map(({id,label,Icon})=>{
                      const active=tab===id;
                      const badge=id==='invoices'?pendCount:id==='trash'?trashCount:0;
                      return (
                        <button key={id} onClick={()=>setTab(id)}
                          style={{display:'flex',alignItems:'center',gap:9,padding:'8px 12px',
                            borderRadius:8,border:'none',cursor:'pointer',textAlign:'left',
                            width:'100%',fontFamily:'inherit',fontSize:13,
                            fontWeight:active?600:400,transition:'background 0.1s',
                            background:active?T.tanLight:'transparent',
                            color:active?T.text:T.muted}}>
                          <Icon size={14} style={{flexShrink:0,opacity:active?1:0.6}}/>
                          <span style={{flex:1}}>{label}</span>
                          {badge>0&&<span style={{fontSize:9,background:T.danger,color:'#fff',
                            borderRadius:8,padding:'1px 5px',fontWeight:700}}>{badge}</span>}
                        </button>
                      );
                    })}
                  </div>
                );
              });
            })()}
          </nav>
        </div>
      )}

      {/* MOBILE — top bar */}
      {isMobile&&(
        <div style={{position:'fixed',top:0,left:0,right:0,zIndex:30,height:52,
          background:T.card,borderBottom:`1px solid ${T.borderLight}`,
          display:'flex',alignItems:'center',justifyContent:'space-between',padding:'0 16px',
          boxShadow:'0 1px 0 rgba(0,0,0,0.04)'}}>
          <div style={{display:'flex',alignItems:'center',gap:8}}>
            <div style={{width:28,height:28,background:T.text,borderRadius:8,
              display:'flex',alignItems:'center',justifyContent:'center'}}>
              <Building size={13} style={{color:'#fff'}}/>
            </div>
            <div style={{fontFamily:'"DM Serif Display",Georgia,serif',fontSize:16,color:T.text}}>
              RenoLedger
            </div>
          </div>
          <div style={{display:'flex',alignItems:'center',gap:6}}>
            {syncStatus==='saving'&&<Loader2 size={11} style={{color:T.muted,animation:'spin 1s linear infinite'}}/>}
            {syncStatus==='saved'&&<CheckCircle size={11} style={{color:T.success}}/>}
            {syncStatus==='error'&&<AlertCircle size={11} style={{color:T.danger,cursor:'pointer'}}
              onClick={()=>alert(`Not synced: ${syncError}`)}/>}
            {!syncStatus&&lastSync&&!syncing&&(
              <span style={{fontSize:10,color:T.dim}}>
                {lastSync.getHours().toString().padStart(2,'0')}:{lastSync.getMinutes().toString().padStart(2,'0')}
              </span>
            )}
            {/* Bell */}
            <button onClick={()=>{setNotifOpen(o=>!o);}}
              style={{background:'none',border:'none',padding:4,cursor:'pointer',
                display:'flex',alignItems:'center',position:'relative',color:T.muted}}>
              <Bell size={16}/>
              {unreadCount>0&&(
                <span style={{position:'absolute',top:0,right:0,width:14,height:14,
                  background:T.danger,borderRadius:'50%',fontSize:8,fontWeight:700,
                  color:'#fff',display:'flex',alignItems:'center',justifyContent:'center'}}>
                  {unreadCount>9?'9+':unreadCount}
                </span>
              )}
            </button>
            {/* Dark mode */}
            <button onClick={()=>setDarkMode(d=>!d)}
              style={{background:'none',border:'none',padding:4,cursor:'pointer',fontSize:14}}>
              {darkMode?'☀️':'🌙'}
            </button>
            <button onClick={()=>loadAllData()} disabled={syncing}
              style={{background:'none',border:'none',padding:4,cursor:'pointer',
                display:'flex',alignItems:'center',color:T.dim}}>
              <RefreshCw size={13} style={syncing?{animation:'spin 1s linear infinite'}:{}}/>
            </button>
          </div>
          <button onClick={logout}
            style={{background:'none',border:'none',cursor:'pointer',
              fontSize:12,color:T.muted,fontFamily:'inherit',fontWeight:500}}>
            Out
          </button>
        </div>
      )}

      {/* MOBILE — bottom nav */}
      {isMobile&&(()=>{
        const primary=['dashboard','projects','payments','invoices','admin'];
        const primaryNav=visibleNav.filter(n=>primary.includes(n.id));
        const overflowNav=visibleNav.filter(n=>!primary.includes(n.id));
        return (
          <>
            {moreOpen&&<div style={{position:'fixed',inset:0,bottom:62,zIndex:39}} onClick={()=>setMoreOpen(false)}/>}
            {moreOpen&&(
              <div style={{position:'fixed',bottom:62,left:0,right:0,zIndex:40,
                background:T.card,borderTop:`1px solid ${T.borderLight}`,padding:'12px 12px 8px',
                boxShadow:'0 -4px 24px rgba(0,0,0,0.08)'}}>
                {['project','expenses','admin'].map(gKey=>{
                  const gItems=overflowNav.filter(n=>n.group===gKey);
                  if(!gItems.length) return null;
                  const g=NAV_GROUPS[gKey];
                  return (
                    <div key={gKey} style={{marginBottom:8}}>
                      <div style={{fontSize:9,fontWeight:700,color:T.dim,textTransform:'uppercase',
                        letterSpacing:'0.12em',padding:'0 4px',marginBottom:5}}>{g.label}</div>
                      <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:4}}>
                        {gItems.map(({id,label,Icon})=>{
                          const active=tab===id;
                          return (
                            <button key={id} onClick={()=>{setTab(id);setMoreOpen(false);}}
                              style={{display:'flex',flexDirection:'column',alignItems:'center',gap:3,
                                padding:'10px 4px',borderRadius:10,border:'none',cursor:'pointer',
                                background:active?T.tanLight:'transparent',
                                color:active?T.text:T.muted,fontFamily:'inherit'}}>
                              <Icon size={18}/>
                              <span style={{fontSize:9,fontWeight:active?600:400,textAlign:'center',lineHeight:1.2}}>{label}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
            <div style={{position:'fixed',bottom:0,left:0,right:0,zIndex:30,
              height:62,background:T.card,borderTop:`1px solid ${T.borderLight}`,
              display:'flex',alignItems:'stretch',
              paddingBottom:'env(safe-area-inset-bottom,0px)'}}>
              {primaryNav.map(({id,label,Icon})=>{
                const active=tab===id;
                const badge=id==='invoices'?pendCount:0;
                return (
                  <button key={id} onClick={()=>{setTab(id);setMoreOpen(false);}}
                    style={{flex:1,display:'flex',flexDirection:'column',alignItems:'center',
                      justifyContent:'center',gap:2,border:'none',
                      background:'none',cursor:'pointer',fontFamily:'inherit',
                      color:active?T.text:T.dim}}>
                    <div style={{position:'relative'}}>
                      {active&&<div style={{position:'absolute',inset:-3,background:T.tanLight,borderRadius:8}}/>}
                      <Icon size={20} style={{position:'relative'}}/>
                      {badge>0&&<span style={{position:'absolute',top:-4,right:-6,fontSize:9,
                        background:T.danger,color:'#fff',borderRadius:8,padding:'1px 4px',fontWeight:700}}>{badge}</span>}
                    </div>
                    <span style={{fontSize:9,fontWeight:active?600:400}}>{label}</span>
                  </button>
                );
              })}
              {overflowNav.length>0&&(
                <button onClick={()=>setMoreOpen(s=>!s)}
                  style={{flex:1,display:'flex',flexDirection:'column',alignItems:'center',
                    justifyContent:'center',gap:2,border:'none',background:'none',
                    cursor:'pointer',fontFamily:'inherit',color:moreOpen?T.text:T.dim}}>
                  <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                    <circle cx="5" cy="12" r="2"/><circle cx="12" cy="12" r="2"/><circle cx="19" cy="12" r="2"/>
                  </svg>
                  <span style={{fontSize:9,fontWeight:moreOpen?600:400}}>More</span>
                </button>
              )}
            </div>
          </>
        );
      })()}

      {/* MAIN CONTENT */}
      <div style={{marginLeft:isMobile?0:240,minHeight:'100vh',paddingTop:isMobile?52:0,paddingBottom:isMobile?80:0}}>
        {!isMobile&&(
          <div style={{padding:'22px 32px 16px',borderBottom:`1px solid ${T.borderLight}`,
            background:T.card,position:'sticky',top:0,zIndex:10}}>
            <div style={{display:'flex',alignItems:'center',justifyContent:'space-between'}}>
              <div>
                <div style={{fontFamily:'"DM Serif Display",Georgia,serif',fontSize:24,color:T.text,lineHeight:1.1}}>
                  {ALL_NAV.find(n=>n.id===tab)?.label}
                </div>
                <div style={{fontSize:12,color:T.dim,marginTop:3}}>{TAB_SUB[tab]}</div>
              </div>
              {isSuperAdmin&&(
                <div style={{display:'flex',alignItems:'center',gap:6,background:'rgba(109,40,217,0.06)',
                  border:'1px solid rgba(109,40,217,0.18)',borderRadius:8,padding:'5px 10px'}}>
                  <Terminal size={11} style={{color:'#6d28d9'}}/>
                  <span style={{fontSize:11,fontWeight:700,color:'#6d28d9'}}>Developer Mode</span>
                </div>
              )}
            </div>
          </div>
        )}
        <div style={{padding:isMobile?'12px 10px 16px':'24px 32px 48px'}}>
          {tab==='dashboard'&&(()=>{
            // Merge any new widget IDs that aren't in the user's saved list yet
            const allIds=DASH_WIDGETS.map(w=>w.id);
            const userWidgets=activeUser?.widgets||allIds;
            const mergedWidgets=[...userWidgets, ...allIds.filter(id=>!userWidgets.includes(id))];
            return <Dashboard projects={userProjects} invoices={invoices.filter(i=>userProjects.some(p=>p.id===i.projectId))} payments={payments.filter(py=>userProjects.some(p=>p.id===py.projectId))} widgets={mergedWidgets} siteWorkers={siteWorkers} onlinePresence={onlinePresence} activeUserId={activeUserId} notices={notices} setNotices={(n)=>{setNotices(n);saveNotices(n);}} isAdmin={isAdmin}/>;
          })()}
          {tab==='projects'&&<Projects projects={userProjects} setProjects={setProjects} invoices={invoices} payments={payments} isAdmin={isAdmin} onSoftDelete={handleSoftDelete} onShowToast={handleShowToast} users={users} acctSettings={acctSettings} logAction={logAction} activeUser={activeUser}/>}
          {tab==='invoices'&&<Invoices invoices={invoices} setInvoices={setInvoices} projects={userProjects} isAdmin={isAdmin} onSoftDelete={handleSoftDelete} onShowToast={handleShowToast} invoiceBatches={invoiceBatches} setInvoiceBatches={setInvoiceBatches} acctSettings={acctSettings} logAction={logAction}/>}
          {tab==='payments'&&<Payments payments={payments} setPayments={setPayments} projects={userProjects} invoices={invoices} isAdmin={isAdmin} onSoftDelete={handleSoftDelete} onShowToast={handleShowToast} acctSettings={acctSettings} logAction={logAction}/>}
          {tab==='contacts'&&<Contacts projects={userProjects} invoices={invoices.filter(i=>userProjects.some(p=>p.id===i.projectId))} payments={payments}/>}
          {tab==='reports'&&<Reports projects={userProjects} invoices={invoices} payments={payments} acctSettings={acctSettings}/>}
          {tab==='commissions'&&<Commissions projects={projects} setProjects={setProjects} invoices={invoices} isAdmin={isAdmin} users={users}/>}
          {tab==='claims'&&<StaffClaims claims={staffClaims} setClaims={setStaffClaims} projects={userProjects} users={users} activeUser={activeUser} isAdmin={isAdmin} invoices={invoices} setInvoices={setInvoices} acctSettings={acctSettings} trash={trash} setTrash={setTrash}/>}
          {tab==='warranty'&&<Warranty warranties={warranties} setWarranties={setWarranties} projects={projects} isAdmin={isAdmin} acctSettings={acctSettings}/>}
          {tab==='workers'&&<WorkerAdmin siteWorkers={siteWorkers} setSiteWorkers={setSiteWorkers} attendance={attendance} setAttendance={setAttendance} projects={projects} invoices={invoices} setInvoices={setInvoices} claims={workerClaims} setClaims={setWorkerClaims} acctSettings={acctSettings} logAction={logAction}/>}
          {tab==='checkin'&&<WorkerLoginScreen siteWorkers={siteWorkers} onLogin={(w)=>setWorkerSession(w)} onAdminLogin={()=>setTab('dashboard')} acctSettings={acctSettings}/>}
          {tab==='accounts'&&isAdmin&&(<CompanyAccounts projects={projects} invoices={invoices} payments={payments} acctSettings={acctSettings} setAcctSettings={setAcctSettings}/>)}
          {tab==='trash'&&isAdmin&&(<TrashBin trash={trash} onRestore={handleRestore} onPermanentDelete={handlePermanentDelete} isSuperAdmin={isSuperAdmin}/>)}
          {tab==='admin'&&isAdmin&&(<Admin users={users.filter(u=>u.id!=='__sa__')} setUsers={setUsers} projects={projects} onSoftDelete={handleSoftDelete} onShowToast={handleShowToast} actionLog={actionLog} onUndoAction={handleRestore} isSuperAdmin={isSuperAdmin}/>)}
          {tab==='system'&&isSuperAdmin&&(<SystemPanel projects={projects} invoices={invoices} payments={payments} siteWorkers={siteWorkers} attendance={attendance} users={users} warranties={warranties} trash={trash} acctSettings={acctSettings} setAcctSettings={setAcctSettings} actionLog={actionLog} setActionLog={setActionLog} logAction={logAction}/>)}
          {tab==='admin'&&!isAdmin&&(
            <div style={{display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',padding:60,gap:12,color:T.muted,textAlign:'center'}}>
              <Lock size={20} style={{color:T.dim}}/>
              <div style={{fontSize:15,fontWeight:600,color:T.secondary}}>Access Denied</div>
            </div>
          )}
        </div>
      </div>

      {toast&&<UndoToast message={toast.message} onUndo={()=>{toast.undoFn&&toast.undoFn();setToast(null);}} onDismiss={()=>setToast(null)}/>}
    </div>
  );
}
