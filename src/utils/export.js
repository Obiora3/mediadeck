import { DEFAULT_APP_SETTINGS } from "../constants/appDefaults";

const escapeHtml = (value) => String(value ?? "")
  .replace(/&/g, "&amp;")
  .replace(/</g, "&lt;")
  .replace(/>/g, "&gt;")
  .replace(/"/g, "&quot;")
  .replace(/'/g, "&#39;");
const safeText = (value) => String(value ?? "").replace(/[\u0000-\u001F\u007F]/g, " ").trim();
export const sanitizeMPOForExport = (mpo) => ({
  ...mpo,
  mpoNo: escapeHtml(safeText(mpo?.mpoNo)),
  date: escapeHtml(safeText(mpo?.date)),
  month: escapeHtml(safeText(mpo?.month)),
  year: escapeHtml(safeText(mpo?.year)),
  vendorName: escapeHtml(safeText(mpo?.vendorName)),
  clientName: escapeHtml(safeText(mpo?.clientName)),
  brand: escapeHtml(safeText(mpo?.brand)),
  campaignName: escapeHtml(safeText(mpo?.campaignName)),
  agencyAddress: escapeHtml(safeText(mpo?.agencyAddress)),
  agencyEmail: escapeHtml(safeText(mpo?.agencyEmail)),
  agencyPhone: escapeHtml(safeText(mpo?.agencyPhone)),
  signedBy: escapeHtml(safeText(mpo?.signedBy)),
  signedTitle: escapeHtml(safeText(mpo?.signedTitle)),
  signedSignature: mpo?.signedSignature || "",
  preparedBy: escapeHtml(safeText(mpo?.preparedBy)),
  preparedContact: escapeHtml(safeText(mpo?.preparedContact)),
  preparedTitle: escapeHtml(safeText(mpo?.preparedTitle)),
  preparedSignature: mpo?.preparedSignature || "",
  medium: escapeHtml(safeText(mpo?.medium)),
  surchLabel: escapeHtml(safeText(mpo?.surchLabel)),
  transmitMsg: escapeHtml(safeText(mpo?.transmitMsg)),
  terms: (Array.isArray(mpo?.terms) ? mpo.terms : DEFAULT_APP_SETTINGS.mpoTerms).map(t => escapeHtml(safeText(t))),
  spots: (mpo?.spots || []).map((s) => ({
    ...s,
    programme: escapeHtml(safeText(s?.programme)),
    wd: escapeHtml(safeText(s?.wd)),
    timeBelt: escapeHtml(safeText(s?.timeBelt)),
    material: escapeHtml(safeText(s?.material)),
    duration: escapeHtml(safeText(s?.duration)),
    scheduleMonth: escapeHtml(safeText(s?.scheduleMonth)),
  })),
});

export const buildProgrammeCostLines = (spots = []) => {
  const grouped = new Map();
  (spots || []).forEach(spot => {
    const programme = String(spot?.programme || '').trim() || 'Untitled Programme';
    const duration = String(spot?.duration || '').trim() || '';
    const rate = parseFloat(spot?.ratePerSpot) || 0;
    const cnt = Array.isArray(spot?.ad) && spot.ad.length
      ? spot.ad.length
      : Array.isArray(spot?.calendarDays) && spot.calendarDays.length
        ? spot.calendarDays.length
        : (parseFloat(spot?.spots) || 0);
    const key = programme.toLowerCase();
    if (!grouped.has(key)) grouped.set(key, { programme, duration, cnt: 0, rate });
    const entry = grouped.get(key);
    entry.cnt += cnt;
    if (!entry.rate && rate) entry.rate = rate;
    if (!entry.duration && duration) entry.duration = duration;
  });
  return Array.from(grouped.values()).map(line => ({ ...line, gross: line.cnt * line.rate }));
};

/* ── EXPORT HELPERS ─────────────────────────────────────── */
export const buildCSV = (rows, headers) => {
  const esc = v => `"${String(v ?? "").replace(/"/g,'""')}"`;
  return [headers.map(esc).join(","), ...rows.map(r => r.map(esc).join(","))].join("\n");
};

export const loadBrowserScript = (src, readyCheck) => new Promise((resolve, reject) => {
  try {
    const ready = typeof readyCheck === "function" ? readyCheck() : readyCheck;
    if (ready) return resolve(ready);
    const existing = document.querySelector(`script[src="${src}"]`);
    if (existing) {
      existing.addEventListener("load", () => resolve(typeof readyCheck === "function" ? readyCheck() : readyCheck), { once: true });
      existing.addEventListener("error", () => reject(new Error(`Failed to load ${src}`)), { once: true });
      return;
    }
    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.onload = () => resolve(typeof readyCheck === "function" ? readyCheck() : readyCheck);
    script.onerror = () => reject(new Error(`Failed to load ${src}`));
    document.head.appendChild(script);
  } catch (error) {
    reject(error);
  }
});

export const loadPreviewPdfLibraries = async () => {
  await loadBrowserScript("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js", () => window.html2canvas);
  await loadBrowserScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js", () => window.jspdf?.jsPDF);
  return { html2canvas: window.html2canvas, jsPDF: window.jspdf.jsPDF };
};

/* ── Lazy-load a CDN script ── */
/* ══════════════════════════════════════════════════════════════════
   PURE JS PDF BUILDER — zero CDN, zero canvas, zero dependencies.
   Generates a valid PDF entirely in-browser from MPO data.
   ══════════════════════════════════════════════════════════════════ */
/* ══════════════════════════════════════════════════════════════════
   PURE-JS PDF BUILDER  —  mirrors buildMPOHTML exactly
   No CDN · No canvas · No network · Works in any sandbox
   ══════════════════════════════════════════════════════════════════ */
export const buildMPOPdf = (mpo) => {
  /* ── shared helpers ─────────────────────────────────────── */
  const esc  = s => String(s??'').replace(/\\/g,'\\\\').replace(/\(/g,'\\(').replace(/\)/g,'\\)');
  const fmtN = n => Number(n||0).toLocaleString('en-NG',{minimumFractionDigits:2,maximumFractionDigits:2});
  const clip = (s,n) => String(s||'').slice(0,n);
  const MN   = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  const FULL = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const DAYS = ['SU','M','T','W','TH','FR','SA'];

  /* month index */
  const rawM = String(mpo.month||'').trim().toUpperCase();
  let mIdx   = MN.indexOf(rawM.slice(0,3));
  if (mIdx<0) mIdx = FULL.findIndex(n=>n.toUpperCase()===rawM);
  const yr   = parseInt(mpo.year)||new Date().getFullYear();
  const dim  = mIdx>=0 ? new Date(yr,mIdx+1,0).getDate() : 31;
  const mLbl = mIdx>=0 ? MN[mIdx]+'-'+String(yr).slice(-2) : rawM;
  const getDN= d => DAYS[new Date(yr,mIdx,d).getDay()];

  /* expand spots to aired dates */
  const WDM={MON:1,TUE:2,WED:3,THU:4,FRI:5,SAT:6,SUN:0};
  const sWD = (mpo.spots||[]).map(s=>{
    let ad=[];
    if (s.calendarDays&&s.calendarDays.length) { ad=s.calendarDays.map(Number); }
    else if (s.wd) {
      const k=s.wd.toUpperCase();
      const set=k==='DAILY'?[0,1,2,3,4,5,6]:k==='WEEKDAYS'?[1,2,3,4,5]:k==='WEEKENDS'?[0,6]:WDM[k]!==undefined?[WDM[k]]:[];
      for(let d=1;d<=dim;d++) if(mIdx>=0&&set.includes(new Date(yr,mIdx,d).getDay())) ad.push(d);
    }
    return {...s,ad};
  });
  const dNums = Array.from({length:dim},(_,i)=>i+1);

  /* costing */
  const costLines = sWD.map(s=>({ programme:s.programme||'', material:s.material||'', duration:s.duration||'', cnt:s.ad.length||parseInt(s.spots)||0, rate:parseFloat(s.ratePerSpot)||0 }));
  costLines.forEach(l=>l.gross=l.cnt*l.rate);
  const subTotal  = costLines.reduce((a,l)=>a+l.gross,0);
  const vdPct     = parseFloat(mpo.discPct)||0;
  const vdAmt     = subTotal*vdPct;
  const afterDisc = subTotal-vdAmt;
  const cPct      = parseFloat(mpo.commPct)||0;
  const cAmt      = afterDisc*cPct;
  const afterComm = afterDisc-cAmt;
  const spPct     = parseFloat(mpo.surchPct)||0;
  const spAmt     = afterComm*spPct;
  const netAmt    = afterComm+spAmt;
  const vatRate   = (parseFloat(mpo.vatPct) || 7.5) / 100;
  const vatAmt    = netAmt * vatRate;
  const totalPayable = netAmt+vatAmt;

  /* ── PDF engine ──────────────────────────────────────────── */
  const PW=595,PH=842;
  const parts=['%PDF-1.4\n'], xref=[];
  const addObj=c=>{ xref.push(parts.reduce((a,b)=>a+b.length,0)); const id=xref.length; parts.push(`${id} 0 obj\n${c}\nendobj\n`); return id; };

  /* fonts — Helvetica built-in, no embed needed */
  const fR=addObj(`<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>`);
  const fB=addObj(`<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>`);
  const RES=`<< /Font << /Hr ${fR} 0 R /Hb ${fB} 0 R >> >>`;

  /* stream helpers */
  const pages=[];
  let   ops=[];
  const o  = t => ops.push(t);
  const BT = (x,y,sz,bold,R,G,B,txt)=>{
    o(`BT /${bold?'Hb':'Hr'} ${sz} Tf ${(R/255).toFixed(3)} ${(G/255).toFixed(3)} ${(B/255).toFixed(3)} rg ${x.toFixed(2)} ${(PH-y).toFixed(2)} Td (${esc(txt)}) Tj ET`);
  };
  // right-aligned: estimate glyph width at ~0.52*size per char for Helvetica
  const BTR=(rx,y,sz,bold,R,G,B,txt)=>{
    const w=String(txt).length*sz*0.52; BT(rx-w,y,sz,bold,R,G,B,txt);
  };
  // center-aligned
  const BTC=(cx,y,sz,bold,R,G,B,txt)=>{
    const w=String(txt).length*sz*0.52; BT(cx-w/2,y,sz,bold,R,G,B,txt);
  };
  const LN=(x1,y1,x2,y2,R=160,G=160,B=160,w=0.4)=>
    o(`${(R/255).toFixed(3)} ${(G/255).toFixed(3)} ${(B/255).toFixed(3)} RG ${w} w ${x1.toFixed(2)} ${(PH-y1).toFixed(2)} m ${x2.toFixed(2)} ${(PH-y2).toFixed(2)} l S`);
  const RECT=(x,y,w,h,fr,fg,fb,sr,sg,sb)=>{
    const yb=(PH-y-h).toFixed(2);
    if(fr!=null) o(`${(fr/255).toFixed(3)} ${(fg/255).toFixed(3)} ${(fb/255).toFixed(3)} rg ${x.toFixed(2)} ${yb} ${w.toFixed(2)} ${h.toFixed(2)} re f`);
    if(sr!=null) o(`${(sr/255).toFixed(3)} ${(sg/255).toFixed(3)} ${(sb/255).toFixed(3)} RG 0.4 w ${x.toFixed(2)} ${yb} ${w.toFixed(2)} ${h.toFixed(2)} re S`);
  };
  const newPage=()=>{ pages.push(ops.join('\n')); ops=[]; };
  const checkY=(need,margin=42)=>{ if(y>PH-margin-need){ newPage(); y=MT; } };

  const ML=28,MR=28,MT=32,MB=32,CW=PW-ML-MR;
  let y=MT;

  /* ══ PAGE 1 ══════════════════════════════════════════════ */

  /* --- header bar --- */
  RECT(ML,y,CW,16, 26,58,107, null,null,null);
  BTC(PW/2, y+11, 11,true,255,255,255, 'MEDIA PURCHASE ORDER');
  y+=20;

  /* agency address */
  BTC(PW/2, y+6, 7,false,80,80,80, mpo.agencyAddress || mpo.agency || '5, Craig Street, Ogudu GRA, Lagos');
  y+=12; LN(ML,y,PW-MR,y,26,58,107,0.7); y+=6;

  /* --- info grid --- */
  const LI=[
    ['CLIENT:',    (mpo.clientName||'—').toUpperCase()],
    ['BRAND:',     (mpo.brand||'—').toUpperCase()],
    ['CAMPAIGN:',  clip(mpo.campaignName||'—',36)],
    ['MEDIUM:',    (mpo.medium||'—').toUpperCase()],
    ['VENDOR:',    clip(mpo.vendorName||'—',32)],
  ];
  const RI=[
    ['MPO No:',   mpo.mpoNo||'—'],
    ['Date:',     mpo.date||'—'],
    ['Period:',   `${mpo.month||''} ${mpo.year||''}`],
    ['Status:',   (mpo.status||'DRAFT').toUpperCase()],
    ['Prepared:', clip(mpo.preparedBy||'—',24)],
  ];
  const iy=y;
  LI.forEach(([l,v],i)=>{ BT(ML,iy+i*9+7,6.5,true,100,100,100,l); BT(ML+36,iy+i*9+7,6.5,false,0,0,0,v); });
  RI.forEach(([l,v],i)=>{ BT(PW/2+4,iy+i*9+7,6.5,true,100,100,100,l); BT(PW/2+36,iy+i*9+7,6.5,false,0,0,0,v); });
  y=iy+LI.length*9+12;

  /* transmit bar */
  RECT(ML,y,CW,13,26,58,107,null,null,null);
  const tx=`PLEASE TRANSMIT SPOTS ON ${(mpo.vendorName||'VENDOR').toUpperCase()} AS SCHEDULED`;
  BTC(PW/2,y+9,7,true,255,255,255,tx);
  y+=17;

  /* ══ CALENDAR GRID ════════════════════════════════════════ */
  /* Group spots by time belt */
  const order=[],groups={};
  sWD.forEach(s=>{ const k=(s.timeBelt||'GENERAL').trim(); if(!groups[k]){groups[k]=[];order.push(k);} groups[k].push(s); });

  /* column widths: Month(20) | Belt(26) | Prog(32) | d1..dN(each 8) | #Spots(14) | Material(30) */
  const dW=7.5;  // per-day column width
  const fixedW=20+26+32+14+30;
  const tableW=fixedW+dim*dW;
  const tScale = tableW>CW ? CW/tableW : 1; // scale down if too wide
  const cMonth =20*tScale, cBelt=26*tScale, cProg=32*tScale, cDay=dW*tScale, cSpots=14*tScale, cMat=30*tScale;
  const tx0=ML, tx1=tx0+cMonth, tx2=tx1+cBelt, tx3=tx2+cProg;
  const txSpots=tx3+dim*cDay, txMat=txSpots+cSpots;
  const RH=9, HH=10;

  /* header row */
  checkY(HH+RH*2);
  RECT(ML,y,CW,HH,26,58,107,null,null,null);
  BT(tx0+1,y+7,5,true,255,255,255,'MONTH');
  BT(tx1+1,y+7,5,true,255,255,255,'Time Belt');
  BT(tx2+1,y+7,5,true,255,255,255,'Programme');
  dNums.forEach(d=>BT(tx3+(d-1)*cDay+1,y+7,4.5,true,255,255,255,String(d)));
  BT(txSpots+1,y+7,5,true,255,255,255,'#');
  BT(txMat+1,y+7,5,true,255,255,255,'Material');
  y+=HH;

  /* day-of-week sub-header */
  RECT(ML,y,CW,7,238,244,252,null,null,null);
  if(mIdx>=0) dNums.forEach(d=>BT(tx3+(d-1)*cDay+1,y+5,4,false,80,80,80,getDN(d)));
  y+=7;

  /* data rows */
  let grandPaid=0;
  order.forEach(belt=>{
    const rows=groups[belt];
    let bTotal=0;
    rows.forEach((s,si)=>{
      const cnt=s.ad.length||parseInt(s.spots)||0;
      bTotal+=cnt; grandPaid+=cnt;
      checkY(RH+2);
      const bg=si%2===0?[252,252,255]:[245,248,253];
      RECT(ML,y,CW,RH,...bg,null,null,null);
      LN(ML,y+RH,ML+CW,y+RH,200,200,200,0.3);
      /* month cell (only first row of belt group) */
      if(si===0){
        BT(tx0+1,y+6,5,true,26,58,107,mLbl);
        BT(tx1+1,y+6,5,false,40,40,40,clip(belt,8));
      }
      BT(tx2+1,y+6,5,false,20,20,20,clip(s.programme||'',12));
      /* day dots */
      dNums.forEach(d => {
  const count = (s.ad || []).reduce(
    (n, value) => n + (Number(value) === Number(d) ? 1 : 0),
    0
  );

  if (count > 0) {
    RECT(tx3 + (d - 1) * cDay + 1, y + 2, cDay - 2, RH - 4, 26, 58, 107, null, null, null);
    BT(tx3 + (d - 1) * cDay + 1.5, y + 7, 4, true, 255, 255, 255, String(count));
  }
});
      /* spot count */
      RECT(txSpots,y,cSpots,RH,220,232,244,null,null,null);
      BTC(txSpots+cSpots/2,y+6,6,true,26,58,107,String(cnt));
      /* material */
      BT(txMat+1,y+6,4.5,false,60,60,60,clip(s.material||'',14));
      /* row border */
      LN(ML,y,ML+CW,y,200,200,200,0.2);
      y+=RH;
    });
    /* belt subtotal if multiple rows */
    if(rows.length>1){
      RECT(ML,y,CW,8,220,232,244,null,null,null);
      BTR(txSpots+cSpots-1,y+5.5,6,true,26,58,107,String(bTotal));
      y+=8;
    }
  });

  /* grand total row */
  RECT(ML,y,CW,10,26,58,107,null,null,null);
  BT(ML+2,y+7,6,true,255,255,255,'GRAND TOTAL');
  BTR(txSpots+cSpots-1,y+7,7,true,255,255,255,String(grandPaid));
  y+=14;

  /* ══ COSTING TABLE ════════════════════════════════════════ */
  checkY(60);
  BTC(PW/2,y+7,9,true,26,58,107,'C  O  S  T  I  N  G');
  y+=12; LN(ML,y,PW-MR,y,26,58,107,0.6); y+=5;

  /* per-spot cost lines */
  const cH=10;
  RECT(ML,y,CW,cH,26,58,107,null,null,null);
  [['Time Belt',96],['Spots',24],['Rate/Spot (N)',44],['Total (N)',44]].reduce((x,[h,w])=>{
    BT(x+1,y+7,5.5,true,255,255,255,h); return x+w;
  },ML);
  y+=cH;

  costLines.forEach((l,ri)=>{
    checkY(cH+2);
    RECT(ML,y,CW,cH,ri%2===0?252:246,ri%2===0?252:248,ri%2===0?255:252,210,210,210);
    let cx=ML;
    [[clip(l.programme,16),50,false],[clip(l.material,20),60,false],[String(l.duration)+'s',20,false],[String(l.cnt),18,true],[fmtN(l.rate),42,true],[fmtN(l.gross),42,true]].forEach(([v,w,b])=>{
      BT(cx+1,y+7,5.5,b,20,20,20,v); cx+=w;
    });
    y+=cH;
  });

  /* costing summary */
  const summaryRows=[
    ['Sub Total',                                          fmtN(subTotal),    false,false],
    ...(vdPct>0?[
      [`Volume Discount (${Math.round(vdPct*100)}%)`,     `- ${fmtN(vdAmt)}`,false,false],
      ['Less Discount',                                    fmtN(afterDisc),   false,false],
    ]:[]),
    ...(cPct>0?[
      [`Agency Commission (${Math.round(cPct*100)}%)`,    `- ${fmtN(cAmt)}`, false,false],
      ['Less Commission',                                  fmtN(afterComm),   false,false],
    ]:[]),
    ...(spPct>0?[
      [mpo.surchLabel||`Surcharge (${Math.round(spPct*100)}%)`, `+ ${fmtN(spAmt)}`,false,false],
      ['Net After Surcharge',                              fmtN(netAmt),      false,false],
    ]:[]),
    [`VAT (${parseFloat(mpo.vatPct) || 7.5}%)`,                    fmtN(vatAmt),      false,false],
    ['TOTAL AMOUNT PAYABLE',                              fmtN(totalPayable),true, true],
  ];
  y+=4;
  summaryRows.forEach(([label,val,bold,tot])=>{
    checkY(11);
    const rh=11;
    if(tot) RECT(ML,y,CW,rh,26,58,107,null,null,null);
    else    RECT(ML,y,CW,rh,248,250,255,215,215,215);
    const [cr,cg,cb]=tot?[255,255,255]:bold?[0,0,0]:[55,55,55];
    BT(ML+4,y+7.5,7,bold||tot,cr,cg,cb,label);
    BTR(PW-MR-3,y+7.5,7,bold||tot,cr,cg,cb,String(val));
    y+=rh;
  });
  y+=8;

  /* ══ CONTRACT TERMS ═══════════════════════════════════════ */
  checkY(30);
  BT(ML,y+6,7.5,true,0,0,0,'Contract Terms & Conditions');
  y+=12;
  const terms = Array.isArray(mpo.terms) && mpo.terms.length ? mpo.terms : DEFAULT_APP_SETTINGS.mpoTerms;
  terms.forEach((t,i)=>{
    checkY(10);
    BT(ML,y+6,6,false,50,50,50,`${i+1}.  ${clip(t,90)}`);
    y+=9;
  });
  y+=8;

  /* ══ SIGNATURES ═══════════════════════════════════════════ */
  checkY(38);
  const sy=y+22;
  LN(ML,sy,ML+52,sy,80,80,80,0.5);
  LN(ML+60,sy,ML+112,sy,80,80,80,0.5);
  LN(ML+120,sy,PW-MR,sy,80,80,80,0.5);
  BT(ML,    sy+5,6,true, 0,0,0,'For (Media House / Supplier)');
  BT(ML+60, sy+5,6,true, 0,0,0,`SIGNED BY: ${clip((mpo.signedBy||'').toUpperCase(),20)}`);
  BT(ML+120,sy+5,6,true, 0,0,0,`PREPARED BY: ${clip((mpo.preparedBy||'').toUpperCase(),18)}`);
  BT(ML,    sy+11,5.5,false,100,100,100,'Name / Signature / Official Stamp');
  BT(ML+60, sy+11,5.5,false,100,100,100,clip(mpo.signedTitle||'',22));
  BT(ML+120,sy+11,5.5,false,100,100,100,clip(mpo.preparedTitle||'',22));
  BT(ML+60, sy+16,5.5,false,100,100,100,clip(mpo.preparedContact||'',24));

  newPage(); // flush last page

  /* ══ ASSEMBLE PDF FILE ════════════════════════════════════ */
  const sids = pages.map(ps=>addObj(`<< /Length ${ps.length} >>\nstream\n${ps}\nendstream`));
  const kids = sids.map(sid=>addObj(`<< /Type /Page /MediaBox [0 0 ${PW} ${PH}] /Contents ${sid} 0 R /Resources ${RES} >>`));
  const pagesId=addObj(`<< /Type /Pages /Kids [${kids.map(i=>`${i} 0 R`).join(' ')}] /Count ${kids.length} >>`);
  /* re-emit pages with /Parent */
  const fkids=kids.map((_,pi)=>addObj(`<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${PW} ${PH}] /Contents ${sids[pi]} 0 R /Resources ${RES} >>`));
  const catId=addObj(`<< /Type /Catalog /Pages ${pagesId} 0 R >>`);

  const body=parts.join('');
  const xpos=body.length;
  const n=xref.length+1;
  const xs=`xref\n0 ${n}\n0000000000 65535 f \n`+xref.map(o=>o.toString().padStart(10,'0')+' 00000 n ').join('\n')+'\n';
  const tr=`trailer\n<< /Size ${n} /Root ${catId} 0 R >>\nstartxref\n${xpos}\n%%EOF`;
  const full=body+xs+tr;
  const bytes=new Uint8Array(full.length);
  for(let i=0;i<full.length;i++) bytes[i]=full.charCodeAt(i)&0xff;
  return bytes;
};

export const buildMPOHTML = (mpo) => {
  const {
  mpoNo, date, month, year, vendorName, clientName, brand, campaignName,
  agencyAddress, agencyEmail, agencyPhone, signedBy, signedTitle, signedSignature, preparedBy, preparedContact, preparedTitle, preparedSignature,
  spots, discPct, commPct, surchPct, surchLabel, transmitMsg, terms = DEFAULT_APP_SETTINGS.mpoTerms, vatPct = 7.5,
} = mpo;

  const fmt = n => Number(n||0).toLocaleString("en-NG",{minimumFractionDigits:2,maximumFractionDigits:2});


  const MN   = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  const FULL = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const DAY  = ["SU","M","T","W","TH","FR","SA"];
  const yr   = parseInt(year) || new Date().getFullYear();

  const resolveMIdx = m => {
    const u = (m||"").trim().toUpperCase();
    let i = MN.indexOf(u.slice(0,3));
    if (i < 0) i = FULL.findIndex(n => n.toUpperCase() === u);
    return i;
  };

  /* Determine which months to render */
  const allMonths = (mpo.months && mpo.months.length > 0) ? mpo.months : (month ? [month] : []);

  /* Build one calendar section for a given month and its spots */
  const buildMonthBlock = (monthName, monthSpots) => {
    const mIdx = resolveMIdx(monthName);
    const dim  = mIdx >= 0 ? new Date(yr, mIdx+1, 0).getDate() : 31;
    const getDN = d => DAY[new Date(yr, mIdx, d).getDay()];
    const monthLabel = mIdx >= 0 ? MN[mIdx]+"-"+String(yr).slice(-2) : (monthName||"").toUpperCase().slice(0,6);

    const WD = {MON:1,TUE:2,WED:3,THU:4,FRI:5,SAT:6,SUN:0};
    const sWD = monthSpots.map(s => {
      let ad = [];
      if (s.calendarDays && s.calendarDays.length) {
        ad = s.calendarDays.map(Number);
      } else if (s.wd) {
        const k = s.wd.toUpperCase();
        const set = k==="DAILY"?[0,1,2,3,4,5,6]:k==="WEEKDAYS"?[1,2,3,4,5]:k==="WEEKENDS"?[0,6]:WD[k]!==undefined?[WD[k]]:[];
        for(let d=1;d<=dim;d++) if(mIdx>=0 && set.includes(new Date(yr,mIdx,d).getDay())) ad.push(d);
      }
      return {...s, ad};
    });
    const dayCountForSpot = (spot, day) =>
  (spot.ad || []).reduce(
    (count, value) => count + (Number(value) === Number(day) ? 1 : 0),
    0
  );

    const dNums   = Array.from({length:dim},(_,i)=>i+1);
    const dateRow = dNums.map(d =>
      '<th style="background:#e8f0f8;color:#000;font-size:6.5px;padding:1px 0;text-align:center;border:1px solid #aaa;min-width:14px;width:14px;font-weight:700">'+d+'</th>').join("");
    const dayRow  = mIdx >= 0
      ? dNums.map(d => '<td style="background:#eef3fa;font-size:6px;padding:1px 0;text-align:center;border:1px solid #aaa;font-weight:500;color:#444">'+getDN(d)+'</td>').join("")
      : dNums.map(()=>'<td style="border:1px solid #aaa"></td>').join("");

    const order = [], groups = {};
    sWD.forEach(s => {
      const k = (s.timeBelt||"GENERAL").trim();
      if (!groups[k]) { groups[k]=[]; order.push(k); }
      groups[k].push(s);
    });

    let calRowsHtml = "";
    let grandPaid   = 0;

    order.forEach(belt => {
      const rows = groups[belt];
      let bTotal = 0;
      rows.forEach((s, si) => {
        const cnt = s.ad.length || parseInt(s.spots)||0;
        bTotal    += cnt;
        grandPaid += cnt;
        const cells = dNums.map(d =>
  '<td style="text-align:center;font-size:7.5px;padding:1px 0;border:1px solid #ddd;font-weight:700">' +
    (dayCountForSpot(s, d) > 0 ? String(dayCountForSpot(s, d)) : '') +
  '</td>'
).join("");
        const isFirst = si === 0;
        const monthTd = isFirst
          ? '<td rowspan="'+rows.length+'" style="font-size:7.5px;padding:2px 3px;border:1px solid #aaa;font-weight:700;text-align:center;vertical-align:middle;white-space:nowrap;background:#f5f8fd">'+monthLabel+'</td><td rowspan="'+rows.length+'" style="font-size:7.5px;padding:2px 5px;border:1px solid #aaa;font-weight:700;vertical-align:middle;white-space:nowrap;background:#f5f8fd">'+belt+'</td>'
          : "";
        calRowsHtml += '<tr>'+monthTd+'<td style="font-size:7.5px;padding:2px 5px;border:1px solid #aaa;white-space:nowrap">'+(s.programme||"")+'</td>'+cells+'<td style="text-align:center;font-weight:700;font-size:8px;padding:2px 3px;border:1px solid #aaa;background:#dce8f4">'+cnt+'</td><td style="font-size:7.5px;padding:2px 5px;border:1px solid #aaa;white-space:nowrap">'+(s.material||"")+'</td></tr>';
      });
      if (rows.length > 1) {
        calRowsHtml += '<tr style="background:#dce8f4"><td colspan="3" style="border:1px solid #aaa;padding:2px 4px;font-weight:700;font-size:7.5px;text-align:right"></td>'+dNums.map(()=>'<td style="border:1px solid #aaa"></td>').join("")+'<td style="text-align:center;font-weight:800;font-size:9px;border:1px solid #aaa;padding:2px 3px;background:#c4d8ee">'+bTotal+'</td><td style="border:1px solid #aaa"></td></tr>';
      }
    });

    const grandRow = '<tr style="background:#fff"><td colspan="'+(dim+5)+'" style="border:1px solid #aaa;padding:3px 6px;font-weight:800;font-size:10px;text-align:right">'+grandPaid+'</td></tr>';

    const headerHtml =
      '<div style="font-family:Arial,Helvetica,sans-serif;font-size:8.5px;font-weight:700;color:#1a3a6b;text-align:center;margin:10px 0 4px;letter-spacing:2px;text-transform:uppercase;border-bottom:1px solid #c0d0e8;padding-bottom:3px">'+monthLabel+' SCHEDULE</div>' +
      '<div class="cal-wrap"><table class="cal"><thead><tr>' +
      '<th style="background:#1a3a6b;color:#fff;font-size:7.5px;padding:3px;text-align:center;min-width:40px">MONTH</th>' +
      '<th style="background:#1a3a6b;color:#fff;font-size:7.5px;padding:3px 5px;text-align:left;min-width:70px">Time Belt</th>' +
      '<th style="background:#1a3a6b;color:#fff;font-size:7.5px;padding:3px 5px;text-align:left;min-width:80px">Programme</th>' +
      '<th colspan="'+dim+'" style="background:#1a3a6b;color:#fff;font-size:8px;padding:3px;text-align:center;letter-spacing:5px">SCHEDULE</th>' +
      '<th style="background:#1a3a6b;color:#fff;font-size:7px;padding:3px 2px;text-align:center;min-width:40px;line-height:1.4">NO OF<br>SPOTS</th>' +
      '<th style="background:#1a3a6b;color:#fff;font-size:7px;padding:3px 4px;text-align:left;min-width:90px;line-height:1.4">MATERIAL TITLE/<br>SPECIFICATION</th>' +
      '</tr><tr>' +
      '<td style="background:#dce8f4;font-size:7px;padding:2px 3px;text-align:center;font-weight:700;border:1px solid #aaa">DATES&#8594;</td>' +
      '<td style="background:#dce8f4;border:1px solid #aaa"></td>' +
      '<td style="background:#dce8f4;border:1px solid #aaa"></td>' +
      dateRow +
      '<td style="background:#dce8f4;border:1px solid #aaa"></td>' +
      '<td style="background:#dce8f4;border:1px solid #aaa"></td>' +
      '</tr><tr>' +
      '<td style="background:#eef3fa;border:1px solid #aaa"></td>' +
      '<td style="background:#eef3fa;border:1px solid #aaa"></td>' +
      '<td style="background:#eef3fa;border:1px solid #aaa"></td>' +
      dayRow +
      '<td style="background:#eef3fa;border:1px solid #aaa"></td>' +
      '<td style="background:#eef3fa;border:1px solid #aaa"></td>' +
      '</tr></thead><tbody>' + calRowsHtml + grandRow + '</tbody></table></div>';

    return { html: headerHtml, sWD };
  };

  /* Accumulate all months */
  let allCalendarHTML = "";
  let allSWD = [];

  if (allMonths.length === 0) {
    allCalendarHTML = '<div style="padding:10px;color:#888;font-style:italic;text-align:center">No schedule months configured.</div>';
  } else {
    allMonths.forEach(monthName => {
      const monthSpots = (spots||[]).filter(s => {
        if (!s.scheduleMonth) return true;
        const sm = (s.scheduleMonth||"").toLowerCase();
        const mn = (monthName||"").toLowerCase();
        return sm.startsWith(mn.slice(0,3)) || sm.includes(mn);
      });
      if (monthSpots.length === 0) return;
      const { html, sWD } = buildMonthBlock(monthName, monthSpots);
      allCalendarHTML += html;
      allSWD = allSWD.concat(sWD);
    });
    if (!allCalendarHTML) {
      allCalendarHTML = '<div style="padding:10px;color:#888;font-style:italic;text-align:center">No spots scheduled yet.</div>';
    }
  }

  const firstDur = allSWD.length > 0 ? (allSWD[0].duration||"30")+"SECS" : "30SECS";

  const costLines = buildProgrammeCostLines(allSWD);

  const subTotal   = costLines.reduce((a,l)=>a+l.gross, 0);
  const vdPct      = parseFloat(discPct)||0;
  const vdAmt      = subTotal * vdPct;
  const afterDisc  = subTotal - vdAmt;
  const cPct       = parseFloat(commPct)||0;
  const cAmt       = afterDisc * cPct;
  const afterComm  = afterDisc - cAmt;
  const VAT_RATE   = (parseFloat(vatPct) || 7.5) / 100;
  const spPct      = parseFloat(surchPct)||0;
  const spAmt      = afterComm * spPct;
  const netAmt     = afterComm + spAmt;
  const vatAmt     = netAmt * VAT_RATE;
  const totalPayable = netAmt + vatAmt;

  const periodLabel = allMonths.length > 1
    ? allMonths.map(m => (m||"").toUpperCase().slice(0,3)).join("/") + " " + yr
    : (allMonths[0]||month||"").toUpperCase() + " " + yr;

  const LOGO_SRC = `data:image/jpeg;base64,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`;

  const vdRow   = vdPct > 0 ? '<tr class="sum"><td colspan="4" style="text-align:right;font-weight:700">Volume Discount ('+Math.round(vdPct*100)+'%)</td><td style="text-align:right;color:#b00">- &#8358; '+fmt(vdAmt)+'</td></tr><tr class="sum"><td colspan="4" style="text-align:right;font-weight:700">Less Discount</td><td style="text-align:right;font-weight:700">&#8358; '+fmt(afterDisc)+'</td></tr>' : "";
  const cRow    = cPct > 0  ? '<tr class="sum"><td colspan="4" style="text-align:right;font-weight:700">Agency Commission ('+Math.round(cPct*100)+'%)</td><td style="text-align:right;color:#b00">- &#8358; '+fmt(cAmt)+'</td></tr><tr class="sum"><td colspan="4" style="text-align:right;font-weight:700">Less Commission</td><td style="text-align:right;font-weight:700">&#8358; '+fmt(afterComm)+'</td></tr>' : "";
  const spRow   = spPct > 0 ? '<tr class="sum"><td colspan="4" style="text-align:right;font-weight:700">'+(surchLabel||("Surcharge ("+Math.round(spPct*100)+"%)"))+' </td><td style="text-align:right;color:#b25400">+ &#8358; '+fmt(spAmt)+'</td></tr><tr class="sum"><td colspan="4" style="text-align:right;font-weight:700">Net After Surcharge</td><td style="text-align:right;font-weight:700">&#8358; '+fmt(netAmt)+'</td></tr>' : "";
  const costBodyRows = costLines.map(l => '<tr><td>'+l.programme+'</td><td style="text-align:center">'+l.duration+'secs</td><td style="text-align:center;font-weight:700">'+l.cnt+'</td><td style="text-align:right">'+fmt(l.rate)+'</td><td style="text-align:right;font-weight:700">'+fmt(l.gross)+'</td></tr>').join("");
  const termsRows = terms.map((t,i) => '<tr><td class="n">'+(i+1)+'</td><td class="t">'+t+'</td></tr>').join("");

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>MPO ${mpoNo}</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:Arial,Helvetica,sans-serif;font-size:9px;color:#111;background:#fff;padding:8mm 10mm}
    @media print{body{padding:5mm 7mm}@page{size:A4;margin:15mm}}
    .logo-wrap{text-align:center;margin:0;padding:0;line-height:1;border:none;background:none}
    .logo-wrap img{max-height:60px;max-width:120px;object-fit:contain;display:block;border:none;outline:none;margin:0 auto;padding:0;background:transparent}
    .agency-addr{text-align:center;font-size:7.5px;color:#7b0000;font-weight:700;text-transform:uppercase;margin:0;padding:0;letter-spacing:.3px;white-space:nowrap;line-height:1.2}
    .header-wrap{text-align:center;margin-bottom:4px;border:none;background:none}
    .header-logo-col{display:flex;flex-direction:column;align-items:center;gap:8px;margin:0;padding:0;border:none;background:none}
    .sub-header-wrap{display:flex;align-items:flex-start;gap:12mm;margin:4mm 0 6px 0;width:100%}
    .rec-box{border:0.5pt solid #000;border-radius:3mm;background:#fff;padding:2.5mm;box-sizing:border-box}
    .rec-box-left{width:48mm;min-height:19mm;flex-shrink:0}
    .rec-box-right{width:62mm;min-height:26mm;flex-shrink:0;margin-left:auto}
    .rec-line{line-height:1.2;color:#000;font-size:6pt;font-family:Arial,Helvetica,sans-serif}
    .rec-line-bold{font-weight:700;font-size:6.5pt;text-transform:uppercase;font-family:Arial,Helvetica,sans-serif;color:#000}
    .det-line{line-height:1.15;color:#000;font-size:5.5pt;font-family:Arial,Helvetica,sans-serif;margin-bottom:0.5mm}
    .det-lbl{font-weight:700;text-transform:uppercase;font-family:Arial,Helvetica,sans-serif}
    .det-val{font-weight:400;font-family:Arial,Helvetica,sans-serif}
    .tx-bar{background:#1a3a6b;color:#fff;text-align:center;font-size:9px;font-weight:700;padding:5px;margin-bottom:8px;letter-spacing:.8px}
    .cal-wrap{overflow-x:auto;margin-bottom:8px}
    .cal{border-collapse:collapse;font-size:8px;width:100%;table-layout:auto}
    .cal th,.cal td{border:1px solid #aaa;padding:2px 3px;white-space:nowrap}
    .cal thead th{background:#1a3a6b;color:#fff;font-size:7.5px;text-align:center}
    .costing-title{text-align:center;font-weight:800;font-size:10px;letter-spacing:6px;margin:12px 0 5px;color:#1a3a6b;text-transform:uppercase}
    .cost{width:100%;border-collapse:collapse;font-size:8.5px;margin-bottom:8px}
    .cost th{background:#1a3a6b;color:#fff;padding:5px 8px;text-align:left;font-size:8px}
    .cost td{padding:4px 8px;border:1px solid #ddd}
    .cost .sum td{border:none;padding:3px 8px}
    .cost .payable td{border:2px solid #888;padding:5px 8px;font-weight:700;font-size:10.5px;background:#f5f8fd}
    .terms-title{font-weight:700;text-decoration:underline;font-size:9px;margin:10px 0 3px;font-style:italic}
    .terms{width:100%;border-collapse:collapse}
    .terms td{padding:3px 5px;font-size:8px;line-height:1.55;vertical-align:top;border:1px solid #e0e0e0}
    .terms .n{width:16px;font-weight:700;text-align:right;border-right:none;white-space:nowrap;color:#222}
    .terms .t{border-left:none}
    .sig{width:100%;border-collapse:collapse;margin-top:24px}
    .sig td{width:33.3%;vertical-align:bottom;padding:0 8px 0 0;font-size:8px}
    .sig .dots{font-size:9px;color:#444;margin-bottom:3px;letter-spacing:1px}
    .sig .role{font-weight:700;font-size:8.5px;margin-bottom:1px}
    .sig .sub{color:#555;font-size:7.5px}
  </style></head><body>

  <div class="header-wrap">
    <table style="margin:0 auto;border-collapse:collapse;border:none">
      <tr><td style="text-align:center;padding:0;border:none">
        <img src="${LOGO_SRC}" alt="QVT Media" style="max-height:60px;max-width:120px;display:block;margin:0 auto;border:none;outline:none">
      </td></tr>
      <tr><td style="text-align:center;padding-top:8px;border:none">
        <span style="font-size:7.5px;color:#7b0000;font-weight:700;text-transform:uppercase;letter-spacing:.3px;white-space:nowrap">${(agencyAddress||"5, CRAIG STREET, OGUDU GRA, LAGOS").replace(/\n/g," | ")} &nbsp;|&nbsp; TEL: ${agencyPhone||preparedContact||"+234 800 000 0000"}${agencyEmail ? ` &nbsp;|&nbsp; EMAIL: ${agencyEmail}` : ""}</span>
      </td></tr>
    </table>
  </div>

  <div class="sub-header-wrap">
    <!-- LEFT BOX: Recipient -->
    <div class="rec-box rec-box-left">
      <div class="rec-line rec-line-bold">THE COMMERCIAL MANAGER,</div>
      <div class="rec-line">${(vendorName||"—").toUpperCase()} MEDIA SALES,</div>
      <div class="rec-line">LAGOS.</div>
    </div>
    <!-- RIGHT BOX: Client Details -->
    <div class="rec-box rec-box-right">
      <div class="det-line"><span class="det-lbl">CLIENT NAME: </span><span class="det-val">${(clientName||"—").toUpperCase()}</span></div>
      <div class="det-line"><span class="det-lbl">BRAND: </span><span class="det-val">${(brand||"—").toUpperCase()}</span></div>
      <div class="det-line"><span class="det-lbl">MEDIA PURCHASE ORDER No: </span><span class="det-val">${mpoNo||"—"}</span></div>
      <div class="det-line"><span class="det-lbl">MEDIUM: </span><span class="det-val">${(mpo.medium||"RADIO").toUpperCase()}</span></div>
      <div class="det-line"><span class="det-lbl">CAMPAIGN TITLE: </span><span class="det-val">${(campaignName||"—").toUpperCase()}</span></div>
      <div class="det-line"><span class="det-lbl">PERIOD: </span><span class="det-val">${periodLabel}</span></div>
      <div class="det-line"><span class="det-lbl">DATE: </span><span class="det-val">${date||""}</span></div>
    </div>
  </div>

  <div class="tx-bar">PLEASE TRANSMIT ${firstDur} SPOTS ON ${(vendorName||"").toUpperCase()} AS SCHEDULED</div>

  ${allCalendarHTML}

  <div class="costing-title">C &nbsp; O &nbsp; S &nbsp; T &nbsp; I &nbsp; N &nbsp; G</div>
  <table class="cost">
    <thead><tr>
      <th style="text-align:left;min-width:150px">PROGRAMME</th>
      <th style="text-align:center;min-width:55px">DURATION</th>
      <th style="text-align:center;min-width:60px">NO OF SPOTS</th>
      <th style="text-align:right;min-width:90px">RATE/SPOT (&#8358;)</th>
      <th style="text-align:right;min-width:110px">TOTAL AMOUNT (&#8358;)</th>
    </tr></thead>
    <tbody>${costBodyRows}</tbody>
    <tfoot>
      <tr class="sum"><td colspan="4" style="text-align:right;font-weight:700;border-top:2px solid #888">Sub Total</td><td style="text-align:right;font-weight:700;border-top:2px solid #888">&#8358; ${fmt(subTotal)}</td></tr>
      ${vdRow}${cRow}${spRow}
      <tr class="sum"><td colspan="4" style="text-align:right;font-weight:700">VAT (${parseFloat(vatPct) || 7.5}%)</td><td style="text-align:right;font-weight:700">&#8358; ${fmt(vatAmt)}</td></tr>
      <tr class="payable"><td colspan="4" style="text-align:right;letter-spacing:.5px">Total Amount Payable &#8596;</td><td style="text-align:right;color:#1a3a6b;font-size:11px">&#8358; ${fmt(totalPayable)}</td></tr>
    </tfoot>
  </table>

  <div class="terms-title">Contract Terms &amp; Condition</div>
  <table class="terms">${termsRows}</table>

  <table class="sig"><tr>
    <td><div class="dots">&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;</div><div class="role">NAME/SIGNATURE/DATE &amp; OFFICIAL STAMP</div><div class="sub">For (Media House / Third party supplier)</div></td>
    <td>${signedSignature ? `<div style="height:42px;display:flex;align-items:flex-end;margin:0 0 2px;"><img src="${signedSignature}" alt="Signed signature" style="max-height:40px;max-width:160px;object-fit:contain"></div>` : `<div style="height:42px;"></div>`}<div class="dots" style="margin-bottom:4px;">&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;</div><div class="role">SIGNED BY: ${(signedBy||"").toUpperCase()}</div><div>${preparedContact||""}</div><div class="sub">${(signedTitle||"").toUpperCase()}</div></td>
    <td>${preparedSignature ? `<div style="height:42px;display:flex;align-items:flex-end;margin:0 0 2px;"><img src="${preparedSignature}" alt="Prepared signature" style="max-height:40px;max-width:160px;object-fit:contain"></div>` : `<div style="height:42px;"></div>`}<div class="dots" style="margin-bottom:4px;">&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;</div><div class="role">PREPARED BY: ${(preparedBy||"").toUpperCase()}</div><div>${preparedContact||""}</div><div class="sub">${(preparedTitle||"").toUpperCase()}</div></td>
  </tr></table>

  <script>
    window.addEventListener('message', function(e) {
      if (e.data === 'print-mpo') { window.focus(); window.print(); }
    });
  </script>
  </body></html>`;
};


/* ── DAILY CALENDAR COMPONENT ──────────────────────────────────────── */
/* ── MPO GENERATOR ──────────────────────────────────────── */
