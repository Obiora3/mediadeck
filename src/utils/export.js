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
  const BTR=(rx,y,sz,bold,R,G,B,txt)=>{
    const w=String(txt).length*sz*0.52; BT(rx-w,y,sz,bold,R,G,B,txt);
  };
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

  RECT(ML,y,CW,16, 26,58,107, null,null,null);
  BTC(PW/2, y+11, 11,true,255,255,255, 'MEDIA PURCHASE ORDER');
  y+=20;

  BTC(PW/2, y+6, 7,false,80,80,80, mpo.agencyAddress || mpo.agency || '5, Craig Street, Ogudu GRA, Lagos');
  y+=12; LN(ML,y,PW-MR,y,26,58,107,0.7); y+=6;

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

  RECT(ML,y,CW,13,26,58,107,null,null,null);
  const tx=`PLEASE TRANSMIT SPOTS ON ${(mpo.vendorName||'VENDOR').toUpperCase()} AS SCHEDULED`;
  BTC(PW/2,y+9,7,true,255,255,255,tx);
  y+=17;

  const order=[],groups={};
  sWD.forEach(s=>{ const k=(s.timeBelt||'GENERAL').trim(); if(!groups[k]){groups[k]=[];order.push(k);} groups[k].push(s); });

  const dW=7.5;
  const fixedW=20+26+32+14+30;
  const tableW=fixedW+dim*dW;
  const tScale = tableW>CW ? CW/tableW : 1;
  const cMonth =20*tScale, cBelt=26*tScale, cProg=32*tScale, cDay=dW*tScale, cSpots=14*tScale, cMat=30*tScale;
  const tx0=ML, tx1=tx0+cMonth, tx2=tx1+cBelt, tx3=tx2+cProg;
  const txSpots=tx3+dim*cDay, txMat=txSpots+cSpots;
  const RH=9, HH=10;

  checkY(HH+RH*2);
  RECT(ML,y,CW,HH,26,58,107,null,null,null);
  BT(tx0+1,y+7,5,true,255,255,255,'MONTH');
  BT(tx1+1,y+7,5,true,255,255,255,'Time Belt');
  BT(tx2+1,y+7,5,true,255,255,255,'Programme');
  dNums.forEach(d=>BT(tx3+(d-1)*cDay+1,y+7,4.5,true,255,255,255,String(d)));
  BT(txSpots+1,y+7,5,true,255,255,255,'#');
  BT(txMat+1,y+7,5,true,255,255,255,'Material');

    y+=HH;

  RECT(ML,y,CW,7,238,244,252,null,null,null);
  if(mIdx>=0) dNums.forEach(d=>BT(tx3+(d-1)*cDay+1,y+5,4,false,80,80,80,getDN(d)));
  y+=7;

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
      if(si===0){
        BT(tx0+1,y+6,5,true,26,58,107,mLbl);
        BT(tx1+1,y+6,5,false,40,40,40,clip(belt,8));
      }
      BT(tx2+1,y+6,5,false,20,20,20,clip(s.programme||'',12));
      dNums.forEach(d=>{
        const count = (s.ad || []).reduce(
          (n, value) => n + (Number(value) === Number(d) ? 1 : 0),
          0
        );
        if(count > 0){
          RECT(tx3+(d-1)*cDay+1,y+2,cDay-2,RH-4,26,58,107,null,null,null);
          BT(tx3+(d-1)*cDay+1.5,y+7,4,true,255,255,255,String(count));
        }
      });
      RECT(txSpots,y,cSpots,RH,220,232,244,null,null,null);
      BTC(txSpots+cSpots/2,y+6,6,true,26,58,107,String(cnt));
      BT(txMat+1,y+6,4.5,false,60,60,60,clip(s.material||'',14));
      LN(ML,y,ML+CW,y,200,200,200,0.2);
      y+=RH;
    });
    if(rows.length>1){
      RECT(ML,y,CW,8,220,232,244,null,null,null);
      BTR(txSpots+cSpots-1,y+5.5,6,true,26,58,107,String(bTotal));
      y+=8;
    }
  });

  RECT(ML,y,CW,10,26,58,107,null,null,null);
  BT(ML+2,y+7,6,true,255,255,255,'GRAND TOTAL');
  BTR(txSpots+cSpots-1,y+7,7,true,255,255,255,String(grandPaid));
  y+=14;

  checkY(60);
  BTC(PW/2,y+7,9,true,26,58,107,'C  O  S  T  I  N  G');
  y+=12; LN(ML,y,PW-MR,y,26,58,107,0.6); y+=5;

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

  newPage();

  const sids = pages.map(ps=>addObj(`<< /Length ${ps.length} >>\nstream\n${ps}\nendstream`));
  const kids = sids.map(sid=>addObj(`<< /Type /Page /MediaBox [0 0 ${PW} ${PH}] /Contents ${sid} 0 R /Resources ${RES} >>`));
  const pagesId=addObj(`<< /Type /Pages /Kids [${kids.map(i=>`${i} 0 R`).join(' ')}] /Count ${kids.length} >>`);
  kids.map((_,pi)=>addObj(`<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${PW} ${PH}] /Contents ${sids[pi]} 0 R /Resources ${RES} >>`));
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
    mpoNo,
    date,
    month,
    year,
    vendorName,
    clientName,
    brand,
    campaignName,
    agencyAddress,
    agencyEmail,
    agencyPhone,
    signedBy,
    signedTitle,
    signedSignature,
    preparedBy,
    preparedContact,
    preparedTitle,
    preparedSignature,
    spots,
    discPct,
    commPct,
    surchPct,
    surchLabel,
    transmitMsg,
    terms = DEFAULT_APP_SETTINGS.mpoTerms,
    vatPct = 7.5,
  } = mpo;

  const fmt = (n) => Number(n || 0).toLocaleString("en-NG", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

  const MN = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  const FULL = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const DAY = ["SU", "M", "T", "W", "TH", "FR", "SA"];
  const yr = parseInt(year, 10) || new Date().getFullYear();

  const resolveMIdx = (m) => {
    const u = (m || "").trim().toUpperCase();
    let i = MN.indexOf(u.slice(0, 3));
    if (i < 0) i = FULL.findIndex((n) => n.toUpperCase() === u);
    return i;
  };

  const allMonths = mpo.months && mpo.months.length > 0 ? mpo.months : (month ? [month] : []);

  const buildMonthBlock = (monthName, monthSpots) => {
    const mIdx = resolveMIdx(monthName);
    const dim = mIdx >= 0 ? new Date(yr, mIdx + 1, 0).getDate() : 31;
    const getDN = (d) => DAY[new Date(yr, mIdx, d).getDay()];
    const monthLabel = mIdx >= 0 ? `${MN[mIdx]}-${String(yr).slice(-2)}` : (monthName || "").toUpperCase().slice(0, 6);

    const WD = { MON: 1, TUE: 2, WED: 3, THU: 4, FRI: 5, SAT: 6, SUN: 0 };
    const sWD = monthSpots.map((s) => {
      let ad = [];
      if (s.calendarDays && s.calendarDays.length) {
        ad = s.calendarDays.map(Number);
      } else if (s.wd) {
        const k = s.wd.toUpperCase();
        const set = k === "DAILY"
          ? [0, 1, 2, 3, 4, 5, 6]
          : k === "WEEKDAYS"
            ? [1, 2, 3, 4, 5]
            : k === "WEEKENDS"
              ? [0, 6]
              : WD[k] !== undefined
                ? [WD[k]]
                : [];
        for (let d = 1; d <= dim; d += 1) {
          if (mIdx >= 0 && set.includes(new Date(yr, mIdx, d).getDay())) ad.push(d);
        }
      }
      return { ...s, ad };
    });

    const dayCountForSpot = (spot, day) =>
      (spot.ad || []).reduce(
        (count, value) => count + (Number(value) === Number(day) ? 1 : 0),
        0
      );

    const dNums = Array.from({ length: dim }, (_, i) => i + 1);
    const dateRow = dNums.map((d) =>
      `<th style="background:#e8f0f8;color:#000;font-size:6.5px;padding:1px 0;text-align:center;border:1px solid #aaa;min-width:14px;width:14px;font-weight:700">${d}</th>`
    ).join("");
    const dayRow = mIdx >= 0
      ? dNums.map((d) => `<td style="background:#eef3fa;font-size:6px;padding:1px 0;text-align:center;border:1px solid #aaa;font-weight:500;color:#444">${getDN(d)}</td>`).join("")
      : dNums.map(() => '<td style="border:1px solid #aaa"></td>').join("");

    const order = [];
    const groups = {};
    sWD.forEach((s) => {
      const k = (s.timeBelt || "GENERAL").trim();
      if (!groups[k]) {
        groups[k] = [];
        order.push(k);
      }
      groups[k].push(s);
    });

    let calRowsHtml = "";
    let grandPaid = 0;

    order.forEach((belt) => {
      const rows = groups[belt];
      let bTotal = 0;

      rows.forEach((s, si) => {
        const cnt = s.ad.length || parseInt(s.spots, 10) || 0;
        bTotal += cnt;
        grandPaid += cnt;

        const cells = dNums.map((d) => {
          const count = dayCountForSpot(s, d);
          return `<td style="text-align:center;font-size:7.5px;padding:1px 0;border:1px solid #ddd;font-weight:700">${count > 0 ? String(count) : ""}</td>`;
        }).join("");

        const isFirst = si === 0;
        const monthTd = isFirst
          ? `<td rowspan="${rows.length}" style="font-size:7.5px;padding:2px 3px;border:1px solid #aaa;font-weight:700;text-align:center;vertical-align:middle;white-space:nowrap;background:#f5f8fd">${monthLabel}</td><td rowspan="${rows.length}" style="font-size:7.5px;padding:2px 5px;border:1px solid #aaa;font-weight:700;vertical-align:middle;white-space:nowrap;background:#f5f8fd">${escapeHtml(belt)}</td>`
          : "";
                  calRowsHtml += `<tr>${monthTd}<td style="font-size:7.5px;padding:2px 5px;border:1px solid #aaa;white-space:nowrap">${s.programme || ""}</td>${cells}<td style="text-align:center;font-weight:700;font-size:8px;padding:2px 3px;border:1px solid #aaa;background:#dce8f4">${cnt}</td><td style="font-size:7.5px;padding:2px 5px;border:1px solid #aaa;white-space:nowrap">${s.material || ""}</td></tr>`;
      });

      if (rows.length > 1) {
        calRowsHtml += `<tr style="background:#dce8f4"><td colspan="3" style="border:1px solid #aaa;padding:2px 4px;font-weight:700;font-size:7.5px;text-align:right"></td>${dNums.map(() => '<td style="border:1px solid #aaa"></td>').join("")}<td style="text-align:center;font-weight:800;font-size:9px;border:1px solid #aaa;padding:2px 3px;background:#c4d8ee">${bTotal}</td><td style="border:1px solid #aaa"></td></tr>`;
      }
    });

    const grandRow = `<tr style="background:#fff"><td colspan="${dim + 5}" style="border:1px solid #aaa;padding:3px 6px;font-weight:800;font-size:10px;text-align:right">${grandPaid}</td></tr>`;

    const headerHtml =
      `<div style="font-family:Arial,Helvetica,sans-serif;font-size:8.5px;font-weight:700;color:#1a3a6b;text-align:center;margin:10px 0 4px;letter-spacing:2px;text-transform:uppercase;border-bottom:1px solid #c0d0e8;padding-bottom:3px">${monthLabel} SCHEDULE</div>` +
      '<div class="cal-wrap"><table class="cal"><thead><tr>' +
      '<th style="background:#1a3a6b;color:#fff;font-size:7.5px;padding:3px;text-align:center;min-width:40px">MONTH</th>' +
      '<th style="background:#1a3a6b;color:#fff;font-size:7.5px;padding:3px 5px;text-align:left;min-width:70px">Time Belt</th>' +
      '<th style="background:#1a3a6b;color:#fff;font-size:7.5px;padding:3px 5px;text-align:left;min-width:80px">Programme</th>' +
      `<th colspan="${dim}" style="background:#1a3a6b;color:#fff;font-size:8px;padding:3px;text-align:center;letter-spacing:5px">SCHEDULE</th>` +
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
      `</tr></thead><tbody>${calRowsHtml}${grandRow}</tbody></table></div>`;

    return { html: headerHtml, sWD };
  };

  let allCalendarHTML = "";
  let allSWD = [];

  if (allMonths.length === 0) {
    allCalendarHTML = '<div style="padding:10px;color:#888;font-style:italic;text-align:center">No schedule months configured.</div>';
  } else {
    allMonths.forEach((monthName) => {
      const monthSpots = (spots || []).filter((s) => {
        if (!s.scheduleMonth) return true;
        const sm = (s.scheduleMonth || "").toLowerCase();
        const mn = (monthName || "").toLowerCase();
        return sm.startsWith(mn.slice(0, 3)) || sm.includes(mn);
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

  const costLines = buildProgrammeCostLines(allSWD);
  const subTotal = costLines.reduce((a, l) => a + l.gross, 0);
  const vdPct = parseFloat(discPct) || 0;
  const vdAmt = subTotal * vdPct;
  const afterDisc = subTotal - vdAmt;
  const cPct = parseFloat(commPct) || 0;
  const cAmt = afterDisc * cPct;
  const afterComm = afterDisc - cAmt;
  const VAT_RATE = (parseFloat(vatPct) || 7.5) / 100;
  const spPct = parseFloat(surchPct) || 0;
  const spAmt = afterComm * spPct;
  const netAmt = afterComm + spAmt;
  const vatAmt = netAmt * VAT_RATE;
  const totalPayable = netAmt + vatAmt;

  const periodLabel = allMonths.length > 1
    ? `${allMonths.map((m) => (m || "").toUpperCase().slice(0, 3)).join("/")} ${yr}`
    : `${(allMonths[0] || month || "").toUpperCase()} ${yr}`;

  const termsHtml = (Array.isArray(terms) && terms.length ? terms : DEFAULT_APP_SETTINGS.mpoTerms)
    .map((t, i) => `<li>${t}</li>`)
    .join("");

  const costRowsHtml = costLines.map((line, i) => `
    <tr class="${i % 2 === 0 ? "alt" : ""}">
      <td>${line.programme}</td>
      <td>${line.duration || "—"}</td>
      <td class="num">${line.cnt}</td>
      <td class="num">${fmt(line.rate)}</td>
      <td class="num">${fmt(line.gross)}</td>
    </tr>
  `).join("");

  const signedSignatureHtml = signedSignature
    ? `<img src="${signedSignature}" alt="Signed signature" style="max-height:48px;max-width:180px;object-fit:contain" />`
    : '<div class="sig-placeholder">Signature / Stamp</div>';

  const preparedSignatureHtml = preparedSignature
    ? `<img src="${preparedSignature}" alt="Prepared signature" style="max-height:48px;max-width:180px;object-fit:contain" />`
    : '<div class="sig-placeholder">Prepared by signature</div>';

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>MPO ${mpoNo || ""}</title>
<style>
  :root{
    --navy:#1a3a6b;
    --navy-2:#24497f;
    --sky:#dce8f4;
    --sky-2:#eef3fa;
    --line:#aab7c8;
    --text:#111827;
    --muted:#5b6472;
    --paper:#ffffff;
    --paper-2:#f8fbff;
    --green:#14532d;
    --red:#991b1b;
    --orange:#9a3412;
  }
  *{box-sizing:border-box}
  html,body{margin:0;padding:0;background:#eef2f7;color:var(--text);font-family:Arial,Helvetica,sans-serif}
  body{padding:18px}
  .sheet{
    width:max-content;
    min-width:1120px;
    margin:0 auto;
    background:var(--paper);
    border:1px solid #cfd8e3;
    box-shadow:0 16px 40px rgba(15,23,42,.10);
    padding:22px 22px 28px;
  }
  .topbar{
    background:var(--navy);
    color:#fff;
    text-align:center;
    font-weight:700;
    letter-spacing:.16em;
    font-size:14px;
    padding:9px 14px;
    border-radius:4px;
  }
  .agency{
    text-align:center;
    font-size:11px;
    color:var(--muted);
    margin:9px 0 10px;
  }
  .rule{height:1px;background:var(--navy);opacity:.55;margin:4px 0 12px}
  .meta{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:18px;
    margin-bottom:10px;
  }
  .meta-col{display:grid;gap:6px}
  .meta-row{display:grid;grid-template-columns:88px 1fr;gap:6px;font-size:11px}
  .meta-row b{color:var(--muted);text-transform:uppercase;letter-spacing:.03em}
  .transmit{
    background:var(--navy);
    color:#fff;
    text-align:center;
    font-size:11px;
    font-weight:700;
    padding:7px 10px;
    border-radius:4px;
    margin:10px 0 14px;
  }
  .section-title{
    text-align:center;
    font-weight:700;
    color:var(--navy);
    letter-spacing:.22em;
    font-size:12px;
    margin:18px 0 8px;
  }
  .cal-wrap{overflow:visible}
  table{border-collapse:collapse;width:100%}
  .cal{width:max-content;min-width:100%}
  .cal th,.cal td{border:1px solid #aaa}
  .cost-table th,.cost-table td{
    border:1px solid #d2dbe6;
    padding:7px 8px;
    font-size:11px;
  }
  .cost-table th{
    background:var(--navy);
    color:#fff;
    text-align:left;
    font-weight:700;
  }
  .cost-table .num{text-align:right}
  .cost-table .alt td{background:var(--paper-2)}
  .summary{
    margin-top:10px;
    width:100%;
    border:1px solid #d2dbe6;
    border-top:none;
  }
  .summary-row{
    display:grid;
    grid-template-columns:1fr 180px;
    gap:8px;
    padding:8px 10px;
    border-top:1px solid #d2dbe6;
    background:#f8fbff;
    font-size:11px;
  }
  .summary-row strong:last-child{text-align:right}
  .summary-row.total{
    background:var(--navy);
    color:#fff;
    font-weight:700;
    font-size:12px;
  }
  .terms{
    margin:8px 0 0;
    padding-left:18px;
    color:#374151;
    font-size:11px;
    line-height:1.55;
  }
  .terms li{margin:4px 0}
  .signatures{
    display:grid;
    grid-template-columns:1fr 1fr 1fr;
    gap:20px;
    margin-top:26px;
    align-items:end;
  }
  .sig{
    min-height:92px;
    display:flex;
    flex-direction:column;
    justify-content:flex-end;
  }
  .sig-line{
    border-top:1px solid #6b7280;
    padding-top:6px;
    font-size:10px;
  }
  .sig-title{font-size:10px;color:#6b7280;margin-top:4px}
  .sig-placeholder{
    height:48px;
    display:flex;
    align-items:flex-end;
    justify-content:flex-start;
    color:#9ca3af;
    font-size:10px;
  }
  .footer{
    margin-top:12px;
    text-align:center;
    color:#6b7280;
    font-size:10px;
  }
</style>
</head>
<body>
  <div class="sheet">
    <div class="topbar">MEDIA PURCHASE ORDER</div>
    <div class="agency">${agencyAddress || ""}${agencyEmail ? ` · ${agencyEmail}` : ""}${agencyPhone ? ` · ${agencyPhone}` : ""}</div>
    <div class="rule"></div>

    <div class="meta">
      <div class="meta-col">
        <div class="meta-row"><b>Client:</b><span>${clientName || "—"}</span></div>
        <div class="meta-row"><b>Brand:</b><span>${brand || "—"}</span></div>
        <div class="meta-row"><b>Campaign:</b><span>${campaignName || "—"}</span></div>
        <div class="meta-row"><b>Medium:</b><span>${mpo.medium || "—"}</span></div>
        <div class="meta-row"><b>Vendor:</b><span>${vendorName || "—"}</span></div>
      </div>
      <div class="meta-col">
        <div class="meta-row"><b>MPO No:</b><span>${mpoNo || "—"}</span></div>
        <div class="meta-row"><b>Date:</b><span>${date || "—"}</span></div>
        <div class="meta-row"><b>Period:</b><span>${periodLabel}</span></div>
        <div class="meta-row"><b>Status:</b><span>${mpo.status || "draft"}</span></div>
        <div class="meta-row"><b>Prepared:</b><span>${preparedBy || "—"}</span></div>
      </div>
    </div>

    <div class="transmit">${transmitMsg || `PLEASE TRANSMIT SPOTS ON ${(vendorName || "VENDOR").toUpperCase()} AS SCHEDULED`}</div>

    ${allCalendarHTML}

    <div class="section-title">C O S T I N G</div>

    <table class="cost-table">
      <thead>
        <tr>
          <th>Programme</th>
          <th>Duration</th>
          <th class="num">Spots</th>
          <th class="num">Rate / Spot (₦)</th>
          <th class="num">Total (₦)</th>
        </tr>
      </thead>
      <tbody>
        ${costRowsHtml || '<tr><td colspan="5" style="text-align:center;color:#6b7280">No costing rows yet.</td></tr>'}
      </tbody>
    </table>

    <div class="summary">
      <div class="summary-row"><strong>Sub Total</strong><strong>${fmt(subTotal)}</strong></div>
      ${vdPct > 0 ? `<div class="summary-row"><span>Volume Discount (${Math.round(vdPct * 100)}%)</span><strong>- ${fmt(vdAmt)}</strong></div><div class="summary-row"><span>Less Discount</span><strong>${fmt(afterDisc)}</strong></div>` : ""}
      ${cPct > 0 ? `<div class="summary-row"><span>Agency Commission (${Math.round(cPct * 100)}%)</span><strong>- ${fmt(cAmt)}</strong></div><div class="summary-row"><span>Less Commission</span><strong>${fmt(afterComm)}</strong></div>` : ""}
      ${spPct > 0 ? `<div class="summary-row"><span>${surchLabel || `Surcharge (${Math.round(spPct * 100)}%)`}</span><strong>+ ${fmt(spAmt)}</strong></div><div class="summary-row"><span>Net After Surcharge</span><strong>${fmt(netAmt)}</strong></div>` : ""}
      <div class="summary-row"><span>VAT (${parseFloat(vatPct) || 7.5}%)</span><strong>${fmt(vatAmt)}</strong></div>
      <div class="summary-row total"><strong>TOTAL AMOUNT PAYABLE</strong><strong>${fmt(totalPayable)}</strong></div>
    </div>

    <div class="section-title">C O N T R A C T&nbsp;&nbsp;T E R M S</div>
    <ol class="terms">${termsHtml}</ol>

    <div class="signatures">
      <div class="sig">
        ${signedSignatureHtml}
        <div class="sig-line">For (Media House / Supplier)</div>
        <div class="sig-title">Name / Signature / Official Stamp</div>
      </div>
      <div class="sig">
        <div style="height:48px;display:flex;align-items:flex-end;">${signedBy || ""}</div>
        <div class="sig-line">SIGNED BY</div>
        <div class="sig-title">${signedTitle || ""}</div>
      </div>
      <div class="sig">
        ${preparedSignatureHtml}
        <div class="sig-line">PREPARED BY</div>
        <div class="sig-title">${preparedBy || ""}${preparedTitle ? ` · ${preparedTitle}` : ""}${preparedContact ? ` · ${preparedContact}` : ""}</div>
      </div>
    </div>

    <div class="footer">Generated directly from MPO data</div>
  </div>
</body>
</html>`;
};