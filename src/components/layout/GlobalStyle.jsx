const GlobalStyle = ({ theme }) => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;1,400&display=swap');
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
    :root{
      ${theme === "dark" ? `
      --bg:#07090f;--bg2:#0e1118;--bg3:#141824;--bg4:#1c2233;
      --border:rgba(255,255,255,0.07);--border2:rgba(255,255,255,0.13);
      --text:#e8ecf4;--text2:#8b93a7;--text3:#4f576b;
      ` : `
      --bg:#f0f2f7;--bg2:#ffffff;--bg3:#f5f7fc;--bg4:#e8ecf4;
      --border:rgba(15,23,42,0.08);--border2:rgba(15,23,42,0.14);
      --text:#0f172a;--text2:#475467;--text3:#667085;
      `}
      --accent:#f0a500;--blue:#3b7ef5;--green:#16a34a;
      --red:#ef4444;--purple:#8b5cf6;--teal:#0d9488;--orange:#f97316;
    }
    html,body,#root{height:100%;font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text)}
    h1,h2,h3,h4,h5,h6,strong,th{color:var(--text)}
    p,span,label,td,li,small{color:inherit}
    input::placeholder,textarea::placeholder{color:var(--text3);opacity:1}
    select,option,input,textarea{color:var(--text)}
    *{scrollbar-width:thin;scrollbar-color:var(--bg4) transparent}
    ::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:var(--bg4);border-radius:3px}
    input,select,textarea,button{font-family:inherit}
    .fade{animation:fadeIn .25s ease}
    @keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}
    @keyframes spin{to{transform:rotate(360deg)}}
    .spin{animation:spin 1s linear infinite}
    textarea{resize:vertical}
    @media print{
      body{background:#fff!important;color:#000!important}
      .no-print{display:none!important}
      .print-area{display:block!important}
    }
  `}</style>
);

export default GlobalStyle;
