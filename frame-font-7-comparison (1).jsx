import { useState } from "react";

const FONTS = [
  { id:"A", key:"titillium", name:"Titillium Web",  family:"'Titillium Web',sans-serif",  tag:"Structured · Authority",        dw:700, weights:[300,400,600,700,900], wl:{300:"Light",400:"Regular",600:"SemiBold",700:"Bold",900:"Black"},             lsD:"0.06em", lsL:"0.10em", lsB:"0" },
  { id:"B", key:"oxanium",   name:"Oxanium",         family:"'Oxanium',sans-serif",         tag:"Circuit DNA · Geometric",       dw:800, weights:[300,400,600,700,800], wl:{300:"Light",400:"Regular",600:"SemiBold",700:"Bold",800:"ExtraBold"},          lsD:"0.06em", lsL:"0.12em", lsB:"0.01em" },
  { id:"C", key:"chakra",    name:"Chakra Petch",    family:"'Chakra Petch',sans-serif",    tag:"Circuit-Trace · FRAME-Native",  dw:700, weights:[300,400,600,700],     wl:{300:"Light",400:"Regular",600:"SemiBold",700:"Bold"},                         lsD:"0.06em", lsL:"0.12em", lsB:"0.01em" },
  { id:"D", key:"spaceg",    name:"Space Grotesk",   family:"'Space Grotesk',sans-serif",   tag:"Tech Grotesque · Modern",       dw:700, weights:[300,400,500,600,700], wl:{300:"Light",400:"Regular",500:"Medium",600:"SemiBold",700:"Bold"},             lsD:"0.04em", lsL:"0.10em", lsB:"0" },
  { id:"E", key:"rajdhani",  name:"Rajdhani",        family:"'Rajdhani',sans-serif",        tag:"Narrow Angular · Compact",      dw:700, weights:[300,400,500,600,700], wl:{300:"Light",400:"Regular",500:"Medium",600:"SemiBold",700:"Bold"},             lsD:"0.08em", lsL:"0.14em", lsB:"0.02em" },
  { id:"F", key:"orbitron",  name:"Orbitron",        family:"'Orbitron',sans-serif",        tag:"Hard Sci-Fi · Zero Ambiguity",  dw:800, weights:[400,500,600,700,800,900], wl:{400:"Regular",500:"Medium",600:"SemiBold",700:"Bold",800:"ExtraBold",900:"Black"}, lsD:"0.04em", lsL:"0.08em", lsB:"0.02em" },
  { id:"G", key:"saira",     name:"Saira",           family:"'Saira',sans-serif",           tag:"Condensed · Data Dense",        dw:700, weights:[300,400,500,600,700,800], wl:{300:"Light",400:"Regular",500:"Medium",600:"SemiBold",700:"Bold",800:"ExtraBold"}, lsD:"0.06em", lsL:"0.10em", lsB:"0.01em" },
];

const PALETTES = [
  { id:"core-black",    name:"Core Black",    bg:"#090909", surface:"#131313", elevated:"#1E1E1E", text:"#FFFFFF", muted:"#888888", accent:"#F5C831", accentFg:"#000000", secondary:"#E0E0E0", border:"#262626" },
  { id:"slate-circuit", name:"Slate Circuit",  bg:"#0A0D14", surface:"#121720", elevated:"#1A2030", text:"#E8EDF8", muted:"#5A6880", accent:"#F5C831", accentFg:"#000000", secondary:"#4A8FE0", border:"#1E2840" },
  { id:"void-purple",   name:"Void Purple",    bg:"#0C080E", surface:"#181018", elevated:"#241828", text:"#EEEAF4", muted:"#706878", accent:"#F5C831", accentFg:"#000000", secondary:"#A060E0", border:"#2A1E38" },
  { id:"deep-forge",    name:"Deep Forge",     bg:"#0C0909", surface:"#181010", elevated:"#231818", text:"#F0ECEC", muted:"#806868", accent:"#F5C831", accentFg:"#000000", secondary:"#E05830", border:"#2E1C1C" },
  { id:"midnight-moss", name:"Midnight Moss",  bg:"#090C08", surface:"#131810", elevated:"#1C2318", text:"#E2E8DA", muted:"#6A7260", accent:"#F5C831", accentFg:"#000000", secondary:"#78B830", border:"#222E1C" },
  { id:"pure-canvas",   name:"Pure Canvas",    bg:"#F6F6F6", surface:"#FFFFFF",  elevated:"#EBEBEB", text:"#080808", muted:"#606060", accent:"#B08800", accentFg:"#FFFFFF", secondary:"#101010", border:"#DCDCDC" },
  { id:"frost-circuit", name:"Frost Circuit",  bg:"#F0F4FA", surface:"#FAFCFF",  elevated:"#E2ECFA", text:"#080E1C", muted:"#4E6080", accent:"#9C7800", accentFg:"#FFFFFF", secondary:"#2860C8", border:"#CCDAEC" },
];

function CircuitMark({ size = 20, color = "#F5C831" }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none">
      <rect x="1" y="1" width="22" height="22" rx="2.5" stroke={color} strokeWidth="1.4"/>
      <circle cx="1"  cy="1"  r="1.8" fill={color}/>
      <circle cx="23" cy="1"  r="1.8" fill={color}/>
      <circle cx="1"  cy="23" r="1.8" fill={color}/>
      <circle cx="23" cy="23" r="1.8" fill={color}/>
      <circle cx="12" cy="12" r="1.5" fill={color}/>
      <line x1="12" y1="12" x2="17" y2="12" stroke={color} strokeWidth="1.2"/>
      <circle cx="17" cy="12" r="1" fill={color} fillOpacity="0.5"/>
      <line x1="12" y1="12" x2="12" y2="17" stroke={color} strokeWidth="1.2"/>
      <circle cx="12" cy="17" r="1" fill={color} fillOpacity="0.5"/>
      <line x1="12" y1="12" x2="7"  y2="12" stroke={color} strokeWidth="1.2"/>
      <circle cx="7"  cy="12" r="1" fill={color} fillOpacity="0.3"/>
    </svg>
  );
}

function AppColumn({ font, p }) {
  const ff = font.family, dw = font.dw;
  const nav   = ["Dashboard","Apps","Build","Deploy","Settings"];
  const stats = [{ l:"Apps", v:"24", hi:true }, { l:"Calls", v:"1.2M", hi:false }, { l:"Up", v:"99.9%", hi:false }];
  const apps  = [{ n:"Analytics Pro", a:true }, { n:"Data Pipeline", a:false }, { n:"Auth Service", a:true }];

  return (
    <div style={{ padding:8, height:"100%", background:"radial-gradient(ellipse at 50% 0%,#111008 0%,#060606 70%)" }}>
      <div style={{ background:p.bg, borderRadius:7, overflow:"hidden", border:`1px solid ${p.border}`, height:"100%", display:"flex", flexDirection:"column" }}>
        {/* Navbar */}
        <div style={{ background:p.surface, borderBottom:`1px solid ${p.border}`, height:34, display:"flex", alignItems:"center", padding:"0 9px", justifyContent:"space-between", flexShrink:0 }}>
          <div style={{ display:"flex", alignItems:"center", gap:5 }}>
            <div style={{ width:12, height:12, background:p.accent, borderRadius:2, display:"flex", alignItems:"center", justifyContent:"center" }}>
              <div style={{ width:4, height:4, background:p.accentFg, borderRadius:1 }}/>
            </div>
            <span style={{ fontFamily:ff, fontWeight:dw, fontSize:9.5, color:p.text, letterSpacing:"0.12em" }}>FRAME</span>
          </div>
          <div style={{ display:"flex", gap:1 }}>
            {["Apps","Docs"].map(l => <span key={l} style={{ fontFamily:ff, fontSize:8, color:p.muted, padding:"1px 4px" }}>{l}</span>)}
          </div>
          <div style={{ width:15, height:15, borderRadius:"50%", background:p.accent, color:p.accentFg, fontSize:6.5, fontWeight:700, fontFamily:ff, display:"flex", alignItems:"center", justifyContent:"center" }}>JD</div>
        </div>

        {/* Body */}
        <div style={{ display:"flex", flex:1, overflow:"hidden" }}>
          {/* Sidebar */}
          <div style={{ width:68, background:p.surface, borderRight:`1px solid ${p.border}`, padding:"6px 3px" }}>
            {nav.map((item, i) => (
              <div key={item} style={{ padding:"3px 5px", borderRadius:3, marginBottom:1, background:i===0?p.elevated:"transparent", borderLeft:`2px solid ${i===0?p.accent:"transparent"}`, color:i===0?p.accent:p.muted, fontSize:8, fontFamily:ff, fontWeight:i===0?600:400, overflow:"hidden", whiteSpace:"nowrap", textOverflow:"ellipsis", letterSpacing:"0.02em" }}>{item}</div>
            ))}
          </div>

          {/* Main */}
          <div style={{ flex:1, padding:9, overflow:"hidden" }}>
            <div style={{ fontFamily:ff, fontWeight:dw, fontSize:14, color:p.text, letterSpacing:font.lsD, marginBottom:1 }}>Dashboard</div>
            <div style={{ fontFamily:ff, fontWeight:300, fontSize:8, color:p.muted, marginBottom:8, letterSpacing:font.lsB }}>March 2026 · All operational</div>

            {/* Stats */}
            <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:3, marginBottom:5 }}>
              {stats.map((s, i) => (
                <div key={i} style={{ background:p.surface, border:`1px solid ${p.border}`, borderRadius:3, padding:"4px 5px" }}>
                  <div style={{ fontFamily:ff, fontWeight:600, fontSize:6, color:p.muted, letterSpacing:"0.1em", textTransform:"uppercase", marginBottom:1 }}>{s.l}</div>
                  <div style={{ fontFamily:ff, fontWeight:dw, fontSize:15, color:s.hi?p.accent:p.text, lineHeight:1 }}>{s.v}</div>
                </div>
              ))}
            </div>

            {/* App list */}
            <div style={{ background:p.surface, border:`1px solid ${p.border}`, borderRadius:3, overflow:"hidden" }}>
              <div style={{ padding:"3px 6px", borderBottom:`1px solid ${p.border}`, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                <span style={{ fontFamily:ff, fontWeight:600, fontSize:8, color:p.text, letterSpacing:"0.04em" }}>Recent Apps</span>
                <span style={{ fontFamily:ff, fontSize:7, fontWeight:700, background:p.accent, color:p.accentFg, padding:"1px 5px", borderRadius:2, letterSpacing:"0.08em" }}>+NEW</span>
              </div>
              {apps.map((app, i) => (
                <div key={i} style={{ padding:"3px 6px", borderBottom:i<2?`1px solid ${p.border}`:"none", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
                  <div style={{ display:"flex", alignItems:"center", gap:5 }}>
                    <div style={{ width:14, height:14, background:p.elevated, border:`1px solid ${p.border}`, borderRadius:3, flexShrink:0 }}/>
                    <span style={{ fontFamily:ff, fontSize:8, fontWeight:600, color:p.text }}>{app.n}</span>
                  </div>
                  <span style={{ fontFamily:ff, fontSize:7, fontWeight:700, padding:"1px 5px", borderRadius:2, background:app.a?`${p.accent}22`:`${p.secondary}22`, color:app.a?p.accent:p.secondary, letterSpacing:"0.08em" }}>{app.a?"ACTIVE":"BUILDING"}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function SpecimenColumn({ font, p }) {
  const ff = font.family, dw = font.dw;
  const sections = [
    { w:dw,  size:17, label:"Heading",    text:"Dashboard Overview" },
    { w:600, size:13, label:"Sub",         text:"Active Apps · Build" },
    { w:400, size:11, label:"Body",        text:"Deploy and manage apps with full observability." },
    { w:300, size:10, label:"Light",       text:"Every app benefits from shared tokens." },
  ];

  return (
    <div style={{ padding:8, display:"flex", flexDirection:"column", gap:6, overflowY:"auto", height:"100%" }}>
      {/* Hero */}
      <div style={{ background:p.bg, border:`1px solid ${p.border}`, borderRadius:7, padding:"11px 12px" }}>
        <div style={{ fontFamily:"'Space Mono',monospace", fontSize:6.5, color:p.muted, letterSpacing:"0.16em", marginBottom:8 }}>DISPLAY · {dw}</div>
        <div style={{ fontFamily:ff, fontWeight:dw, fontSize:24, color:p.text, letterSpacing:font.lsD, lineHeight:1.1, marginBottom:5 }}>FRAME Portal</div>
        <div style={{ fontFamily:ff, fontWeight:300, fontSize:11, color:p.muted, lineHeight:1.7, letterSpacing:font.lsB }}>Build, deploy and manage apps with confidence.</div>
      </div>

      {/* Scale */}
      {sections.map((s, i) => (
        <div key={i} style={{ background:p.surface, border:`1px solid ${p.border}`, borderRadius:6, padding:"9px 11px" }}>
          <div style={{ fontFamily:"'Space Mono',monospace", fontSize:6.5, color:p.muted, letterSpacing:"0.12em", marginBottom:5 }}>{s.label} · {s.w} · {s.size}px</div>
          <div style={{ fontFamily:ff, fontWeight:s.w, fontSize:s.size, color:p.text, lineHeight:1.5, letterSpacing:s.w>=600?font.lsL:font.lsB }}>{s.text}</div>
        </div>
      ))}

      {/* UI elements + charset */}
      <div style={{ background:p.surface, border:`1px solid ${p.border}`, borderRadius:6, padding:"9px 11px" }}>
        <div style={{ display:"flex", gap:5, flexWrap:"wrap", alignItems:"center", marginBottom:8 }}>
          <div style={{ background:p.accent, color:p.accentFg, padding:"4px 10px", borderRadius:3, fontFamily:ff, fontWeight:700, fontSize:8.5, letterSpacing:font.lsL }}>LAUNCH</div>
          <span style={{ fontFamily:ff, fontWeight:700, fontSize:7.5, padding:"2px 6px", borderRadius:2, background:`${p.accent}22`, color:p.accent, letterSpacing:font.lsL }}>ACTIVE</span>
          <span style={{ fontFamily:ff, fontWeight:700, fontSize:7.5, padding:"2px 6px", borderRadius:2, background:`${p.secondary}22`, color:p.secondary, letterSpacing:font.lsL }}>BUILDING</span>
        </div>
        <div style={{ borderTop:`1px solid ${p.border}`, paddingTop:7 }}>
          <div style={{ fontFamily:ff, fontWeight:dw,  fontSize:10, color:p.text, letterSpacing:"0.08em", lineHeight:1.9 }}>A B C D E F G H I J K L M</div>
          <div style={{ fontFamily:ff, fontWeight:400, fontSize:10, color:p.muted, letterSpacing:"0.05em", lineHeight:1.9 }}>N O P Q R S T U V W X Y Z</div>
          <div style={{ fontFamily:ff, fontWeight:600, fontSize:9.5, color:p.text, letterSpacing:"0.05em", lineHeight:1.9 }}>0 1 2 3 4 5 6 7 8 9 · ! @</div>
        </div>
      </div>
    </div>
  );
}

function WeightsColumn({ font, p }) {
  const ff = font.family;
  return (
    <div style={{ padding:"8px 10px", display:"flex", flexDirection:"column", gap:3, overflowY:"auto", height:"100%" }}>
      {font.weights.map(w => (
        <div key={w} style={{ background:p.surface, border:`1px solid ${p.border}`, borderRadius:5, padding:"8px 10px", display:"flex", alignItems:"baseline", gap:8 }}>
          <div style={{ width:60, flexShrink:0 }}>
            <div style={{ fontFamily:"'Space Mono',monospace", fontSize:7, color:p.accent, letterSpacing:"0.1em" }}>{font.wl[w]}</div>
            <div style={{ fontFamily:"'Space Mono',monospace", fontSize:7, color:p.muted }}>{w}</div>
          </div>
          <div style={{ fontFamily:ff, fontWeight:w, fontSize:16, color:p.text, letterSpacing:"0.04em", lineHeight:1 }}>FRAME — Build</div>
        </div>
      ))}

      {/* Size scale */}
      <div style={{ background:p.surface, border:`1px solid ${p.border}`, borderRadius:5, padding:"8px 10px", marginTop:2 }}>
        <div style={{ fontFamily:"'Space Mono',monospace", fontSize:7, color:p.muted, letterSpacing:"0.16em", marginBottom:7 }}>SIZE SCALE</div>
        {[28,22,18,14,12,10].map(px => (
          <div key={px} style={{ display:"flex", alignItems:"baseline", gap:6, marginBottom:3, paddingBottom:3, borderBottom:`1px solid ${p.border}` }}>
            <div style={{ fontFamily:"'Space Mono',monospace", fontSize:7, color:p.muted, width:22, flexShrink:0 }}>{px}</div>
            <div style={{ fontFamily:ff, fontWeight:font.dw, fontSize:px, color:p.text, letterSpacing:"0.04em", lineHeight:1.1 }}>FRAME</div>
          </div>
        ))}
      </div>
    </div>
  );
}

export default function App() {
  const [palIdx,  setPalIdx]  = useState(0);
  const [tab,     setTab]     = useState(0);
  const [chosen,  setChosen]  = useState(null);
  const p = PALETTES[palIdx];

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Titillium+Web:wght@300;400;600;700;900&family=Oxanium:wght@300;400;600;700;800&family=Chakra+Petch:wght@300;400;600;700&family=Space+Grotesk:wght@300;400;500;600;700&family=Rajdhani:wght@300;400;500;600;700&family=Orbitron:wght@400;500;600;700;800;900&family=Saira:wght@300;400;500;600;700;800&family=Space+Mono&display=swap');
        * { box-sizing:border-box; margin:0; padding:0; }
        ::-webkit-scrollbar { width:3px; height:3px; }
        ::-webkit-scrollbar-track { background:transparent; }
        ::-webkit-scrollbar-thumb { background:#2A2A2A; border-radius:2px; }
      `}</style>

      <div style={{ background:"#060606", height:"100vh", display:"flex", flexDirection:"column", color:"#fff", overflow:"hidden", fontFamily:"'Titillium Web',sans-serif" }}>

        {/* ── Header ── */}
        <header style={{ background:"#0A0A0A", borderBottom:"1px solid #181818", height:52, display:"flex", alignItems:"center", padding:"0 14px", justifyContent:"space-between", flexShrink:0, gap:8 }}>
          {/* Logo */}
          <div style={{ display:"flex", alignItems:"center", gap:8, flexShrink:0 }}>
            <CircuitMark size={20} color="#F5C831"/>
            <div>
              <div style={{ fontFamily:"'Titillium Web',sans-serif", fontWeight:700, fontSize:14, letterSpacing:"0.16em" }}>FRAME</div>
              <div style={{ fontFamily:"'Titillium Web',sans-serif", fontSize:7, color:"#F5C831", letterSpacing:"0.18em", fontWeight:600, marginTop:-1 }}>7 FONT PERSONALITIES</div>
            </div>
          </div>

          {/* Palette switcher */}
          <div style={{ display:"flex", gap:2, background:"#111", border:"1px solid #1A1A1A", borderRadius:6, padding:2, flex:1, justifyContent:"center" }}>
            {PALETTES.map((pal, i) => (
              <button key={pal.id} onClick={() => setPalIdx(i)} style={{ padding:"3px 8px", borderRadius:4, border:"none", cursor:"pointer", background:palIdx===i?"#F5C831":"transparent", color:palIdx===i?"#000":"#444", fontFamily:"'Titillium Web',sans-serif", fontWeight:700, fontSize:8.5, letterSpacing:"0.06em", whiteSpace:"nowrap", transition:"all 0.12s" }}>{pal.name}</button>
            ))}
          </div>

          {/* Tab switcher */}
          <div style={{ display:"flex", gap:2, background:"#111", border:"1px solid #1A1A1A", borderRadius:6, padding:2, flexShrink:0 }}>
            {["App UI","Specimen","Weights"].map((label, i) => (
              <button key={label} onClick={() => setTab(i)} style={{ padding:"4px 10px", borderRadius:4, border:"none", cursor:"pointer", background:tab===i?"#F5C831":"transparent", color:tab===i?"#000":"#444", fontFamily:"'Titillium Web',sans-serif", fontWeight:700, fontSize:9.5, letterSpacing:"0.06em", transition:"all 0.12s" }}>{label}</button>
            ))}
          </div>
        </header>

        {/* ── Columns ── */}
        <div style={{ display:"grid", gridTemplateColumns:"repeat(7,1fr)", flex:1, overflow:"hidden" }}>
          {FONTS.map((font, idx) => (
            <div key={font.key} style={{ display:"flex", flexDirection:"column", borderRight:idx<FONTS.length-1?"1px solid #141414":"none", overflow:"hidden" }}>

              {/* Column header */}
              <div style={{ background:"#0C0C0C", borderBottom:"1px solid #181818", padding:"9px 11px", flexShrink:0 }}>
                <div style={{ display:"inline-flex", alignItems:"center", justifyContent:"center", width:16, height:16, background:"#F5C831", color:"#000", borderRadius:2, fontFamily:"'Titillium Web',sans-serif", fontWeight:700, fontSize:8, marginBottom:5 }}>{font.id}</div>
                <div style={{ fontFamily:font.family, fontWeight:font.dw, fontSize:12, color:"#FFF", letterSpacing:"0.05em", lineHeight:1.2, marginBottom:2 }}>{font.name}</div>
                <div style={{ fontFamily:"'Titillium Web',sans-serif", fontSize:7.5, color:"#383838", letterSpacing:"0.06em", lineHeight:1.3, marginBottom:7 }}>{font.tag}</div>
                <button
                  onClick={() => setChosen(font.key)}
                  style={{ width:"100%", padding:"3px 0", borderRadius:3, border:`1px solid ${chosen===font.key?"#F5C831":"#222"}`, color:chosen===font.key?"#000":"#3A3A3A", background:chosen===font.key?"#F5C831":"transparent", cursor:"pointer", fontFamily:"'Titillium Web',sans-serif", fontWeight:700, fontSize:7.5, letterSpacing:"0.1em", transition:"all 0.12s" }}>
                  {chosen===font.key ? "✓ SELECTED" : "SELECT ✓"}
                </button>
              </div>

              {/* Column body */}
              <div style={{ flex:1, overflow:"hidden", background:tab===0?"#060606":"#080808" }}>
                {tab === 0 && <AppColumn      font={font} p={p}/>}
                {tab === 1 && <SpecimenColumn font={font} p={p}/>}
                {tab === 2 && <WeightsColumn  font={font} p={p}/>}
              </div>
            </div>
          ))}
        </div>

        {/* ── Winner bar ── */}
        {chosen && (
          <div style={{ background:"#F5C831", padding:"9px 18px", display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0 }}>
            <span style={{ fontFamily:"'Titillium Web',sans-serif", fontWeight:700, fontSize:11, color:"#000", letterSpacing:"0.1em" }}>
              ✓ PERSONALITY {FONTS.find(f => f.key===chosen)?.id} — {FONTS.find(f => f.key===chosen)?.name.toUpperCase()} SELECTED AS FRAME STANDARD FONT
            </span>
            <button onClick={() => setChosen(null)} style={{ background:"transparent", border:"1px solid #000", color:"#000", padding:"2px 10px", borderRadius:3, cursor:"pointer", fontFamily:"'Titillium Web',sans-serif", fontWeight:700, fontSize:8, letterSpacing:"0.08em" }}>DISMISS</button>
          </div>
        )}
      </div>
    </>
  );
}
