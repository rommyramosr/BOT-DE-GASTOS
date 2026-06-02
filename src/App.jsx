import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

const CATEGORIES = {
  comida: ["comida", "almuerzo", "desayuno", "cena", "restaurante", "cafe", "café", "pizza", "burger", "pollo", "mercado", "supermercado", "snack", "helado", "bebida", "menú"],
  transporte: ["taxi", "uber", "bus", "metro", "micro", "gasolina", "combustible", "pasaje", "moto", "cabify", "peaje"],
  entretenimiento: ["cine", "netflix", "spotify", "juego", "concierto", "bar", "discoteca", "fiesta", "salida", "evento", "trago", "tragos", "mike's", "mikes", "cerveza", "licor", "ron", "pisco", "shots"],
  salud: ["farmacia", "medicina", "doctor", "médico", "consulta", "pastilla", "vitamina", "clinica", "hospital"],
  ropa: ["ropa", "zapatos", "camisa", "polo", "pantalon", "vestido", "zapatilla", "accesorio"],
  servicios: ["internet", "luz", "recibo", "factura", "recibo de agua", "recibo de luz", "recibo de gas", "planilla", "alquiler", "renta", "seguro", "cable", "telefonia", "telefonía"],
  educacion: ["curso", "libro", "universidad", "colegio", "matrícula", "matricula", "útiles", "utiles", "clases", "taller", "capacitación", "capacitacion", "carrera"],
  mascotas: ["veterinario", "veterinaria", "mascota", "perro", "gato", "comida de perro", "comida de chichu", "antiparasitario", "vacuna mascota", "pienso", "chichu", "arena para gato", "collar", "correa", "juguete mascota", "guarderia mascota"],
  viajes: ["hotel", "hospedaje", "vuelo", "pasaje aéreo", "pasaje aereo", "tour", "maleta", "airbnb", "hostal", "agencia", "excursión", "excursion"],
  deporte: ["gimnasio", "gym", "cancha", "suplemento", "proteína", "proteina", "creatina", "equipo deportivo", "bicicleta", "natación", "natacion", "entrenador", "agua de mesa", "agua cielo", "agua san luis", "agua san mateo", "hidratacion", "hidratación"],
  higiene: ["shampoo", "champú", "champu", "jabón", "jabon", "desodorante", "pasta dental", "cepillo", "papel higienico", "papel higiénico", "toalla", "crema", "loción", "locion", "afeitadora", "rasuradora", "hilo dental", "enjuague", "gel", "higiene"],
  otros: []
};

const MONTHS = {
  enero:1,febrero:2,marzo:3,abril:4,mayo:5,junio:6,
  julio:7,agosto:8,septiembre:9,octubre:10,noviembre:11,diciembre:12,
  ene:1,feb:2,mar:3,abr:4,jun:6,jul:7,ago:8,sep:9,oct:10,nov:11,dic:12
};

function fmtDate(d) {
  return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`;
}

// Extract date and return { date, cleanMsg } — removes date tokens from msg so amount parser won't hit them
function extractDateAndClean(msg) {
  const lower = msg.toLowerCase();
  const today = new Date();
  let date = null;
  let clean = msg;

  // "ayer"
  if (/\bayer\b/.test(lower)) {
    const d = new Date(today); d.setDate(d.getDate()-1);
    date = fmtDate(d);
    clean = clean.replace(/\bayer\b/gi, "");
  }
  // "anteayer"
  else if (/\bante\s*ayer\b/.test(lower)) {
    const d = new Date(today); d.setDate(d.getDate()-2);
    date = fmtDate(d);
    clean = clean.replace(/\bante\s*ayer\b/gi, "");
  }
  // weekdays
  else {
    const weekdays = [
      {re:/\bdomingo\b/i, dow:0},{re:/\blunes\b/i, dow:1},{re:/\bmartes\b/i, dow:2},
      {re:/\bmi[eé]rcoles\b/i, dow:3},{re:/\bjueves\b/i, dow:4},{re:/\bviernes\b/i, dow:5},
      {re:/\bs[aá]bado\b/i, dow:6}
    ];
    for (const {re, dow} of weekdays) {
      if (re.test(lower)) {
        const d = new Date(today);
        const diff = (d.getDay() - dow + 7) % 7 || 7;
        d.setDate(d.getDate()-diff);
        date = fmtDate(d);
        clean = clean.replace(re, "");
        break;
      }
    }
  }

  // "28 de mayo" / "el 28 de mayo" / "28 de mayo de 2025"
  const longRe = /\b(?:el\s+)?(\d{1,2})\s+de\s+([a-záéíóúü]+)(?:\s+(?:de\s+)?(\d{4}))?\b/gi;
  clean = clean.replace(longRe, (match, day, monthStr, year) => {
    const month = MONTHS[monthStr.toLowerCase()];
    if (month && parseInt(day) >= 1 && parseInt(day) <= 31) {
      const y = year ? parseInt(year) : today.getFullYear();
      date = `${String(parseInt(day)).padStart(2,"0")}/${String(month).padStart(2,"0")}/${y}`;
      return ""; // remove from clean
    }
    return match;
  });

  // "28/05" or "28/05/2025" or "28-05" (only if no date found yet)
  if (!date) {
    const shortRe = /\b(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?\b/g;
    clean = clean.replace(shortRe, (match, d, m, y) => {
      const day = parseInt(d), month = parseInt(m);
      if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
        let year = y ? parseInt(y) : today.getFullYear();
        if (year < 100) year += 2000;
        date = `${String(day).padStart(2,"0")}/${String(month).padStart(2,"0")}/${year}`;
        return "";
      }
      return match;
    });
  }

  return { date: date || fmtDate(today), clean: clean.replace(/\s+/g, " ").trim() };
}

function detectCategory(text) {
  const lower = text.toLowerCase();
  for (const [cat, keywords] of Object.entries(CATEGORIES)) {
    if (cat === "otros") continue;
    if (keywords.some(k => lower.includes(k))) return cat;
  }
  return "otros";
}

// Keywords that indicate "monto" context — amount usually follows these
const AMOUNT_SIGNALS = /(\b(son|fue|pagué|pague|gasté|gaste|costó|costo|vale|cuesta|precio|monto|total|por|s\/)\s*)(\d+(?:\.\d+)?)/gi;

function parseMessage(original) {
  // Replace commas used as text separators before parsing
  const noCommas = original.replace(/,/g, " ");
  const { date, clean } = extractDateAndClean(noCommas);
  const category = detectCategory(clean);

  // Try to find amount after signal words first
  let amount = null;
  const signalMatch = [...clean.matchAll(AMOUNT_SIGNALS)];
  if (signalMatch.length > 0) {
    amount = parseFloat(signalMatch[signalMatch.length-1][3]);
  }

  // Fallback: last standalone number (not part of a word)
  if (amount === null) {
    const nums = [...clean.matchAll(/\b\d+(?:\.\d+)?\b/g)];
    if (nums.length === 0) return null;
    amount = parseFloat(nums[nums.length-1][0]);
  }

  // Build description — also remove commas used as separators
  let desc = clean
    .replace(/\d+(?:\.\d+)?/g, "")
    .replace(/,/g, " ")
    .replace(/gasté|gaste|compré|compre|pagué|pague|costó|costo|son|fue|vale|cuesta|precio|monto|total|por|en|el|la|los|las|sol(es)?|s\/|pe[ñn]u?/gi, "")
    .replace(/\s+/g, " ").trim() || "Gasto";
  desc = desc.charAt(0).toUpperCase() + desc.slice(1);

  return { amount, category, desc, date };
}

const CAT_COLORS = {
  comida:"#FF6B6B",transporte:"#4ECDC4",entretenimiento:"#FFE66D",
  salud:"#A8E6CF",ropa:"#C3B1E1",servicios:"#F7B733",
  educacion:"#80DEEA",mascotas:"#FFCC80",viajes:"#80CBC4",deporte:"#EF9A9A",
  higiene:"#CE93D8",otros:"#B0BEC5"
};
const CAT_EMOJI = {
  comida:"🍽️",transporte:"🚌",entretenimiento:"🎬",
  salud:"💊",ropa:"👕",servicios:"⚡",
  educacion:"🎓",mascotas:"🐾",viajes:"✈️",deporte:"🏋️",
  higiene:"🧴",otros:"📦"
};
const CAT_LIST = Object.keys(CAT_COLORS);

function dateToSortable(d) {
  if (!d) return "";
  const [dd,mm,yyyy] = d.split("/");
  return `${yyyy}${mm}${dd}`;
}
function isoToDisplay(iso) {
  if (!iso) return "";
  const [y,m,d] = iso.split("-");
  return `${d}/${m}/${y}`;
}
function displayToIso(disp) {
  if (!disp) return "";
  const [d,m,y] = disp.split("/");
  return `${y}-${m}-${d}`;
}

const inputStyle = {
  border:"1.5px solid #e0e0e0",borderRadius:10,padding:"8px 11px",
  fontSize:13,outline:"none",background:"#fff",width:"100%",boxSizing:"border-box"
};

function EditForm({expense, onSave, onCancel}) {
  const [form, setForm] = useState({
    desc: String(expense.desc || ""),
    amount: String(expense.amount || ""),
    category: expense.category || "otros",
    date: expense.date || fmtDate(new Date())
  });
  return (
    <div style={{padding:12}}>
      <div style={{fontSize:11,fontWeight:700,color:"#128C7E",marginBottom:8}}>✏️ Cambia solo lo que quieras</div>
      <div style={{fontSize:10,color:"#999",marginBottom:4}}>Descripción</div>
      <input value={form.desc} onChange={ev=>setForm(f=>({...f,desc:ev.target.value}))} style={{border:"1.5px solid #e0e0e0",borderRadius:10,padding:"8px 11px",fontSize:13,outline:"none",background:"#fff",width:"100%",boxSizing:"border-box",marginBottom:8}}/>
      <div style={{display:"flex",gap:8,marginBottom:8}}>
        <div style={{flex:1}}>
          <div style={{fontSize:10,color:"#999",marginBottom:4}}>Monto (S/)</div>
          <input type="text" inputMode="decimal" value={form.amount} onChange={ev=>setForm(f=>({...f,amount:ev.target.value}))} style={{border:"1.5px solid #e0e0e0",borderRadius:10,padding:"8px 11px",fontSize:13,outline:"none",background:"#fff",width:"100%",boxSizing:"border-box"}}/>
        </div>
        <div style={{flex:1}}>
          <div style={{fontSize:10,color:"#999",marginBottom:4}}>Fecha</div>
          <input type="date" value={displayToIso(form.date)} onChange={ev=>setForm(f=>({...f,date:isoToDisplay(ev.target.value)}))} style={{border:"1.5px solid #e0e0e0",borderRadius:10,padding:"8px 11px",fontSize:13,outline:"none",background:"#fff",width:"100%",boxSizing:"border-box"}}/>
        </div>
      </div>
      <div style={{fontSize:10,color:"#999",marginBottom:4}}>Categoría</div>
      <select value={form.category} onChange={ev=>setForm(f=>({...f,category:ev.target.value}))} style={{border:"1.5px solid #e0e0e0",borderRadius:10,padding:"8px 11px",fontSize:13,outline:"none",background:"#fff",width:"100%",boxSizing:"border-box",marginBottom:10}}>
        {CAT_LIST.map(c=><option key={c} value={c}>{CAT_EMOJI[c]} {c.charAt(0).toUpperCase()+c.slice(1)}</option>)}
      </select>
      <div style={{display:"flex",gap:8}}>
        <button onClick={()=>onSave({...form,amount:parseFloat(form.amount)||0})} style={{flex:1,background:"#25D366",color:"#fff",border:"none",borderRadius:8,padding:"8px",fontWeight:700,cursor:"pointer"}}>✅ Guardar</button>
        <button onClick={onCancel} style={{flex:1,background:"#eee",color:"#555",border:"none",borderRadius:8,padding:"8px",cursor:"pointer"}}>Cancelar</button>
      </div>
    </div>
  );
}

function MonthShortcuts({fDateFrom,fDateTo,setFDateFrom,setFDateTo}) {
  const months=[];
  for(let i=0;i<6;i++){
    const d=new Date(); d.setDate(1); d.setMonth(d.getMonth()-i);
    const y=d.getFullYear(), m=String(d.getMonth()+1).padStart(2,"0");
    const label=d.toLocaleString("es-PE",{month:"short",year:"numeric"});
    const from=`${y}-${m}-01`;
    const lastDay=new Date(y,d.getMonth()+1,0).getDate();
    const to=`${y}-${m}-${String(lastDay).padStart(2,"0")}`;
    const isActive=fDateFrom===from&&fDateTo===to;
    months.push({from,to,label,isActive});
  }
  return (
    <div style={{marginTop:12}}>
      <div style={{fontSize:11,fontWeight:700,color:"#555",marginBottom:6}}>📅 EXPORTAR POR MES RÁPIDO</div>
      <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
        {months.map(({from,to,label,isActive})=>(
          <button key={from} onClick={()=>{setFDateFrom(from);setFDateTo(to);}} style={{background:isActive?"#128C7E":"#f0f0f0",color:isActive?"#fff":"#555",border:"none",borderRadius:16,padding:"5px 12px",fontSize:11,cursor:"pointer",fontWeight:isActive?700:400,marginBottom:4}}>{label}</button>
        ))}
      </div>
    </div>
  );
}

export default function GastosTracker() {
  const [messages, setMessages] = useState([
    { from:"bot", text:"¡Hola! 👋 Escríbeme lo que gastaste y detecto fecha, monto y categoría automáticamente:\n\n• \"Taxi a Sodoma 29 de mayo 15 soles\"\n• \"Almuerzo 30 ayer\"\n• \"Netflix 35 el lunes\"\n• \"Farmacia 80 28/05\"\n\nSin fecha → uso hoy 📅" }
  ]);
  const [input, setInput] = useState("");
  const [expenses, setExpenses] = useState(() => {
    try {
      const saved = localStorage.getItem("gastosbot_expenses");
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch(e) {}
    return [];
  });
  const [tab, setTab] = useState("chat");
  const [exporting, setExporting] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);

  const [fDateFrom, setFDateFrom] = useState("");
  const [fDateTo, setFDateTo] = useState("");
  const [fCats, setFCats] = useState([]);
  const [fMinAmount, setFMinAmount] = useState("");
  const [fMaxAmount, setFMaxAmount] = useState("");
  const [showFilters, setShowFilters] = useState(false);

  const messagesEndRef = useRef(null);
  const inputRef = useRef(null);

  // Save to localStorage whenever expenses change
  useEffect(() => {
    try { localStorage.setItem("gastosbot_expenses", JSON.stringify(expenses)); } catch(e) {}
  }, [expenses]);

  useEffect(() => { messagesEndRef.current?.scrollIntoView({ behavior:"smooth" }); }, [messages]);

  const filteredExpenses = expenses.filter(e => {
    const ds = dateToSortable(e.date);
    if (fDateFrom && ds < dateToSortable(isoToDisplay(fDateFrom))) return false;
    if (fDateTo && ds > dateToSortable(isoToDisplay(fDateTo))) return false;
    if (fCats.length > 0 && !fCats.includes(e.category)) return false;
    if (fMinAmount !== "" && e.amount < parseFloat(fMinAmount)) return false;
    if (fMaxAmount !== "" && e.amount > parseFloat(fMaxAmount)) return false;
    return true;
  });

  const activeFilterCount = [fDateFrom,fDateTo,fCats.length>0?"cat":"",fMinAmount,fMaxAmount].filter(Boolean).length;
  function clearFilters() { setFDateFrom("");setFDateTo("");setFCats([]);setFMinAmount("");setFMaxAmount(""); }
  function toggleCat(cat) { setFCats(prev => prev.includes(cat)?prev.filter(c=>c!==cat):[...prev,cat]); }

  function sendMessage() {
    const text = input.trim();
    if (!text) return;
    const parsed = parseMessage(text);
    const userMsg = { from:"user", text };
    let botMsg;
    if (parsed) {
      const newExpense = { ...parsed, id:Date.now() };
      setExpenses(prev => [...prev, newExpense]);
      const isToday = parsed.date === fmtDate(new Date());
      botMsg = { from:"bot", text:`✅ Registrado!\n\n${CAT_EMOJI[parsed.category]} *${parsed.desc}*\n💰 S/ ${parsed.amount.toFixed(2)}\n📂 ${parsed.category.charAt(0).toUpperCase()+parsed.category.slice(1)}\n📅 ${parsed.date}${isToday?" (hoy)":""}` };
    } else {
      botMsg = { from:"bot", text:"Hmm, no pude detectar el monto 🤔 Prueba: \"Taxi 25\" o \"Almuerzo 30 ayer\"" };
    }
    setMessages(prev => [...prev, userMsg, botMsg]);
    setInput("");
    inputRef.current?.focus();
  }

  function deleteExpense(id) { setExpenses(prev=>prev.filter(e=>e.id!==id)); setConfirmDeleteId(null); }
  function startEdit(id) { setConfirmDeleteId(null); setEditingId(id); }
  function saveEdit(id, form) {
    setExpenses(prev=>prev.map(e=>e.id===id?{...e,...form}:e));
    setEditingId(null);
  }

  function exportExcel() {
    const toExport = activeFilterCount>0 ? filteredExpenses : expenses;
    if (toExport.length===0) return;
    setExporting(true);
    setTimeout(() => {
      const wb = XLSX.utils.book_new();
      const data = [["Fecha","Descripción","Categoría","Monto (S/)"]];
      toExport.forEach(e=>data.push([e.date,e.desc,e.category.charAt(0).toUpperCase()+e.category.slice(1),e.amount]));
      const tot = toExport.reduce((s,e)=>s+e.amount,0);
      data.push([],["TOTAL","","",tot.toFixed(2)]);
      const ws1=XLSX.utils.aoa_to_sheet(data); ws1["!cols"]=[{wch:12},{wch:28},{wch:18},{wch:14}];
      XLSX.utils.book_append_sheet(wb,ws1,"Gastos");
      const byCat={};
      toExport.forEach(e=>{byCat[e.category]=(byCat[e.category]||0)+e.amount;});
      const catData=[["Categoría","Total (S/)","% del total"]];
      Object.entries(byCat).sort((a,b)=>b[1]-a[1]).forEach(([cat,amt])=>catData.push([cat.charAt(0).toUpperCase()+cat.slice(1),amt.toFixed(2),((amt/tot)*100).toFixed(1)+"%"]));
      const ws2=XLSX.utils.aoa_to_sheet(catData); ws2["!cols"]=[{wch:18},{wch:14},{wch:14}];
      XLSX.utils.book_append_sheet(wb,ws2,"Por Categoría");
      const byDate={};
      toExport.forEach(e=>{byDate[e.date]=(byDate[e.date]||0)+e.amount;});
      const dateData=[["Fecha","Total (S/)"]];
      Object.entries(byDate).sort().forEach(([d,a])=>dateData.push([d,a.toFixed(2)]));
      const ws3=XLSX.utils.aoa_to_sheet(dateData); ws3["!cols"]=[{wch:14},{wch:14}];
      XLSX.utils.book_append_sheet(wb,ws3,"Por Fecha");
      XLSX.writeFile(wb,"mis_gastos.xlsx");
      setExporting(false);
    },300);
  }

  const ftotal = filteredExpenses.reduce((s,e)=>s+e.amount,0);
  const fbyCat={},fbyDate={};
  filteredExpenses.forEach(e=>{fbyCat[e.category]=(fbyCat[e.category]||0)+e.amount; fbyDate[e.date]=(fbyDate[e.date]||0)+e.amount;});

  const tabBtn=(id,label)=>(
    <button onClick={()=>setTab(id)} style={{background:tab===id?"rgba(255,255,255,0.3)":"transparent",border:"none",color:"#fff",borderRadius:8,padding:"4px 10px",cursor:"pointer",fontSize:11,fontWeight:600}}>{label}</button>
  );

  return (
    <div style={{fontFamily:"'Segoe UI',sans-serif",background:"#0B1426",minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",padding:"16px",boxSizing:"border-box"}}>
      <div style={{width:"100%",maxWidth:"420px",display:"flex",flexDirection:"column",height:"calc(100vh - 32px)",borderRadius:"20px",overflow:"hidden",boxShadow:"0 20px 60px rgba(0,0,0,0.5)",background:"#fff"}}>

        {/* Header */}
        <div style={{background:"linear-gradient(135deg,#25D366 0%,#128C7E 100%)",padding:"14px 16px",display:"flex",alignItems:"center",gap:"12px"}}>
          <div style={{width:40,height:40,borderRadius:"50%",background:"rgba(255,255,255,0.25)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20}}>💸</div>
          <div>
            <div style={{color:"#fff",fontWeight:700,fontSize:15}}>GastosBot</div>
            <div style={{color:"rgba(255,255,255,0.8)",fontSize:11}}>{expenses.length} registros · S/ {expenses.reduce((s,e)=>s+e.amount,0).toFixed(2)} total</div>
          </div>
          <div style={{marginLeft:"auto",display:"flex",gap:4}}>
            {tabBtn("chat","Chat")} {tabBtn("list","Lista")} {tabBtn("stats","Stats")}
          </div>
        </div>

        {/* ── CHAT ── */}
        {tab==="chat" && (<>
          <div style={{flex:1,overflowY:"auto",padding:"12px",background:"#ECE5DD",backgroundImage:"radial-gradient(circle at 1px 1px,rgba(0,0,0,0.04) 1px,transparent 0)",backgroundSize:"20px 20px"}}>
            {messages.map((m,i)=>(
              <div key={i} style={{display:"flex",justifyContent:m.from==="user"?"flex-end":"flex-start",marginBottom:8}}>
                <div style={{maxWidth:"82%",background:m.from==="user"?"#DCF8C6":"#fff",borderRadius:m.from==="user"?"16px 16px 4px 16px":"16px 16px 16px 4px",padding:"8px 12px",boxShadow:"0 1px 2px rgba(0,0,0,0.1)",fontSize:13,lineHeight:1.5,whiteSpace:"pre-wrap",color:"#111"}}>
                  {m.text.split(/(\*[^*]+\*)/).map((p,j)=>p.startsWith("*")&&p.endsWith("*")?<strong key={j}>{p.slice(1,-1)}</strong>:p)}
                </div>
              </div>
            ))}
            <div ref={messagesEndRef}/>
          </div>
          <div style={{background:"#ECE5DD",padding:"0 12px 8px",display:"flex",gap:6,flexWrap:"wrap"}}>
            {["Taxi 25 ayer","Almuerzo 30 el lunes","Netflix 35 28/05"].map(hint=>(
              <button key={hint} onClick={()=>setInput(hint)} style={{background:"#fff",border:"1px solid #ccc",borderRadius:16,padding:"4px 10px",fontSize:11,cursor:"pointer",color:"#555"}}>{hint}</button>
            ))}
          </div>
          <div style={{background:"#F0F0F0",padding:"10px 12px",display:"flex",gap:8,alignItems:"center",borderTop:"1px solid #ddd"}}>
            <input ref={inputRef} value={input} onChange={e=>setInput(e.target.value)} onKeyDown={e=>e.key==="Enter"&&sendMessage()} placeholder='Ej: "Taxi a Lima 29 de mayo 15 soles"' style={{flex:1,border:"none",borderRadius:24,padding:"10px 16px",fontSize:13,outline:"none",background:"#fff",boxShadow:"0 1px 3px rgba(0,0,0,0.1)"}}/>
            <button onClick={sendMessage} style={{width:40,height:40,borderRadius:"50%",background:"linear-gradient(135deg,#25D366,#128C7E)",border:"none",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>➤</button>
          </div>
          {expenses.length>0&&(
            <button onClick={exportExcel} disabled={exporting} style={{margin:"0 12px 12px",background:exporting?"#aaa":"linear-gradient(135deg,#25D366,#128C7E)",color:"#fff",border:"none",borderRadius:12,padding:"12px",fontSize:14,fontWeight:700,cursor:"pointer"}}>
              {exporting?"Generando...":`📥 Exportar Excel (${expenses.length} gastos)`}
            </button>
          )}
        </>)}

        {/* ── LISTA ── */}
        {tab==="list"&&(
          <div style={{flex:1,overflowY:"auto",background:"#f8f8f8"}}>
            {expenses.length===0?(
              <div style={{textAlign:"center",color:"#aaa",marginTop:60,fontSize:14}}><div style={{fontSize:48}}>📋</div><p>No hay gastos todavía.</p></div>
            ):(<>
              <div style={{padding:"10px 12px 4px",fontSize:11,color:"#999",fontWeight:600}}>{expenses.length} GASTOS</div>
              {[...expenses].sort((a,b)=>b.id-a.id).map(e=>(
                <div key={e.id} style={{margin:"0 10px 8px",borderRadius:14,background:"#fff",boxShadow:"0 1px 6px rgba(0,0,0,0.07)",overflow:"hidden"}}>
                  {editingId===e.id?(
                    <EditForm expense={e} onSave={(form)=>saveEdit(e.id,form)} onCancel={()=>setEditingId(null)}/>
                  ):confirmDeleteId===e.id?(
                    /* confirm delete inline */
                    <div style={{padding:"12px 14px",background:"#fff5f5",display:"flex",alignItems:"center",gap:10}}>
                      <div style={{flex:1,fontSize:13,color:"#c62828",fontWeight:600}}>¿Borrar "{e.desc}"?</div>
                      <button onClick={()=>deleteExpense(e.id)} style={{background:"#e53935",color:"#fff",border:"none",borderRadius:8,padding:"6px 14px",fontWeight:700,cursor:"pointer",fontSize:13}}>Sí, borrar</button>
                      <button onClick={()=>setConfirmDeleteId(null)} style={{background:"#eee",color:"#555",border:"none",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:13}}>No</button>
                    </div>
                  ):(
                    <div style={{display:"flex",alignItems:"center",padding:"10px 12px",gap:10}}>
                      <div style={{width:36,height:36,borderRadius:10,background:CAT_COLORS[e.category]+"30",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>{CAT_EMOJI[e.category]}</div>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:13,fontWeight:600,color:"#222",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{e.desc}</div>
                        <div style={{fontSize:11,color:"#999"}}>{e.date} · {e.category.charAt(0).toUpperCase()+e.category.slice(1)}</div>
                      </div>
                      <div style={{fontWeight:800,color:"#128C7E",fontSize:14,flexShrink:0}}>S/ {e.amount.toFixed(2)}</div>
                      <button onClick={()=>{startEdit(e.id);setConfirmDeleteId(null);}} style={{background:"#EEF7FF",border:"none",borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:15,flexShrink:0}}>✏️</button>
                      <button onClick={()=>{setConfirmDeleteId(e.id);setEditingId(null);}} style={{background:"#FFF0F0",border:"none",borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:15,flexShrink:0}}>🗑️</button>
                    </div>
                  )}
                </div>
              ))}
              <button onClick={exportExcel} disabled={exporting} style={{margin:"4px 10px 16px",width:"calc(100% - 20px)",background:"linear-gradient(135deg,#25D366,#128C7E)",color:"#fff",border:"none",borderRadius:12,padding:"12px",fontSize:14,fontWeight:700,cursor:"pointer"}}>📥 Exportar Excel</button>
            </>)}
          </div>
        )}

        {/* ── STATS ── */}
        {tab==="stats"&&(
          <div style={{flex:1,overflowY:"auto",background:"#f4f6f8"}}>
            <div style={{padding:"10px 12px 0"}}>
              <button onClick={()=>setShowFilters(f=>!f)} style={{width:"100%",background:showFilters?"#128C7E":"#fff",color:showFilters?"#fff":"#128C7E",border:"2px solid #128C7E",borderRadius:12,padding:"9px 14px",fontSize:13,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                <span>🔍 Filtros{activeFilterCount>0?` (${activeFilterCount} activos)`:""}</span>
                <span style={{fontSize:11,opacity:0.8}}>{showFilters?"▲ Ocultar":"▼ Ver filtros"}</span>
              </button>
              {showFilters&&(
                <div style={{background:"#fff",borderRadius:14,padding:14,marginTop:8,boxShadow:"0 2px 10px rgba(0,0,0,0.08)"}}>
                  <div style={{marginBottom:12}}>
                    <div style={{fontSize:11,fontWeight:700,color:"#555",marginBottom:6}}>📅 RANGO DE FECHAS</div>
                    <div style={{display:"flex",gap:8}}>
                      <div style={{flex:1}}><div style={{fontSize:10,color:"#999",marginBottom:3}}>Desde</div><input type="date" value={fDateFrom} onChange={e=>setFDateFrom(e.target.value)} style={{...inputStyle}}/></div>
                      <div style={{flex:1}}><div style={{fontSize:10,color:"#999",marginBottom:3}}>Hasta</div><input type="date" value={fDateTo} onChange={e=>setFDateTo(e.target.value)} style={{...inputStyle}}/></div>
                    </div>
                  </div>
                  <div style={{marginBottom:12}}>
                    <div style={{fontSize:11,fontWeight:700,color:"#555",marginBottom:6}}>💰 RANGO DE MONTO (S/)</div>
                    <div style={{display:"flex",gap:8}}>
                      <div style={{flex:1}}><div style={{fontSize:10,color:"#999",marginBottom:3}}>Mínimo</div><input type="number" min="0" value={fMinAmount} onChange={e=>setFMinAmount(e.target.value)} placeholder="0" style={{...inputStyle}}/></div>
                      <div style={{flex:1}}><div style={{fontSize:10,color:"#999",marginBottom:3}}>Máximo</div><input type="number" min="0" value={fMaxAmount} onChange={e=>setFMaxAmount(e.target.value)} placeholder="Sin límite" style={{...inputStyle}}/></div>
                    </div>
                  </div>
                  <div style={{marginBottom:12}}>
                    <div style={{fontSize:11,fontWeight:700,color:"#555",marginBottom:6}}>📂 CATEGORÍAS {fCats.length===0?"(todas)":`(${fCats.length} sel.)`}</div>
                    <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                      {CAT_LIST.map(cat=>{const active=fCats.includes(cat);return<button key={cat} onClick={()=>toggleCat(cat)} style={{background:active?CAT_COLORS[cat]:"#f0f0f0",color:active?"#222":"#666",border:active?`2px solid ${CAT_COLORS[cat]}`:"2px solid #e0e0e0",borderRadius:20,padding:"5px 11px",fontSize:12,cursor:"pointer",fontWeight:active?700:400}}>{CAT_EMOJI[cat]} {cat.charAt(0).toUpperCase()+cat.slice(1)}</button>;})}
                    </div>
                  </div>
                  {activeFilterCount>0&&<button onClick={clearFilters} style={{width:"100%",background:"#fff0f0",color:"#e53935",border:"1.5px solid #ffcdd2",borderRadius:10,padding:"8px",fontSize:13,fontWeight:700,cursor:"pointer"}}>🗑️ Limpiar filtros</button>}
                  <MonthShortcuts fDateFrom={fDateFrom} fDateTo={fDateTo} setFDateFrom={setFDateFrom} setFDateTo={setFDateTo}/>
                </div>
              )}
            </div>
            <div style={{padding:"10px 12px 16px"}}>
              {expenses.length===0?(
                <div style={{textAlign:"center",color:"#aaa",marginTop:40}}><div style={{fontSize:48}}>📊</div><p>Ve al chat y empieza a registrar!</p></div>
              ):filteredExpenses.length===0?(
                <div style={{textAlign:"center",color:"#aaa",marginTop:30,background:"#fff",borderRadius:16,padding:24}}>
                  <div style={{fontSize:36}}>🔍</div><p>Ningún gasto coincide con los filtros.</p>
                  <button onClick={clearFilters} style={{background:"#128C7E",color:"#fff",border:"none",borderRadius:10,padding:"8px 18px",cursor:"pointer",fontSize:13,fontWeight:700}}>Limpiar filtros</button>
                </div>
              ):(<>
                <div style={{background:"linear-gradient(135deg,#25D366,#128C7E)",borderRadius:16,padding:"14px 18px",color:"#fff",marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div><div style={{fontSize:11,opacity:0.85}}>{activeFilterCount>0?"TOTAL FILTRADO":"TOTAL GASTADO"}</div><div style={{fontSize:28,fontWeight:800}}>S/ {ftotal.toFixed(2)}</div></div>
                  <div style={{textAlign:"right"}}><div style={{fontSize:20,fontWeight:800}}>{filteredExpenses.length}</div><div style={{fontSize:11,opacity:0.85}}>transacciones</div></div>
                </div>
                <div style={{background:"#fff",borderRadius:16,padding:14,marginBottom:12,boxShadow:"0 2px 8px rgba(0,0,0,0.06)"}}>
                  <div style={{fontWeight:700,fontSize:13,marginBottom:10,color:"#333"}}>📂 Por categoría</div>
                  {Object.entries(fbyCat).sort((a,b)=>b[1]-a[1]).map(([cat,amt])=>(
                    <div key={cat} style={{marginBottom:10}}>
                      <div style={{display:"flex",justifyContent:"space-between",marginBottom:3,fontSize:13}}>
                        <span>{CAT_EMOJI[cat]} {cat.charAt(0).toUpperCase()+cat.slice(1)}</span>
                        <span style={{fontWeight:700,color:CAT_COLORS[cat]}}>S/ {amt.toFixed(2)}</span>
                      </div>
                      <div style={{background:"#eee",borderRadius:8,height:8,overflow:"hidden"}}>
                        <div style={{height:"100%",borderRadius:8,background:CAT_COLORS[cat],width:`${(amt/ftotal)*100}%`}}/>
                      </div>
                      <div style={{fontSize:10,color:"#999",textAlign:"right"}}>{((amt/ftotal)*100).toFixed(0)}%</div>
                    </div>
                  ))}
                </div>
                <div style={{background:"#fff",borderRadius:16,padding:14,marginBottom:12,boxShadow:"0 2px 8px rgba(0,0,0,0.06)"}}>
                  <div style={{fontWeight:700,fontSize:13,marginBottom:10,color:"#333"}}>📅 Por fecha</div>
                  {Object.entries(fbyDate).sort().map(([date,amt])=>{
                    const maxDay=Math.max(...Object.values(fbyDate));
                    return(
                      <div key={date} style={{marginBottom:8}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:3,fontSize:12}}>
                          <span style={{color:"#555"}}>{date}</span>
                          <span style={{fontWeight:700,color:"#128C7E"}}>S/ {amt.toFixed(2)}</span>
                        </div>
                        <div style={{background:"#eee",borderRadius:8,height:6,overflow:"hidden"}}>
                          <div style={{height:"100%",borderRadius:8,background:"linear-gradient(90deg,#25D366,#128C7E)",width:`${(amt/maxDay)*100}%`}}/>
                        </div>
                      </div>
                    );
                  })}
                </div>
                <button onClick={exportExcel} disabled={exporting} style={{width:"100%",background:"linear-gradient(135deg,#25D366,#128C7E)",color:"#fff",border:"none",borderRadius:12,padding:"12px",fontSize:14,fontWeight:700,cursor:"pointer"}}>
                  {exporting?"Generando...":`📥 Exportar Excel${activeFilterCount>0?" (filtrado)":""}`}
                </button>
              </>)}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
