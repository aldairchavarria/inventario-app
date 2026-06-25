import { useState, useMemo } from "react";
import * as XLSX from "xlsx";

const INITIAL_PRODUCTS = [
  { id:1,  categoria:"TECH",           producto:"Case iPhone",             sku:"TCH-001", stock:25, stockMin:10, stockMax:60,  precioCompra:8,   precioVenta:18,  notas:"" },
  { id:2,  categoria:"TECH",           producto:"Pop Socket",              sku:"TCH-002", stock:9,  stockMin:15, stockMax:80,  precioCompra:3.5, precioVenta:9,   notas:"" },
  { id:3,  categoria:"TECH",           producto:"Protector de Cable",      sku:"TCH-003", stock:30, stockMin:10, stockMax:70,  precioCompra:2,   precioVenta:6,   notas:"" },
  { id:4,  categoria:"TECH",           producto:"Luz Portátil",            sku:"TCH-004", stock:20, stockMin:8,  stockMax:50,  precioCompra:9,   precioVenta:22,  notas:"" },
  { id:5,  categoria:"DAILY CHIC",     producto:"Aretes",                  sku:"CHC-001", stock:50, stockMin:20, stockMax:100, precioCompra:4,   precioVenta:12,  notas:"" },
  { id:6,  categoria:"DAILY CHIC",     producto:"Anillos",                 sku:"CHC-002", stock:45, stockMin:15, stockMax:90,  precioCompra:3,   precioVenta:10,  notas:"" },
  { id:7,  categoria:"DAILY CHIC",     producto:"Pañuelos",                sku:"CHC-003", stock:35, stockMin:12, stockMax:80,  precioCompra:5,   precioVenta:14,  notas:"" },
  { id:8,  categoria:"DAILY CHIC",     producto:"Peines",                  sku:"CHC-004", stock:60, stockMin:20, stockMax:100, precioCompra:2.5, precioVenta:7,   notas:"" },
  { id:9,  categoria:"DAILY CHIC",     producto:"Ganchos",                 sku:"CHC-005", stock:80, stockMin:25, stockMax:150, precioCompra:1.5, precioVenta:4.5, notas:"" },
  { id:10, categoria:"BAG ESSENTIALS", producto:"Neceser",                 sku:"BAG-001", stock:20, stockMin:10, stockMax:50,  precioCompra:12,  precioVenta:30,  notas:"" },
  { id:11, categoria:"BAG ESSENTIALS", producto:"Porta Toalla Higiénica",  sku:"BAG-002", stock:18, stockMin:8,  stockMax:40,  precioCompra:7,   precioVenta:18,  notas:"" },
  { id:12, categoria:"BAG ESSENTIALS", producto:"Espejos",                 sku:"BAG-003", stock:30, stockMin:10, stockMax:60,  precioCompra:5,   precioVenta:14,  notas:"" },
  { id:13, categoria:"BAG ESSENTIALS", producto:"Peines",                  sku:"BAG-004", stock:25, stockMin:10, stockMax:55,  precioCompra:2.5, precioVenta:7,   notas:"" },
  { id:14, categoria:"BAG ESSENTIALS", producto:"Organizador de Cables",   sku:"BAG-005", stock:15, stockMin:8,  stockMax:40,  precioCompra:8,   precioVenta:20,  notas:"" },
  { id:15, categoria:"BAG ESSENTIALS", producto:"Ligas Set",               sku:"BAG-006", stock:40, stockMin:15, stockMax:80,  precioCompra:3,   precioVenta:8,   notas:"" },
  { id:16, categoria:"BAG ESSENTIALS", producto:"EJEMPLO",                 sku:"BAG-007", stock:40, stockMin:15, stockMax:80,  precioCompra:3,   precioVenta:8,   notas:"" },
  { id:17, categoria:"BAG ESSENTIALS", producto:"Pastilleros",             sku:"BAG-008", stock:22, stockMin:10, stockMax:50,  precioCompra:6,   precioVenta:15,  notas:"" },
];

const CATEGORIAS = ["TECH", "DAILY CHIC", "BAG ESSENTIALS"];

function getEstado(stock, min, max) {
  if (stock === 0)  return { label: "⛔ SIN STOCK",   color: "#ef4444", bg: "#fef2f2" };
  if (stock <= min) return { label: "⚠️ BAJO STOCK",  color: "#f59e0b", bg: "#fffbeb" };
  if (stock >= max) return { label: "📦 SOBRE STOCK", color: "#8b5cf6", bg: "#f5f3ff" };
  return                   { label: "✅ OK",           color: "#10b981", bg: "#f0fdf4" };
}

function margen(compra, venta) {
  if (!venta) return 0;
  return ((venta - compra) / venta) * 100;
}

function nextId(products) {
  return products.length ? Math.max(...products.map(p => p.id)) + 1 : 1;
}

function suggestSku(categoria, products) {
  const prefixes = { "TECH": "TCH", "DAILY CHIC": "CHC", "BAG ESSENTIALS": "BAG" };
  const prefix = prefixes[categoria] || "PRD";
  const existing = products.filter(p => p.sku.startsWith(prefix));
  const nums = existing.map(p => parseInt(p.sku.split("-")[1] || "0", 10)).filter(Boolean);
  const next = nums.length ? Math.max(...nums) + 1 : 1;
  return `${prefix}-${String(next).padStart(3, "0")}`;
}

function exportToExcel(products) {
  const today = new Date().toLocaleDateString("es-PE");
  const wb = XLSX.utils.book_new();
  const n = products.length;

  const invData = [
    ["📦  SISTEMA DE INVENTARIO"],
    [`Gestión de Productos por Categoría · Actualizado: ${today}`],
    ["CATEGORÍA","PRODUCTO","SKU","STOCK\nACTUAL","STOCK\nMÍNIMO","STOCK\nMÁXIMO","PRECIO\nCOMPRA (S/)","PRECIO\nVENTA (S/)","MARGEN\n%","VALOR\nINVENTARIO","ESTADO","NOTAS"],
    ...products.map((p, i) => {
      const row = i + 4;
      return [
        p.categoria, p.producto, p.sku,
        p.stock, p.stockMin, p.stockMax,
        p.precioCompra, p.precioVenta,
        { f: `=(H${row}-G${row})/H${row}` },
        { f: `=D${row}*G${row}` },
        { f: `=IF(D${row}=0,"⛔ SIN STOCK",IF(D${row}<=E${row},"⚠️ BAJO STOCK",IF(D${row}>=F${row},"📦 SOBRE STOCK","✅ OK")))` },
        p.notas || "",
      ];
    }),
    ["TOTALES","","",
      { f: `=SUM(D4:D${n+3})` },"","","","","",
      { f: `=SUM(J4:J${n+3})` },
      { f: `=COUNTIF(K4:K${n+3},"*BAJO*")&" productos con bajo stock"` },""],
  ];
  const wsInv = XLSX.utils.aoa_to_sheet(invData);
  wsInv["!cols"] = [20,24,12,10,10,10,14,14,10,14,16,20].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsInv, "Inventario");

  const cats = [...new Set(products.map(p => p.categoria))];
  const dashData = [
    [null, "📊  DASHBOARD · RESUMEN DE INVENTARIO"],
    [],
    [null, "TOTAL\nPRODUCTOS", null, "TOTAL\nUNIDADES"],
    [null, { f: `=COUNTA(Inventario!B4:B${n+3})` }, null, { f: `=Inventario!D${n+4}` }],
    [],
    [null, "VALOR TOTAL\nINVENTARIO", null, "BAJO STOCK / SIN STOCK"],
    [null, { f: `=Inventario!J${n+4}` }, null, { f: `=COUNTIF(Inventario!K4:K${n+3},"*BAJO*")+COUNTIF(Inventario!K4:K${n+3},"*SIN*")` }],
    [],
    [null, "DESGLOSE POR CATEGORÍA"],
    [null, "CATEGORÍA","PRODUCTOS","UNIDADES","VALOR (S/)","ESTADO"],
    ...cats.map(cat => [
      null, cat,
      { f: `=COUNTIF(Inventario!A4:A${n+3},"${cat}")` },
      { f: `=SUMIF(Inventario!A4:A${n+3},"${cat}",Inventario!D4:D${n+3})` },
      { f: `=SUMIF(Inventario!A4:A${n+3},"${cat}",Inventario!J4:J${n+3})` },
      { f: `=COUNTIF(Inventario!K4:K${n+3},"*BAJO*")&" alertas"` },
    ]),
    [],
    [null, "LEYENDA DE ESTADOS"],
    [null, "✅ OK", "Stock dentro del rango establecido"],
    [null, "⚠️ BAJO STOCK", "Stock igual o por debajo del mínimo"],
    [null, "⛔ SIN STOCK", "Stock en cero, requiere reposición urgente"],
    [null, "📦 SOBRE STOCK", "Stock igual o mayor al máximo definido"],
  ];
  const wsDash = XLSX.utils.aoa_to_sheet(dashData);
  XLSX.utils.book_append_sheet(wb, wsDash, "Dashboard");

  const guiaData = [
    ["➕  GUÍA PARA AGREGAR NUEVOS PRODUCTOS"],[], ["PASO","ACCIÓN","DETALLE"],
    [1,"Ir a hoja «Inventario»","Haz clic en la pestaña «Inventario» en la parte inferior"],
    [2,"Ir a la última fila de datos","Desplázate hasta debajo del último producto ingresado"],
    [3,"Ingresar Categoría (Col A)","Escribe: TECH · DAILY CHIC · BAG ESSENTIALS (o nueva)"],
    [4,"Ingresar Producto (Col B)","Nombre descriptivo del producto"],
    [5,"Ingresar SKU (Col C)","Código único: TCH-XXX / CHC-XXX / BAG-XXX"],
    [6,"Stock Actual (Col D)","Unidades actuales en almacén"],
    [7,"Stock Mínimo (Col E)","Nivel mínimo antes de reabastecer"],
    [8,"Stock Máximo (Col F)","Nivel máximo que deseas almacenar"],
    [9,"Precio Compra (Col G)","Costo de adquisición en S/"],
    [10,"Precio Venta (Col H)","Precio de venta al público en S/"],
    [11,"Copiar fórmulas (Col I-K)","Copia las celdas I·J·K de la fila anterior y pégalas"],
    [12,"Dashboard se actualiza solo","El resumen en «Dashboard» refleja los cambios automáticamente"],
  ];
  const wsGuia = XLSX.utils.aoa_to_sheet(guiaData);
  XLSX.utils.book_append_sheet(wb, wsGuia, "Agregar Producto");

  XLSX.writeFile(wb, "Inventario_Escalable.xlsx");
}

function Badge({ estado }) {
  return (
    <span style={{
      display:"inline-block", padding:"2px 10px", borderRadius:99,
      fontSize:12, fontWeight:600, color:estado.color, background:estado.bg,
      border:`1px solid ${estado.color}33`,
    }}>{estado.label}</span>
  );
}

function StatCard({ icon, label, value, sub, accent }) {
  return (
    <div style={{
      background:"#fff", borderRadius:14, padding:"18px 22px",
      border:`1.5px solid ${accent}22`, boxShadow:"0 1px 6px #0001",
      minWidth:140, flex:1,
    }}>
      <div style={{fontSize:22, marginBottom:4}}>{icon}</div>
      <div style={{fontSize:26, fontWeight:800, color:accent, lineHeight:1}}>{value}</div>
      <div style={{fontSize:12, color:"#64748b", marginTop:3}}>{label}</div>
      {sub && <div style={{fontSize:11, color:"#94a3b8", marginTop:2}}>{sub}</div>}
    </div>
  );
}

function Modal({ open, onClose, children }) {
  if (!open) return null;
  return (
    <div onClick={onClose} style={{
      position:"fixed", inset:0, background:"#0007", zIndex:1000,
      display:"flex", alignItems:"center", justifyContent:"center",
    }}>
      <div onClick={e => e.stopPropagation()} style={{
        background:"#fff", borderRadius:18, padding:28,
        width:"min(540px, 96vw)", maxHeight:"90vh", overflowY:"auto",
        boxShadow:"0 8px 40px #0003",
      }}>{children}</div>
    </div>
  );
}

const EMPTY_FORM = {
  categoria:"TECH", producto:"", sku:"", stock:"",
  stockMin:"", stockMax:"", precioCompra:"", precioVenta:"", notas:"",
};

function ProductForm({ initial, products, onSave, onCancel, isEdit }) {
  const [form, setForm] = useState(initial || EMPTY_FORM);
  const [errors, setErrors] = useState({});

  function set(k, v) {
    setForm(f => {
      const next = { ...f, [k]: v };
      if (k === "categoria" && !isEdit) {
        next.sku = suggestSku(v, products.filter(p => p.id !== initial?.id));
      }
      return next;
    });
    setErrors(e => ({ ...e, [k]: undefined }));
  }

  function validate() {
    const e = {};
    if (!form.producto.trim()) e.producto = "Requerido";
    if (!form.sku.trim())      e.sku = "Requerido";
    if (form.stock === "")     e.stock = "Requerido";
    if (form.stockMin === "")  e.stockMin = "Requerido";
    if (form.stockMax === "")  e.stockMax = "Requerido";
    if (form.precioCompra === "") e.precioCompra = "Requerido";
    if (form.precioVenta === "")  e.precioVenta = "Requerido";
    if (Number(form.precioCompra) >= Number(form.precioVenta)) e.precioVenta = "Debe ser mayor al precio de compra";
    if (Number(form.stockMin) >= Number(form.stockMax)) e.stockMax = "Debe ser mayor al stock mínimo";
    const duplicate = products.find(p => p.sku === form.sku.trim() && p.id !== initial?.id);
    if (duplicate) e.sku = "SKU ya existe";
    return e;
  }

  function handleSave() {
    const e = validate();
    if (Object.keys(e).length) { setErrors(e); return; }
    onSave({
      ...form,
      stock: Number(form.stock), stockMin: Number(form.stockMin),
      stockMax: Number(form.stockMax), precioCompra: Number(form.precioCompra),
      precioVenta: Number(form.precioVenta),
      sku: form.sku.trim().toUpperCase(),
      producto: form.producto.trim(),
    });
  }

  const inp = (err) => ({
    width:"100%", padding:"9px 12px", borderRadius:8,
    border:`1.5px solid ${err ? "#ef4444" : "#e2e8f0"}`,
    fontSize:14, outline:"none", boxSizing:"border-box",
    background: err ? "#fff5f5" : "#f8fafc", fontFamily:"inherit",
  });

  const mg = margen(Number(form.precioCompra), Number(form.precioVenta));
  const est = form.stock !== "" ? getEstado(Number(form.stock), Number(form.stockMin), Number(form.stockMax)) : null;

  return (
    <div>
      <h2 style={{margin:"0 0 20px", fontSize:20, fontWeight:800, color:"#1e293b"}}>
        {isEdit ? "✏️ Editar Producto" : "➕ Nuevo Producto"}
      </h2>
      {est && (
        <div style={{background:"#f1f5f9", borderRadius:10, padding:"10px 14px", marginBottom:18, display:"flex", gap:16, flexWrap:"wrap", alignItems:"center"}}>
          <Badge estado={est} />
          {form.precioCompra && form.precioVenta && (
            <span style={{fontSize:13, color:"#475569"}}>
              Margen: <b style={{color: mg > 50 ? "#10b981" : "#f59e0b"}}>{mg.toFixed(1)}%</b>
            </span>
          )}
          {form.stock && form.precioCompra && (
            <span style={{fontSize:13, color:"#475569"}}>
              Valor inv: <b style={{color:"#6366f1"}}>S/ {(Number(form.stock)*Number(form.precioCompra)).toFixed(2)}</b>
            </span>
          )}
        </div>
      )}
      <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:"0 18px"}}>
        <div style={{gridColumn:"1 / -1", marginBottom:14}}>
          <div style={{fontSize:12, fontWeight:600, color:"#475569", marginBottom:4}}>Categoría</div>
          <select value={form.categoria} onChange={e => set("categoria", e.target.value)}
            style={{...inp(), width:"100%", cursor:"pointer"}}>
            {CATEGORIAS.map(c => <option key={c}>{c}</option>)}
            <option value="__nueva__">+ Nueva categoría…</option>
          </select>
          {form.categoria === "__nueva__" && (
            <input placeholder="Ej: HOME DECO" value={form._catCustom || ""}
              onChange={e => { set("_catCustom", e.target.value); set("categoria", e.target.value); }}
              style={{...inp(), marginTop:8}} />
          )}
        </div>
        {[
          ["producto","Producto","Ej: Cartera Mediana",false],
          ["sku","SKU","Ej: BAG-009",false],
        ].map(([k,lbl,ph]) => (
          <div key={k} style={{marginBottom:14}}>
            <div style={{fontSize:12, fontWeight:600, color:"#475569", marginBottom:4}}>{lbl}</div>
            {errors[k] && <div style={{fontSize:11, color:"#ef4444", marginBottom:3}}>⚠ {errors[k]}</div>}
            <input value={form[k]} onChange={e => set(k, e.target.value)} placeholder={ph} style={inp(errors[k])} />
          </div>
        ))}
        <div style={{marginBottom:14}}>
          <div style={{fontSize:12, fontWeight:600, color:"#475569", marginBottom:4}}>Stock Actual</div>
          {errors.stock && <div style={{fontSize:11, color:"#ef4444", marginBottom:3}}>⚠ {errors.stock}</div>}
          <input type="number" min={0} value={form.stock} onChange={e => set("stock", e.target.value)} style={inp(errors.stock)} />
        </div>
        <div></div>
        {[
          ["stockMin","Stock Mínimo",errors.stockMin],
          ["stockMax","Stock Máximo",errors.stockMax],
          ["precioCompra","Precio Compra (S/)",errors.precioCompra],
          ["precioVenta","Precio Venta (S/)",errors.precioVenta],
        ].map(([k,lbl,err]) => (
          <div key={k} style={{marginBottom:14}}>
            <div style={{fontSize:12, fontWeight:600, color:"#475569", marginBottom:4}}>{lbl}</div>
            {err && <div style={{fontSize:11, color:"#ef4444", marginBottom:3}}>⚠ {err}</div>}
            <input type="number" min={0} step={0.5} value={form[k]} onChange={e => set(k, e.target.value)} style={inp(err)} />
          </div>
        ))}
        <div style={{gridColumn:"1 / -1", marginBottom:14}}>
          <div style={{fontSize:12, fontWeight:600, color:"#475569", marginBottom:4}}>Notas (opcional)</div>
          <input value={form.notas} onChange={e => set("notas", e.target.value)}
            placeholder="Observaciones adicionales..." style={inp()} />
        </div>
      </div>
      <div style={{display:"flex", gap:10, marginTop:10}}>
        <button onClick={handleSave} style={{
          flex:1, padding:"11px 0", borderRadius:9, border:"none",
          background:"linear-gradient(135deg,#6366f1,#8b5cf6)", color:"#fff",
          fontWeight:700, fontSize:15, cursor:"pointer", fontFamily:"inherit",
        }}>
          {isEdit ? "Guardar cambios" : "Agregar producto"}
        </button>
        <button onClick={onCancel} style={{
          padding:"11px 20px", borderRadius:9, border:"1.5px solid #e2e8f0",
          background:"#fff", fontWeight:600, fontSize:14, cursor:"pointer",
          color:"#475569", fontFamily:"inherit",
        }}>Cancelar</button>
      </div>
    </div>
  );
}

export default function App() {
  const [products, setProducts] = useState(INITIAL_PRODUCTS);
  const [modal, setModal] = useState(null);
  const [search, setSearch] = useState("");
  const [filterCat, setFilterCat] = useState("Todas");
  const [filterEst, setFilterEst] = useState("Todos");
  const [toast, setToast] = useState(null);
  const [deleteConfirm, setDeleteConfirm] = useState(null);
  const [sortCol, setSortCol] = useState(null);
  const [sortDir, setSortDir] = useState("asc");

  const allCategorias = useMemo(() => ["Todas", ...new Set(products.map(p => p.categoria))], [products]);

  function showToast(msg, color = "#10b981") {
    setToast({ msg, color });
    setTimeout(() => setToast(null), 3000);
  }

  function handleAdd(data) {
    setProducts(p => [...p, { ...data, id: nextId(products) }]);
    setModal(null);
    showToast(`✅ "${data.producto}" agregado al inventario`);
  }

  function handleEdit(data) {
    setProducts(p => p.map(x => x.id === data.id ? data : x));
    setModal(null);
    showToast(`✏️ "${data.producto}" actualizado`);
  }

  function handleDelete(id) {
    const p = products.find(x => x.id === id);
    setProducts(prev => prev.filter(x => x.id !== id));
    setDeleteConfirm(null);
    showToast(`🗑️ "${p?.producto}" eliminado`, "#ef4444");
  }

  const filtered = useMemo(() => {
    let list = products;
    if (filterCat !== "Todas") list = list.filter(p => p.categoria === filterCat);
    if (filterEst !== "Todos") list = list.filter(p => getEstado(p.stock, p.stockMin, p.stockMax).label.includes(filterEst === "OK" ? "OK" : filterEst));
    if (search) {
      const q = search.toLowerCase();
      list = list.filter(p => p.producto.toLowerCase().includes(q) || p.sku.toLowerCase().includes(q) || p.categoria.toLowerCase().includes(q));
    }
    if (sortCol) {
      list = [...list].sort((a, b) => {
        let va = a[sortCol], vb = b[sortCol];
        if (typeof va === "string") va = va.toLowerCase();
        if (typeof vb === "string") vb = vb.toLowerCase();
        return sortDir === "asc" ? (va > vb ? 1 : -1) : (va < vb ? 1 : -1);
      });
    }
    return list;
  }, [products, filterCat, filterEst, search, sortCol, sortDir]);

  function toggleSort(col) {
    if (sortCol === col) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortCol(col); setSortDir("asc"); }
  }

  const kpis = useMemo(() => ({
    totalVal: products.reduce((s,p) => s + p.stock * p.precioCompra, 0),
    bajoStock: products.filter(p => getEstado(p.stock,p.stockMin,p.stockMax).label.includes("BAJO")).length,
    sinStock:  products.filter(p => p.stock === 0).length,
    avgMargen: products.length ? products.reduce((s,p) => s + margen(p.precioCompra,p.precioVenta), 0) / products.length : 0,
  }), [products]);

  const thS = (col) => ({
    padding:"10px 12px", textAlign:"left", fontSize:11,
    fontWeight:700, color:"#64748b", textTransform:"uppercase",
    letterSpacing:".5px", cursor: col ? "pointer" : "default",
    userSelect:"none", whiteSpace:"nowrap",
    background: sortCol === col ? "#f1f5f9" : "transparent",
  });

  const SortIcon = ({ col }) => sortCol === col
    ? <span style={{marginLeft:4}}>{sortDir === "asc" ? "↑" : "↓"}</span>
    : <span style={{marginLeft:4, opacity:.3}}>↕</span>;

  return (
    <div style={{fontFamily:"'Segoe UI',system-ui,sans-serif", background:"#f8fafc", minHeight:"100vh", paddingBottom:40}}>
      {toast && (
        <div style={{
          position:"fixed", top:18, right:18, zIndex:9999,
          background:toast.color, color:"#fff", borderRadius:10,
          padding:"12px 20px", fontWeight:600, fontSize:14,
          boxShadow:"0 4px 20px #0003",
        }}>{toast.msg}</div>
      )}

      {/* Header */}
      <div style={{background:"linear-gradient(135deg,#6366f1 0%,#8b5cf6 60%,#a78bfa 100%)", padding:"28px 32px 24px", color:"#fff"}}>
        <div style={{display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:14}}>
          <div>
            <div style={{fontSize:24, fontWeight:800, letterSpacing:"-0.5px"}}>📦 Sistema de Inventario</div>
            <div style={{fontSize:13, opacity:.8, marginTop:3}}>
              {products.length} productos · Actualizado: {new Date().toLocaleDateString("es-PE")}
            </div>
          </div>
          <div style={{display:"flex", gap:10}}>
            <button onClick={() => exportToExcel(products)} style={{
              padding:"10px 18px", borderRadius:9, border:"2px solid rgba(255,255,255,.4)",
              background:"rgba(255,255,255,.15)", color:"#fff",
              fontWeight:700, fontSize:14, cursor:"pointer", fontFamily:"inherit",
            }}>⬇️ Exportar Excel</button>
            <button onClick={() => setModal("add")} style={{
              padding:"10px 18px", borderRadius:9, border:"none",
              background:"#fff", color:"#6366f1",
              fontWeight:700, fontSize:14, cursor:"pointer", fontFamily:"inherit",
            }}>➕ Nuevo Producto</button>
          </div>
        </div>
        <div style={{display:"flex", gap:12, marginTop:20, flexWrap:"wrap"}}>
          <StatCard icon="📦" label="Total productos"   value={products.length}                                                                      accent="#6366f1" />
          <StatCard icon="📊" label="Unidades en stock" value={products.reduce((s,p)=>s+p.stock,0)}                                                  accent="#8b5cf6" />
          <StatCard icon="💰" label="Valor inventario"  value={`S/ ${kpis.totalVal.toLocaleString("es-PE",{minimumFractionDigits:2})}`}              accent="#a78bfa" />
          <StatCard icon="📈" label="Margen promedio"   value={`${kpis.avgMargen.toFixed(1)}%`}                                                      accent="#7c3aed" />
          {(kpis.bajoStock+kpis.sinStock) > 0 && (
            <StatCard icon="⚠️" label="Alertas de stock" value={kpis.bajoStock+kpis.sinStock} sub={`${kpis.sinStock} sin stock`} accent="#ef4444" />
          )}
        </div>
      </div>

      {/* Filtros */}
      <div style={{background:"#fff", borderBottom:"1px solid #e2e8f0", padding:"14px 32px", display:"flex", gap:12, flexWrap:"wrap", alignItems:"center"}}>
        <input placeholder="🔍 Buscar producto, SKU o categoría..." value={search}
          onChange={e => setSearch(e.target.value)}
          style={{padding:"8px 14px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:14, width:260, outline:"none", background:"#f8fafc", fontFamily:"inherit"}} />
        <select value={filterCat} onChange={e => setFilterCat(e.target.value)}
          style={{padding:"8px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:14, background:"#f8fafc", cursor:"pointer", fontFamily:"inherit"}}>
          {allCategorias.map(c => <option key={c}>{c}</option>)}
        </select>
        <select value={filterEst} onChange={e => setFilterEst(e.target.value)}
          style={{padding:"8px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:14, background:"#f8fafc", cursor:"pointer", fontFamily:"inherit"}}>
          {["Todos","OK","BAJO STOCK","SIN STOCK","SOBRE STOCK"].map(e => <option key={e}>{e}</option>)}
        </select>
        {(search || filterCat !== "Todas" || filterEst !== "Todos") && (
          <button onClick={() => { setSearch(""); setFilterCat("Todas"); setFilterEst("Todos"); }}
            style={{padding:"8px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:13, cursor:"pointer", color:"#64748b", background:"#fff", fontFamily:"inherit"}}>
            ✕ Limpiar filtros
          </button>
        )}
        <span style={{marginLeft:"auto", fontSize:13, color:"#94a3b8"}}>{filtered.length} de {products.length} productos</span>
      </div>

      {/* Tabla */}
      <div style={{padding:"24px 32px"}}>
        <div style={{background:"#fff", borderRadius:14, border:"1px solid #e2e8f0", overflow:"hidden", boxShadow:"0 1px 8px #0001"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%", borderCollapse:"collapse"}}>
              <thead>
                <tr style={{borderBottom:"2px solid #e2e8f0"}}>
                  {[["categoria","Categoría"],["producto","Producto"],["sku","SKU"],["stock","Stock"],["stockMin","Mín"],["stockMax","Máx"],["precioCompra","Compra (S/)"],["precioVenta","Venta (S/)"]].map(([col,lbl]) => (
                    <th key={col} onClick={() => toggleSort(col)} style={thS(col)}>
                      {lbl}<SortIcon col={col} />
                    </th>
                  ))}
                  {["Margen%","Valor Inv.","Estado"].map(lbl => <th key={lbl} style={thS(null)}>{lbl}</th>)}
                  <th style={thS(null)}>Acciones</th>
                </tr>
              </thead>
              <tbody>
                {filtered.length === 0 && (
                  <tr><td colSpan={12} style={{textAlign:"center", padding:40, color:"#94a3b8", fontSize:15}}>No se encontraron productos</td></tr>
                )}
                {filtered.map((p, i) => {
                  const est = getEstado(p.stock, p.stockMin, p.stockMax);
                  const mg  = margen(p.precioCompra, p.precioVenta);
                  return (
                    <tr key={p.id}
                      style={{borderBottom:"1px solid #f1f5f9", background: i%2===0?"#fff":"#fafbfc"}}
                      onMouseEnter={e => e.currentTarget.style.background="#f0f9ff"}
                      onMouseLeave={e => e.currentTarget.style.background=i%2===0?"#fff":"#fafbfc"}>
                      <td style={{padding:"11px 12px"}}>
                        <span style={{background:"#f1f5f9", borderRadius:6, padding:"2px 8px", fontSize:11, fontWeight:700, color:"#475569"}}>{p.categoria}</span>
                      </td>
                      <td style={{padding:"11px 12px", fontSize:14, fontWeight:600, color:"#1e293b"}}>{p.producto}</td>
                      <td style={{padding:"11px 12px", fontSize:12, color:"#64748b", fontFamily:"monospace"}}>{p.sku}</td>
                      <td style={{padding:"11px 12px", fontSize:14, fontWeight:700, color:p.stock===0?"#ef4444":"#1e293b", textAlign:"right"}}>{p.stock}</td>
                      <td style={{padding:"11px 12px", fontSize:13, color:"#64748b", textAlign:"right"}}>{p.stockMin}</td>
                      <td style={{padding:"11px 12px", fontSize:13, color:"#64748b", textAlign:"right"}}>{p.stockMax}</td>
                      <td style={{padding:"11px 12px", fontSize:13, textAlign:"right"}}>S/ {p.precioCompra.toFixed(2)}</td>
                      <td style={{padding:"11px 12px", fontSize:13, textAlign:"right"}}>S/ {p.precioVenta.toFixed(2)}</td>
                      <td style={{padding:"11px 12px", fontSize:13, textAlign:"right", color:mg>=60?"#10b981":mg>=40?"#f59e0b":"#ef4444", fontWeight:600}}>{mg.toFixed(1)}%</td>
                      <td style={{padding:"11px 12px", fontSize:13, textAlign:"right", color:"#6366f1", fontWeight:600}}>S/ {(p.stock*p.precioCompra).toFixed(2)}</td>
                      <td style={{padding:"11px 12px"}}><Badge estado={est} /></td>
                      <td style={{padding:"11px 12px"}}>
                        <div style={{display:"flex", gap:6}}>
                          <button onClick={() => setModal({edit:p})}
                            style={{padding:"5px 10px", borderRadius:7, border:"1px solid #e2e8f0", background:"#fff", cursor:"pointer", fontSize:13}}>✏️</button>
                          <button onClick={() => setDeleteConfirm(p.id)}
                            style={{padding:"5px 10px", borderRadius:7, border:"1px solid #fee2e2", background:"#fff5f5", cursor:"pointer", fontSize:13}}>🗑️</button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              {filtered.length > 0 && (
                <tfoot>
                  <tr style={{borderTop:"2px solid #e2e8f0", background:"#f8fafc"}}>
                    <td colSpan={3} style={{padding:"10px 12px", fontSize:13, fontWeight:700, color:"#475569"}}>TOTALES ({filtered.length} productos)</td>
                    <td style={{padding:"10px 12px", fontWeight:800, color:"#1e293b", textAlign:"right"}}>{filtered.reduce((s,p)=>s+p.stock,0)}</td>
                    <td colSpan={5}></td>
                    <td style={{padding:"10px 12px", fontWeight:800, color:"#6366f1", textAlign:"right"}}>S/ {filtered.reduce((s,p)=>s+p.stock*p.precioCompra,0).toFixed(2)}</td>
                    <td colSpan={2}></td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
        </div>
      </div>

      <Modal open={!!modal} onClose={() => setModal(null)}>
        {modal === "add" && (
          <ProductForm products={products}
            initial={{...EMPTY_FORM, sku: suggestSku("TECH", products)}}
            onSave={handleAdd} onCancel={() => setModal(null)} isEdit={false} />
        )}
        {modal?.edit && (
          <ProductForm products={products} initial={modal.edit}
            onSave={handleEdit} onCancel={() => setModal(null)} isEdit={true} />
        )}
      </Modal>

      <Modal open={!!deleteConfirm} onClose={() => setDeleteConfirm(null)}>
        <div style={{textAlign:"center"}}>
          <div style={{fontSize:40, marginBottom:12}}>🗑️</div>
          <h3 style={{margin:"0 0 8px", color:"#1e293b"}}>¿Eliminar producto?</h3>
          <p style={{color:"#64748b", margin:"0 0 24px", fontSize:14}}>Esta acción no se puede deshacer.</p>
          <div style={{display:"flex", gap:10, justifyContent:"center"}}>
            <button onClick={() => handleDelete(deleteConfirm)} style={{padding:"10px 24px", borderRadius:9, border:"none", background:"#ef4444", color:"#fff", fontWeight:700, cursor:"pointer", fontFamily:"inherit"}}>Eliminar</button>
            <button onClick={() => setDeleteConfirm(null)} style={{padding:"10px 24px", borderRadius:9, border:"1.5px solid #e2e8f0", background:"#fff", fontWeight:600, cursor:"pointer", color:"#475569", fontFamily:"inherit"}}>Cancelar</button>
          </div>
        </div>
      </Modal>

      <style>{`* { box-sizing: border-box; } select, input, button { font-family: inherit; }`}</style>
    </div>
  );
}
