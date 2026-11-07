 "use client";
import {useEffect,useMemo,useState} from "react";
import {Activity,CheckCircle2,ClipboardList,Download,ExternalLink,FileSpreadsheet,FolderOpen,History,LoaderCircle,MessageCircle,Plus,Search,ShieldCheck,Trash2,X,RefreshCw} from "lucide-react";
import {buildWorkbook,safeFileName} from "@/lib/excel";
import {emptyItem,Inspection,InspectionItem,Line,MODEL_PREFIX,Model,STORAGE_DRAFT} from "@/lib/types";

const today=()=>{const d=new Date();return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`};
const blank=():Inspection=>({id:crypto.randomUUID(),inspector:"",inspectionDate:today(),line:["LH"],lotDate:today(),shift:"1",items:[emptyItem()],createdAt:new Date().toISOString()});
const allModels:Model[]=["SU2ID LH","SU2ID RH","KS LH","KS RH"];
const gasUrl=process.env.NEXT_PUBLIC_GAS_WEB_APP_URL||"";
const fmt=(value:string)=>value?new Date(`${value}T00:00:00`).toLocaleDateString("id-ID",{day:"2-digit",month:"long",year:"numeric"}):"-";
type SavedRecord=Inspection & {fileName?:string;driveUrl?:string;downloadUrl?:string;folderUrl?:string};

export default function Home(){
 const [tab,setTab]=useState<"form"|"history">("form");
 const [form,setForm]=useState<Inspection>(blank());
 const [records,setRecords]=useState<SavedRecord[]>([]);
 const [search,setSearch]=useState("");
 const [filterDate,setFilterDate]=useState("");
 const [filterShift,setFilterShift]=useState("");
 const [filterLine,setFilterLine]=useState("");
 const [filterModel,setFilterModel]=useState("");
 const [busy,setBusy]=useState(false);
 const [loading,setLoading]=useState(false);
 const [error,setError]=useState("");
 const [notice,setNotice]=useState("");
 const [successModal,setSuccessModal]=useState(false);
 const [successFilename,setSuccessFilename]=useState("");
 const [confirm,setConfirm]=useState<"one"|"all"|null>(null);
 const [selected,setSelected]=useState<SavedRecord|null>(null);
 const [draftReady,setDraftReady]=useState(false);

 async function loadHistory(showLoading=true){
  if(!gasUrl){setRecords([]);return false;}
  if(showLoading)setLoading(true);
  try{
   const separator=gasUrl.includes("?")?"&":"?";
   const response=await fetch(`${gasUrl}${separator}action=list&_=${Date.now()}`,{method:"GET",cache:"no-store"});
   const result=await response.json();
   if(!response.ok||!result.success)throw new Error(result.message||"Gagal mengambil History dari Google Sheets.");
   setRecords(Array.isArray(result.records)?result.records:[]);
   return true;
  }catch(e){
   setError(e instanceof Error?e.message:"Gagal mengambil History dari Google Sheets.");
   return false;
  }finally{
   if(showLoading)setLoading(false);
  }
 }

 useEffect(()=>{
  try{
   const d=localStorage.getItem(STORAGE_DRAFT);
   if(d)setForm(JSON.parse(d));
  }catch(e){
   console.error("Gagal memulihkan draft PDI",e);
  }finally{
   setDraftReady(true);
  }
  void loadHistory(false);
 },[]);
 useEffect(()=>{if(draftReady)localStorage.setItem(STORAGE_DRAFT,JSON.stringify(form))},[form,draftReady]);

 const totalSample=useMemo(()=>form.items.reduce((s,x)=>s+(Number(x.sampleQty)||0),0),[form.items]);
 const totalNg=useMemo(()=>form.items.reduce((s,x)=>s+(Number(x.ng)||0),0),[form.items]);
 const total=totalSample+totalNg;
 const allowedModels=allModels.filter(m=>form.line.includes(m.endsWith("LH")?"LH":"RH"));
 const filtered=useMemo(()=>records.filter(r=>{
   const itemModels=r.items.map(x=>x.model);
   const searchable=[r.inspector,r.inspectionDate,...itemModels,...r.items.map(x=>x.suffix)].join(" ").toLowerCase();
   return searchable.includes(search.trim().toLowerCase())&&(!filterDate||r.inspectionDate===filterDate)&&(!filterShift||r.shift===filterShift)&&(!filterLine||r.line.includes(filterLine as Line))&&(!filterModel||itemModels.includes(filterModel as Model));
 }),[records,search,filterDate,filterShift,filterLine,filterModel]);

 const update=(patch:Partial<Inspection>)=>setForm(p=>({...p,...patch}));
 const updateItem=(id:string,patch:Partial<InspectionItem>)=>update({items:form.items.map(x=>x.id===id?{...x,...patch}:x)});
 function addItem(){if(form.items.length>=20){setError("Maksimal 20 Inspection Item sesuai kapasitas template Excel.");return;}const prev=form.items[form.items.length-1]||emptyItem();update({items:[...form.items,{...emptyItem(),model:prev.model,suffix:prev.suffix}]});setError("");}
 function reset(){
  if(!window.confirm("Reset seluruh isian form?")) return;
  setForm(blank());
  setError("");
  setNotice("");
}
 function resetFilters(){setSearch("");setFilterDate("");setFilterShift("");setFilterLine("");setFilterModel("");}
 function validate(){
  const issues:string[]=[];
  if(!form.inspector.trim())issues.push("Inspector Name belum diisi.");
  if(!form.inspectionDate)issues.push("Inspection Date belum dipilih.");
  if(!form.line.length)issues.push("Pilih minimal satu Inspection Line (LH/RH).");
  if(!form.lotDate)issues.push("Inspection Lot belum dipilih.");
  if(!form.shift)issues.push("Shift belum dipilih.");
  if(!form.items.length||form.items.length>20)issues.push("Jumlah Inspection Item harus 1 sampai 20.");
  form.items.forEach((x,i)=>{
   if(!x.model)issues.push(`Item ${i+1}: Model belum dipilih.`);
   if(!/^[A-Z0-9]{6}$/.test(x.suffix.toUpperCase()))issues.push(`Item ${i+1}: Part Number harus 6 karakter. Contoh: 020NB9.`);
   if(x.sampleQty.trim()===""||!Number.isInteger(Number(x.sampleQty))||Number(x.sampleQty)<=0)issues.push(`Item ${i+1}: SAMPLE QTY wajib diisi dan lebih dari 0.`);
   if(x.ng!==""&&(!Number.isInteger(Number(x.ng))||Number(x.ng)<0))issues.push(`Item ${i+1}: NG harus bilangan bulat nol atau lebih.`);
  });
  return issues;
 }
 const blobToBase64=(blob:Blob)=>new Promise<string>((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(String(r.result).split(",")[1]);r.onerror=reject;r.readAsDataURL(blob)});
 async function generate(){
  setError("");setNotice("");const issues=validate();if(issues.length){setError(issues.join("\n"));return;}
  setBusy(true);
  try{
   const normalizedForm={...form,items:form.items.map(item=>({...item,ng:item.ng.trim()===""?"0":item.ng}))};
   const blob=await buildWorkbook(normalizedForm);
   const lineName=form.line.join("-");
   const fileName=safeFileName(`${form.inspector}-PDI-${lineName}-${form.inspectionDate}.xlsx`);
   const base64=await blobToBase64(blob);
   let driveUrl="",downloadUrl="",folderUrl="";
   if(gasUrl){
    const response=await fetch(gasUrl,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({action:"save",fileName,base64,inspection:{...normalizedForm,totalSample,totalNg,totalQty:total}})});
    const result=await response.json();if(!result.success)throw new Error(result.message||"Gagal menyimpan ke Google Drive.");
    driveUrl=result.fileUrl||"";downloadUrl=result.downloadUrl||"";folderUrl=result.folderUrl||"";
   }
   const objectUrl=URL.createObjectURL(blob);const a=document.createElement("a");a.href=objectUrl;a.download=fileName;a.click();setTimeout(()=>URL.revokeObjectURL(objectUrl),10000);
   if(gasUrl)await loadHistory(false);
   setSuccessFilename(fileName);setSuccessModal(true);
   if(gasUrl)setNotice("Laporan berhasil dibuat dan disimpan ke Google Drive.");
   else setNotice("File Excel berhasil dibuat dan diunduh. Google Drive belum terhubung.");
   localStorage.removeItem(STORAGE_DRAFT);
  }catch(e){setError(e instanceof Error?e.message:"Gagal membuat laporan Excel.");}
  finally{setBusy(false)}
 }
 async function deleteOne(){
  if(!selected)return;
  if(!gasUrl){setError("Google Apps Script belum dikonfigurasi.");return;}
  setLoading(true);
  try{
   const response=await fetch(gasUrl,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({action:"delete",recordId:selected.id})});
   const result=await response.json();
   if(!result.success)throw new Error(result.message||"Gagal menghapus History dari Google Sheets.");
   await loadHistory(false);
   setSelected(null);
   setConfirm(null);
   setNotice("Data History berhasil dihapus dari Google Sheets. File Excel di Google Drive tetap tersimpan.");
  }catch(e){
   setError(e instanceof Error?e.message:"Gagal menghapus History.");
  }finally{
   setLoading(false);
  }
 }
 async function deleteAll(){
  if(!gasUrl){setError("Google Apps Script belum dikonfigurasi.");return;}
  setLoading(true);
  try{
   const response=await fetch(gasUrl,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({action:"deleteAll"})});
   const result=await response.json();
   if(!result.success)throw new Error(result.message||"Gagal menghapus seluruh History dari Google Sheets.");
   await loadHistory(false);
   setConfirm(null);
   setNotice("Seluruh History berhasil dihapus dari Google Sheets. File Excel di Google Drive tetap tersimpan.");
  }catch(e){
   setError(e instanceof Error?e.message:"Gagal menghapus seluruh History.");
  }finally{
   setLoading(false);
  }
 }
 
function share(r:SavedRecord){
  const url = r.downloadUrl || r.driveUrl || "";

  const msg = `DAILY INSPECTION AUTOMOTIVE (PDI)

Inspector: ${r.inspector}
Inspection Date: ${fmt(r.inspectionDate)}
Inspection Line: ${r.line.join(" / ")}
Shift: ${r.shift}

Download File Excel: ${url || "Link Google Drive belum tersedia"}`;

  window.open(
    `https://wa.me/?text=${encodeURIComponent(msg)}`,
    "_blank",
    "noopener,noreferrer"
  );
}

 function openLink(url:string){if(url)window.open(url,"_blank","noopener,noreferrer");else alert("Link Google Drive belum tersedia.")}
 return <main className="min-h-screen bg-slate-50">
  <header className="topbar"><div className="topbar-inner"><div className="brand"><div className="brand-icon"><ShieldCheck size={25}/></div><div><div className="brand-title">PDI INSPECTION</div><div className="brand-subtitle">Inspection Report Management System</div></div></div><div className="topbar-status"><span/> OQC</div></div></header>
  <div className="page-shell">
   <div className="page-heading"><div><div className="eyebrow">QUALITY CONTROL / OQC</div><h1>Daily PDI Inspection</h1><p>Buat laporan inspeksi dan kelola arsip Excel.</p></div><div className="tabs"><button className={tab==="form"?"tab active":"tab"} onClick={()=>setTab("form")}><Plus size={16}/> Form Inspection</button><button className={tab==="history"?"tab active":"tab"} onClick={()=>setTab("history")}><History size={16}/> History <span className="count">{records.length}</span></button></div></div>
   {!gasUrl&&<div className="alert warning">Integrasi Google belum dikonfigurasi. Tambahkan <b>NEXT_PUBLIC_GAS_WEB_APP_URL</b> pada Environment Variables Vercel setelah deploy Apps Script.</div>}
   {error&&<div className="alert danger-alert"><span style={{whiteSpace:"pre-line"}}>{error}</span><button onClick={()=>setError("")}><X size={16}/></button></div>}
   {notice&&<div className="alert success-alert"><CheckCircle2 size={17}/><span>{notice}</span><button onClick={()=>setNotice("")}><X size={16}/></button></div>}
   {tab==="form"?<section className="main-layout">
    <div className="panel form-panel">
     <div className="section-title"><ClipboardList size={20}/><h2>Inspection Information</h2></div>
     <div className="info-grid">
      <div className="field"><label>Inspector Name <b>*</b></label><input className="control" placeholder="Nama inspector" value={form.inspector} onChange={e=>update({inspector:e.target.value})}/></div>
      <div className="field"><label>Inspection Date <b>*</b></label><input type="date" className="control" value={form.inspectionDate} onChange={e=>update({inspectionDate:e.target.value})}/></div>
      <div className="field"><label>Inspection Line <b>*</b></label><div className="choice-group">{(["LH","RH"] as Line[]).map(x=><label className="choice" key={x}><input type="checkbox" checked={form.line.includes(x)} onChange={e=>update({line:e.target.checked?[...form.line,x]:form.line.filter(v=>v!==x)})}/><span>{x}</span></label>)}</div></div>
      <div className="field"><label>Inspection Lot <b>*</b></label><input type="date" className="control" value={form.lotDate} onChange={e=>update({lotDate:e.target.value})}/></div>
      <div className="field"><label>Shift <b>*</b></label><div className="choice-group">{(["1","2","3"] as const).map(x=><label className="choice" key={x}><input type="radio" name="shift" checked={form.shift===x} onChange={()=>update({shift:x})}/><span>{x}</span></label>)}</div></div>
     </div>
     <div className="section-divider"/>
     <div className="section-head"><div><h2>A. Inspection Item</h2><p>Tambahkan item sesuai kebutuhan inspeksi. Maksimal 20 item.</p></div><span className="item-count">{form.items.length} / 20 Items</span></div>
     <div className="item-list">{form.items.map((item,index)=>{
       const options=Array.from(new Set([...allowedModels,...(item.model?[item.model]:[])]));
       return <article className="item-card" key={item.id}>
        <div className="item-card-head"><div className="item-heading"><span className="item-number">{index+1}</span><h3>Inspection Item {index+1}</h3></div><button className="remove-item" disabled={form.items.length===1} onClick={()=>update({items:form.items.filter(x=>x.id!==item.id)})}><Trash2 size={15}/> Hapus</button></div>
        <div className="item-fields">
         <div className="field"><label>Model <b>*</b></label><select className="control" value={item.model} onChange={e=>updateItem(item.id,{model:e.target.value as Model})}><option value="">Pilih Model</option>{options.map(m=><option key={m} value={m}>{m}</option>)}</select></div>
         <div className="field"><label>Part Number <b>*</b></label><div className="part-input"><span>{item.model?MODEL_PREFIX[item.model]:"Prefix"}</span><input className="control mono" maxLength={6} placeholder="020NB9" value={item.suffix} onChange={e=>updateItem(item.id,{suffix:e.target.value.toUpperCase().replace(/[^A-Z0-9]/g,"")})}/></div><small>Masukkan 6 karakter: 3 kode produk + 3 kode warna.</small></div>
         <div className="field"><label>SAMPLE QTY <b>*</b></label><input type="number" min="1" step="1" className="control" placeholder="0" value={item.sampleQty} onChange={e=>{const value=e.target.value;updateItem(item.id,{sampleQty:value!==""&&Number(value)<=0?"":value})}}/></div>
         <div className="field"><label>NG <b>*</b></label><input type="number" min="0" step="1" className="control" placeholder="0" value={item.ng} onChange={e=>updateItem(item.id,{ng:e.target.value})}/></div>
         <div className="field"><label>Total QTY (Otomatis)</label><input className="control read-only" readOnly value={(Number(item.sampleQty)||0)+(Number(item.ng)||0)}/></div>
         <div className="field detail-field"><label>Detail NG</label><textarea className="control detail-area" placeholder="Tuliskan detail defect / NG..." value={item.detailNg} onChange={e=>updateItem(item.id,{detailNg:e.target.value})}/></div>
        </div>
       </article>
     })}</div>
     <button className="add-item" onClick={addItem} disabled={form.items.length>=20}><Plus size={18}/> Tambah Inspection Item</button>
     <div className="total-summary"><h3>Total QTY Summary</h3><div className="summary-metrics"><div><span>Total QTY</span><strong className="blue">{total}</strong></div><div><span>SAMPLE QTY</span><strong className="green">{totalSample}</strong></div><div><span>Total NG</span><strong className="red">{totalNg}</strong></div></div></div>
     <div className="section-divider"/>
     <div className="form-actions"><button className="btn btn-secondary" onClick={reset} disabled={busy}><RefreshCw size={16}/> Reset Form</button><button className="btn btn-primary" onClick={generate} disabled={busy}>{busy?<LoaderCircle size={16} className="spin"/>:<FileSpreadsheet size={16}/>} {busy?"Menyimpan...":"Generate & Simpan Excel"}</button></div>
    </div>
    <aside className="sidebar">
     <div className="panel side-card"><div className="section-title"><Activity size={19}/><h3>Inspection Summary</h3></div><div className="mini-metrics"><div className="metric green-bg"><span>Total Sample QTY</span><strong>{totalSample}</strong></div><div className="metric red-bg"><span>Total NG</span><strong>{totalNg}</strong></div></div><div className="metric blue-bg total-metric"><span>Total QTY</span><strong>{total}</strong></div><div className="side-foot">Jumlah Inspection Item: <b>{form.items.length} / 20</b></div></div>
     <div className="panel side-card"><h3 className="side-heading">Report Configuration</h3><div className="config-list"><div><span>Template</span><b>PDI rev.3</b></div><div><span>Prepared</span><b>{form.inspector||"Inspector Name"}</b></div><div><span>Inspection Line</span><b>{form.line.length?form.line.join(" / "):"-"}</b></div><div><span>Inspection Item</span><b>{form.items.length} / 20</b></div><div><span>Format</span><b>.xlsx</b></div><div><span>Draft / History</span><b className="green-text">Local / Google Sheets</b></div></div></div>
    </aside>
   </section>:<section className="panel history-panel">
    <div className="history-head"><div><h2>History Inspection</h2><p>{records.length} laporan tersimpan di Google Sheets.</p></div><button className="btn btn-secondary" onClick={()=>loadHistory(true)} disabled={loading}>{loading?<LoaderCircle size={15} className="spin"/>:<Activity size={15}/>} Refresh</button></div>
    <div className="filters"><div className="filter-grid">
     <div className="field"><label>Tanggal Inspeksi</label><input type="date" className="control" value={filterDate} onChange={e=>setFilterDate(e.target.value)}/></div>
     <div className="field"><label>Shift</label><select className="control" value={filterShift} onChange={e=>setFilterShift(e.target.value)}><option value="">Semua Shift</option><option value="1">Shift 1</option><option value="2">Shift 2</option><option value="3">Shift 3</option></select></div>
     <div className="field"><label>Inspection Line</label><select className="control" value={filterLine} onChange={e=>setFilterLine(e.target.value)}><option value="">Semua Line</option><option value="LH">LH</option><option value="RH">RH</option></select></div>
     <div className="field"><label>Model</label><select className="control" value={filterModel} onChange={e=>setFilterModel(e.target.value)}><option value="">Semua Model</option>{allModels.map(m=><option key={m} value={m}>{m}</option>)}</select></div>
    </div><div className="search-row"><div className="searchbox"><Search size={17}/><input className="control" placeholder="Cari inspector, tanggal, model, atau part number..." value={search} onChange={e=>setSearch(e.target.value)}/></div><button className="btn btn-secondary" onClick={resetFilters}><X size={15}/> Reset Filter</button><button className="btn btn-danger" onClick={()=>setConfirm("all")} disabled={!records.length}><Trash2 size={15}/> Hapus Semua History</button></div><p className="filter-count">Menampilkan <b>{filtered.length}</b> dari <b>{records.length}</b> data histori.</p></div>
    {loading?<div className="empty-state"><LoaderCircle className="spin"/>Memuat data...</div>:filtered.length===0?<div className="empty-state"><History size={30}/>{records.length?"Data tidak ditemukan.":"Belum ada laporan yang tersimpan."}</div>:<div className="table-wrap"><table className="history-table"><thead><tr><th>INSPECTOR / FILE</th><th>INSPECTION DATE</th><th>INSPECTION ITEMS</th><th>SHIFT / LINE</th><th>OK / NG</th><th>ACTIONS</th></tr></thead><tbody>{filtered.map(r=><tr key={r.id}><td><b>{r.inspector}</b><a className="file-link" onClick={()=>openLink(r.driveUrl||"")}>{r.fileName||"Buka file Excel"} <ExternalLink size={11}/></a><small>{r.items.length} Inspection Item</small></td><td>{fmt(r.inspectionDate)}</td><td><div className="history-items">{Array.from(new Set(r.items.map(it=>it.model).filter(Boolean))).map((model,i)=><div key={model}><b>{i+1}. {model}</b></div>)}</div></td><td>{r.shift} / {r.line.join(", ")}</td><td><span className="green-text">{r.items.reduce((s,x)=>s+(Number(x.sampleQty)||0),0)} OK</span><span className="muted-sep"> / </span><span className="red-text">{r.items.reduce((s,x)=>s+(Number(x.ng)||0),0)} NG</span></td><td><div className="action-buttons"><button title="Buka file Excel" onClick={()=>openLink(r.driveUrl||"")} disabled={!r.driveUrl}><ExternalLink size={15}/></button><button title="Download Excel" onClick={()=>openLink(r.downloadUrl||"")} disabled={!r.downloadUrl}><Download size={15}/></button><button title="Buka folder" onClick={()=>openLink(r.folderUrl||"")} disabled={!r.folderUrl}><FolderOpen size={15}/></button><button title="Bagikan WhatsApp" className="whatsapp-action" onClick={()=>share(r)}><MessageCircle size={15}/></button><button title="Hapus dari History" className="delete-action" onClick={()=>{setSelected(r);setConfirm("one")}}><Trash2 size={15}/></button></div></td></tr>)}</tbody></table></div>}
   </section>}
   <footer className="footer">PDI Inspection System · Quality Control</footer>
  </div>
  {error&&<div className="modal-backdrop"><div className="modal error-modal"><div className="error-icon"><X size={30}/></div><h2>Periksa Kembali Data</h2><p className="error-message">{error}</p><button className="btn btn-primary full" onClick={()=>setError("")}>Mengerti</button></div></div>}
  {successModal&&<div className="modal-backdrop"><div className="modal"><div className="success-icon"><CheckCircle2 size={34}/></div><h2>Proses Berhasil!</h2><p>Laporan PDI berhasil dibuat{gasUrl?" dan disimpan ke Google Drive.":"."}</p><div className="filename"><span>Nama file:</span><b>{successFilename}</b></div><p className="modal-hint">Buka menu History untuk melihat arsip laporan.</p><button className="btn btn-primary full" onClick={()=>setSuccessModal(false)}>Tutup</button></div></div>}
  {confirm&&<div className="modal-backdrop"><div className="modal confirm-modal"><div className="modal-top"><h2>{confirm==="all"?"Hapus Semua History?":"Hapus History?"}</h2><button onClick={()=>{setConfirm(null);setSelected(null)}}><X size={18}/></button></div><p>{confirm==="all"?"Seluruh catatan History akan dihapus dari Google Sheets. File Excel di Google Drive tidak akan dihapus.":"Catatan ini akan dihapus dari Google Sheets. File Excel di Google Drive tetap tersimpan."}</p><div className="modal-actions"><button className="btn btn-secondary" onClick={()=>{setConfirm(null);setSelected(null)}}>Batal</button><button className="btn btn-danger" onClick={confirm==="all"?deleteAll:deleteOne}>Hapus</button></div></div></div>}
 </main>
}