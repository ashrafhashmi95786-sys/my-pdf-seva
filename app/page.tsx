"use client";
import { useCallback, useMemo, useState } from "react";
import { PDFDocument, degrees } from "pdf-lib";
import mammoth from "mammoth";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import { UploadCloud, FileText, Files, Scissors, RotateCw, Image as ImageIcon, ShieldCheck, Zap, LockKeyhole, ArrowRight, CheckCircle2, X, Download, Menu, Sparkles, ChevronDown } from "lucide-react";

type ToolId = "merge" | "split" | "rotate" | "jpg" | "word" | "excel";
type PdfFile = { file: File; id: string };
const tools = [
  { id: "merge" as ToolId, name: "Merge PDF", desc: "Combine multiple PDFs into one organized file.", icon: Files, tag: "Most popular" },
  { id: "split" as ToolId, name: "Split PDF", desc: "Extract a page range into a new PDF.", icon: Scissors, tag: "Quick & easy" },
  { id: "rotate" as ToolId, name: "Rotate PDF", desc: "Fix sideways pages in just a few clicks.", icon: RotateCw, tag: "Useful" },
  { id: "jpg" as ToolId, name: "JPG to PDF", desc: "Turn your images into a shareable PDF.", icon: ImageIcon, tag: "Image tools" },
  { id: "word" as ToolId, name: "Word to PDF", desc: "Convert DOCX documents into PDF files.", icon: FileText, tag: "Document" },
  { id: "excel" as ToolId, name: "Excel to PDF", desc: "Export spreadsheet tabs as a readable PDF.", icon: Files, tag: "Spreadsheet" }
];
export default function Home() {
  const [active, setActive] = useState<ToolId>("merge");
  const [files, setFiles] = useState<PdfFile[]>([]);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");
  const [range, setRange] = useState("1");
  const [angle, setAngle] = useState("90");
  const [mobileMenu, setMobileMenu] = useState(false);
  const accept = active === "jpg" ? "image/jpeg,image/png,image/webp" : active === "word" ? ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" : active === "excel" ? ".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" : "application/pdf,.pdf";
  const title = tools.find(t => t.id === active)?.name ?? "PDF tools";
  const addFiles = useCallback((list: FileList | null) => {
    if (!list) return;
    const incoming = Array.from(list).filter(f => active === "jpg" ? ["image/jpeg","image/png","image/webp"].includes(f.type) : active === "word" ? f.name.toLowerCase().endsWith(".docx") : active === "excel" ? /\.(xlsx|xls)$/i.test(f.name) : f.type === "application/pdf" || f.name.toLowerCase().endsWith(".pdf"));
    if (!incoming.length) { setMessage(active === "jpg" ? "Please choose JPG, PNG or WebP images." : active === "word" ? "Please choose a .docx Word document." : active === "excel" ? "Please choose an .xlsx or .xls Excel workbook." : "Please choose valid PDF files."); return; }
    if (incoming.some(f => f.size > 30 * 1024 * 1024)) { setMessage("Each file must be 30 MB or smaller."); return; }
    setFiles(old => [...old, ...incoming.map(file => ({ file, id: crypto.randomUUID() }))]);
    setMessage("");
  }, [active]);
  const removeFile = (id: string) => setFiles(old => old.filter(f => f.id !== id));
  const reset = () => { setFiles([]); setMessage(""); };
  const download = (bytes: Uint8Array, name: string) => {
    const blob = new Blob([bytes as BlobPart], { type: "application/pdf" });
    const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = name; a.click(); URL.revokeObjectURL(url);
  };
  const runTool = async () => {
    if (!files.length) { setMessage("Upload at least one file to continue."); return; }
    setBusy(true); setMessage("");
    try {
      if (active === "merge") {
        if (files.length < 2) throw new Error("Add at least two PDF files to merge.");
        const out = await PDFDocument.create();
        for (const item of files) { const src = await PDFDocument.load(await item.file.arrayBuffer()); const pages = await out.copyPages(src, src.getPageIndices()); pages.forEach(p => out.addPage(p)); }
        download(await out.save(), "my-pdf-seva-merged.pdf");
      } else if (active === "split") {
        const src = await PDFDocument.load(await files[0].file.arrayBuffer());
        const total = src.getPageCount(); const parts = range.split(",").flatMap(v => { const [a,b] = v.trim().split("-").map(Number); return Array.from({length: Math.max(0,(b || a)-a+1)},(_,i)=>a+i); });
        if (!parts.length || parts.some(n => !Number.isInteger(n) || n < 1 || n > total)) throw new Error(`Enter valid page numbers from 1 to ${total}, e.g. 1-3,5.`);
        const out = await PDFDocument.create(); const copied = await out.copyPages(src, [...new Set(parts)].map(n=>n-1)); copied.forEach(p=>out.addPage(p)); download(await out.save(), "my-pdf-seva-pages.pdf");
      } else if (active === "rotate") {
        const src = await PDFDocument.load(await files[0].file.arrayBuffer()); const rotation = Number(angle);
        if (![90,180,270].includes(rotation)) throw new Error("Choose 90, 180 or 270 degrees.");
        src.getPages().forEach(p=>p.setRotation(degrees((p.getRotation().angle + rotation) % 360))); download(await src.save(), "my-pdf-seva-rotated.pdf");
      } else {
        const out = await PDFDocument.create();
        for (const item of files) { const bytes = await item.file.arrayBuffer(); const img = item.file.type === "image/png" ? await out.embedPng(bytes) : await out.embedJpg(bytes); const page = out.addPage([img.width, img.height]); page.drawImage(img,{x:0,y:0,width:img.width,height:img.height}); }
        download(await out.save(), "my-pdf-seva-images.pdf");
      }
      setMessage("Done! Your PDF download has started.");
    } catch (e) { setMessage(e instanceof Error ? e.message : "Something went wrong. Please check your file and try again."); }
    finally { setBusy(false); }
  };
  const switchTool = (id: ToolId) => { setActive(id); reset(); };
  const selectedTool = useMemo(() => tools.find(t=>t.id===active)!, [active]);
  return <main>
    <div className="topbar"><div className="container topinner"><a className="brand" href="#"><span className="brandmark"><FileText size={22}/></span><span>My PDF <b>Seva</b></span></a><nav className={mobileMenu?"nav open":"nav"}><a href="#tools" onClick={()=>setMobileMenu(false)}>All tools</a><a href="#how" onClick={()=>setMobileMenu(false)}>How it works</a><a href="#privacy" onClick={()=>setMobileMenu(false)}>Privacy</a></nav><button className="menuBtn" onClick={()=>setMobileMenu(!mobileMenu)} aria-label="Toggle menu"><Menu/></button><a className="navCta" href="#tools">Explore tools <ArrowRight size={16}/></a></div></div>
    <section className="hero"><div className="heroGlow"/><div className="container heroGrid"><div className="heroCopy"><div className="eyebrow"><Sparkles size={14}/> YOUR EVERYDAY PDF TOOLKIT</div><h1>PDF work made<br/><span>simple & seamless.</span></h1><p>Everyday PDF tools to merge, split, rotate and convert documents and images. Free, straightforward and right in your browser.</p><div className="heroActions"><a className="primaryBtn" href="#tools">Explore PDF tools <ArrowRight size={18}/></a><span className="trust"><ShieldCheck size={17}/> No signup required</span></div><div className="heroStats"><div><b>100%</b><small>Free to use</small></div><i/><div><b>Private</b><small>Browser processing</small></div><i/><div><b>Simple</b><small>Made for everyone</small></div></div></div><div className="heroVisual"><div className="floatCard floatOne"><span className="miniIcon purple"><Files size={19}/></span><div><b>Merge PDF</b><small>Combine documents</small></div><CheckCircle2 className="check"/></div><div className="documentArt"><div className="docBack"/><div className="docFront"><div className="pdfBadge">PDF</div><div className="docLine wide"/><div className="docLine"/><div className="docLine short"/><div className="docGraph"><span/><span/><span/><span/><span/></div><div className="docLine wide"/></div><div className="spark s1">✦</div><div className="spark s2">✧</div></div><div className="floatCard floatTwo"><span className="miniIcon green"><LockKeyhole size={18}/></span><div><b>Your files stay yours</b><small>Processed on your device</small></div></div></div></div></section>
    <section className="toolSection" id="tools"><div className="container"><div className="sectionHead"><div><div className="eyebrow dark"><span/> SIMPLE TOOLS, BIG HELP</div><h2>What would you like to <span>do today?</span></h2><p>Pick a tool and get your PDF task done in a few clicks.</p></div><div className="toolCount"><b>06</b><small>TOOLS TO START</small></div></div><div className="toolGrid">{tools.map(t=>{const Icon=t.icon;return <button key={t.id} className={active===t.id?"toolCard selected":"toolCard"} onClick={()=>switchTool(t.id)}><div className="toolIcon"><Icon size={23}/></div><span className="toolTag">{t.tag}</span><h3>{t.name}</h3><p>{t.desc}</p><span className="toolLink">Use tool <ArrowRight size={15}/></span></button>})}</div>
    <div className="workspace"><div className="workHead"><div><div className="eyebrow dark">PDF WORKSPACE</div><h3>{title}</h3><p>{active==="merge"?"Bring multiple PDF documents together in the order you choose.":active==="split"?"Choose which pages you want in your new PDF.":active==="rotate"?"Rotate every page in your PDF to the correct orientation.":active==="jpg"?"Convert your images into one neat PDF document.":active==="word"?"Convert a Word DOCX document into a PDF on this device.":"Convert every non-empty worksheet in an Excel workbook into PDF tables."}</p></div><span className="securePill"><ShieldCheck size={15}/> Secure & private</span></div><div className="workBody"><label className="dropzone" htmlFor="fileInput"><input id="fileInput" type="file" accept={accept} multiple={active==="merge"||active==="jpg"||active==="word"||active==="excel"} onChange={e=>{addFiles(e.target.files);e.currentTarget.value="";}}/><span className="uploadIcon"><UploadCloud size={27}/></span><b>Click to upload or drag files here</b><small>{active==="jpg"?"JPG, PNG or WebP images":active==="word"?"Word .docx document":active==="excel"?"Excel .xlsx or .xls workbook":"PDF files"} · Max 30 MB per file</small></label>
      {files.length>0&&<div className="fileList"><div className="fileListTitle"><b>{files.length} file{files.length>1?"s":""} selected</b><button onClick={reset}>Remove all</button></div>{files.map((f,i)=><div className="fileRow" key={f.id}><span className="fileIcon"><FileText size={18}/></span><span className="fileName"><b>{f.file.name}</b><small>{(f.file.size/1024/1024).toFixed(2)} MB</small></span><span className="fileOrder">{i+1}</span><button className="removeBtn" onClick={()=>removeFile(f.id)} aria-label="Remove file"><X size={17}/></button></div>)}</div>}
      {active==="split"&&<label className="optionField">Pages to extract <input value={range} onChange={e=>setRange(e.target.value)} placeholder="Example: 1-3, 5, 8-10"/></label>}
      {active==="rotate"&&<label className="optionField">Rotate pages by <select value={angle} onChange={e=>setAngle(e.target.value)}><option value="90">90° clockwise</option><option value="180">180°</option><option value="270">270° clockwise</option></select></label>}
      <div className="workActions"><button className="primaryBtn runBtn" disabled={busy||files.length===0} onClick={runTool}>{busy?<><span className="spinner"/> Processing…</>:<>Process PDF <ArrowRight size={17}/></>}</button><span className="privacyNote"><LockKeyhole size={14}/> Files are processed locally in your browser</span></div>
      {message&&<div className={message.startsWith("Done")?"successMsg":"errorMsg"} role="status">{message}</div>}
    </div></div></div></section>
    <section className="stepsSection" id="how"><div className="container"><div className="centerHead"><div className="eyebrow dark"><span/> EASY AS 1, 2, 3</div><h2>Get it done in <span>three steps.</span></h2><p>No complicated setup. No learning curve.</p></div><div className="stepsGrid"><div className="step"><div className="stepNum">01</div><UploadCloud/><h3>Choose your file</h3><p>Upload a PDF or image from your device.</p></div><div className="step"><div className="stepNum">02</div><Zap/><h3>Use your tool</h3><p>Select the action and set your options.</p></div><div className="step"><div className="stepNum">03</div><Download/><h3>Download result</h3><p>Your new file is ready to save and share.</p></div></div></div></section>
    <section className="privacySection" id="privacy"><div className="container privacyBox"><div className="privacyIcon"><ShieldCheck size={30}/></div><div><div className="eyebrow">YOUR PRIVACY MATTERS</div><h2>Your files stay on your device.</h2><p>PDF processing happens in your browser. My PDF Seva does not upload your selected files to our server for these tools. Close the page and your selected files are cleared from this workspace.</p></div><div className="privacyStamp"><LockKeyhole size={24}/><b>Private by design</b><small>No account needed</small></div></div></section>
    <footer><div className="container footerInner"><a className="brand" href="#"><span className="brandmark"><FileText size={20}/></span><span>My PDF <b>Seva</b></span></a><span>Free PDF tools, made simple for everyone.</span><span>© {new Date().getFullYear()} My PDF Seva</span></div></footer>
  </main>;
}