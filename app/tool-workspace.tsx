"use client";
import { useCallback, useEffect, useRef, useState } from "react";
import Link from "next/link";
import { PDFDocument, degrees, StandardFonts, rgb } from "pdf-lib";
import JSZip from "jszip";
import mammoth from "mammoth";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import { Document, Packer, Paragraph, TextRun } from "docx";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;
import { UploadCloud, FileText, Files, Scissors, RotateCw, Image as ImageIcon, ShieldCheck, ArrowRight, CheckCircle2, X, Download, LockKeyhole, ChevronUp, ChevronDown } from "lucide-react";

type ToolId = "merge" | "split" | "remove" | "rotate" | "jpg" | "word" | "excel" | "pdfword" | "pdfexcel" | "pdfjpg" | "html" | "pagenumbers" | "watermark" | "crop";
type PdfFile = { file: File; id: string };
const tools = [
  { id: "merge" as ToolId, name: "Merge PDF", desc: "Combine multiple PDFs into one organized file.", icon: Files, tag: "Most popular", slug:"merge-pdf" },
  { id: "split" as ToolId, name: "Split PDF", desc: "Extract selected pages into a new PDF file.", icon: Scissors, tag: "Quick & easy", slug:"split-pdf" },
  { id: "remove" as ToolId, name: "Remove Pages", desc: "Delete unwanted pages from a PDF.", icon: Scissors, tag: "Organize", slug:"remove-pages" },
  { id: "pdfjpg" as ToolId, name: "PDF to JPG", desc: "Convert every PDF page into a JPG image.", icon: ImageIcon, tag: "Convert", slug:"pdf-to-jpg" },
  { id: "html" as ToolId, name: "HTML to PDF", desc: "Convert an HTML file into a PDF document.", icon: FileText, tag: "Convert", slug:"html-to-pdf" },
  { id: "pagenumbers" as ToolId, name: "Add Page Numbers", desc: "Add page numbers to every page of a PDF.", icon: FileText, tag: "Edit PDF", slug:"add-page-numbers" },
  { id: "watermark" as ToolId, name: "Add Watermark", desc: "Stamp text across every page of a PDF.", icon: ShieldCheck, tag: "Edit PDF", slug:"add-watermark" },
  { id: "crop" as ToolId, name: "Crop PDF", desc: "Crop the same edge margin on every page.", icon: FileText, tag: "Edit PDF", slug:"crop-pdf" },
  { id: "rotate" as ToolId, name: "Rotate PDF", desc: "Fix sideways pages in just a few clicks.", icon: RotateCw, tag: "Useful", slug:"rotate-pdf" },
  { id: "jpg" as ToolId, name: "JPG to PDF", desc: "Turn your images into a shareable PDF.", icon: ImageIcon, tag: "Image tools", slug:"jpg-to-pdf" },
  { id: "word" as ToolId, name: "Word to PDF", desc: "Convert DOCX documents into PDF files.", icon: FileText, tag: "Document", slug:"word-to-pdf" },
  { id: "excel" as ToolId, name: "Excel to PDF", desc: "Export spreadsheet tabs as PDF tables.", icon: Files, tag: "Spreadsheet", slug:"excel-to-pdf" },
  { id: "pdfword" as ToolId, name: "PDF to Word", desc: "Extract selectable PDF text into DOCX.", icon: FileText, tag: "Convert", slug:"pdf-to-word" },
  { id: "pdfexcel" as ToolId, name: "PDF to Excel", desc: "Extract PDF text into spreadsheet rows.", icon: Files, tag: "Convert", slug:"pdf-to-excel" }
];
function MergePageGallery({ file, selectable = false, selectedValue = "", onSelectPage }: { file: File; selectable?: boolean; selectedValue?: string; onSelectPage?: (page: number) => void }) {
  const [pages, setPages] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const selectedPages = new Set(selectedValue.split(",").map(value => Number(value.trim())).filter(value => Number.isInteger(value) && value > 0));
  useEffect(() => {
    let cancelled = false;
    let loadingTask: ReturnType<typeof pdfjsLib.getDocument> | undefined;
    (async () => {
      try {
        setLoading(true);
        setError("");
        const bytes = new Uint8Array(await file.arrayBuffer());
        loadingTask = pdfjsLib.getDocument({ data: bytes });
        const pdf = await loadingTask.promise;
        const previews: string[] = [];
        for (let pageNo = 1; pageNo <= pdf.numPages; pageNo++) {
          if (cancelled) return;
          const page = await pdf.getPage(pageNo);
          const base = page.getViewport({ scale: 1 });
          const viewport = page.getViewport({ scale: Math.min(150 / base.width, 190 / base.height) });
          const canvas = document.createElement("canvas");
          const context = canvas.getContext("2d");
          if (!context) continue;
          canvas.width = Math.ceil(viewport.width);
          canvas.height = Math.ceil(viewport.height);
          await page.render({ canvasContext: context, viewport }).promise;
          previews.push(canvas.toDataURL("image/jpeg", 0.78));
        }
        if (!cancelled) setPages(previews);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : "Could not create page previews.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
      loadingTask?.destroy();
    };
  }, [file]);
  return <div className={`${selectable ? "splitPageGallery " : ""}mergePageGallery`}>
    <div className="mergePageGalleryHead"><b>Pages in this PDF</b><span>{loading ? "Preparing previews…" : pages.length ? `${pages.length} page${pages.length === 1 ? "" : "s"}` : "Preview unavailable"}</span></div>
    {error ? <p className="galleryError">Page preview could not be loaded. You can still continue with this file.</p> : null}
    {loading ? <div className="galleryLoading">Loading page thumbnails…</div> : <div className="mergePageGrid">{pages.map((src, index) => <button type="button" className={`mergePageThumb ${selectable && selectedPages.has(index + 1) ? "selectedPage" : ""}`} key={index} onClick={() => selectable && onSelectPage?.(index + 1)} aria-pressed={selectable ? selectedPages.has(index + 1) : undefined}><img src={src} alt={`Page ${index + 1}`}/><span>Page {index + 1}{selectable && selectedPages.has(index + 1) ? " ✓" : ""}</span></button>)}</div>}
  </div>;
}

function UploadedPdfPreview({ file, showAllPages = false, hideSinglePreview = false, selectablePages = false, selectedValue = "", onSelectPage }: { file: File; showAllPages?: boolean; hideSinglePreview?: boolean; selectablePages?: boolean; selectedValue?: string; onSelectPage?: (page: number) => void }) {
  const [url, setUrl] = useState("");
  const [expanded, setExpanded] = useState(false);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const isPdf = file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");
  const isImage = file.type.startsWith("image/");
  useEffect(() => {
    const nextUrl = URL.createObjectURL(file);
    setUrl(nextUrl);
    return () => URL.revokeObjectURL(nextUrl);
  }, [file]);
  useEffect(() => {
    if (!isPdf || !canvasRef.current) return;
    let cancelled = false;
    let renderTask: { cancel: () => void } | undefined;
    let loadingTask: ReturnType<typeof pdfjsLib.getDocument> | undefined;
    (async () => {
      try {
        loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) });
        const pdf = await loadingTask.promise;
        const page = await pdf.getPage(1);
        if (cancelled || !canvasRef.current) return;
        const base = page.getViewport({ scale: 1 });
        const scale = Math.min(166 / base.width, 202 / base.height);
        const viewport = page.getViewport({ scale });
        const canvas = canvasRef.current;
        const context = canvas.getContext("2d");
        if (!context) return;
        const pixelRatio = Math.min(window.devicePixelRatio || 1, 2);
        canvas.width = Math.floor(viewport.width * pixelRatio);
        canvas.height = Math.floor(viewport.height * pixelRatio);
        canvas.style.width = `${Math.floor(viewport.width)}px`;
        canvas.style.height = `${Math.floor(viewport.height)}px`;
        context.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
        renderTask = page.render({ canvasContext: context, viewport });
        await (renderTask as any).promise;
      } catch (error) {
        if (!cancelled) console.error("PDF thumbnail preview failed", error);
      }
    })();
    return () => {
      cancelled = true;
      renderTask?.cancel();
      loadingTask?.destroy();
    };
  }, [file, isPdf]);
  if (hideSinglePreview && isPdf) {
    return <div className="uploadedPreview splitOnlyPreview">
      <MergePageGallery file={file} selectable={selectablePages} selectedValue={selectedValue} onSelectPage={onSelectPage}/>
    </div>;
  }
  return <div className="uploadedPreview">
    <div className="uploadedPreviewLabel">File preview <span>Click thumbnail to enlarge</span></div>
    <button type="button" className="previewThumb" onClick={() => setExpanded(true)} aria-label={`Open large preview of ${file.name}`}>
      {isPdf ? <canvas ref={canvasRef} className="pdfThumbnailCanvas" aria-label={`First page of ${file.name}`}/> : null}
      {url && isImage ? <img src={url} alt={file.name} /> : null}
      {!isPdf && !isImage ? <span className="documentThumb"><FileText size={42}/><b>{file.name.toLowerCase().endsWith(".docx") ? "WORD DOCUMENT" : "EXCEL WORKBOOK"}</b></span> : null}
      <span className="previewThumbHint">Click to view large</span>
    </button>
    {showAllPages && isPdf ? <MergePageGallery file={file}/> : null}
    {expanded && <div className="previewModal" role="dialog" aria-modal="true" aria-label={`Preview of ${file.name}`}>
      <div className="previewModalHeader"><div><b>{file.name}</b><small>Large file preview</small></div><button type="button" onClick={() => setExpanded(false)} aria-label="Close preview"><X size={22}/></button></div>
      <div className="previewModalBody">
        {url && isPdf ? <iframe title={`Full preview of ${file.name}`} src={url}/> : null}
        {url && isImage ? <img src={url} alt={file.name}/> : null}
        {!isPdf && !isImage ? <div className="previewUnavailable"><FileText size={36}/><b>Document preview is not available yet</b><span>This file is selected and ready for conversion.</span></div> : null}
      </div>
    </div>}
  </div>;
}

export default function ToolWorkspace({ initialTool }: { initialTool: ToolId }) {
  const [active, setActive] = useState<ToolId>(initialTool);
  const [files, setFiles] = useState<PdfFile[]>([]);
  const [showWorkspace, setShowWorkspace] = useState(false);
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<{ blob: Blob; name: string; type: string; url: string } | null>(null);
  const [message, setMessage] = useState("");
  const [range, setRange] = useState("");
  const [angle, setAngle] = useState("90");
  const [pageNumberStart, setPageNumberStart] = useState("1");
  const [pageNumberPosition, setPageNumberPosition] = useState("center");
  const [watermarkText, setWatermarkText] = useState("My PDF Seva");
  const [cropMargin, setCropMargin] = useState("5");
  
  const accept = active === "jpg" ? "image/jpeg,image/png,image/webp" : active === "word" ? ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" : active === "excel" ? ".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" : active === "html" ? ".html,.htm,text/html" : "application/pdf,.pdf";
  const title = tools.find(t => t.id === active)?.name ?? "PDF tools";
  const addFiles = useCallback((list: FileList | null) => {
    if (!list) return;
    const incoming = Array.from(list).filter(f => active === "jpg" ? ["image/jpeg","image/png","image/webp"].includes(f.type) : active === "word" ? f.name.toLowerCase().endsWith(".docx") : active === "excel" ? /\.(xlsx|xls)$/i.test(f.name) : active === "html" ? /\.html?$/i.test(f.name) : f.type === "application/pdf" || f.name.toLowerCase().endsWith(".pdf"));
    if (!incoming.length) { setMessage(active === "jpg" ? "Please choose JPG, PNG or WebP images." : active === "word" ? "Please choose a .docx Word document." : active === "excel" ? "Please choose an .xlsx or .xls Excel workbook." : active === "html" ? "Please choose an .html or .htm file." : "Please choose a valid PDF file."); return; }
    if (incoming.some(f => f.size > 30 * 1024 * 1024)) { setMessage("Each file must be 30 MB or smaller."); return; }
    setFiles(old => [...old, ...incoming.map(file => ({ file, id: crypto.randomUUID() }))]);
    setShowWorkspace(true);
    setMessage("");
  }, [active]);
  const removeFile = (id: string) => setFiles(old => { const next = old.filter(f => f.id !== id); if (!next.length) setShowWorkspace(false); return next; });
  const moveFile = (index: number, direction: -1 | 1) => setFiles(old => { const next = [...old]; const target = index + direction; if (target < 0 || target >= next.length) return old; [next[index], next[target]] = [next[target], next[index]]; return next; });
  const reset = () => { setFiles([]); setShowWorkspace(false); setMessage(""); setResult(null); };
  const saveResult = (blob: Blob, name: string, type: string) => { if (result?.url) URL.revokeObjectURL(result.url); setResult({ blob, name, type, url: URL.createObjectURL(blob) }); };
  const download = (bytes: Uint8Array, name: string) => {
    const blob = new Blob([bytes as BlobPart], { type: "application/pdf" });
    saveResult(blob, name, "application/pdf");
  };
  const runTool = async () => {
    if (!files.length) { setMessage("Upload at least one file to continue."); return; }
    setBusy(true); setMessage("");
    try {
      if (active === "html") {
        if (files.length !== 1) throw new Error("Please upload one HTML file at a time.");
        const source = await files[0].file.text();
        const parsed = new DOMParser().parseFromString(source, "text/html");
        parsed.querySelectorAll("script,iframe,object,embed,form,base,link[rel=stylesheet]").forEach(node => node.remove());
        parsed.querySelectorAll("style").forEach(style => { style.textContent = (style.textContent || "").replace(/@import[^;]+;?/gi, "").replace(/url\\((?![\\\x27\\\x22]?data:)[^)]+\\)/gi, "none"); });
        parsed.querySelectorAll("[style]").forEach(node => { const el = node as HTMLElement; el.style.backgroundImage = "none"; });
        parsed.querySelectorAll("img").forEach(node => { const image = node as HTMLImageElement; if (image.src && !image.src.startsWith("data:")) image.removeAttribute("src"); });
        parsed.querySelectorAll("*").forEach(node => Array.from(node.attributes).forEach(attribute => {
          const name = attribute.name.toLowerCase();
          if (name.startsWith("on") || ((name === "href" || name === "src" || name === "xlink:href") && /^\\s*javascript:/i.test(attribute.value))) node.removeAttribute(attribute.name);
        }));
        const holder = document.createElement("div");
        holder.style.cssText = "position:fixed;left:-12000px;top:0;width:794px;padding:36px;box-sizing:border-box;background:#fff;color:#111;font-family:Arial,sans-serif;font-size:14px;line-height:1.5;overflow:visible;";
        parsed.head.querySelectorAll("style").forEach(style => holder.appendChild(style.cloneNode(true)));
        const content = document.createElement("div");
        content.innerHTML = parsed.body.innerHTML || parsed.documentElement.innerHTML;
        content.querySelectorAll("img").forEach(img => { (img as HTMLImageElement).style.maxWidth = "100%"; });
        holder.appendChild(content);
        document.body.appendChild(holder);
        try {
          const canvas = await html2canvas(holder, { scale: 1.35, backgroundColor: "#ffffff", useCORS: true, allowTaint: false, logging: false, windowWidth: 794 });
          if (canvas.height > 30000) throw new Error("This HTML file is too long to convert safely in one run. Please split the content into smaller files.");
          const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
          const pageHeightPx = Math.floor(canvas.width * 297 / 210);
          let offset = 0; let page = 0;
          while (offset < canvas.height) {
            const slice = document.createElement("canvas");
            slice.width = canvas.width; slice.height = Math.min(pageHeightPx, canvas.height - offset);
            const ctx = slice.getContext("2d");
            if (!ctx) throw new Error("Could not prepare an HTML page for the PDF.");
            ctx.fillStyle = "#ffffff"; ctx.fillRect(0, 0, slice.width, slice.height);
            ctx.drawImage(canvas, 0, offset, canvas.width, slice.height, 0, 0, canvas.width, slice.height);
            if (page > 0) pdf.addPage();
            pdf.addImage(slice.toDataURL("image/jpeg", 0.9), "JPEG", 0, 0, 210, slice.height * 210 / canvas.width);
            offset += pageHeightPx; page += 1;
            slice.width = 0; slice.height = 0;
          }
          saveResult(pdf.output("blob"), files[0].file.name.replace(/\\.html?$/i, "") + ".pdf", "application/pdf");
        } finally { document.body.removeChild(holder); }
      } else if (active === "pdfjpg") {
        if (files.length !== 1) throw new Error("Please upload one PDF file at a time.");
        const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(await files[0].file.arrayBuffer()) });
        const pdf = await loadingTask.promise;
        if (pdf.numPages > 80) throw new Error("For browser stability, PDF to JPG currently supports up to 80 pages per file.");
        const zip = new JSZip();
        for (let pageNo = 1; pageNo <= pdf.numPages; pageNo++) {
          const page = await pdf.getPage(pageNo);
          const base = page.getViewport({ scale: 1 });
          const scale = Math.min(1.5, 1800 / Math.max(base.width, base.height));
          const viewport = page.getViewport({ scale });
          const canvas = document.createElement("canvas");
          const context = canvas.getContext("2d");
          if (!context) throw new Error("Could not prepare a page image.");
          canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
          await page.render({ canvasContext: context, viewport }).promise;
          const imageBlob = await new Promise<Blob>((resolve, reject) => canvas.toBlob(blob => blob ? resolve(blob) : reject(new Error("Could not encode page " + pageNo + " as JPG.")), "image/jpeg", 0.9));
          zip.file(`page-${String(pageNo).padStart(3, "0")}.jpg`, imageBlob);
          page.cleanup(); canvas.width = 0; canvas.height = 0;
        }
        const archive = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 3 } });
        const baseName = files[0].file.name.replace(/\\.pdf$/i, "");
        saveResult(archive, `${baseName}-jpg-pages.zip`, "application/zip");
      } else if (active === "pagenumbers") {
        if (files.length !== 1) throw new Error("Please upload one PDF file at a time.");
        const start = Number(pageNumberStart);
        if (!Number.isInteger(start) || start < 1 || start > 999999) throw new Error("Enter a starting page number between 1 and 999999.");
        const pdf = await PDFDocument.load(await files[0].file.arrayBuffer());
        const font = await pdf.embedFont(StandardFonts.Helvetica);
        pdf.getPages().forEach((page, index) => {
          const label = String(start + index); const size = 11; const textWidth = font.widthOfTextAtSize(label, size);
          const margin = 24; const x = pageNumberPosition === "left" ? margin : pageNumberPosition === "right" ? Math.max(margin, page.getWidth() - margin - textWidth) : Math.max(margin, (page.getWidth() - textWidth) / 2);
          page.drawText(label, { x, y: 18, size, font, color: rgb(0.32, 0.38, 0.46) });
        });
        download(await pdf.save(), files[0].file.name.replace(/\\.pdf$/i, "") + "-numbered.pdf");
      } else if (active === "watermark") {
        if (files.length !== 1) throw new Error("Please upload one PDF file at a time.");
        const text = watermarkText.trim();
        if (!text || text.length > 80) throw new Error("Enter watermark text between 1 and 80 characters.");
        if (!/^[\\x20-\\x7E]+$/.test(text)) throw new Error("For compatibility, use English letters, numbers and common keyboard symbols in the watermark.");
        const pdf = await PDFDocument.load(await files[0].file.arrayBuffer());
        const font = await pdf.embedFont(StandardFonts.Helvetica);
        pdf.getPages().forEach(page => {
          const size = Math.min(48, Math.max(20, page.getWidth() / 12));
          const textWidth = font.widthOfTextAtSize(text, size);
          page.drawText(text, { x: Math.max(18, (page.getWidth() - textWidth) / 2), y: page.getHeight() / 2, size, font, color: rgb(0.48, 0.55, 0.64), opacity: 0.24, rotate: degrees(35) });
        });
        download(await pdf.save(), files[0].file.name.replace(/\\.pdf$/i, "") + "-watermarked.pdf");
      } else if (active === "crop") {
        if (files.length !== 1) throw new Error("Please upload one PDF file at a time.");
        const marginPercent = Number(cropMargin);
        if (!Number.isFinite(marginPercent) || marginPercent < 0 || marginPercent > 30) throw new Error("Choose a crop margin between 0 and 30 percent.");
        const pdf = await PDFDocument.load(await files[0].file.arrayBuffer());
        pdf.getPages().forEach((page, index) => {
          const width = page.getWidth(); const height = page.getHeight();
          const marginX = width * marginPercent / 100; const marginY = height * marginPercent / 100;
          if (width - marginX * 2 < 72 || height - marginY * 2 < 72) throw new Error(`Crop margin is too large for page ${index + 1}.`);
          page.setCropBox(marginX, marginY, width - marginX * 2, height - marginY * 2);
        });
        download(await pdf.save(), files[0].file.name.replace(/\\.pdf$/i, "") + "-cropped.pdf");
      } else if (active === "word") {
        const item = files[0];
        const converted = await mammoth.convertToHtml({ arrayBuffer: await item.file.arrayBuffer() });
        if (!converted.value.trim()) throw new Error("No readable document text was found in this DOCX file.");
        const holder = document.createElement("div");
        holder.style.cssText = "position:fixed;left:-12000px;top:0;width:794px;padding:48px;background:#fff;color:#111;font-family:Arial,sans-serif;font-size:14px;line-height:1.5;overflow:visible;";
        holder.innerHTML = converted.value;
        holder.querySelectorAll("img").forEach(img => { (img as HTMLImageElement).style.maxWidth = "100%"; });
        document.body.appendChild(holder);
        try {
          const canvas = await html2canvas(holder, { scale: 1.5, backgroundColor: "#ffffff", useCORS: true, logging: false });
          const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
          const pageHeightPx = Math.floor(canvas.width * 297 / 210);
          let offset = 0; let page = 0;
          while (offset < canvas.height) {
            const slice = document.createElement("canvas");
            slice.width = canvas.width; slice.height = Math.min(pageHeightPx, canvas.height - offset);
            const ctx = slice.getContext("2d");
            if (!ctx) throw new Error("Could not prepare the PDF page.");
            ctx.fillStyle = "#ffffff"; ctx.fillRect(0, 0, slice.width, slice.height);
            ctx.drawImage(canvas, 0, offset, canvas.width, slice.height, 0, 0, canvas.width, slice.height);
            if (page > 0) pdf.addPage();
            pdf.addImage(slice.toDataURL("image/jpeg", 0.92), "JPEG", 0, 0, 210, slice.height * 210 / canvas.width);
            offset += pageHeightPx; page += 1;
          }
          saveResult(pdf.output("blob"), item.file.name.replace(/\.docx$/i, "") + ".pdf", "application/pdf");
        } finally { document.body.removeChild(holder); }
      } else if (active === "excel") {
        const item = files[0];
        const workbook = XLSX.read(await item.file.arrayBuffer(), { type: "array", cellDates: true });
        if (!workbook.SheetNames.length) throw new Error("No worksheets were found in this workbook.");
        const pdf = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
        let renderedSheets = 0;
        for (const sheetName of workbook.SheetNames) {
          const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, defval: "" }) as unknown[][];
          const nonEmpty = rows.filter(row => row.some(cell => String(cell ?? "").trim() !== ""));
          if (!nonEmpty.length) continue;
          if (renderedSheets > 0) pdf.addPage();
          pdf.setFontSize(13); pdf.text(sheetName.slice(0, 80), 14, 14);
          const head = [nonEmpty[0].map(cell => String(cell ?? ""))];
          const body = nonEmpty.slice(1).map(row => row.map(cell => String(cell ?? "")));
          autoTable(pdf, { head, body, startY: 20, theme: "grid", styles: { fontSize: 7, cellPadding: 2, overflow: "linebreak" }, headStyles: { fillColor: [37, 99, 235] }, margin: { left: 10, right: 10 }, horizontalPageBreak: true });
          renderedSheets += 1;
        }
        if (!renderedSheets) throw new Error("The workbook has no non-empty worksheet data.");
        saveResult(pdf.output("blob"), item.file.name.replace(/\.(xlsx|xls)$/i, "") + ".pdf", "application/pdf");
      } else if (active === "pdfword" || active === "pdfexcel") {
        if (files.length !== 1) throw new Error("Please upload one PDF file at a time.");
        const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(await files[0].file.arrayBuffer()) });
        const pdf = await loadingTask.promise;
        const extracted: { page: number; y: number; x: number; text: string }[] = [];
        for (let pageNo = 1; pageNo <= pdf.numPages; pageNo++) {
          const page = await pdf.getPage(pageNo);
          const content = await page.getTextContent();
          for (const item of content.items as any[]) {
            if (!("str" in item) || !item.str.trim()) continue;
            const transform = item.transform as number[];
            extracted.push({ page: pageNo, x: transform[4], y: transform[5], text: item.str.trim() });
          }
        }
        if (!extracted.length) throw new Error("No selectable text found. This may be a scanned PDF; OCR support is not available yet.");
        const pageLines: { page: number; y: number; chunks: { x: number; text: string }[] }[] = [];
        for (const item of extracted) {
          let line = pageLines.find(row => row.page === item.page && Math.abs(row.y - item.y) < 3);
          if (!line) { line = { page: item.page, y: item.y, chunks: [] }; pageLines.push(line); }
          line.chunks.push({ x: item.x, text: item.text });
        }
        pageLines.sort((a,b) => a.page-b.page || b.y-a.y);
        const lines = pageLines.map(line => ({ page: line.page, chunks: line.chunks.sort((a,b)=>a.x-b.x), text: line.chunks.sort((a,b)=>a.x-b.x).map(c=>c.text).join(" ").replace(/\\s+/g," ").trim() })).filter(line=>line.text);
        if (active === "pdfword") {
          const children: Paragraph[] = [];
          let lastPage = 1;
          for (const line of lines) {
            if (line.page !== lastPage) { children.push(new Paragraph({ children: [new TextRun({ text: "", break: 1 })] })); lastPage = line.page; }
            children.push(new Paragraph({ children: [new TextRun({ text: line.text })], spacing: { after: 100 } }));
          }
          const doc = new Document({ sections: [{ children }] });
          const blob = await Packer.toBlob(doc);
          saveResult(blob, files[0].file.name.replace(/\\.pdf$/i,"") + ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        } else {
          const rows = lines.map(line => {
            const chunks = line.chunks;
            if (chunks.length < 2) return [line.text];
            const cells: string[] = [];
            let current = chunks[0].text;
            let previousX = chunks[0].x;
            for (let i=1;i<chunks.length;i++) {
              const gap = chunks[i].x - previousX;
              if (gap > 28) { cells.push(current); current = chunks[i].text; }
              else current += " " + chunks[i].text;
              previousX = chunks[i].x;
            }
            cells.push(current);
            return cells;
          });
          const sheet = XLSX.utils.aoa_to_sheet(rows);
          const book = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(book, sheet, "PDF Data");
          const output = XLSX.write(book, { bookType: "xlsx", type: "array" });
          const blob = new Blob([output], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
          saveResult(blob, files[0].file.name.replace(/\\.pdf$/i,"") + ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }
      } else if (active === "merge") {
        if (files.length < 2) throw new Error("Add at least two PDF files to merge.");
        const out = await PDFDocument.create();
        for (const item of files) { const src = await PDFDocument.load(await item.file.arrayBuffer()); const pages = await out.copyPages(src, src.getPageIndices()); pages.forEach(p => out.addPage(p)); }
        download(await out.save(), "my-pdf-seva-merged.pdf");
      } else if (active === "split") {
        const src = await PDFDocument.load(await files[0].file.arrayBuffer());
        const total = src.getPageCount(); const parts = range.split(",").flatMap(v => { const [a,b] = v.trim().split("-").map(Number); return Array.from({length: Math.max(0,(b || a)-a+1)},(_,i)=>a+i); });
        if (!parts.length || parts.some(n => !Number.isInteger(n) || n < 1 || n > total)) throw new Error(`Enter valid page numbers from 1 to ${total}, e.g. 1-3,5.`);
        const out = await PDFDocument.create(); const copied = await out.copyPages(src, [...new Set(parts)].map(n=>n-1)); copied.forEach(p=>out.addPage(p)); download(await out.save(), "my-pdf-seva-pages.pdf");
      } else if (active === "remove") {
        const src = await PDFDocument.load(await files[0].file.arrayBuffer());
        const total = src.getPageCount();
        const parts = range.split(",").flatMap(v => { const value = v.trim(); if (!value) return []; const [a,b] = value.split("-").map(Number); if (!Number.isInteger(a) || (value.includes("-") && !Number.isInteger(b))) return [NaN]; return Array.from({length: Math.max(0,(b || a)-a+1)},(_,i)=>a+i); });
        const unique = [...new Set(parts)];
        if (!unique.length || unique.some(n => !Number.isInteger(n) || n < 1 || n > total)) throw new Error(`Select valid page numbers to remove from 1 to ${total}.`);
        if (unique.length >= total) throw new Error("You cannot remove every page. Keep at least one page in the PDF.");
        const keep = src.getPageIndices().filter(index => !unique.includes(index + 1));
        const out = await PDFDocument.create();
        const copied = await out.copyPages(src, keep);
        copied.forEach(page => out.addPage(page));
        download(await out.save(), "my-pdf-seva-pages-removed.pdf");
      } else if (active === "rotate") {
        const src = await PDFDocument.load(await files[0].file.arrayBuffer()); const rotation = Number(angle);
        if (![90,180,270].includes(rotation)) throw new Error("Choose 90, 180 or 270 degrees.");
        src.getPages().forEach(p=>p.setRotation(degrees((p.getRotation().angle + rotation) % 360))); download(await src.save(), "my-pdf-seva-rotated.pdf");
      } else {
        const out = await PDFDocument.create();
        for (const item of files) { const bytes = await item.file.arrayBuffer(); const img = item.file.type === "image/png" ? await out.embedPng(bytes) : await out.embedJpg(bytes); const page = out.addPage([img.width, img.height]); page.drawImage(img,{x:0,y:0,width:img.width,height:img.height}); }
        download(await out.save(), "my-pdf-seva-images.pdf");
      }
      setMessage("Processing complete. Your result is ready to preview and download.");
    } catch (e) { setMessage(e instanceof Error ? e.message : "Something went wrong. Please check your file and try again."); }
    finally { setBusy(false); }
  };

  const tool = tools.find(t=>t.id===active)!;
  const description = active==="merge" ? "Bring multiple PDF documents together in the order you choose." : active==="split" ? "Choose which pages you want in your new PDF." : active==="remove" ? "Select the pages you want to delete. The remaining pages will be saved as a new PDF." : active==="pdfjpg" ? "Convert each PDF page to a JPG image and download all pages in one ZIP file." : active==="html" ? "Convert a local HTML file into a PDF. External scripts and embedded frames are removed for safety." : active==="pagenumbers" ? "Add sequential page numbers at the bottom of every page." : active==="watermark" ? "Add a light text watermark across every page of your PDF." : active==="crop" ? "Crop the same percentage from all four edges of every page." : active==="rotate" ? "Rotate every page in your PDF to the correct orientation." : active==="jpg" ? "Convert your images into one neat PDF document." : active==="word" ? "Convert a Word DOCX document into a PDF on this device." : active==="excel" ? "Convert every non-empty worksheet in an Excel workbook into PDF tables." : active==="pdfword" ? "Convert selectable text in a PDF into a Word document." : "Extract selectable PDF text into an Excel workbook.";
  return <main className="toolPage">
    <div className="topbar"><div className="container topinner"><Link className="brand" href="/"><span className="brandmark"><FileText size={22}/></span><span>My PDF <b>Seva</b></span></Link><Link className="navCta" href="/#tools">All tools <ArrowRight size={16}/></Link></div></div>
    <section className="toolPageSection"><div className="container">
      <div className="toolPageIntro"><Link href="/#tools" className="backTools">← All PDF tools</Link><div className="eyebrow dark">MY PDF SEVA · ONLINE TOOL</div><h1>{tool.name}</h1><p>{description}</p></div>
          {!showWorkspace ? <div className="workspace fileSelectStart">
            <div className="workHead"><div><div className="eyebrow dark">START HERE</div><h3>{title}</h3><p>Select the file you want to work with. Your file will open in its workspace preview.</p></div><span className="securePill"><ShieldCheck size={15}/> Secure & private</span></div>
            <div className="workBody selectOnlyBody">
              <div className="selectDropArea" onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();addFiles(e.dataTransfer.files);}}><label className="selectFileButton" htmlFor="fileInput"><input id="fileInput" type="file" accept={accept} multiple={active==="merge"||active==="jpg"} onChange={e=>{addFiles(e.target.files);e.currentTarget.value="";}}/><UploadCloud size={19}/><b>Select {active==="jpg"?"Images":active==="word"?"Word File":active==="excel"?"Excel File":active==="merge"?"PDF Files":active==="html"?"HTML File":"PDF File"}</b></label><div className="dropPrompt">or drop {active==="jpg"?"images":active==="word"?"Word files":active==="excel"?"Excel files":active==="html"?"HTML files":"PDFs"} here</div><small className="selectFileHint">{active==="jpg"?"JPG, PNG or WebP images":active==="word"?"Word .docx document":active==="excel"?"Excel .xlsx or .xls workbook":active==="html"?"HTML .html or .htm file":"PDF file"} · Max 30 MB per file</small></div>
              {message&&<div className="errorMsg" role="status">{message}</div>}
              <div className="privacyNote selectPrivacy"><LockKeyhole size={14}/> Files are processed locally in your browser</div>
            </div>
          </div> : <div className="workspace"><div className="workHead"><div><div className="eyebrow dark">YOUR WORKSPACE</div><h3>{title}</h3><p>{description}</p></div><span className="securePill"><ShieldCheck size={15}/> Secure & private</span></div><div className="workBody">
            <div className="fileSelectToolbar"><span>Selected file preview and processing</span><label className="secondarySelectBtn" htmlFor="fileInput">+ Add file<input id="fileInput" type="file" accept={accept} multiple={active==="merge"||active==="jpg"} onChange={e=>{addFiles(e.target.files);e.currentTarget.value="";}}/></label></div>
      {files.length>0&&<div className="fileList"><div className="fileListTitle"><b>Selected for processing · {files.length} file{files.length>1?"s":""}</b><button onClick={reset}>Remove all</button></div>{files.map((f,i)=><div className="fileItem" key={f.id}><div className="fileRow"><span className="fileIcon"><FileText size={18}/></span><span className="fileName"><b>{f.file.name}</b><small>{(f.file.size/1024/1024).toFixed(2)} MB · {f.file.type||"Document"}</small></span><span className="fileOrder">{i+1}</span><div className="fileOrderActions"><button className="moveFileBtn" onClick={()=>moveFile(i,-1)} disabled={i===0} aria-label="Move file up"><ChevronUp size={16}/></button><button className="moveFileBtn" onClick={()=>moveFile(i,1)} disabled={i===files.length-1} aria-label="Move file down"><ChevronDown size={16}/></button></div><button className="removeBtn" onClick={()=>removeFile(f.id)} aria-label="Remove file"><X size={17}/></button></div><UploadedPdfPreview file={f.file} showAllPages={!["jpg","word","excel"].includes(active)} hideSinglePreview={!["jpg","word","excel"].includes(active)} selectablePages={active==="split"||active==="remove"} selectedValue={range} onSelectPage={(page)=>setRange(old=>{const nums=new Set(old.split(",").map(v=>Number(v.trim())).filter(n=>Number.isInteger(n)&&n>0));if(nums.has(page))nums.delete(page);else nums.add(page);return [...nums].sort((a,b)=>a-b).join(",");})}/></div>)}</div>}
      {(active==="split"||active==="remove")&&<label className="optionField">{active==="split"?"Pages to extract":"Pages to remove"} <input value={range} onChange={e=>setRange(e.target.value)} placeholder="Example: 1-3, 5, 8-10"/></label>}
      {active==="rotate"&&<label className="optionField">Rotate pages by <select value={angle} onChange={e=>setAngle(e.target.value)}><option value="90">90° clockwise</option><option value="180">180°</option><option value="270">270° clockwise</option></select></label>}
      {active==="pagenumbers"&&<div className="optionField">Page numbering <div className="optionControls"><label>Start from <input type="number" min="1" max="999999" value={pageNumberStart} onChange={e=>setPageNumberStart(e.target.value)}/></label><label>Position <select value={pageNumberPosition} onChange={e=>setPageNumberPosition(e.target.value)}><option value="left">Bottom left</option><option value="center">Bottom center</option><option value="right">Bottom right</option></select></label></div></div>}
      {active==="watermark"&&<label className="optionField">Watermark text <input value={watermarkText} maxLength={80} onChange={e=>setWatermarkText(e.target.value)} placeholder="Enter watermark text"/></label>}
      {active==="crop"&&<label className="optionField">Crop edge margin (%) <input type="number" min="0" max="30" step="1" value={cropMargin} onChange={e=>setCropMargin(e.target.value)}/><small>Same margin is applied to all four edges of every page. Cropping hides the outer area; it does not permanently erase it.</small></label>}
      <div className="workActions"><button className="primaryBtn runBtn" disabled={busy||files.length===0} onClick={runTool}>{busy?<><span className="spinner"/> Processing…</>:<>Process PDF <ArrowRight size={17}/></>}</button><span className="privacyNote"><LockKeyhole size={14}/> Files are processed locally in your browser</span></div>
      {busy&&<div className="processingStatus" role="status"><span className="spinner blueSpinner"/><div><b>Processing your selected file{files.length>1?"s":""}…</b><small>Your files are being processed in this browser. They are not being uploaded to a server.</small></div></div>}
      {message&&<div className={result?"successMsg":"errorMsg"} role="status">{message}</div>}
      {result&&<div className="resultPanel"><div className="resultTop"><span className="resultIcon"><CheckCircle2 size={20}/></span><div className="resultDetails"><b>Your result is ready</b><small>{result.name} · {(result.blob.size/1024).toFixed(1)} KB</small></div><button className="downloadResult" onClick={()=>{const a=document.createElement("a");a.href=result.url;a.download=result.name;a.click();}}><Download size={16}/> Download</button></div>{result.type==="application/pdf"?<div className="previewWrap"><div className="previewHeading"><FileText size={15}/> {active==="merge"?"Final merged PDF preview":"Generated PDF preview"}</div><iframe title="Generated PDF preview" src={result.url} className="pdfPreview"/></div>:<div className="nonPdfPreview"><FileText size={20}/><span>{result.type==="application/zip"?"Your JPG pages are packed into a ZIP archive. Download and extract it to view every page.":"Preview is not available for this file type. Download the generated file to open it."}</span></div>}</div>}
</div></div>}
    </div></section>
    <footer><div className="container footerInner"><Link className="brand" href="/"><span className="brandmark"><FileText size={20}/></span><span>My PDF <b>Seva</b></span></Link><span>Free PDF tools, made simple for everyone.</span><span>© {new Date().getFullYear()} My PDF Seva</span></div></footer>
  </main>;
}
