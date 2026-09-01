import * as pdfjsLib from "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";

const filesInput = document.querySelector("#pdf-files");
const dropZone = document.querySelector(".drop-zone");
const fileList = document.querySelector("#file-list");
const emptyState = document.querySelector("#empty-state");
const clearButton = document.querySelector("#clear-files");
const generateButton = document.querySelector("#generate");
const status = document.querySelector("#status");
let invoices = [];

function parseInvoiceName(file) {
  const stem = file.name.replace(/\.pdf$/i, "").trim();
  const formats = [
    /(?<day>\d{2})[-_.](?<month>\d{2})[-_.](?<year>\d{4})/,
    /(?<year>\d{4})[-_.](?<month>\d{2})[-_.](?<day>\d{2})/,
  ];
  for (const pattern of formats) {
    const match = stem.match(pattern);
    if (!match?.groups) continue;
    const { day, month, year } = match.groups;
    const date = new Date(Number(year), Number(month) - 1, Number(day));
    if (date.getFullYear() !== Number(year) || date.getMonth() !== Number(month) - 1 || date.getDate() !== Number(day)) continue;
    const remainder = stem.slice(match.index + match[0].length).trim().replace(/^[_ -]+|\s*\(\d+\)\s*$/g, "").trim();
    return { file, date, password: remainder || undefined };
  }
  return { file, error: "No contiene una fecha válida." };
}

function setFiles(fileItems) {
  const unique = new Map([...invoices.map(({ file }) => [file.name + file.size + file.lastModified, file]), ...fileItems.map((file) => [file.name + file.size + file.lastModified, file])]);
  invoices = [...unique.values()].map(parseInvoiceName).sort((a, b) => (a.date?.valueOf() ?? Infinity) - (b.date?.valueOf() ?? Infinity) || a.file.name.localeCompare(b.file.name));
  renderFiles();
}

function renderFiles() {
  fileList.replaceChildren();
  const valid = invoices.filter((invoice) => !invoice.error);
  emptyState.hidden = invoices.length > 0;
  clearButton.disabled = invoices.length === 0;
  generateButton.disabled = valid.length === 0;
  invoices.forEach((invoice, index) => {
    const item = document.createElement("li"); item.className = "file";
    const dateText = invoice.date ? invoice.date.toLocaleDateString("es-ES") : "Fecha no reconocida";
    item.innerHTML = `<span class="file-number">${index + 1}</span><span class="file-name"></span><span class="file-date">${dateText}</span>${invoice.error ? `<span class="file-error">${invoice.error}</span>` : ""}`;
    item.querySelector(".file-name").textContent = invoice.file.name;
    fileList.append(item);
  });
}

function updateStatus(message, isError = false) { status.textContent = message; status.classList.toggle("error", isError); }
function dataUrlToBytes(dataUrl) { return Uint8Array.from(atob(dataUrl.split(",")[1]), (char) => char.charCodeAt(0)); }

async function renderInvoice(invoice) {
  const data = new Uint8Array(await invoice.file.arrayBuffer());
  const loadingTask = pdfjsLib.getDocument({ data, password: invoice.password });
  const pdf = await loadingTask.promise;
  const pages = [];
  for (let number = 1; number <= pdf.numPages; number += 1) {
    const page = await pdf.getPage(number);
    const viewport = page.getViewport({ scale: 2 });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
    const context = canvas.getContext("2d", { alpha: false });
    context.fillStyle = "#fff"; context.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: context, viewport }).promise;
    pages.push({ data: dataUrlToBytes(canvas.toDataURL("image/png")), width: viewport.width, height: viewport.height });
  }
  await pdf.destroy();
  return pages;
}

async function createWord() {
  const validInvoices = invoices.filter((invoice) => !invoice.error);
  if (!validInvoices.length) return;
  if (!window.docx) throw new Error("No se pudo cargar la biblioteca para crear Word. Comprueba tu conexión e inténtalo de nuevo.");
  generateButton.disabled = true; clearButton.disabled = true;
  const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, AlignmentType, PageBreak, PageOrientation } = window.docx;
  const children = [new Paragraph({ text: "FACTURAS ORDENADAS POR FECHA", heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER }), new Paragraph(`Facturas incluidas: ${validInvoices.length}`), new Paragraph("Orden: de la fecha más antigua a la más reciente.")];
  const errors = [];
  for (const [index, invoice] of validInvoices.entries()) {
    updateStatus(`Procesando ${index + 1} de ${validInvoices.length}: ${invoice.file.name}`);
    try {
      const pages = await renderInvoice(invoice);
      children.push(new Paragraph({ children: [new PageBreak()] }), new Paragraph({ text: `FACTURA ${index + 1}`, heading: HeadingLevel.HEADING_2, alignment: AlignmentType.CENTER }), new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `Fecha: ${invoice.date.toLocaleDateString("es-ES")}\n`, bold: true }), new TextRun(`Archivo: ${invoice.file.name}`)] }));
      for (const image of pages) {
        const maxWidth = 700; const scale = Math.min(maxWidth / image.width, 1);
        children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: image.data, transformation: { width: Math.round(image.width * scale), height: Math.round(image.height * scale) }, type: "png" })] }));
      }
    } catch (error) { errors.push(`${invoice.file.name}: ${error.message || "No se pudo abrir el PDF."}`); }
  }
  if (children.length <= 3) throw new Error("No se pudo incluir ninguna factura.");
  updateStatus("Creando el archivo Word…");
  const wordDocument = new Document({ sections: [{ properties: { page: { margin: { top: 710, right: 710, bottom: 710, left: 710 }, size: { orientation: PageOrientation.PORTRAIT } } }, children }] });
  const blob = await Packer.toBlob(wordDocument);
  const link = document.createElement("a"); link.href = URL.createObjectURL(blob); link.download = "Facturas_ordenadas.docx"; link.click(); URL.revokeObjectURL(link.href);
  updateStatus(errors.length ? `Word descargado. ${errors.length} archivo(s) no se pudieron incluir: ${errors.join(" · ")}` : "¡Word creado y descargado correctamente!");
  status.classList.toggle("error", errors.length > 0);
  generateButton.disabled = false; clearButton.disabled = false;
}

filesInput.addEventListener("change", () => setFiles([...filesInput.files]));
clearButton.addEventListener("click", () => { invoices = []; filesInput.value = ""; updateStatus(""); renderFiles(); });
generateButton.addEventListener("click", () => createWord().catch((error) => { updateStatus(error.message || "No se pudo crear el Word.", true); generateButton.disabled = false; clearButton.disabled = false; }));
["dragenter", "dragover"].forEach((event) => dropZone.addEventListener(event, (e) => { e.preventDefault(); dropZone.classList.add("dragging"); }));
["dragleave", "drop"].forEach((event) => dropZone.addEventListener(event, (e) => { e.preventDefault(); dropZone.classList.remove("dragging"); }));
dropZone.addEventListener("drop", (event) => setFiles([...event.dataTransfer.files].filter((file) => file.name.toLowerCase().endsWith(".pdf"))));
