import { convertDocxToPdf } from "./docx-to-pdf.mjs";

const bridge = window.wordPdfConverter;
const status = document.getElementById("converter-status");
let started = false;
let finished = false;

async function failConversion() {
  if (finished) return;
  finished = true;
  if (status) status.textContent = "The Word document could not be converted.";
  try {
    await bridge?.fail();
  } catch {
    // The main process also enforces the deadline and will close this window.
  }
}

async function startConversion(data) {
  if (started || finished) return;
  started = true;
  try {
    const libraries = {
      docx: window.docx,
      html2canvas: window.html2canvas,
      PDFLib: window.PDFLib,
      JSZip: window.JSZip,
    };
    if (!bridge || !libraries.docx || !libraries.html2canvas || !libraries.PDFLib || !libraries.JSZip) {
      throw new Error("The converter libraries are unavailable.");
    }
    if (status) status.textContent = "Converting the Word document…";
    const output = await convertDocxToPdf(data, () => {}, libraries);
    if (!(output instanceof Uint8Array) || !output.byteLength) {
      throw new Error("The converter returned no PDF.");
    }
    await bridge.complete(output);
    finished = true;
    if (status) status.textContent = "The PDF is ready.";
  } catch (error) {
    console.error("The Word-to-PDF page conversion failed.", error);
    await failConversion();
  }
}

if (!bridge || !window.docx || !window.html2canvas || !window.PDFLib || !window.JSZip) {
  await failConversion();
} else {
  bridge.onStart(startConversion);
  bridge.ready();
}
