import { PDFDocument } from "pdf-lib";
import * as pdfjsLib from "pdfjs-dist";
import {
  Document,
  HorizontalPositionRelativeFrom,
  ImageRun,
  Packer,
  Paragraph,
  TextRun,
  TextWrappingType,
  VerticalPositionRelativeFrom,
} from "docx";
import { ShapeRun as DocxShapeRun } from "docx/shapes";
import { baseName, canvasToBlob, drawImage, loadImage, readFile } from "./utils.js";

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL("./assets/pdf.worker.min.mjs", document.baseURI).href;

const emuPerPixel = 9525;
const wordScale = 96 / 72;

function getEditableTextItems(page, viewport) {
  return page.items.flatMap((item) => {
    if (!("str" in item) || !item.str.trim()) return [];
    const style = page.styles[item.fontName];
    const [left, baseline] = viewport.convertToViewportPoint(item.transform[4], item.transform[5]);
    const fontSize = Math.hypot(item.transform[0], item.transform[1]);
    const width = Math.max(1, item.width * wordScale);
    const height = Math.max(1, (item.height || fontSize) * wordScale);
    const [nextX, nextY] = viewport.convertToViewportPoint(
      item.transform[4] + item.transform[0],
      item.transform[5] + item.transform[1],
    );
    return [{
      text: item.str,
      x: left * wordScale,
      y: baseline * wordScale - height,
      width,
      height,
      fontSize: Math.max(1, fontSize * 2),
      rotation: Math.atan2(nextY - baseline, nextX - left) * 180 / Math.PI,
      font: style?.fontFamily || "Arial",
      bold: /bold/i.test(style?.fontFamily || ""),
      italics: /italic|oblique/i.test(style?.fontFamily || ""),
    }];
  });
}

async function addImagePage(document, file) {
  const image = await loadImage(file);
  try {
    const canvas = drawImage(image, image.width || image.naturalWidth, image.height || image.naturalHeight);
    const png = await canvasToBlob(canvas, "image/png");
    const embedded = await document.embedPng(new Uint8Array(await png.arrayBuffer()));
    const pageWidth = embedded.width >= embedded.height ? 841.89 : 595.28;
    const pageHeight = embedded.width >= embedded.height ? 595.28 : 841.89;
    const margin = 28.35;
    const scale = Math.min((pageWidth - margin * 2) / embedded.width, (pageHeight - margin * 2) / embedded.height);
    const width = embedded.width * scale;
    const height = embedded.height * scale;
    const page = document.addPage([pageWidth, pageHeight]);
    page.drawImage(embedded, {
      x: (pageWidth - width) / 2,
      y: (pageHeight - height) / 2,
      width,
      height,
    });
  } finally {
    if (typeof image.close === "function") image.close();
  }
}

async function mergePdfs(files, combining, report) {
  const merged = await PDFDocument.create();
  for (let index = 0; index < files.length; index += 1) {
    const file = files[index];
    report(Math.round(index / files.length * 90), `Adding ${file.name}.`);
    if (combining && !file.name.toLowerCase().endsWith(".pdf")) {
      await addImagePage(merged, file);
    } else {
      const source = await PDFDocument.load(await readFile(file));
      const pages = await merged.copyPages(source, source.getPageIndices());
      pages.forEach((page) => merged.addPage(page));
    }
  }
  report(95, "Saving the combined document.");
  return [{
    name: combining ? "combined-files.pdf" : "merged-document.pdf",
    data: await merged.save(),
  }];
}

async function splitPdf(file, report) {
  const source = await PDFDocument.load(await readFile(file));
  const outputs = [];
  const indices = source.getPageIndices();
  for (let index = 0; index < indices.length; index += 1) {
    const document = await PDFDocument.create();
    const [page] = await document.copyPages(source, [indices[index]]);
    document.addPage(page);
    outputs.push({
      name: `${baseName(file.name)}-page-${String(index + 1).padStart(2, "0")}.pdf`,
      data: await document.save(),
    });
    report(Math.round((index + 1) / indices.length * 100), `Created page ${index + 1} of ${indices.length}.`);
  }
  return outputs;
}

async function renderPage(page) {
  const viewport = page.getViewport({ scale: 2 });
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  const context = canvas.getContext("2d", { alpha: false });
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: context, viewport }).promise;
  return { canvas, viewport };
}

async function pdfToWord(file, report) {
  const loading = pdfjsLib.getDocument({ data: await readFile(file), useSystemFonts: true, isEvalSupported: false });
  const pdf = await loading.promise;
  const sections = [];
  let foundText = false;
  try {
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      report(Math.round((pageNumber - 1) / pdf.numPages * 90), `Preparing page ${pageNumber} of ${pdf.numPages}.`);
      const page = await pdf.getPage(pageNumber);
      const content = await page.getTextContent();
      const editableText = getEditableTextItems(content, page.getViewport({ scale: 1 }));
      if (editableText.length) foundText = true;
      const { canvas, viewport } = await renderPage(page);
      const context = canvas.getContext("2d");
      for (const item of editableText) {
        const x = item.x / wordScale * canvas.width / viewport.width;
        const y = item.y / wordScale * canvas.height / viewport.height;
        const width = item.width / wordScale * canvas.width / viewport.width;
        const height = item.height / wordScale * canvas.height / viewport.height;
        context.fillStyle = "#ffffff";
        context.fillRect(Math.max(0, x - 1), Math.max(0, y - 1), width + 2, height + 2);
      }
      const pageImage = await canvasToBlob(canvas, "image/png");
      const children = [
        new ImageRun({
          type: "png",
          data: new Uint8Array(await pageImage.arrayBuffer()),
          transformation: {
            width: viewport.width * wordScale,
            height: viewport.height * wordScale,
          },
          floating: {
            horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: 0 },
            verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: 0 },
            behindDocument: true,
            allowOverlap: true,
            wrap: { type: TextWrappingType.NONE },
          },
        }),
        ...editableText.map((item) => new DocxShapeRun({
          type: "rectangle",
          transformation: {
            width: item.width,
            height: item.height,
            rotation: item.rotation,
          },
          fill: "FFFFFF",
          line: "none",
          floating: {
            horizontalPosition: {
              relative: HorizontalPositionRelativeFrom.PAGE,
              offset: Math.round(item.x * emuPerPixel),
            },
            verticalPosition: {
              relative: VerticalPositionRelativeFrom.PAGE,
              offset: Math.round(item.y * emuPerPixel),
            },
            allowOverlap: true,
            wrap: { type: TextWrappingType.NONE },
            zIndex: 2,
          },
          textOptions: {
            verticalAlignment: "top",
            margins: { top: 0, right: 0, bottom: 0, left: 0 },
            wrap: false,
            shrinkTextOnOverflow: true,
          },
          children: [new Paragraph({
            children: [new TextRun({
              text: item.text,
              font: item.font,
              size: item.fontSize,
              bold: item.bold,
              italics: item.italics,
            })],
            spacing: { before: 0, after: 0 },
          })],
        })),
      ];
      sections.push({
        properties: {
          page: {
            size: {
              width: Math.round(viewport.width * 20),
              height: Math.round(viewport.height * 20),
            },
            margin: { top: 0, right: 0, bottom: 0, left: 0 },
          },
        },
        children: [new Paragraph({ children, spacing: { before: 0, after: 0 } })],
      });
      page.cleanup();
    }
    if (!foundText) {
      throw new Error("This PDF contains no selectable text. Scanned or image-only PDFs need OCR first.");
    }
    report(95, "Creating the editable Word document.");
    const document = new Document({ sections });
    return [{ name: `${baseName(file.name)}.docx`, data: await Packer.toBlob(document) }];
  } finally {
    await pdf.destroy();
  }
}

export async function processPdfs(files, options, report = () => {}) {
  try {
    if (options.action === "merge" || options.action === "combine") {
      return await mergePdfs(files, options.action === "combine", report);
    }
    if (options.action === "split") return await splitPdf(files[0], report);
    return await pdfToWord(files[0], report);
  } catch (error) {
    if (/password|encrypted/i.test(error.message)) {
      throw new Error("This PDF is password-protected. Unlock it first, then try again.");
    }
    if (/selectable text|OCR/i.test(error.message)) throw error;
    if (options.action === "to-word") {
      throw new Error(`PDF to Word conversion failed: ${error.message}`);
    }
    throw new Error(`PDF processing failed: ${error.message}`);
  }
}
