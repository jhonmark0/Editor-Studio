const fs = require("node:fs/promises");
const fsSync = require("node:fs");
const path = require("node:path");
const { spawn } = require("node:child_process");
const sharp = require("sharp");
const { PDFDocument } = require("pdf-lib");
const ffmpegPath = require("ffmpeg-static");
const { PDFParse } = require("pdf-parse");
const {
  Document,
  HorizontalPositionRelativeFrom,
  ImageRun,
  Packer,
  Paragraph,
  TextRun,
  TextWrappingType,
  VerticalPositionRelativeFrom,
} = require("docx");
const { ShapeRun } = require("docx/shapes");

const emuPerPixel = 9525;
const allowedExtensions = {
  image: new Set([".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp", ".gif"]),
  pdf: new Set([".pdf"]),
  video: new Set([".mp4", ".mov", ".m4v", ".mkv", ".webm", ".avi"]),
};
const combinedPdfExtensions = new Set([...allowedExtensions.image, ...allowedExtensions.pdf]);

function uniqueOutputPath(outputDir, fileName) {
  const extension = path.extname(fileName);
  const base = path.basename(fileName, extension);
  let candidate = path.join(outputDir, fileName);
  let counter = 2;
  while (fsSync.existsSync(candidate)) {
    candidate = path.join(outputDir, `${base} (${counter})${extension}`);
    counter += 1;
  }
  return candidate;
}

async function validateJob(job) {
  if (!job || !["image", "pdf", "video"].includes(job.type)) {
    throw new Error("Choose a supported media specialist.");
  }
  if (!Array.isArray(job.files) || job.files.length < 1 || job.files.length > 30) {
    throw new Error("Choose between 1 and 30 input files.");
  }
  if (typeof job.outputDir !== "string" || !path.isAbsolute(job.outputDir)) {
    throw new Error("Choose a valid output folder.");
  }
  const outputStat = await fs.stat(job.outputDir).catch(() => null);
  if (!outputStat || !outputStat.isDirectory()) throw new Error("The output folder is unavailable.");

  if (job.type === "image") {
    if (job.action !== "resize" && job.files.length !== 1) {
      throw new Error("Choose exactly one image for this photo task.");
    }
    if (job.action === "resize") {
      const width = Number(job.width);
      const height = Number(job.height);
      if (!Number.isInteger(width) || width < 1 || width > 12000 ||
          !Number.isInteger(height) || height < 1 || height > 12000) {
        throw new Error("Enter a width and height between 1 and 12000 pixels.");
      }
    } else if (job.action === "transform") {
      if (![0, 90, 180, 270].includes(Number(job.rotation)) ||
          !["none", "horizontal", "vertical"].includes(job.flip) ||
          (Number(job.rotation) === 0 && job.flip === "none")) {
        throw new Error("Choose a rotation, a flip, or both for the photo.");
      }
    } else if (job.action === "passport") {
      const widthMm = Number(job.widthMm);
      const heightMm = Number(job.heightMm);
      const dpi = Number(job.dpi);
      if (!Number.isFinite(widthMm) || widthMm < 15 || widthMm > 100 ||
          !Number.isFinite(heightMm) || heightMm < 15 || heightMm > 100 ||
          ![150, 300, 600].includes(dpi)) {
        throw new Error("Use dimensions from 15 to 100 mm and choose 150, 300, or 600 DPI.");
      }
    } else if (job.action === "convert") {
      if (!["jpeg", "png", "webp", "pdf"].includes(job.format)) {
        throw new Error("Choose JPG, PNG, WebP, or PDF as the output format.");
      }
    } else if (job.action !== "grayscale") {
      throw new Error("Choose a supported photo edit.");
    }
  } else if (job.type === "pdf") {
    if (job.action === "merge" && job.files.length < 2) throw new Error("Select at least two PDFs to merge.");
    if (job.action === "split" && job.files.length !== 1) throw new Error("Choose exactly one PDF to split.");
    if (job.action === "to-word" && job.files.length !== 1) {
      throw new Error("Choose exactly one PDF to convert to Word.");
    }
    if (!["merge", "split", "combine", "to-word"].includes(job.action)) {
      throw new Error("Choose merge, split, combine, or convert PDF to Word.");
    }
  } else if (job.type === "video") {
    if (job.files.length !== 1) throw new Error("Choose exactly one video for this task.");
    if (job.action === "trim") {
      if (!Number.isFinite(Number(job.start)) || Number(job.start) < 0 ||
          !Number.isFinite(Number(job.duration)) || Number(job.duration) <= 0) {
        throw new Error("Enter a start time of zero or more and a duration greater than zero, in seconds.");
      }
    } else if (!["convert", "extract-audio"].includes(job.action)) {
      throw new Error("Choose trim, convert, or extract audio for the video task.");
    }
  }

  for (const file of job.files) {
    if (typeof file !== "string" || !path.isAbsolute(file)) {
      throw new Error("An input file path is invalid.");
    }
    const inputStat = await fs.stat(file).catch(() => null);
    if (!inputStat || !inputStat.isFile()) throw new Error(`Input file not found: ${path.basename(file || "")}`);
    const extensions = job.type === "pdf" && job.action === "combine"
      ? combinedPdfExtensions : allowedExtensions[job.type];
    if (!extensions.has(path.extname(file).toLowerCase())) {
      throw new Error(`Unsupported ${job.type} file: ${path.basename(file)}`);
    }
  }
}

async function addImagePageToPdf(document, input) {
  const imageData = await sharp(input, { animated: false }).rotate().png().toBuffer({ resolveWithObject: true });
  const image = await document.embedPng(imageData.data);
  const pageWidth = image.width >= image.height ? 841.89 : 595.28;
  const pageHeight = image.width >= image.height ? 595.28 : 841.89;
  const margin = 28.35;
  const scale = Math.min(
    (pageWidth - margin * 2) / image.width,
    (pageHeight - margin * 2) / image.height,
  );
  const width = image.width * scale;
  const height = image.height * scale;
  const page = document.addPage([pageWidth, pageHeight]);
  page.drawImage(image, {
    x: (pageWidth - width) / 2,
    y: (pageHeight - height) / 2,
    width,
    height,
  });
}

async function processImage(job, report) {
  const action = job.action;
  const outputs = [];
  let output;
  try {
    for (const input of job.files) {
      if (action === "convert" && job.format === "pdf") {
        output = uniqueOutputPath(job.outputDir, `${path.parse(input).name}-converted.pdf`);
        report(`Converting ${path.basename(input)} to PDF locally.`, "Photo editor");
        const document = await PDFDocument.create();
        await addImagePageToPdf(document, input);
        await fs.writeFile(output, await document.save());
        outputs.push(output);
        continue;
      }

      const sourceFormat = path.extname(input).slice(1).toLowerCase();
      const format = action === "passport" ? "jpeg" : action === "convert" ? job.format : sourceFormat;
      const outputFormat = ({ jpg: "jpeg", tif: "tiff", bmp: "jpeg" })[format] || format;
      const extension = outputFormat === "jpeg" ? ".jpg" : `.${outputFormat}`;
      const suffix = action === "passport" ? "passport-style"
        : action === "resize" ? "resized"
          : action === "grayscale" ? "grayscale"
            : action === "transform" ? "transformed" : "converted";
      output = uniqueOutputPath(job.outputDir, `${path.parse(input).name}-${suffix}${extension}`);
      report(`Processing ${path.basename(input)} locally with ${action}.`, "Photo editor");
      let pipeline = sharp(input, { animated: false });
      if (action === "transform" && Number(job.rotation) > 0) {
        pipeline = pipeline.rotate(Number(job.rotation));
      } else {
        pipeline = pipeline.rotate();
      }
      if (action === "resize") {
        pipeline = pipeline.resize(Number(job.width), Number(job.height), {
          fit: "inside",
          withoutEnlargement: true,
        });
      } else if (action === "passport") {
        const width = Math.round(Number(job.widthMm) / 25.4 * Number(job.dpi));
        const height = Math.round(Number(job.heightMm) / 25.4 * Number(job.dpi));
        pipeline = pipeline
          .flatten({ background: "#ffffff" })
          .resize(width, height, { fit: "cover", position: sharp.gravity.centre })
          .withMetadata({ density: Number(job.dpi) });
      } else if (action === "grayscale") {
        pipeline = pipeline.grayscale();
      } else if (action === "transform") {
        if (job.flip === "horizontal") pipeline = pipeline.flop();
        if (job.flip === "vertical") pipeline = pipeline.flip();
      }
      await pipeline.toFormat(outputFormat).toFile(output);
      outputs.push(output);
    }
    return outputs;
  } catch (error) {
    await Promise.all([...outputs, output].filter(Boolean).map((file) => fs.rm(file, { force: true })));
    throw new Error(`Photo edit failed: ${error.message}`);
  }
}

function getEditableTextItems(page, viewport) {
  const textItems = [];
  const scale = 96 / 72;

  for (const item of page.items) {
    if (!("str" in item) || !item.str.trim()) continue;

    const style = page.styles[item.fontName];
    const [left, baseline] = viewport.convertToViewportPoint(item.transform[4], item.transform[5]);
    const fontSize = Math.hypot(item.transform[0], item.transform[1]);
    const width = Math.max(1, item.width * scale);
    const height = Math.max(1, (item.height || fontSize) * scale);
    const [nextX, nextY] = viewport.convertToViewportPoint(
      item.transform[4] + item.transform[0],
      item.transform[5] + item.transform[1],
    );

    textItems.push({
      text: item.str,
      x: left * scale,
      y: baseline * scale - height,
      width,
      height,
      fontSize: Math.max(1, fontSize * 2),
      rotation: Math.atan2(nextY - baseline, nextX - left) * 180 / Math.PI,
      font: style?.fontFamily || "Arial",
      bold: /bold/i.test(style?.fontFamily || ""),
      italics: /italic|oblique/i.test(style?.fontFamily || ""),
    });
  }

  return textItems;
}

async function processPdf(job, report) {
  if (job.action === "to-word") {
    const input = job.files[0];
    const output = uniqueOutputPath(job.outputDir, `${path.parse(input).name}.docx`);
    report(`Preserving page layouts and preparing editable text from ${path.basename(input)} locally.`, "PDF editor");

    const parser = new PDFParse({ data: await fs.readFile(input) });
    try {
      let firstPageScreenshot = await parser.getScreenshot({
        partial: [1],
        scale: 3,
        imageBuffer: true,
        imageDataUrl: false,
      });
      const pdfDocument = parser.doc;
      if (!pdfDocument) throw new Error("The PDF could not be opened for layout-preserving conversion.");

      const sections = [];
      let foundText = false;
      const wordScale = 96 / 72;

      for (let pageNumber = 1; pageNumber <= pdfDocument.numPages; pageNumber += 1) {
        const page = await pdfDocument.getPage(pageNumber);
        const viewport = page.getViewport({ scale: 1 });
        const content = await page.getTextContent();
        const editableText = getEditableTextItems(content, viewport);
        if (editableText.length) foundText = true;

        const screenshotResult = pageNumber === 1
          ? firstPageScreenshot
          : await parser.getScreenshot({
            partial: [pageNumber],
            scale: 3,
            imageBuffer: true,
            imageDataUrl: false,
          });
        firstPageScreenshot = undefined;
        const screenshot = screenshotResult.pages[0];
        if (!screenshot?.data) {
          throw new Error(`Could not render page ${pageNumber} for the Word document.`);
        }
        const screenshotBuffer = Buffer.from(screenshot.data);
        const screenshotMetadata = await sharp(screenshotBuffer).metadata();
        const imageWidth = screenshotMetadata.width;
        const imageHeight = screenshotMetadata.height;
        if (!imageWidth || !imageHeight) {
          throw new Error(`Could not read the rendered dimensions of page ${pageNumber}.`);
        }

        const masks = editableText.map((item) => {
          const x = item.x / wordScale * imageWidth / viewport.width;
          const y = item.y / wordScale * imageHeight / viewport.height;
          const width = item.width / wordScale * imageWidth / viewport.width;
          const height = item.height / wordScale * imageHeight / viewport.height;
          return `<rect x="${Math.max(0, x - 1)}" y="${Math.max(0, y - 1)}" width="${width + 2}" height="${height + 2}" fill="#ffffff"/>`;
        }).join("");
        const pageImage = masks
          ? await sharp(screenshotBuffer)
            .composite([{
              input: Buffer.from(
                `<svg xmlns="http://www.w3.org/2000/svg" width="${imageWidth}" height="${imageHeight}" viewBox="0 0 ${imageWidth} ${imageHeight}">${masks}</svg>`,
              ),
            }])
            .png()
            .toBuffer()
          : screenshotBuffer;

        const children = [
          new ImageRun({
            type: "png",
            data: pageImage,
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
          ...editableText.map((item) => new ShapeRun({
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
        throw new Error("This PDF contains no selectable text. Scanned or image-only PDFs need OCR.");
      }

      const document = new Document({
        sections,
      });
      await fs.writeFile(output, await Packer.toBuffer(document));
      report(`Created ${path.basename(output)} with the original page layout and editable text.`, "PDF editor");
      return [output];
    } catch (error) {
      await fs.rm(output, { force: true });
      if (/scanned|selectable text|OCR/i.test(error.message)) throw error;
      throw new Error(`PDF to Word conversion failed: ${error.message}`);
    } finally {
      await parser.destroy();
    }
  }

  if (job.action === "merge" || job.action === "combine") {
    const merged = await PDFDocument.create();
    const combining = job.action === "combine";
    const output = uniqueOutputPath(job.outputDir, combining ? "combined-files.pdf" : "merged-document.pdf");
    try {
      for (const file of job.files) {
        if (combining && path.extname(file).toLowerCase() !== ".pdf") {
          report(`Adding ${path.basename(file)} as a PDF page.`, "PDF editor");
          await addImagePageToPdf(merged, file);
        } else {
          report(`Adding ${path.basename(file)} to the PDF.`, "PDF editor");
          const source = await PDFDocument.load(await fs.readFile(file));
          const pages = await merged.copyPages(source, source.getPageIndices());
          pages.forEach((page) => merged.addPage(page));
        }
      }
      await fs.writeFile(output, await merged.save());
      return [output];
    } catch (error) {
      await fs.rm(output, { force: true });
      throw new Error(`PDF processing failed: ${error.message}`);
    }
  }

  const input = job.files[0];
  report(`Splitting ${path.basename(input)} into pages.`, "PDF editor");
  const source = await PDFDocument.load(await fs.readFile(input));
  const outputs = [];
  try {
    for (const pageIndex of source.getPageIndices()) {
      const pageDocument = await PDFDocument.create();
      const [page] = await pageDocument.copyPages(source, [pageIndex]);
      pageDocument.addPage(page);
      const output = uniqueOutputPath(
        job.outputDir,
        `${path.parse(input).name}-page-${String(pageIndex + 1).padStart(2, "0")}.pdf`,
      );
      await fs.writeFile(output, await pageDocument.save());
      outputs.push(output);
      report(`Created page ${pageIndex + 1} of ${source.getPageCount()}.`, "PDF editor");
    }
    return outputs;
  } catch (error) {
    await Promise.all(outputs.map((output) => fs.rm(output, { force: true })));
    throw error;
  }
}

function runFfmpeg(args) {
  if (!ffmpegPath || !fsSync.existsSync(ffmpegPath)) {
    return Promise.reject(new Error("The bundled FFmpeg tool is unavailable on this system."));
  }
  return new Promise((resolve, reject) => {
    const child = spawn(ffmpegPath, args, { windowsHide: true });
    let errorOutput = "";
    child.stderr.on("data", (chunk) => {
      errorOutput = `${errorOutput}${chunk}`.slice(-4000);
    });
    child.once("error", reject);
    child.once("close", (code) => {
      if (code === 0) resolve();
      else reject(new Error(errorOutput.trim().split(/\r?\n/).slice(-3).join(" ") || `FFmpeg exited with code ${code}.`));
    });
  });
}

async function processVideo(job, report) {
  const input = job.files[0];
  const extractingAudio = job.action === "extract-audio";
  const suffix = extractingAudio ? "audio" : job.action;
  const extension = extractingAudio ? ".m4a" : ".mp4";
  const output = uniqueOutputPath(job.outputDir, `${path.parse(input).name}-${suffix}${extension}`);
  const args = ["-nostdin", "-hide_banner", "-loglevel", "error"];
  if (job.action === "trim") {
    const start = Number(job.start);
    const duration = Number(job.duration);
    args.push("-ss", String(start), "-i", input, "-t", String(duration));
    report(`Trimming ${path.basename(input)} from ${start}s for ${duration}s locally.`, "Video editor");
  } else if (extractingAudio) {
    args.push("-i", input, "-map", "0:a:0", "-vn", "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart",
      "-n", output);
    report(`Extracting audio from ${path.basename(input)} locally.`, "Video editor");
  } else {
    args.push("-i", input);
    report(`Converting ${path.basename(input)} to MP4 locally.`, "Video editor");
  }
  if (!extractingAudio) {
    args.push("-map", "0:v:0", "-map", "0:a?", "-c:v", "libx264", "-preset", "medium", "-crf", "23",
      "-c:a", "aac", "-movflags", "+faststart", "-n", output);
  }
  try {
    await runFfmpeg(args);
    return [output];
  } catch (error) {
    await fs.rm(output, { force: true });
    throw new Error(`Video processing failed: ${error.message}`);
  }
}

async function processJob(job, report = () => {}) {
  await validateJob(job);
  let outputs;
  if (job.type === "image") outputs = await processImage(job, report);
  else if (job.type === "pdf") outputs = await processPdf(job, report);
  else outputs = await processVideo(job, report);
  return outputs;
}

module.exports = { processJob, uniqueOutputPath };
