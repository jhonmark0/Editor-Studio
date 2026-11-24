import { PDFDocument } from "pdf-lib";
import { baseName, canvasToBlob, drawImage, loadImage } from "./utils.js";

const outputTypes = {
  jpeg: { mime: "image/jpeg", extension: ".jpg" },
  png: { mime: "image/png", extension: ".png" },
  webp: { mime: "image/webp", extension: ".webp" },
};

function createCanvas(width, height, background = null) {
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(width));
  canvas.height = Math.max(1, Math.round(height));
  const context = canvas.getContext("2d", { alpha: true });
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";
  if (background) {
    context.fillStyle = background;
    context.fillRect(0, 0, canvas.width, canvas.height);
  }
  return canvas;
}

function sourceFormat(file) {
  const extension = file.name.split(".").pop().toLowerCase();
  if (["jpg", "jpeg"].includes(extension)) return "jpeg";
  if (["png", "webp"].includes(extension)) return extension;
  return "png";
}

function transformedCanvas(image, width, height, rotation, flip) {
  const canvas = createCanvas(width, height);
  const context = canvas.getContext("2d");
  context.translate(canvas.width / 2, canvas.height / 2);
  context.rotate(rotation * Math.PI / 180);
  context.scale(flip === "horizontal" ? -1 : 1, flip === "vertical" ? -1 : 1);
  context.drawImage(image, -image.width / 2, -image.height / 2);
  return canvas;
}

function passportCanvas(image, width, height) {
  const canvas = createCanvas(width, height, "#ffffff");
  const context = canvas.getContext("2d");
  const scale = Math.max(canvas.width / image.width, canvas.height / image.height);
  const drawWidth = image.width * scale;
  const drawHeight = image.height * scale;
  context.drawImage(image, (canvas.width - drawWidth) / 2, (canvas.height - drawHeight) / 2, drawWidth, drawHeight);
  return canvas;
}

async function imageToPdf(canvas, fileName) {
  const png = await canvasToBlob(canvas, "image/png");
  const document = await PDFDocument.create();
  const image = await document.embedPng(new Uint8Array(await png.arrayBuffer()));
  const pageWidth = image.width >= image.height ? 841.89 : 595.28;
  const pageHeight = image.width >= image.height ? 595.28 : 841.89;
  const margin = 28.35;
  const scale = Math.min((pageWidth - margin * 2) / image.width, (pageHeight - margin * 2) / image.height);
  const width = image.width * scale;
  const height = image.height * scale;
  const page = document.addPage([pageWidth, pageHeight]);
  page.drawImage(image, {
    x: (pageWidth - width) / 2,
    y: (pageHeight - height) / 2,
    width,
    height,
  });
  return { name: `${baseName(fileName)}-converted.pdf`, data: await document.save() };
}

export async function processPhotos(files, options, report = () => {}) {
  const outputs = [];
  for (let index = 0; index < files.length; index += 1) {
    const file = files[index];
    report(Math.round(index / files.length * 100), `Processing ${file.name} locally.`);
    const image = await loadImage(file);
    try {
      const width = image.width || image.naturalWidth;
      const height = image.height || image.naturalHeight;
      let canvas;
      if (options.action === "resize") {
        const targetWidth = Number(options.width);
        const targetHeight = Number(options.height);
        const scale = Math.min(1, targetWidth / width, targetHeight / height);
        canvas = drawImage(image, width * scale, height * scale);
      } else if (options.action === "passport") {
        const dpi = Number(options.dpi);
        canvas = passportCanvas(
          image,
          Math.round(Number(options.widthMm) / 25.4 * dpi),
          Math.round(Number(options.heightMm) / 25.4 * dpi),
        );
      } else if (options.action === "transform") {
        const rotation = Number(options.rotation);
        const swapped = rotation === 90 || rotation === 270;
        canvas = transformedCanvas(
          image,
          swapped ? height : width,
          swapped ? width : height,
          rotation,
          options.flip,
        );
      } else {
        canvas = drawImage(image, width, height);
      }

      if (options.action === "grayscale") {
        const grayscale = createCanvas(canvas.width, canvas.height);
        const context = grayscale.getContext("2d");
        context.filter = "grayscale(1)";
        context.drawImage(canvas, 0, 0);
        canvas = grayscale;
      }

      const format = options.action === "convert" ? options.format : sourceFormat(file);
      if (format === "pdf") {
        outputs.push(await imageToPdf(canvas, file.name));
      } else {
        const output = outputTypes[format] || outputTypes.png;
        if (format === "jpeg") {
          const flattened = createCanvas(canvas.width, canvas.height, "#ffffff");
          flattened.getContext("2d").drawImage(canvas, 0, 0);
          canvas = flattened;
        }
        const blob = await canvasToBlob(canvas, output.mime);
        if (blob.type !== output.mime) {
          throw new Error(`This browser does not support ${format.toUpperCase()} export. Choose PNG or JPG instead.`);
        }
        const suffix = options.action === "resize" ? "resized"
          : options.action === "grayscale" ? "grayscale"
            : options.action === "transform" ? "transformed"
              : options.action === "passport" ? "passport-style" : "converted";
        outputs.push({ name: `${baseName(file.name)}-${suffix}${output.extension}`, data: blob });
      }
    } finally {
      if (typeof image.close === "function") image.close();
    }
    report(Math.round((index + 1) / files.length * 100), `Finished ${file.name}.`);
  }
  return outputs;
}
