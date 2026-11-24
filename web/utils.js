export function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(units.length - 1, Math.floor(Math.log(bytes) / Math.log(1024)));
  const value = bytes / 1024 ** index;
  return `${value >= 10 || index === 0 ? value.toFixed(0) : value.toFixed(1)} ${units[index]}`;
}

export function baseName(fileName) {
  return fileName.replace(/\.[^.]+$/, "").replace(/[<>:"/\\|?*\u0000-\u001f]/g, "-") || "editor-studio";
}

export async function readFile(file) {
  return new Uint8Array(await file.arrayBuffer());
}

export async function loadImage(file) {
  if ("createImageBitmap" in window) {
    try {
      return await createImageBitmap(file, { imageOrientation: "from-image" });
    } catch {
      // Fall back for browsers with incomplete image bitmap support.
    }
  }
  const url = URL.createObjectURL(file);
  try {
    const image = new Image();
    image.src = url;
    await image.decode();
    return image;
  } finally {
    URL.revokeObjectURL(url);
  }
}

export function canvasToBlob(canvas, type, quality = 0.92) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error("The browser could not create the output image."));
    }, type, quality);
  });
}

export function drawImage(image, width, height) {
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(width));
  canvas.height = Math.max(1, Math.round(height));
  const context = canvas.getContext("2d", { alpha: true });
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";
  context.drawImage(image, 0, 0, canvas.width, canvas.height);
  return canvas;
}

export async function fetchWithProgress(url, onProgress, label) {
  let response;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      response = await fetch(url);
      break;
    } catch (error) {
      if (attempt === 1) throw error;
      await new Promise((resolve) => setTimeout(resolve, 600));
    }
  }
  if (!response.ok) throw new Error(`${label} could not be downloaded. Please check your connection and try again.`);
  const total = Number(response.headers.get("content-length")) || 0;
  if (!response.body) {
    onProgress(100, `${label} downloaded.`);
    return await response.blob();
  }
  const reader = response.body.getReader();
  const chunks = [];
  let transferred = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    transferred += value.length;
    const percent = total ? Math.min(99, Math.round(transferred / total * 100)) : 0;
    onProgress(percent, total
      ? `${label}: ${formatBytes(transferred)} of ${formatBytes(total)}`
      : `${label}: ${formatBytes(transferred)}`);
  }
  onProgress(100, `${label} ready.`);
  return new Blob(chunks, { type: response.headers.get("content-type") || "application/octet-stream" });
}
