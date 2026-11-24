import { zipSync } from "fflate";
import { processPhotos } from "./photo.js";
import { processPdfs } from "./pdf.js";
import { processVideo } from "./video.js";
import { formatBytes } from "./utils.js";

const tools = {
  photo: {
    name: "Photo editor",
    eyebrow: "PHOTO TOOL",
    title: "Edit a photo",
    accept: ".png,.jpg,.jpeg,.webp,.tif,.tiff,.bmp,.gif",
    multiple: false,
    fileTitle: "Choose an image",
    tipTitle: "A quicker photo workflow",
    tipCopy: "Choose Resize image to process several photos at once. The editor keeps their aspect ratios.",
    help: "Your original file stays untouched.",
    maxFileSize: 40 * 1024 * 1024,
    actions: [
      { value: "resize", label: "Resize image" },
      { value: "transform", label: "Rotate or flip" },
      { value: "convert", label: "Convert format" },
      { value: "grayscale", label: "Make grayscale" },
      { value: "passport", label: "Passport-style crop (35 × 45 mm)" },
    ],
  },
  pdf: {
    name: "PDF editor",
    eyebrow: "PDF TOOL",
    title: "Edit a PDF",
    accept: ".pdf,.png,.jpg,.jpeg,.webp,.tif,.tiff,.bmp,.gif",
    multiple: false,
    fileTitle: "Choose a PDF",
    tipTitle: "Keep the page order you want",
    tipCopy: "For Merge PDFs and Combine images and PDFs, select files in the order they should appear.",
    help: "Pages are combined in the order you select them.",
    maxFileSize: 100 * 1024 * 1024,
    actions: [
      { value: "merge", label: "Merge PDFs" },
      { value: "combine", label: "Combine images and PDFs" },
      { value: "split", label: "Split into pages" },
      { value: "to-word", label: "Convert PDF to Word" },
    ],
  },
  video: {
    name: "Video editor",
    eyebrow: "VIDEO TOOL",
    title: "Edit a video",
    accept: ".mp4,.mov,.m4v,.mkv,.webm,.avi",
    multiple: false,
    fileTitle: "Choose a video",
    tipTitle: "A quicker video workflow",
    tipCopy: "Trim only the section you need to reduce processing time and keep the result file small.",
    help: "Video processing can take longer on phones and tablets.",
    maxFileSize: 300 * 1024 * 1024,
    actions: [
      { value: "trim", label: "Trim a clip" },
      { value: "convert", label: "Convert to MP4" },
      { value: "extract-audio", label: "Extract audio" },
    ],
  },
};

const state = {
  tool: "photo",
  files: [],
  busy: false,
  resultUrls: [],
};

const toolForm = document.getElementById("tool-form");
const actionSelect = document.getElementById("action");
const actionFields = document.getElementById("action-fields");
const fileInput = document.getElementById("file-input");
const filePicker = document.getElementById("choose-files");
const filePickerTitle = document.getElementById("file-picker-title");
const filePickerHint = document.getElementById("file-picker-hint");
const selectedFiles = document.getElementById("selected-files");
const processButton = document.getElementById("process-button");
const formError = document.getElementById("form-error");
const progressPanel = document.getElementById("progress-panel");
const progressTitle = document.getElementById("progress-title");
const progressPercent = document.getElementById("progress-percent");
const progressTrack = document.getElementById("progress-track");
const progressBar = document.getElementById("progress-bar");
const progressDetail = document.getElementById("progress-detail");
const resultsPanel = document.getElementById("results");
const resultLinks = document.getElementById("result-links");
const editorPanel = document.getElementById("editor-panel");

function allowsMultiple() {
  return (state.tool === "photo" && actionSelect.value === "resize") ||
    (state.tool === "pdf" && ["merge", "combine"].includes(actionSelect.value));
}

function acceptedExtensions() {
  if (state.tool === "pdf" && actionSelect.value === "combine") {
    return [...tools.pdf.accept.split(","), ...tools.photo.accept.split(",")];
  }
  if (state.tool === "pdf" && actionSelect.value !== "combine") return [".pdf"];
  return tools[state.tool].accept.split(",");
}

function numberField(id, label, value, min, max, step = "1") {
  return `<div class="input-group"><label for="${id}">${label}</label><input class="text-input" id="${id}" type="number" value="${value}" min="${min}" max="${max}" step="${step}"></div>`;
}

function renderActionFields() {
  const action = actionSelect.value;
  if (state.tool === "photo" && action === "resize") {
    actionFields.innerHTML = `<div class="dynamic-fields input-row">${numberField("width", "Width (pixels)", 1200, 1, 12000)}${numberField("height", "Height (pixels)", 1200, 1, 12000)}</div>`;
  } else if (state.tool === "photo" && action === "transform") {
    actionFields.innerHTML = `<div class="dynamic-fields input-row">
      <div class="input-group"><label for="rotation">Rotate</label><select id="rotation"><option value="90">90° clockwise</option><option value="180">180°</option><option value="270">270° clockwise</option><option value="0">No rotation</option></select></div>
      <div class="input-group"><label for="flip">Flip</label><select id="flip"><option value="none">No flip</option><option value="horizontal">Horizontal</option><option value="vertical">Vertical</option></select></div>
    </div>`;
  } else if (state.tool === "photo" && action === "convert") {
    actionFields.innerHTML = `<div class="dynamic-fields input-group"><label for="format">Output format</label><select id="format"><option value="jpeg">JPG</option><option value="png">PNG</option><option value="webp">WebP</option><option value="pdf">PDF</option></select></div>`;
  } else if (state.tool === "photo" && action === "passport") {
    actionFields.innerHTML = `<div class="dynamic-fields input-row">${numberField("width-mm", "Width (mm)", 35, 15, 100, "0.1")}${numberField("height-mm", "Height (mm)", 45, 15, 100, "0.1")}</div>
      <div class="input-group"><label for="dpi">Print resolution</label><select id="dpi"><option value="300">300 DPI</option><option value="600">600 DPI</option><option value="150">150 DPI</option></select></div>
      <p class="field-note">This creates a centered crop at the selected print size. It does not change clothing and is not certified for official passport use.</p>`;
  } else if (state.tool === "video" && action === "trim") {
    actionFields.innerHTML = `<div class="dynamic-fields input-row">${numberField("start", "Start (seconds)", 0, 0, Number.MAX_SAFE_INTEGER, "0.1")}${numberField("duration", "Duration (seconds)", 10, 0.1, Number.MAX_SAFE_INTEGER, "0.1")}</div>`;
  } else if (state.tool === "pdf" && action === "combine") {
    actionFields.innerHTML = `<p class="field-note">Images become A4 pages; existing PDF pages keep their original size.</p>`;
  } else if (state.tool === "pdf" && action === "to-word") {
    actionFields.innerHTML = `<p class="field-note">Keeps page layout and artwork in place while making selectable text editable in Word. Scanned PDFs need OCR first.</p>`;
  } else {
    actionFields.replaceChildren();
  }
}

function updateFileControl() {
  const tool = tools[state.tool];
  const multiple = allowsMultiple();
  fileInput.accept = acceptedExtensions().join(",");
  fileInput.multiple = multiple;
  filePickerTitle.textContent = state.tool === "pdf" && actionSelect.value === "merge" ? "Choose PDF files"
    : state.tool === "pdf" && actionSelect.value === "combine" ? "Choose image/PDF files"
      : state.tool === "pdf" ? "Choose a PDF" : tool.fileTitle;
  filePickerHint.textContent = multiple ? "Tap to browse or drop files here" : "Tap to browse or drop one file here";
}

function clearResults() {
  state.resultUrls.forEach((url) => URL.revokeObjectURL(url));
  state.resultUrls = [];
  resultLinks.replaceChildren();
  resultsPanel.hidden = true;
}

function clearError() {
  formError.textContent = "";
}

function renderSelectedFiles() {
  selectedFiles.replaceChildren();
  state.files.forEach((file, index) => {
    const row = document.createElement("div");
    row.className = "selected-file";
    const icon = document.createElement("span");
    icon.className = "selected-file-icon";
    icon.textContent = file.name.split(".").pop().slice(0, 3).toUpperCase() || "FILE";
    const copy = document.createElement("div");
    copy.className = "selected-file-copy";
    const name = document.createElement("strong");
    name.textContent = file.name;
    name.title = file.name;
    const size = document.createElement("small");
    size.textContent = formatBytes(file.size);
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "remove-file";
    remove.textContent = "×";
    remove.setAttribute("aria-label", `Remove ${file.name}`);
    remove.disabled = state.busy;
    remove.addEventListener("click", () => {
      state.files.splice(index, 1);
      renderSelectedFiles();
    });
    copy.append(name, size);
    row.append(icon, copy, remove);
    selectedFiles.append(row);
  });
}

function filesAreCompatible(files) {
  const extensions = acceptedExtensions();
  return files.every((file) => extensions.includes(`.${file.name.split(".").pop().toLowerCase()}`));
}

function addFiles(files) {
  if (state.busy) return;
  clearError();
  clearResults();
  const incoming = Array.from(files);
  if (!incoming.length) return;
  const next = allowsMultiple() ? [...state.files, ...incoming] : incoming.slice(0, 1);
  const unique = next.filter((file, index) => next.findIndex((item) => item.name === file.name && item.size === file.size) === index);
  if (!filesAreCompatible(unique)) {
    showError("Choose supported files for the selected edit.");
    return;
  }
  const tool = tools[state.tool];
  if (unique.some((file) => file.size > tool.maxFileSize)) {
    showError(`Each ${state.tool === "pdf" ? "PDF or image" : state.tool} must be ${formatBytes(tool.maxFileSize)} or smaller. Try a smaller file or use the Windows app.`);
    return;
  }
  if (unique.reduce((total, file) => total + file.size, 0) > 600 * 1024 * 1024) {
    showError("The selected files are too large for browser memory. Process them in smaller batches.");
    return;
  }
  state.files = unique;
  renderSelectedFiles();
}

function selectTool(toolName, focus = false) {
  if (state.busy || toolName === state.tool) return;
  state.tool = toolName;
  state.files = [];
  fileInput.value = "";
  const tool = tools[toolName];
  document.querySelectorAll(".tool-card").forEach((card) => {
    const active = card.dataset.tool === toolName;
    card.classList.toggle("active", active);
    card.setAttribute("aria-selected", String(active));
    card.tabIndex = active ? 0 : -1;
    if (active && focus) card.focus();
  });
  document.getElementById("tool-eyebrow").textContent = tool.eyebrow;
  document.getElementById("tool-title").textContent = tool.title;
  document.getElementById("tip-title").textContent = tool.tipTitle;
  document.getElementById("tip-copy").textContent = tool.tipCopy;
  document.getElementById("tool-help").textContent = tool.help;
  editorPanel.setAttribute("aria-labelledby", `tab-${toolName}`);
  actionSelect.replaceChildren();
  tool.actions.forEach((action) => {
    const option = document.createElement("option");
    option.value = action.value;
    option.textContent = action.label;
    actionSelect.append(option);
  });
  renderActionFields();
  updateFileControl();
  renderSelectedFiles();
  clearError();
  clearResults();
  progressPanel.hidden = true;
}

function readOptions() {
  const value = (id, fallback = "") => {
    const element = document.getElementById(id);
    return element ? element.value : fallback;
  };
  return {
    action: actionSelect.value,
    width: value("width"),
    height: value("height"),
    rotation: value("rotation"),
    flip: value("flip"),
    format: value("format"),
    widthMm: value("width-mm"),
    heightMm: value("height-mm"),
    dpi: value("dpi"),
    start: value("start"),
    duration: value("duration"),
  };
}

function validateJob() {
  if (!state.files.length) {
    return state.tool === "pdf" && actionSelect.value === "merge"
      ? "Choose at least two PDF files to merge."
      : `Choose ${allowsMultiple() ? "one or more files" : "a file"} to continue.`;
  }
  if (!allowsMultiple() && state.files.length !== 1) return "Choose exactly one file for this edit.";
  if (state.tool === "pdf" && actionSelect.value === "merge" && state.files.length < 2) {
    return "Select at least two PDFs to merge.";
  }
  const options = readOptions();
  if (state.tool === "photo" && options.action === "resize") {
    const width = Number(options.width);
    const height = Number(options.height);
    if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1 || width > 12000 || height > 12000) {
      return "Enter a width and height between 1 and 12000 pixels.";
    }
    if (width * height > 100_000_000) return "Those dimensions are too large for browser memory. Try a smaller size.";
  }
  if (state.tool === "photo" && options.action === "transform" && Number(options.rotation) === 0 && options.flip === "none") {
    return "Choose a rotation, a flip, or both for the photo.";
  }
  if (state.tool === "photo" && options.action === "passport") {
    const width = Number(options.widthMm);
    const height = Number(options.heightMm);
    if (!Number.isFinite(width) || !Number.isFinite(height) || width < 15 || height < 15 || width > 100 || height > 100) {
      return "Use dimensions from 15 to 100 mm.";
    }
  }
  if (state.tool === "video" && options.action === "trim") {
    const start = Number(options.start);
    const duration = Number(options.duration);
    if (!Number.isFinite(start) || start < 0 || !Number.isFinite(duration) || duration <= 0) {
      return "Enter a start time of zero or more and a duration greater than zero.";
    }
  }
  return "";
}

function showError(message) {
  formError.textContent = message;
  formError.scrollIntoView({ block: "nearest", behavior: "smooth" });
}

function friendlyError(error) {
  const message = error?.message || "";
  if (/^(Choose|Enter|Select|Use|This PDF contains|Each |The selected files|Those dimensions)/i.test(message)) return message;
  if (/password|encrypted/i.test(message)) return "This PDF is password-protected. Unlock it first, then try again.";
  if (/memory|allocation|quota|too large/i.test(message)) return "This file is too large for your browser to process. Try a smaller file or use the Windows app.";
  if (/network|failed to fetch|could not be downloaded/i.test(message)) return "The browser video engine could not be downloaded. Check your connection and try again.";
  if (/ffmpeg|video engine|video processing|aborted/i.test(message)) return "This video could not be processed. Try a shorter clip, another video, or the Windows app.";
  if (/pdf to word|pdf processing|pdfjs/i.test(message)) return "This PDF could not be processed. Check the file and try again.";
  if (/photo|image|canvas|webp export/i.test(message)) return "This image could not be edited. Try another image or output format.";
  return "Something went wrong while processing the file. Please try again or use the Windows app.";
}

function showProgress(percent, title, detail) {
  const safePercent = Math.max(0, Math.min(100, Math.round(Number(percent) || 0)));
  progressPanel.hidden = false;
  progressTitle.textContent = title;
  progressPercent.textContent = `${safePercent}%`;
  progressBar.style.width = `${safePercent}%`;
  progressTrack.setAttribute("aria-valuenow", String(safePercent));
  progressDetail.textContent = detail;
}

function setBusy(busy) {
  state.busy = busy;
  processButton.disabled = busy;
  processButton.innerHTML = busy ? "Working locally…" : `Create download <span aria-hidden="true">→</span>`;
  toolForm.querySelectorAll("select, input, button").forEach((element) => {
    if (element !== processButton) element.disabled = busy;
  });
  renderSelectedFiles();
}

function mimeType(name) {
  const extension = name.split(".").pop().toLowerCase();
  return ({
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    png: "image/png",
    webp: "image/webp",
    pdf: "application/pdf",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    mp4: "video/mp4",
    m4a: "audio/mp4",
    zip: "application/zip",
  })[extension] || "application/octet-stream";
}

async function renderResults(outputs) {
  clearResults();
  let name;
  let blob;
  if (outputs.length === 1) {
    name = outputs[0].name;
    const data = outputs[0].data;
    blob = data instanceof Blob ? data : new Blob([data], { type: mimeType(name) });
  } else {
    name = "editor-studio-results.zip";
    const entries = {};
    for (const output of outputs) {
      let bytes = output.data;
      if (bytes instanceof Blob) bytes = new Uint8Array(await bytes.arrayBuffer());
      entries[output.name] = bytes;
    }
    blob = new Blob([zipSync(entries, { level: 6 })], { type: "application/zip" });
  }
  const url = URL.createObjectURL(blob);
  state.resultUrls.push(url);
  const link = document.createElement("a");
  link.className = "result-link";
  link.href = url;
  link.download = name;
  const title = document.createElement("strong");
  title.textContent = name;
  const action = document.createElement("span");
  action.textContent = "Download";
  link.append(title, action);
  resultLinks.append(link);
  resultsPanel.hidden = false;
  resultsPanel.scrollIntoView({ block: "nearest", behavior: "smooth" });
}

document.querySelectorAll(".tool-card").forEach((card) => {
  card.addEventListener("click", () => selectTool(card.dataset.tool));
  card.addEventListener("keydown", (event) => {
    const order = ["photo", "pdf", "video"];
    const index = order.indexOf(state.tool);
    let next = index;
    if (event.key === "ArrowRight") next = (index + 1) % order.length;
    else if (event.key === "ArrowLeft") next = (index - 1 + order.length) % order.length;
    else if (event.key === "Home") next = 0;
    else if (event.key === "End") next = order.length - 1;
    else return;
    event.preventDefault();
    selectTool(order[next], true);
  });
});

actionSelect.addEventListener("change", () => {
  if (state.busy) return;
  if (!allowsMultiple() && state.files.length > 1) state.files = [];
  if (state.files.length && !filesAreCompatible(state.files)) state.files = [];
  renderActionFields();
  updateFileControl();
  renderSelectedFiles();
  clearError();
  clearResults();
});

filePicker.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", () => addFiles(fileInput.files));
["dragenter", "dragover"].forEach((eventName) => {
  filePicker.addEventListener(eventName, (event) => {
    event.preventDefault();
    if (!state.busy) filePicker.classList.add("dragging");
  });
});
["dragleave", "drop"].forEach((eventName) => {
  filePicker.addEventListener(eventName, (event) => {
    event.preventDefault();
    filePicker.classList.remove("dragging");
  });
});
filePicker.addEventListener("drop", (event) => addFiles(event.dataTransfer.files));

toolForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (state.busy) return;
  clearError();
  clearResults();
  const validationError = validateJob();
  if (validationError) {
    showError(validationError);
    return;
  }
  const options = readOptions();
  setBusy(true);
  progressPanel.hidden = false;
  showProgress(0, `Working with the ${tools[state.tool].name.toLowerCase()}`, "Processing locally in your browser.");
  try {
    const report = (percent, detail) => showProgress(percent, `Working with the ${tools[state.tool].name.toLowerCase()}`, detail);
    const outputs = state.tool === "photo"
      ? await processPhotos(state.files, options, report)
      : state.tool === "pdf"
        ? await processPdfs(state.files, options, report)
        : await processVideo(state.files[0], options, report);
    showProgress(100, "Finished", "Preparing your download.");
    await renderResults(outputs);
    progressPanel.hidden = true;
  } catch (error) {
    console.error("Browser edit failed.", error);
    progressPanel.hidden = true;
    showError(friendlyError(error));
  } finally {
    setBusy(false);
  }
});

state.tool = "";
selectTool("photo");
renderSelectedFiles();
