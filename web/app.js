import { zipSync } from "fflate";
import { processPhotos } from "./photo.js";
import { processPdfs } from "./pdf.js";
import { processVideo } from "./video.js";
import { processWord } from "./word.js";
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
  word: {
    name: "Word tools",
    eyebrow: "WORD TOOL",
    title: "Work with Word documents",
    accept: ".docx",
    multiple: false,
    fileTitle: "Choose a Word document",
    tipTitle: "Arrange Word files your way",
    tipCopy: "For Merge DOCX files, choose 2–30 documents, then drag the rows or use Move up and Move down. Word tools never mix in PDFs or images.",
    help: "DOCX files are processed locally. Merging and PDF conversion are separate actions.",
    maxFileSize: 100 * 1024 * 1024,
    actions: [
      { value: "merge", label: "Merge DOCX files" },
      { value: "to-pdf", label: "Convert DOCX to PDF" },
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
const fileSelectionHelp = document.getElementById("file-selection-help");
const fileOrderStatus = document.getElementById("file-order-status");
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
    (state.tool === "pdf" && ["merge", "combine"].includes(actionSelect.value)) ||
    (state.tool === "word" && actionSelect.value === "merge");
}

function acceptedExtensions() {
  if (state.tool === "pdf" && actionSelect.value === "combine") {
    return [...tools.pdf.accept.split(","), ...tools.photo.accept.split(",")];
  }
  if (state.tool === "pdf" && actionSelect.value !== "combine") return [".pdf"];
  return tools[state.tool].accept.split(",");
}

function maximumFileCount() {
  return state.tool === "word" && actionSelect.value === "merge" ? 30 : Infinity;
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
  } else if (state.tool === "word" && action === "merge") {
    actionFields.innerHTML = `<p class="field-note">Choose 2–30 .docx files. Arrange them below; PDFs and images are not supported in Word merges.</p>`;
  } else if (state.tool === "word" && action === "to-pdf") {
    actionFields.innerHTML = `<p class="field-note">Choose exactly one .docx file. Conversion runs in this browser; complex layouts can differ from Microsoft Word. Use the Windows app for Word's native PDF export.</p>`;
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
      : state.tool === "word" && actionSelect.value === "merge" ? "Choose DOCX files"
        : state.tool === "word" ? "Choose a DOCX file"
          : state.tool === "pdf" ? "Choose a PDF" : tool.fileTitle;
  filePickerHint.textContent = state.tool === "word" && actionSelect.value === "merge"
    ? "Choose 2–30 files; tap to browse or drop"
    : multiple ? "Tap to browse or drop files here" : "Tap to browse or drop one file here";
  updateFileSelectionHelp();
}

function updateFileSelectionHelp() {
  const count = state.files.length;
  const multiple = allowsMultiple();
  let message;
  if (state.tool === "word" && actionSelect.value === "merge") {
    message = count === 0
      ? "Choose 2–30 .docx files to merge. PDFs and images cannot be added here."
      : count === 1
        ? "1 document selected. Add at least one more .docx file."
        : `${count} documents selected. Drag rows or use Move up/down to set the merge order.`;
  } else if (state.tool === "word") {
    message = count === 0
      ? "Choose exactly one .docx file to convert to PDF."
      : "1 document selected. The PDF will use the original file name.";
  } else if (state.tool === "pdf" && actionSelect.value === "merge") {
    message = count === 0
      ? "Choose at least two PDF files to merge."
      : count === 1
        ? "1 PDF selected. Add at least one more PDF."
        : `${count} PDFs selected. Drag rows or use Move up/down to set the page order.`;
  } else if (multiple) {
    message = count === 0
      ? "Choose one or more supported files. You can arrange them below."
      : `${count} file${count === 1 ? "" : "s"} selected. Drag rows or use Move up/down to change the order.`;
  } else {
    message = count === 0 ? "Choose a supported file to begin." : "1 file selected and ready.";
  }
  fileSelectionHelp.textContent = state.busy ? "Processing locally. File controls are paused until this edit finishes." : message;
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

let draggedFileIndex = null;

function focusFileControl(index, preferredAction) {
  const row = selectedFiles.querySelector(`[data-index="${index}"]`);
  if (!row) {
    filePicker.focus();
    return;
  }
  const preferred = row.querySelector(`[data-action="${preferredAction}"]:not(:disabled)`);
  const fallback = row.querySelector("[data-action]:not(:disabled)");
  (preferred || fallback || filePicker).focus();
}

function moveFile(from, to) {
  if (state.busy || from === to || from < 0 || to < 0 || from >= state.files.length || to >= state.files.length) return;
  const [file] = state.files.splice(from, 1);
  state.files.splice(to, 0, file);
  clearError();
  clearResults();
  renderSelectedFiles();
  fileOrderStatus.textContent = `${file.name} moved to position ${to + 1} of ${state.files.length}.`;
}

function renderSelectedFiles() {
  fileOrderStatus.textContent = "";
  selectedFiles.replaceChildren();
  const multiple = allowsMultiple();
  state.files.forEach((file, index) => {
    const row = document.createElement("li");
    row.className = "selected-file";
    row.dataset.index = String(index);
    row.draggable = multiple && state.files.length > 1 && !state.busy;

    const number = document.createElement("span");
    number.className = "selected-file-number";
    number.textContent = String(index + 1);
    number.setAttribute("aria-hidden", "true");

    const handle = document.createElement("span");
    handle.className = "drag-handle";
    handle.textContent = multiple ? "↕" : "";
    handle.setAttribute("aria-hidden", "true");
    handle.title = multiple ? "Drag to reorder" : "";

    const copy = document.createElement("div");
    copy.className = "selected-file-copy";
    const name = document.createElement("strong");
    name.textContent = file.name;
    name.title = file.name;
    const size = document.createElement("small");
    size.textContent = formatBytes(file.size);
    copy.append(name, size);

    const actions = document.createElement("div");
    actions.className = "selected-file-actions";
    if (multiple) {
      const up = document.createElement("button");
      up.type = "button";
      up.className = "file-control move-file";
      up.dataset.action = "up";
      up.textContent = "↑";
      up.title = "Move up";
      up.setAttribute("aria-label", `Move ${file.name} up`);
      up.disabled = state.busy || index === 0;
      up.addEventListener("click", () => {
        moveFile(index, index - 1);
        focusFileControl(index - 1, "up");
      });

      const down = document.createElement("button");
      down.type = "button";
      down.className = "file-control move-file";
      down.dataset.action = "down";
      down.textContent = "↓";
      down.title = "Move down";
      down.setAttribute("aria-label", `Move ${file.name} down`);
      down.disabled = state.busy || index === state.files.length - 1;
      down.addEventListener("click", () => {
        moveFile(index, index + 1);
        focusFileControl(index + 1, "down");
      });
      actions.append(up, down);
    }

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "file-control remove-file";
    remove.dataset.action = "remove";
    remove.textContent = "×";
    remove.title = "Remove file";
    remove.setAttribute("aria-label", `Remove ${file.name}`);
    remove.disabled = state.busy;
    remove.addEventListener("click", () => {
      if (state.busy) return;
      state.files.splice(index, 1);
      clearError();
      clearResults();
      renderSelectedFiles();
      focusFileControl(Math.min(index, state.files.length - 1), "remove");
    });
    actions.append(remove);

    row.append(number, handle, copy, actions);
    row.addEventListener("dragstart", (event) => {
      if (state.busy || !multiple || state.files.length < 2) {
        event.preventDefault();
        return;
      }
      draggedFileIndex = index;
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", String(index));
      row.classList.add("is-dragging");
    });
    row.addEventListener("dragend", () => {
      draggedFileIndex = null;
      row.classList.remove("is-dragging");
      selectedFiles.querySelectorAll(".drop-before, .drop-after").forEach((item) => {
        item.classList.remove("drop-before", "drop-after");
      });
    });
    selectedFiles.append(row);
  });

  updateFileSelectionHelp();
}

selectedFiles.addEventListener("dragover", (event) => {
  if (draggedFileIndex === null || state.busy) return;
  event.preventDefault();
  event.dataTransfer.dropEffect = "move";
  selectedFiles.querySelectorAll(".drop-before, .drop-after").forEach((item) => {
    item.classList.remove("drop-before", "drop-after");
  });
  const row = event.target.closest(".selected-file");
  if (!row) return;
  const bounds = row.getBoundingClientRect();
  row.classList.add(event.clientY < bounds.top + bounds.height / 2 ? "drop-before" : "drop-after");
});

selectedFiles.addEventListener("drop", (event) => {
  if (draggedFileIndex === null || state.busy) return;
  event.preventDefault();
  const row = event.target.closest(".selected-file");
  if (!row) return;
  const targetIndex = Number(row.dataset.index);
  const placeAfter = row.classList.contains("drop-after");
  let destination = targetIndex + (placeAfter ? 1 : 0);
  if (draggedFileIndex < destination) destination -= 1;
  const from = draggedFileIndex;
  draggedFileIndex = null;
  moveFile(from, destination);
});

selectedFiles.addEventListener("dragleave", (event) => {
  if (selectedFiles.contains(event.relatedTarget)) return;
  selectedFiles.querySelectorAll(".drop-before, .drop-after").forEach((item) => {
    item.classList.remove("drop-before", "drop-after");
  });
});

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
  if (!allowsMultiple() && incoming.length !== 1) {
    showError(state.tool === "word"
      ? "Choose exactly one .docx file to convert to PDF."
      : "Choose exactly one file for this edit.");
    return;
  }
  const next = allowsMultiple() ? [...state.files, ...incoming] : incoming.slice(0, 1);
  const unique = next.filter((file, index) => next.findIndex((item) => item.name === file.name && item.size === file.size) === index);
  if (!filesAreCompatible(unique)) {
    showError(state.tool === "word"
      ? "Choose .docx files only. PDFs and images cannot be merged or converted with the Word tools."
      : "Choose supported files for the selected edit.");
    return;
  }
  if (state.tool === "word" && unique.some((file) => !file.size)) {
    showError("This Word file is empty. Choose another .docx file.");
    return;
  }
  if (unique.length > maximumFileCount()) {
    showError("Choose no more than 30 .docx files for a Word merge.");
    return;
  }
  const tool = tools[state.tool];
  if (unique.some((file) => file.size > tool.maxFileSize)) {
    const fileLabel = state.tool === "word" ? "Word document" : state.tool === "pdf" ? "PDF or image" : state.tool;
    showError(`Each ${fileLabel} must be ${formatBytes(tool.maxFileSize)} or smaller. Try a smaller file or use the Windows app.`);
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
  if (state.busy) return;
  if (toolName === state.tool) {
    if (focus) document.getElementById(`tab-${toolName}`)?.focus();
    return;
  }
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
  tool.actions.forEach((action, index) => {
    const option = document.createElement("option");
    option.value = action.value;
    option.textContent = action.label;
    option.selected = index === 0;
    actionSelect.append(option);
  });
  actionSelect.selectedIndex = 0;
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
    if (state.tool === "word" && actionSelect.value === "merge") return "Choose at least two .docx files to merge.";
    if (state.tool === "word") return "Choose one .docx file to convert to PDF.";
    return state.tool === "pdf" && actionSelect.value === "merge"
      ? "Choose at least two PDF files to merge."
      : `Choose ${allowsMultiple() ? "one or more files" : "a file"} to continue.`;
  }
  if (!filesAreCompatible(state.files)) {
    return state.tool === "word"
      ? "Choose .docx files only. PDFs and images cannot be used with the Word tools."
      : "Choose supported files for the selected edit.";
  }
  if (state.tool === "word" && state.files.some((file) => !file.size)) {
    return "This Word file is empty. Choose another .docx file.";
  }
  if (!allowsMultiple() && state.files.length !== 1) {
    return state.tool === "word" ? "Choose exactly one .docx file to convert to PDF." : "Choose exactly one file for this edit.";
  }
  if (state.tool === "word" && actionSelect.value === "merge" && (state.files.length < 2 || state.files.length > 30)) {
    return "Choose between 2 and 30 .docx files to merge.";
  }
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
  if (/^(Choose|Enter|Select|Use|This PDF contains|Each |The selected files|Those dimensions|This Word file is empty|One or more Word documents|This Word document could not)/i.test(message)) return message;
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
  document.querySelectorAll(".tool-card").forEach((card) => {
    card.disabled = busy;
  });
  editorPanel.setAttribute("aria-busy", String(busy));
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
    const order = ["photo", "pdf", "video", "word"];
    const index = order.indexOf(card.dataset.tool);
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
  if (state.files.length > maximumFileCount()) state.files = [];
  fileInput.value = "";
  renderActionFields();
  updateFileControl();
  renderSelectedFiles();
  clearError();
  clearResults();
});

filePicker.addEventListener("click", () => {
  if (!state.busy) fileInput.click();
});
fileInput.addEventListener("change", () => {
  addFiles(fileInput.files);
  fileInput.value = "";
});
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
        : state.tool === "word"
          ? await processWord(state.files, options, report)
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
