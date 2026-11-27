const office = window.automationOffice;
const appVersionElement = document.getElementById("app-version");
const teamSelect = document.getElementById("team");
const actionSelect = document.getElementById("action");
const actionFields = document.getElementById("action-fields");
const chooseFilesButton = document.getElementById("choose-files");
const fileTitle = document.getElementById("file-picker-title");
const fileHint = document.getElementById("file-picker-hint");
const selectedFilesElement = document.getElementById("selected-files");
const fileOrderHelp = document.getElementById("file-order-help");
const fileOrderStatus = document.getElementById("file-order-status");
const folderButton = document.getElementById("choose-folder");
const folderName = document.getElementById("output-folder-name");
const activityList = document.getElementById("activity-list");
const workStatus = document.getElementById("work-status");
const resultPanel = document.getElementById("results");
const formError = document.getElementById("form-error");
const submitButton = document.getElementById("submit-job");
const updateBanner = document.getElementById("update-banner");
const updateIcon = document.getElementById("update-icon");
const updateTitle = document.getElementById("update-title");
const updateDetail = document.getElementById("update-detail");
const updatePercent = document.getElementById("update-percent");
const updateProgress = document.getElementById("update-progress");
const updateProgressBar = document.getElementById("update-progress-bar");

let selectedFiles = [];
let outputDir = null;
let jobIsRunning = false;

office.getAppVersion().then((version) => {
  if (typeof version !== "string" || !version) {
    throw new Error("The application version is unavailable.");
  }
  appVersionElement.textContent = `v${version}`;
}).catch((error) => {
  appVersionElement.textContent = "Version unavailable";
  console.error("Failed to load the application version.", error);
});

const stationIds = {
  "Photo editor": "station-photo",
  "PDF editor": "station-pdf",
  "Word editor": "station-word",
  "Video editor": "station-video",
};

const teams = {
  image: {
    name: "Photo editor",
    fileType: "image",
    extensions: [".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp", ".gif"],
    multiple: false,
    fileTitle: "Choose an image",
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
    fileType: "pdf",
    extensions: [".pdf"],
    fileTitle: "Choose PDF file(s)",
    actions: [
      { value: "merge", label: "Merge PDFs" },
      { value: "combine", label: "Combine images and PDFs" },
      { value: "split", label: "Split into pages" },
      { value: "to-word", label: "Convert PDF to Word" },
    ],
  },
  word: {
    name: "Word editor",
    fileType: "word",
    extensions: [".docx"],
    fileTitle: "Choose Word documents",
    actions: [
      { value: "merge", label: "Merge Word documents (.docx)" },
      { value: "to-pdf", label: "Convert one Word document to PDF" },
    ],
  },
  video: {
    name: "Video editor",
    fileType: "video",
    extensions: [".mp4", ".mov", ".m4v", ".mkv", ".webm", ".avi"],
    multiple: false,
    fileTitle: "Choose a video",
    actions: [
      { value: "trim", label: "Trim a clip" },
      { value: "convert", label: "Convert to MP4" },
      { value: "extract-audio", label: "Extract audio" },
    ],
  },
};

function allowsMultipleFiles() {
  return (teamSelect.value === "image" && actionSelect.value === "resize") ||
    (teamSelect.value === "pdf" && ["merge", "combine"].includes(actionSelect.value)) ||
    (teamSelect.value === "word" && actionSelect.value === "merge");
}

function extensionsForCurrentTask() {
  const team = teams[teamSelect.value];
  if (teamSelect.value === "pdf" && actionSelect.value === "combine") {
    return [...team.extensions, ...teams.image.extensions];
  }
  return team.extensions;
}

function setWorkStatus(status) {
  workStatus.classList.toggle("busy", status === "WORKING");
  workStatus.innerHTML = `<i></i> ${status}`;
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(units.length - 1, Math.floor(Math.log(bytes) / Math.log(1024)));
  const value = bytes / 1024 ** index;
  return `${value >= 10 || index === 0 ? value.toFixed(0) : value.toFixed(1)} ${units[index]}`;
}

function showUpdateStatus(status) {
  updateBanner.hidden = false;
  updateBanner.classList.toggle("error", status.status === "error");
  updateBanner.classList.toggle("downloaded", status.status === "downloaded");
  updatePercent.hidden = status.status === "error";
  updateProgress.hidden = status.status === "error";

  if (status.status === "downloading") {
    const percent = Math.max(0, Math.min(100, Math.round(Number(status.percent) || 0)));
    const transferred = formatBytes(status.transferred);
    const total = formatBytes(status.total);
    const speed = formatBytes(status.bytesPerSecond);
    updateIcon.textContent = "↓";
    updateTitle.textContent = status.version ? `Downloading Editor Studio ${status.version}` : "Downloading update";
    updateDetail.textContent = [transferred && total ? `${transferred} of ${total}` : "", speed ? `${speed}/s` : ""]
      .filter(Boolean).join(" · ");
    updatePercent.textContent = `${percent}%`;
    updateProgressBar.style.width = `${percent}%`;
    updateProgress.setAttribute("aria-valuenow", String(percent));
    return;
  }

  if (status.status === "downloaded") {
    updateIcon.textContent = "✓";
    updateTitle.textContent = status.version ? `Editor Studio ${status.version} is ready` : "Update ready";
    updateDetail.textContent = "Restart the app to install the update.";
    updatePercent.textContent = "100%";
    updateProgressBar.style.width = "100%";
    updateProgress.setAttribute("aria-valuenow", "100");
    return;
  }

  updateIcon.textContent = "!";
  updateTitle.textContent = "Update download failed";
  updateDetail.textContent = status.message || "The update could not be downloaded. Please try again later.";
}

function setStationStatus(agent, status) {
  const station = document.getElementById(stationIds[agent]);
  if (!station) return;
  station.textContent = status;
  station.className = `status ${status.toLowerCase()}`;
}

function addActivity({ agent, message, time = new Date().toISOString() }) {
  const empty = activityList.querySelector(".empty-state");
  if (empty) empty.remove();

  const initials = agent.split(/\s+/).map((part) => part[0]).join("").slice(0, 2).toUpperCase();
  const item = document.createElement("article");
  item.className = "activity-item";
  const avatar = document.createElement("div");
  avatar.className = "activity-avatar";
  avatar.textContent = initials;
  const body = document.createElement("div");
  body.className = "activity-body";
  const meta = document.createElement("div");
  meta.className = "activity-meta";
  const name = document.createElement("strong");
  name.textContent = agent;
  const timestamp = document.createElement("time");
  timestamp.textContent = new Date(time).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  const text = document.createElement("p");
  text.className = "activity-message";
  text.textContent = message;
  const status = /could not be completed|failed|error/i.test(message) ? "ATTENTION"
    : /complete|returned|saved|ready/i.test(message) ? "DONE" : "WORKING";
  setStationStatus(agent, status);
  meta.append(name, timestamp);
  body.append(meta, text);
  item.append(avatar, body);
  activityList.append(item);
  activityList.scrollTop = activityList.scrollHeight;
}

function addNumberField(parent, id, labelText, value, min, max, step = "1") {
  const group = document.createElement("div");
  group.className = "input-group";
  const label = document.createElement("label");
  label.htmlFor = id;
  label.textContent = labelText;
  const input = document.createElement("input");
  input.id = id;
  input.type = "number";
  input.min = String(min);
  input.max = String(max);
  input.step = step;
  input.value = String(value);
  group.append(label, input);
  parent.append(group);
}

function renderActionFields() {
  const teamType = teamSelect.value;
  const action = actionSelect.value;
  const multiple = allowsMultipleFiles();
  const wordAction = teamType === "word";
  fileTitle.textContent = wordAction
    ? action === "merge" ? "Choose Word documents" : "Choose a Word document"
    : teamType === "image" && action === "resize" ? "Choose image files"
      : teamType === "pdf" && action === "merge" ? "Choose PDF files"
        : teamType === "pdf" && action === "combine" ? "Choose image/PDF files"
          : teamType === "pdf" ? "Choose a PDF" : teams[teamType].fileTitle;
  fileHint.textContent = wordAction && action === "merge"
    ? "Choose 2–30 .docx files, or drop them here"
    : multiple ? "Drop files here, or choose them" : "Drop one file here, or choose it";
  fileOrderHelp.hidden = !multiple;
  actionFields.replaceChildren();

  if (teamType === "image" && action === "resize") {
    const row = document.createElement("div");
    row.className = "dynamic-fields input-row";
    addNumberField(row, "width", "Width (pixels)", 1200, 1, 12000);
    addNumberField(row, "height", "Height (pixels)", 1200, 1, 12000);
    actionFields.append(row);
  } else if (teamType === "image" && action === "transform") {
    actionFields.innerHTML = `<div class="dynamic-fields input-row">
      <div class="input-group"><label for="rotation">Rotate</label><select id="rotation">
        <option value="90">90° clockwise</option><option value="180">180°</option>
        <option value="270">270° clockwise</option><option value="0">No rotation</option>
      </select></div>
      <div class="input-group"><label for="flip">Flip</label><select id="flip">
        <option value="none">No flip</option><option value="horizontal">Horizontal</option>
        <option value="vertical">Vertical</option>
      </select></div>
    </div>`;
  } else if (teamType === "image" && action === "convert") {
    actionFields.innerHTML = `<div class="dynamic-fields input-group"><label for="format">Output format</label><select id="format"><option value="jpeg">JPG</option><option value="png">PNG</option><option value="webp">WebP</option><option value="pdf">PDF</option></select></div>`;
  } else if (teamType === "image" && action === "passport") {
    const row = document.createElement("div");
    row.className = "dynamic-fields input-row";
    addNumberField(row, "width-mm", "Width (mm)", 35, 15, 100, "0.1");
    addNumberField(row, "height-mm", "Height (mm)", 45, 15, 100, "0.1");
    const dpi = document.createElement("div");
    dpi.className = "input-group";
    dpi.innerHTML = `<label for="dpi">Print resolution</label><select id="dpi"><option value="300">300 DPI</option><option value="600">600 DPI</option><option value="150">150 DPI</option></select>`;
    actionFields.append(row, dpi);
    const note = document.createElement("p");
    note.className = "passport-size-note";
    note.textContent = "This creates a centered crop at the selected print size. It does not change clothing and is not certified for official passport use.";
    actionFields.append(note);
  } else if (teamType === "video" && action === "trim") {
    const row = document.createElement("div");
    row.className = "dynamic-fields input-row";
    addNumberField(row, "start", "Start (seconds)", 0, 0, Number.MAX_SAFE_INTEGER, "0.1");
    addNumberField(row, "duration", "Duration (seconds)", 10, 0.1, Number.MAX_SAFE_INTEGER, "0.1");
    actionFields.append(row);
  } else if (teamType === "pdf" && action === "combine") {
    const note = document.createElement("p");
    note.className = "passport-size-note";
    note.textContent = "Images become A4 pages; existing PDF pages keep their original size.";
    actionFields.append(note);
  } else if (teamType === "pdf" && action === "to-word") {
    const note = document.createElement("p");
    note.className = "passport-size-note";
    note.textContent = "Keeps page layout and artwork in place while making selectable text editable in Word. Scanned PDFs need OCR first.";
    actionFields.append(note);
  } else if (wordAction) {
    const help = document.createElement("div");
    help.className = "word-help";
    const heading = document.createElement("strong");
    heading.textContent = "Word help";
    const copy = document.createElement("p");
    copy.textContent = action === "merge"
      ? "Choose 2–30 .docx files. They are combined from top to bottom, keeping images already inside each document. PDFs and separate images are not accepted. The result is saved as merged-documents.docx."
      : "Choose exactly one .docx file. When Microsoft Word is available, the Windows app uses its native PDF export to preserve page layout and selectable text. The result is saved as <input-name>.pdf. Conversion can take a moment.";
    help.append(heading, copy);
    actionFields.append(help);
  }

  if (multiple) {
    const orderNote = document.createElement("p");
    orderNote.className = "passport-size-note";
    orderNote.textContent = "The top-to-bottom order below is the order used for the finished file.";
    actionFields.append(orderNote);
  }
}

let draggedFileIndex = null;

function moveSelectedFile(fromIndex, toIndex, focusDirection) {
  if (jobIsRunning || !allowsMultipleFiles() || fromIndex === toIndex) return;
  if (!Number.isInteger(fromIndex) || !Number.isInteger(toIndex) ||
      fromIndex < 0 || toIndex < 0 || fromIndex >= selectedFiles.length || toIndex >= selectedFiles.length) {
    return;
  }
  const [movedFile] = selectedFiles.splice(fromIndex, 1);
  selectedFiles.splice(toIndex, 0, movedFile);
  draggedFileIndex = null;
  updateSelectedFiles({ index: toIndex, direction: focusDirection });
  fileOrderStatus.textContent = `${movedFile.split(/[\\/]/).pop()} moved to position ${toIndex + 1} of ${selectedFiles.length}.`;
}

function updateSelectedFiles(focus) {
  selectedFilesElement.replaceChildren();
  selectedFilesElement.classList.toggle("arrangement-disabled", jobIsRunning);
  selectedFilesElement.setAttribute("aria-busy", String(jobIsRunning));
  if (!selectedFiles.length) return;

  const multiple = allowsMultipleFiles();
  selectedFiles.forEach((file, index) => {
    const name = file.split(/[\\/]/).pop();
    const row = document.createElement("li");
    row.className = "selected-file-row";
    row.dataset.index = String(index);
    row.draggable = multiple && !jobIsRunning;

    const number = document.createElement("span");
    number.className = "selected-file-number";
    number.setAttribute("aria-hidden", "true");
    number.textContent = String(index + 1).padStart(2, "0");

    const fileName = document.createElement("span");
    fileName.className = "selected-file-name";
    fileName.textContent = name;
    fileName.title = file;

    row.append(number, fileName);
    if (multiple) {
      const controls = document.createElement("div");
      controls.className = "file-move-controls";
      [
        { direction: "up", label: "Move up", offset: -1 },
        { direction: "down", label: "Move down", offset: 1 },
      ].forEach(({ direction, label, offset }) => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "file-move-button";
        button.dataset.index = String(index);
        button.dataset.direction = direction;
        button.draggable = false;
        button.disabled = jobIsRunning || index + offset < 0 || index + offset >= selectedFiles.length;
        button.setAttribute("aria-label", `${label}: ${name}`);
        button.textContent = label;
        button.addEventListener("click", () => moveSelectedFile(index, index + offset, direction));
        controls.append(button);
      });
      row.append(controls);

      row.addEventListener("dragstart", (event) => {
        if (jobIsRunning || event.target.closest("button")) {
          event.preventDefault();
          return;
        }
        draggedFileIndex = index;
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData("text/plain", String(index));
        row.classList.add("dragging");
      });
      row.addEventListener("dragover", (event) => {
        if (jobIsRunning || draggedFileIndex === null) return;
        event.preventDefault();
        event.dataTransfer.dropEffect = "move";
        row.classList.add("drag-over");
      });
      row.addEventListener("dragleave", () => row.classList.remove("drag-over"));
      row.addEventListener("drop", (event) => {
        event.preventDefault();
        event.stopPropagation();
        row.classList.remove("drag-over");
        const transferIndex = Number(event.dataTransfer.getData("text/plain"));
        const sourceIndex = draggedFileIndex === null ? transferIndex : draggedFileIndex;
        moveSelectedFile(sourceIndex, index, sourceIndex < index ? "down" : "up");
      });
      row.addEventListener("dragend", () => {
        draggedFileIndex = null;
        row.classList.remove("dragging", "drag-over");
        selectedFilesElement.querySelectorAll(".drag-over").forEach((item) => item.classList.remove("drag-over"));
      });
    }
    selectedFilesElement.append(row);
  });

  if (focus && Number.isInteger(focus.index)) {
    const preferred = selectedFilesElement.querySelector(
      `.file-move-button[data-index="${focus.index}"][data-direction="${focus.direction}"]:not(:disabled)`,
    );
    const fallback = selectedFilesElement.querySelector(
      `.file-move-button[data-index="${focus.index}"]:not(:disabled)`,
    );
    (preferred || fallback)?.focus();
  }
}

function validateSelectedPaths(paths) {
  const team = teams[teamSelect.value];
  const action = actionSelect.value;
  const multiple = allowsMultipleFiles();
  if (!Array.isArray(paths) || paths.length === 0) return "Choose the source file(s) for this job.";
  if (paths.length > 30) return "Choose no more than 30 files at a time.";
  if (!multiple && paths.length !== 1) return `Choose exactly one file for ${team.name.toLowerCase()}.`;

  const supportedExtensions = extensionsForCurrentTask();
  const unsupported = paths.find((file) => {
    const name = typeof file === "string" ? file.split(/[\\/]/).pop().toLowerCase() : "";
    return !supportedExtensions.some((extension) => name.endsWith(extension));
  });
  if (unsupported) {
    return teamSelect.value === "word"
      ? "Word tasks accept .docx files only. PDFs and images cannot be used for Word tasks."
      : `That file type is not supported for ${team.name.toLowerCase()}.`;
  }
  if (teamSelect.value === "word" && action === "merge" && paths.length < 2) {
    return "Choose between 2 and 30 Word documents (.docx) to merge.";
  }
  if (teamSelect.value === "pdf" && action === "merge" && paths.length < 2) {
    return "Choose at least two PDFs to merge.";
  }
  return null;
}

function updateTeam() {
  const team = teams[teamSelect.value];
  actionSelect.replaceChildren();
  team.actions.forEach((action) => {
    const option = document.createElement("option");
    option.value = action.value;
    option.textContent = action.label;
    actionSelect.append(option);
  });
  selectedFiles = [];
  draggedFileIndex = null;
  fileOrderStatus.textContent = "";
  updateSelectedFiles();
  renderActionFields();
}

function updateAction() {
  if (selectedFiles.length && validateSelectedPaths(selectedFiles)) {
    selectedFiles = [];
    draggedFileIndex = null;
    fileOrderStatus.textContent = "";
    updateSelectedFiles();
  }
  renderActionFields();
}

function buildJob() {
  const type = teamSelect.value;
  const action = actionSelect.value;
  const job = { type, action, files: [...selectedFiles], outputDir };
  if (type === "image" && action === "resize") {
    job.width = document.getElementById("width").value;
    job.height = document.getElementById("height").value;
  }
  if (type === "image" && action === "convert") {
    job.format = document.getElementById("format").value;
  }
  if (type === "image" && action === "transform") {
    job.rotation = document.getElementById("rotation").value;
    job.flip = document.getElementById("flip").value;
  }
  if (type === "image" && action === "passport") {
    job.widthMm = document.getElementById("width-mm").value;
    job.heightMm = document.getElementById("height-mm").value;
    job.dpi = document.getElementById("dpi").value;
  }
  if (type === "video" && action === "trim") {
    job.start = document.getElementById("start").value;
    job.duration = document.getElementById("duration").value;
  }
  return job;
}

function renderResults(files) {
  resultPanel.replaceChildren();
  const card = document.createElement("div");
  card.className = "result-card";
  const title = document.createElement("p");
  title.className = "result-title";
  title.textContent = "RESULTS READY";
  card.append(title);
  files.forEach((file) => {
    const button = document.createElement("button");
    button.className = "result-file";
    button.type = "button";
    const name = document.createElement("span");
    name.textContent = file.split(/[\\/]/).pop();
    const reveal = document.createElement("span");
    reveal.textContent = "Show folder ↗";
    button.append(name, reveal);
    button.addEventListener("click", () => office.revealFile(file).catch(showError));
    card.append(button);
  });
  resultPanel.append(card);
}

function formatErrorMessage(error) {
  const rawMessage = error && error.message ? error.message : "";
  const message = rawMessage
    .replace(/^Error invoking remote method '[^']+':\s*/i, "")
    .replace(/^Error:\s*/, "")
    .trim();
  if (!message) return "Something went wrong. Please try again.";
  if (/^(Choose|Enter|Select|Use|Drop|Unsupported|Input file not found:|The output folder|An input file path|This PDF|The requested result file|No files were found|That file type)/i.test(message)) {
    return message;
  }
  if (/ffmpeg|spawn|ENOENT/i.test(message)) {
    return "The video processor could not start. Close and reopen Editor Studio, then try again.";
  }
  if (/video processing failed/i.test(message)) {
    return "This video could not be processed. Try another video or convert it to MP4 first.";
  }
  if (/photo edit failed/i.test(message)) {
    return "This image could not be edited. Try another image or output format.";
  }
  if (/timed out|timeout/i.test(message) && /word|docx/i.test(message)) {
    return "The Word-to-PDF conversion timed out. Try a smaller document or close other apps, then try again.";
  }
  if (/word.*(merge|merged)|merge.*word/i.test(message)) {
    return "The Word documents could not be merged. Make sure each file is a valid .docx that opens in Word, then try again.";
  }
  if (/word.*pdf|pdf.*word|docx.*convert|convert.*docx/i.test(message)) {
    return "This Word document could not be converted to PDF. Make sure it is a valid .docx, try a smaller document, or close other apps and try again.";
  }
  if (/pdf/i.test(message)) {
    return "This PDF could not be processed. Check the file and try again.";
  }
  if (teamSelect.value === "word") {
    return actionSelect.value === "merge"
      ? "The Word documents could not be merged. Make sure each file is a valid .docx that opens in Word, then try again."
      : "This Word document could not be converted to PDF. Make sure it is a valid .docx and try again.";
  }
  return "Something went wrong. Please try again.";
}

function showError(error) {
  formError.textContent = typeof error === "string" ? error : formatErrorMessage(error);
}

teamSelect.addEventListener("change", updateTeam);
actionSelect.addEventListener("change", updateAction);
document.getElementById("focus-assignment").addEventListener("click", () => {
  document.getElementById("task-form").scrollIntoView({ behavior: "smooth", block: "center" });
});

chooseFilesButton.addEventListener("click", async () => {
  formError.textContent = "";
  if (jobIsRunning) return;
  const team = teams[teamSelect.value];
  try {
    const paths = await office.chooseFiles({
      type: team.fileType,
      multiple: allowsMultipleFiles(),
      action: actionSelect.value,
    });
    if (!paths.length) return;
    const validationError = validateSelectedPaths(paths);
    if (validationError) {
      showError(validationError);
      return;
    }
    selectedFiles = paths;
    draggedFileIndex = null;
    fileOrderStatus.textContent = "";
    updateSelectedFiles();
  } catch (error) {
    showError(error);
  }
});

let filePickerDragDepth = 0;
chooseFilesButton.addEventListener("dragenter", (event) => {
  event.preventDefault();
  if (jobIsRunning) return;
  filePickerDragDepth += 1;
  chooseFilesButton.classList.add("drag-active");
});
chooseFilesButton.addEventListener("dragover", (event) => {
  event.preventDefault();
  event.dataTransfer.dropEffect = "copy";
});
chooseFilesButton.addEventListener("dragleave", (event) => {
  event.preventDefault();
  filePickerDragDepth = Math.max(0, filePickerDragDepth - 1);
  if (filePickerDragDepth === 0) chooseFilesButton.classList.remove("drag-active");
});
chooseFilesButton.addEventListener("drop", (event) => {
  event.preventDefault();
  filePickerDragDepth = 0;
  chooseFilesButton.classList.remove("drag-active");
});

office.onFilesDropped((paths) => {
  formError.textContent = "";
  if (jobIsRunning) {
    formError.textContent = "Please wait for the current task to finish before changing files.";
    return;
  }
  if (!paths.length) {
    formError.textContent = "No files were found in that drop.";
    return;
  }
  const validationError = validateSelectedPaths(paths);
  if (validationError) {
    showError(validationError);
    return;
  }
  selectedFiles = paths;
  draggedFileIndex = null;
  fileOrderStatus.textContent = "";
  updateSelectedFiles();
});

folderButton.addEventListener("click", async () => {
  formError.textContent = "";
  try {
    const selected = await office.chooseOutputFolder();
    if (selected) {
      outputDir = selected;
      folderName.textContent = selected;
      folderButton.querySelector("span:first-child").textContent = "Output folder selected";
    }
  } catch (error) {
    showError(error);
  }
});

function setJobRunning(running) {
  jobIsRunning = running;
  submitButton.disabled = running;
  submitButton.textContent = running ? "Working…" : "";
  if (!running) submitButton.innerHTML = 'Assign task <span>→</span>';
  teamSelect.disabled = running;
  actionSelect.disabled = running;
  chooseFilesButton.disabled = running;
  folderButton.disabled = running;
  actionFields.querySelectorAll("input, select").forEach((control) => {
    control.disabled = running;
  });
  updateSelectedFiles();
}

document.getElementById("task-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  if (jobIsRunning) return;
  formError.textContent = "";
  resultPanel.replaceChildren();
  const validationError = validateSelectedPaths(selectedFiles);
  if (validationError) {
    showError(validationError);
    return;
  }
  const job = buildJob();
  const specialist = teams[job.type].name;
  if (!job.outputDir) {
    formError.textContent = "Choose a folder for the result files.";
    return;
  }
  setJobRunning(true);
  setWorkStatus("WORKING");
  Object.keys(stationIds).forEach((agent) => setStationStatus(agent, "READY"));
  activityList.replaceChildren();
  try {
    const outputs = await office.submitJob(job);
    setStationStatus(specialist, "DONE");
    addActivity({ agent: specialist, message: "The finished files are ready." });
    renderResults(outputs);
    setWorkStatus("COMPLETE");
  } catch (error) {
    console.error("The edit could not be completed.", error);
    showError(error);
    setStationStatus(specialist, "ATTENTION");
    addActivity({ agent: specialist, message: `The task could not be completed: ${formatErrorMessage(error)}` });
    setWorkStatus("NEEDS ATTENTION");
  } finally {
    setJobRunning(false);
  }
});

office.onActivity(addActivity);
office.onUpdateStatus(showUpdateStatus);
updateTeam();
