const office = window.automationOffice;
const appVersionElement = document.getElementById("app-version");
const teamSelect = document.getElementById("team");
const actionSelect = document.getElementById("action");
const actionFields = document.getElementById("action-fields");
const chooseFilesButton = document.getElementById("choose-files");
const fileTitle = document.getElementById("file-picker-title");
const fileHint = document.getElementById("file-picker-hint");
const selectedFilesElement = document.getElementById("selected-files");
const folderButton = document.getElementById("choose-folder");
const folderName = document.getElementById("output-folder-name");
const activityList = document.getElementById("activity-list");
const workStatus = document.getElementById("work-status");
const resultPanel = document.getElementById("results");
const formError = document.getElementById("form-error");
const submitButton = document.getElementById("submit-job");

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
    (teamSelect.value === "pdf" && ["merge", "combine"].includes(actionSelect.value));
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
  fileTitle.textContent = teamType === "image" && action === "resize" ? "Choose image files"
    : teamType === "pdf" && action === "merge" ? "Choose PDF files"
      : teamType === "pdf" && action === "combine" ? "Choose image/PDF files"
        : teamType === "pdf" ? "Choose a PDF" : teams[teamType].fileTitle;
  fileHint.textContent = multiple ? "Drop files here, or choose them" : "Drop one file here, or choose it";
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
  }
}

function updateSelectedFiles() {
  selectedFilesElement.textContent = selectedFiles.map((file) => file.split(/[\\/]/).pop()).join(", ");
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
  updateSelectedFiles();
  renderActionFields();
}

function updateAction() {
  const multiple = allowsMultipleFiles();
  const supportedExtensions = extensionsForCurrentTask();
  const unsupported = selectedFiles.some((file) => {
    const name = file.split(/[\\/]/).pop().toLowerCase();
    return !supportedExtensions.some((extension) => name.endsWith(extension));
  });
  const missingMergeInput = teamSelect.value === "pdf" && actionSelect.value === "merge" && selectedFiles.length < 2;
  if ((!multiple && selectedFiles.length > 1) || unsupported || missingMergeInput) {
    selectedFiles = [];
    updateSelectedFiles();
  }
  renderActionFields();
}

function buildJob() {
  const type = teamSelect.value;
  const action = actionSelect.value;
  const job = { type, action, files: selectedFiles, outputDir };
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

function showError(error) {
  formError.textContent = error && error.message ? error.message : "Something went wrong. Please try again.";
}

teamSelect.addEventListener("change", updateTeam);
actionSelect.addEventListener("change", updateAction);
document.getElementById("focus-assignment").addEventListener("click", () => {
  document.getElementById("task-form").scrollIntoView({ behavior: "smooth", block: "center" });
});

chooseFilesButton.addEventListener("click", async () => {
  formError.textContent = "";
  const team = teams[teamSelect.value];
  try {
    selectedFiles = await office.chooseFiles({
      type: team.fileType,
      multiple: allowsMultipleFiles(),
      action: actionSelect.value,
    });
    updateSelectedFiles();
  } catch (error) {
    showError(error);
  }
});

let filePickerDragDepth = 0;
chooseFilesButton.addEventListener("dragenter", (event) => {
  event.preventDefault();
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
  const team = teams[teamSelect.value];
  const multiple = allowsMultipleFiles();
  const supportedExtensions = extensionsForCurrentTask();
  if (!paths.length) {
    formError.textContent = "No files were found in that drop.";
    return;
  }
  if (paths.length > 30) {
    formError.textContent = "Drop no more than 30 files at a time.";
    return;
  }
  const unsupported = paths.find((file) => {
    const name = file.split(/[\\/]/).pop().toLowerCase();
    return !supportedExtensions.some((extension) => name.endsWith(extension));
  });
  if (unsupported) {
    formError.textContent = `That file type is not supported for ${team.name.toLowerCase()}.`;
    return;
  }
  if (!multiple && paths.length !== 1) {
    formError.textContent = "Drop exactly one file for this task.";
    return;
  }
  if (teamSelect.value === "pdf" && actionSelect.value === "merge" && paths.length < 2) {
    formError.textContent = "Drop at least two PDFs to merge.";
    return;
  }
  selectedFiles = paths;
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

document.getElementById("task-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  if (jobIsRunning) return;
  formError.textContent = "";
  resultPanel.replaceChildren();
  if (!selectedFiles.length) {
    formError.textContent = "Choose the source file(s) for this job.";
    return;
  }
  const job = buildJob();
  if (!job.outputDir) {
    formError.textContent = "Choose a folder for the result files.";
    return;
  }
  jobIsRunning = true;
  submitButton.disabled = true;
  submitButton.textContent = "Working…";
  setWorkStatus("WORKING");
  Object.keys(stationIds).forEach((agent) => setStationStatus(agent, "READY"));
  activityList.replaceChildren();
  try {
    const outputs = await office.submitJob(job);
    const specialist = teams[teamSelect.value].name;
    setStationStatus(specialist, "DONE");
    addActivity({ agent: specialist, message: "The finished files are ready." });
    renderResults(outputs);
    setWorkStatus("COMPLETE");
  } catch (error) {
    showError(error);
    const specialist = teams[teamSelect.value].name;
    setStationStatus(specialist, "ATTENTION");
    addActivity({ agent: specialist, message: `The task could not be completed: ${error.message}` });
    setWorkStatus("NEEDS ATTENTION");
  } finally {
    jobIsRunning = false;
    submitButton.disabled = false;
    submitButton.innerHTML = 'Assign task <span>→</span>';
  }
});

office.onActivity(addActivity);
updateTeam();
