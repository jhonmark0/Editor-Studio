import { FFmpeg } from "@ffmpeg/ffmpeg";
import { baseName, fetchWithProgress } from "./utils.js";

const coreBase = "https://cdn.jsdelivr.net/npm/@ffmpeg/core@0.12.10/dist/umd";
let ffmpeg = null;
let loading = null;

async function loadEngine(report) {
  if (ffmpeg?.loaded) return ffmpeg;
  if (loading) return loading;
  const instance = new FFmpeg();
  loading = (async () => {
    report(2, "Downloading the browser video engine.");
    const coreBlob = await fetchWithProgress(
      `${coreBase}/ffmpeg-core.js`,
      (percent, detail) => report(Math.round(2 + percent * 0.08), detail),
      "Video engine",
    );
    const coreModule = new Blob([`${await coreBlob.text()}\nexport default createFFmpegCore;`], { type: "text/javascript" });
    const coreURL = URL.createObjectURL(coreModule);
    let wasmURL = "";
    try {
      const wasmBlob = await fetchWithProgress(
        `${coreBase}/ffmpeg-core.wasm`,
        (percent, detail) => report(Math.round(10 + percent * 0.75), detail),
        "Video processing data",
      );
      wasmURL = URL.createObjectURL(wasmBlob);
      const classWorkerURL = new URL("./assets/ffmpeg-worker.js", document.baseURI).href;
      await instance.load({ classWorkerURL, coreURL, wasmURL });
      report(88, "Video engine ready.");
      ffmpeg = instance;
      return instance;
    } finally {
      URL.revokeObjectURL(coreURL);
      if (wasmURL) URL.revokeObjectURL(wasmURL);
    }
  })();
  try {
    return await loading;
  } catch (error) {
    instance.terminate();
    loading = null;
    throw error;
  }
}

export async function processVideo(file, options, report = () => {}) {
  if (!window.WebAssembly || !window.Worker) {
    throw new Error("This browser does not support the video engine. Use a current browser or the Windows app.");
  }
  const engine = await loadEngine(report);
  const extension = `.${file.name.split(".").pop().toLowerCase()}`;
  const inputName = `input${extension}`;
  const extractingAudio = options.action === "extract-audio";
  const suffix = extractingAudio ? "audio" : options.action;
  const outputExtension = extractingAudio ? ".m4a" : ".mp4";
  const outputName = `${baseName(file.name)}-${suffix}${outputExtension}`;
  const temporaryOutput = `output${outputExtension}`;
  const args = ["-nostdin", "-hide_banner", "-loglevel", "error"];

  if (options.action === "trim") {
    args.push("-ss", String(Number(options.start)), "-i", inputName, "-t", String(Number(options.duration)));
  } else {
    args.push("-i", inputName);
  }

  if (extractingAudio) {
    args.push("-map", "0:a:0", "-vn", "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart", temporaryOutput);
  } else {
    args.push("-map", "0:v:0", "-map", "0:a?", "-c:v", "libx264", "-preset", "medium", "-crf", "23",
      "-c:a", "aac", "-movflags", "+faststart", temporaryOutput);
  }

  const progressHandler = ({ progress }) => {
    const percent = Number.isFinite(progress) ? Math.max(0, Math.min(1, progress)) : 0;
    report(Math.round(88 + percent * 12), extractingAudio ? "Extracting audio locally." : "Processing video locally.");
  };
  let lastVideoError = "";
  const logHandler = ({ type, message }) => {
    if (type === "stderr" && message !== "Aborted()") lastVideoError = message;
  };
  engine.on("progress", progressHandler);
  engine.on("log", logHandler);
  try {
    report(89, "Adding the video to the local workspace.");
    await engine.writeFile(inputName, new Uint8Array(await file.arrayBuffer()));
    const code = await engine.exec(args);
    if (code !== 0) throw new Error(`Video processing failed in the browser: ${lastVideoError}`);
    report(100, "Reading the finished file.");
    const data = await engine.readFile(temporaryOutput);
    if (typeof data === "string" || !data.length) throw new Error("Video processing did not create an output file.");
    return [{ name: outputName, data }];
  } finally {
    engine.off("progress", progressHandler);
    engine.off("log", logHandler);
    await engine.deleteFile(inputName).catch(() => {});
    await engine.deleteFile(temporaryOutput).catch(() => {});
  }
}
