import { build } from "esbuild";
import { copyFile, mkdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const docs = path.join(root, "docs");
const assets = path.join(docs, "assets");

await mkdir(assets, { recursive: true });

await Promise.all([
  build({
    entryPoints: [path.join(root, "web", "app.js")],
    outfile: path.join(assets, "app.js"),
    bundle: true,
    format: "esm",
    platform: "browser",
    target: ["es2020"],
    minify: true,
    sourcemap: false,
    alias: {
      "pdfjs-dist": path.join(root, "node_modules", "pdfjs-dist", "build", "pdf.min.mjs"),
    },
    define: {
      "process.env.NODE_ENV": "\"production\"",
    },
    logLevel: "info",
  }),
  build({
    entryPoints: [path.join(root, "node_modules", "@ffmpeg", "ffmpeg", "dist", "esm", "worker.js")],
    outfile: path.join(assets, "ffmpeg-worker.js"),
    bundle: true,
    format: "esm",
    platform: "browser",
    target: ["es2020"],
    minify: true,
    sourcemap: false,
    logLevel: "info",
  }),
]);

await Promise.all([
  copyFile(path.join(root, "web", "index.html"), path.join(docs, "index.html")),
  copyFile(path.join(root, "web", "styles.css"), path.join(assets, "styles.css")),
  copyFile(path.join(root, "src", "logo.svg"), path.join(assets, "logo.svg")),
  copyFile(
    path.join(root, "node_modules", "pdfjs-dist", "build", "pdf.worker.min.mjs"),
    path.join(assets, "pdf.worker.min.mjs"),
  ),
]);

console.log("Browser workspace built into docs/.");
