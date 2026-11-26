import * as docx from "docx-preview";
import html2canvas from "html2canvas";
import * as PDFLib from "pdf-lib";
import JSZip from "jszip";
import { mergeDocxDocuments } from "../src/docx-merge.mjs";
import { convertDocxToPdf } from "../src/docx-to-pdf.mjs";
import { baseName, readFile } from "./utils.js";

const libraries = { docx, html2canvas, PDFLib, JSZip };

function validateWordFiles(files, action) {
  if (!files.length) {
    throw new Error(action === "merge"
      ? "Choose at least two .docx files to merge."
      : "Choose one .docx file to convert to PDF.");
  }
  if (files.some((file) => !file.name.toLowerCase().endsWith(".docx"))) {
    throw new Error("Choose .docx files only. PDFs and images are not supported by the Word tools.");
  }
  if (files.some((file) => !file.size)) {
    throw new Error("This Word file is empty. Choose another .docx file.");
  }
  if (action === "merge" && (files.length < 2 || files.length > 30)) {
    throw new Error("Choose between 2 and 30 .docx files to merge.");
  }
  if (action === "to-pdf" && files.length !== 1) {
    throw new Error("Choose exactly one .docx file to convert to PDF.");
  }
}

export async function processWord(files, options, report = () => {}) {
  const action = options.action;
  validateWordFiles(files, action);

  try {
    if (action === "merge") {
      const inputs = [];
      for (let index = 0; index < files.length; index += 1) {
        const file = files[index];
        report(Math.round(index / files.length * 10), `Reading ${file.name}.`);
        inputs.push({ name: file.name, data: await readFile(file) });
      }
      report(10, "Merging Word documents in the selected order.");
      let mergePercent = 15;
      const mergeReport = (message) => {
        if (/^Merging /.test(message)) {
          mergePercent = Math.min(90, mergePercent + Math.round(70 / (files.length - 1)));
        } else if (/merged/i.test(message)) {
          mergePercent = 95;
        }
        report(mergePercent, message);
      };
      const data = await mergeDocxDocuments(inputs, mergeReport);
      report(100, "The merged Word document is ready.");
      return [{ name: "merged-documents.docx", data }];
    }

    if (action === "to-pdf") {
      const file = files[0];
      report(5, `Opening ${file.name}.`);
      const data = await readFile(file);
      const pdf = await convertDocxToPdf(data, report, libraries);
      return [{ name: `${baseName(file.name)}.pdf`, data: pdf }];
    }

    throw new Error("Choose a Word action to continue.");
  } catch (error) {
    if (/^(Choose |Select |This Word file is empty)/.test(error?.message || "")) throw error;
    throw new Error(action === "merge"
      ? "One or more Word documents could not be merged. Check that each file is a valid .docx and try again."
      : "This Word document could not be converted to PDF. Check that it is a valid .docx and try again.");
  }
}
