const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function wordChild(element, name) {
  return element && [...element.children]
    .find((child) => child.namespaceURI === WORD_NS && child.localName === name);
}

function inheritSectionReferences(section, previous) {
  const clone = section.cloneNode(true);
  if (!previous) return clone;

  const references = new Map();
  for (const source of [previous, section]) {
    for (const name of ["headerReference", "footerReference"]) {
      for (const reference of [...source.children]
        .filter((child) => child.namespaceURI === WORD_NS && child.localName === name)) {
        const type = reference.getAttributeNS(WORD_NS, "type") || "default";
        references.set(`${name}:${type}`, reference.cloneNode(true));
      }
    }
  }
  for (const child of [...clone.children]) {
    if (child.namespaceURI === WORD_NS &&
        ["headerReference", "footerReference"].includes(child.localName)) {
      clone.removeChild(child);
    }
  }
  let index = 0;
  for (const name of ["headerReference", "footerReference"]) {
    for (const [key, reference] of references) {
      if (!key.startsWith(`${name}:`)) continue;
      clone.insertBefore(reference, clone.children[index] || null);
      index += 1;
    }
  }
  return clone;
}

function wordDescendants(element, name) {
  return element ? [...element.getElementsByTagNameNS(WORD_NS, name)] : [];
}

async function ensureDefaultFontSize(zip) {
  const stylesPart = zip.file("word/styles.xml");
  if (!stylesPart) return;
  const xml = await stylesPart.async("string");
  const root = new DOMParser().parseFromString(xml, "application/xml").documentElement;
  if (!root || root.localName !== "styles") return;

  const defaults = wordChild(root, "docDefaults");
  const runDefaults = wordChild(wordChild(defaults, "rPrDefault"), "rPr");
  const defaultStyles = wordDescendants(root, "style").filter((style) => {
    const type = style.getAttributeNS(WORD_NS, "type");
    const isDefault = style.getAttributeNS(WORD_NS, "default");
    return (type === "paragraph" || type === "character") &&
      (isDefault === "1" || isDefault === "true" || isDefault === "on");
  });
  if (wordDescendants(runDefaults, "sz").length ||
      defaultStyles.some((style) => wordDescendants(style, "sz").length)) return;

  const docDefaults = defaults || document.createElementNS(WORD_NS, "docDefaults");
  if (!defaults) root.insertBefore(docDefaults, root.firstChild);
  const rPrDefault = wordChild(docDefaults, "rPrDefault") || document.createElementNS(WORD_NS, "rPrDefault");
  if (!wordChild(docDefaults, "rPrDefault")) docDefaults.append(rPrDefault);
  const rPr = wordChild(rPrDefault, "rPr") || document.createElementNS(WORD_NS, "rPr");
  if (!wordChild(rPrDefault, "rPr")) rPrDefault.append(rPr);
  for (const name of ["sz", "szCs"]) {
    const size = document.createElementNS(WORD_NS, name);
    size.setAttributeNS(WORD_NS, "val", "20");
    rPr.append(size);
  }
  zip.file("word/styles.xml", new XMLSerializer().serializeToString(root));
}

function fieldMarker(instruction = "") {
  const field = instruction.trim().split(/\s+/)[0]?.toUpperCase();
  const markers = {
    PAGE: "[[ES_PAGE]]",
    NUMPAGES: "[[ES_NUMPAGES]]",
    SECTIONPAGES: "[[ES_SECTIONPAGES]]",
    SECTION: "[[ES_SECTION]]",
  };
  return markers[field];
}

function markerRun(sourceRun, marker) {
  const run = document.createElementNS(WORD_NS, "r");
  const properties = wordChild(sourceRun, "rPr");
  if (properties) run.append(properties.cloneNode(true));
  const text = document.createElementNS(WORD_NS, "t");
  text.textContent = marker;
  run.append(text);
  return run;
}

function replaceWordFields(root) {
  for (const paragraph of wordDescendants(root, "p")) {
    let child = paragraph.firstElementChild;
    while (child) {
      const next = child.nextElementSibling;
      if (child.namespaceURI === WORD_NS && child.localName === "fldSimple") {
        const marker = fieldMarker(child.getAttributeNS(WORD_NS, "instr"));
        if (marker) {
          const sourceRun = wordDescendants(child, "r")[0];
          paragraph.replaceChild(markerRun(sourceRun || child, marker), child);
        }
        child = next;
        continue;
      }
      if (child.namespaceURI !== WORD_NS || child.localName !== "r") {
        child = next;
        continue;
      }

      const beginsField = wordDescendants(child, "fldChar")
        .some((field) => field.getAttributeNS(WORD_NS, "fldCharType") === "begin");
      if (!beginsField) {
        child = next;
        continue;
      }

      const fieldRuns = [child];
      let instruction = "";
      let cursor = next;
      while (cursor) {
        fieldRuns.push(cursor);
        instruction += wordDescendants(cursor, "instrText")
          .map((text) => text.textContent).join("");
        const endsField = cursor.namespaceURI === WORD_NS && cursor.localName === "r" &&
          wordDescendants(cursor, "fldChar")
            .some((field) => field.getAttributeNS(WORD_NS, "fldCharType") === "end");
        cursor = cursor.nextElementSibling;
        if (endsField) break;
      }
      const marker = fieldMarker(instruction);
      const following = cursor?.nextElementSibling || null;
      if (marker) {
        for (const run of fieldRuns) paragraph.removeChild(run);
        paragraph.insertBefore(markerRun(child, marker), following);
      }
      child = following;
    }
  }
}

async function prepareFieldMarkers(zip) {
  for (const file of Object.values(zip.files)) {
    if (!/(^|\/)(header|footer)\d+\.xml$/.test(file.name)) continue;
    const xml = await file.async("string");
    const parsed = new DOMParser().parseFromString(xml, "application/xml");
    if (parsed.querySelector("parsererror")) continue;
    replaceWordFields(parsed.documentElement);
    zip.file(file.name, new XMLSerializer().serializeToString(parsed.documentElement));
  }
}

function explicitPageStart(section) {
  const numbering = wordChild(section, "pgNumType");
  const value = Number(numbering?.getAttributeNS(WORD_NS, "start"));
  return Number.isInteger(value) && value > 0 ? value : null;
}

function sectionStartType(section) {
  const type = wordChild(section, "type");
  return type?.getAttributeNS(WORD_NS, "val") || "nextPage";
}

async function generateDocxPackage(zip) {
  return zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

function relativeRect(element, pageBounds) {
  const rect = element.getBoundingClientRect();
  return {
    left: rect.left - pageBounds.left,
    top: rect.top - pageBounds.top,
    right: rect.right - pageBounds.left,
    bottom: rect.bottom - pageBounds.top,
    width: rect.width,
    height: rect.height,
  };
}

function alignContentEnd(start, desiredEnd, blocks) {
  const containing = blocks
    .filter((block) => block.top < desiredEnd - 0.5 && block.bottom > desiredEnd + 0.5)
    .sort((first, second) => first.height - second.height)[0];
  if (!containing || containing.top <= start + 0.5) return desiredEnd;
  return Math.max(start + 1, containing.top);
}

function normalizeParagraphTypography(root) {
  for (const paragraph of root.querySelectorAll("p")) {
    let run = [...paragraph.querySelectorAll("span")]
      .find((span) => span.style.fontSize);
    if (!run) {
      const table = paragraph.closest("table");
      if (table) {
        const cell = paragraph.closest("td, th");
        const row = paragraph.closest("tr");
        for (const scope of [cell, row, table]) {
          if (!scope) continue;
          run = [...scope.querySelectorAll("span")]
            .find((span) => span.style.fontSize);
          if (run) break;
        }
      }
    }
    if (!run) {
      run = [...paragraph.querySelectorAll("span")]
        .find((span) => span.textContent.trim());
    }
    if (!run) continue;
    const runStyle = getComputedStyle(run);
    if (runStyle.fontFamily) paragraph.style.fontFamily = runStyle.fontFamily;
    paragraph.style.fontSize = runStyle.fontSize;
  }
}

function buildPageLayout(page, sectionIndex) {
  page.style.margin = "0";
  page.style.boxShadow = "none";
  page.style.overflow = "visible";
  const bounds = page.getBoundingClientRect();
  const width = Math.ceil(bounds.width);
  const height = bounds.height;
  if (width < 1 || height < 1) throw new Error("The Word document has no printable area.");

  const computed = getComputedStyle(page);
  const padding = {
    top: parseFloat(computed.paddingTop) || 0,
    right: parseFloat(computed.paddingRight) || 0,
    bottom: parseFloat(computed.paddingBottom) || 0,
    left: parseFloat(computed.paddingLeft) || 0,
  };
  const sheetHeight = parseFloat(computed.minHeight) || height;
  const bodyHeight = Math.max(1, sheetHeight - padding.top - padding.bottom);
  const contentLimit = Math.max(padding.top, height - padding.bottom);
  const article = [...page.children].find((child) => child.tagName === "ARTICLE");
  const blocks = article ? [...article.children]
    .map((child) => relativeRect(child, bounds))
    .filter((rect) => rect.height > 0) : [];
  const tables = [...(article?.querySelectorAll("table") || [])].map((table) => {
    const rect = relativeRect(table, bounds);
    const rows = [...table.querySelectorAll("tr")].map((row) => relativeRect(row, bounds));
    blocks.push(...rows.filter((row) => row.height > 0));
    return { rect, rows, header: rows[0] || null };
  });
  const header = [...page.children].find((child) => child.tagName === "HEADER") || null;
  const footer = [...page.children].find((child) => child.tagName === "FOOTER") || null;
  const sheets = [];
  let offset = 0;

  while (sheets.length < 1000) {
    const start = padding.top + offset;
    const table = tables.find((candidate) =>
      start >= candidate.rect.top - 0.5 && start < candidate.rect.bottom - 0.5);
    const repeatHeader = Boolean(table?.header && start > table.rect.top + 0.5);
    const headerHeight = repeatHeader ? table.header.height : 0;
    const availableHeight = Math.max(1, bodyHeight - headerHeight);
    const desiredEnd = Math.min(contentLimit, start + availableHeight);
    const end = start >= contentLimit - 0.5
      ? start
      : Math.max(start + 1, alignContentEnd(start, desiredEnd, blocks));
    sheets.push({ start, end, repeatHeader, table, headerHeight });
    if (end >= contentLimit - 0.5) break;
    offset = end - padding.top;
  }
  if (sheets.length >= 1000) throw new Error("The Word document is too long to paginate.");

  return {
    page,
    sectionIndex,
    width,
    height,
    sheetHeight,
    padding,
    bodyHeight,
    contentLimit,
    bounds,
    header,
    footer,
    headerTemplate: header?.innerHTML || "",
    footerTemplate: footer?.innerHTML || "",
    headerRect: header ? relativeRect(header, bounds) : null,
    footerRect: footer ? relativeRect(footer, bounds) : null,
    headerMarginTop: header ? parseFloat(getComputedStyle(header).marginTop) || 0 : 0,
    footerMarginBottom: footer ? parseFloat(getComputedStyle(footer).marginBottom) || 0 : 0,
    sheets,
  };
}

function applyFieldValues(element, values) {
  if (!element) return;
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  const replacements = {
    "[[ES_PAGE]]": String(values.page),
    "[[ES_NUMPAGES]]": String(values.totalPages),
    "[[ES_SECTIONPAGES]]": String(values.sectionPages),
    "[[ES_SECTION]]": String(values.section),
  };
  let node = walker.nextNode();
  while (node) {
    const text = node.nodeValue;
    const updated = Object.entries(replacements)
      .reduce((value, [marker, replacement]) => value.split(marker).join(replacement), text);
    if (updated !== text) node.nodeValue = updated;
    node = walker.nextNode();
  }
}

function drawRegion(context, canvas, source, destination, scale) {
  const sourceY = Math.max(0, Math.round(source.y * scale));
  const sourceHeight = Math.max(0, Math.min(Math.round(source.height * scale), canvas.height - sourceY));
  if (!sourceHeight) return;
  context.drawImage(
    canvas,
    0, sourceY, canvas.width, sourceHeight,
    0, Math.round(destination.y * scale), canvas.width, sourceHeight,
  );
}

async function prepareDocxSections(bytes, JSZip) {
  if (!JSZip) return [{ bytes, sectionIndex: 1, explicitPageStart: null }];
  const zip = await JSZip.loadAsync(bytes);
  const documentPart = zip.file("word/document.xml");
  if (!documentPart) return [{ bytes, sectionIndex: 1, explicitPageStart: null }];

  const xml = await documentPart.async("string");
  const root = new DOMParser().parseFromString(xml, "application/xml").documentElement;
  if (!root || root.localName !== "document") {
    throw new Error("The Word document package is invalid.");
  }
  const body = wordChild(root, "body");
  if (!body) throw new Error("The Word document package is invalid.");
  await ensureDefaultFontSize(zip);
  await prepareFieldMarkers(zip);
  const finalSection = wordChild(body, "sectPr");
  const sections = [];
  let children = [];
  let previousSection = null;

  for (const child of [...body.children]) {
    if (child === finalSection) continue;
    const properties = wordChild(child, "pPr");
    const section = wordChild(properties, "sectPr");
    if (!section) {
      children.push(child);
      continue;
    }

    const paragraph = child.cloneNode(true);
    const paragraphProperties = wordChild(paragraph, "pPr");
    paragraphProperties.removeChild(wordChild(paragraphProperties, "sectPr"));
    if (!paragraphProperties.children.length && !paragraphProperties.attributes.length) {
      paragraph.removeChild(paragraphProperties);
    }
    const hasContent = [
      "t", "delText", "instrText", "drawing", "pict", "object", "br", "tab", "cr",
      "sym", "noBreakHyphen", "softHyphen", "footnoteReference", "endnoteReference", "ffData",
    ].some((name) => wordDescendants(paragraph, name).length);
    if (hasContent) children.push(paragraph);
    previousSection = inheritSectionReferences(section, previousSection);
    sections.push({
      children,
      section: previousSection,
      explicitPageStart: explicitPageStart(previousSection),
    });
    children = [];
  }

  if (!sections.length) {
    return [{
      bytes: await generateDocxPackage(zip),
      sectionIndex: 1,
      explicitPageStart: explicitPageStart(finalSection),
    }];
  }
  previousSection = finalSection
    ? inheritSectionReferences(finalSection, previousSection)
    : previousSection;
  sections.push({
    children,
    section: previousSection,
    explicitPageStart: explicitPageStart(previousSection),
  });

  const groups = [];
  for (const definition of sections) {
    const current = groups[groups.length - 1];
    if (current && sectionStartType(definition.section) === "continuous") {
      current.push(definition);
    } else {
      groups.push([definition]);
    }
  }

  const output = [];
  for (let index = 0; index < groups.length; index += 1) {
    const group = groups[index];
    const definition = group[group.length - 1];
    const sectionRoot = root.cloneNode(true);
    const sectionBody = wordChild(sectionRoot, "body");
    const previewSection = definition.section?.cloneNode(true);
    const sectionType = wordChild(previewSection, "type");
    if (sectionType) previewSection.removeChild(sectionType);
    sectionBody.replaceChildren(
      ...group.flatMap((item) => item.children.map((child) => child.cloneNode(true))),
      ...(previewSection ? [previewSection] : []),
    );
    zip.file("word/document.xml", new XMLSerializer().serializeToString(sectionRoot));
    output.push({
      bytes: await generateDocxPackage(zip),
      sectionIndex: index + 1,
      explicitPageStart: group.map((item) => item.explicitPageStart).find((value) => value !== null) ?? null,
    });
  }
  return output;
}

export async function convertDocxToPdf(data, report = () => {}, libraries) {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (!bytes.byteLength) throw new Error("The Word file is empty.");

  const container = document.createElement("div");
  const styleContainer = document.createElement("div");
  styleContainer.hidden = true;
  container.style.cssText = "position:fixed;left:-10000px;top:0;background:#ffffff;";
  document.body.append(styleContainer);
  document.body.append(container);

  const temporaryContainers = [];

  try {
    report(10, "Opening the Word document");
    const previewDocuments = await prepareDocxSections(bytes, libraries.JSZip);
    const renderedSections = [];
    for (let index = 0; index < previewDocuments.length; index += 1) {
      const sectionContainer = document.createElement("div");
      const sectionStyleContainer = document.createElement("div");
      sectionStyleContainer.hidden = true;
      document.body.append(sectionStyleContainer, sectionContainer);
      temporaryContainers.push(sectionContainer, sectionStyleContainer);
      await libraries.docx.renderAsync(
        new Blob([previewDocuments[index].bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
        sectionContainer,
        sectionStyleContainer,
        {
          className: "docx-page",
          inWrapper: false,
          breakPages: true,
          ignoreLastRenderedPageBreak: false,
          renderHeaders: true,
          renderFooters: true,
          renderFootnotes: true,
          renderEndnotes: true,
        },
      );
      const sectionPages = [...sectionContainer.querySelectorAll("section.docx-page")];
      for (const page of sectionPages) normalizeParagraphTypography(page);
      container.append(...sectionPages);
      styleContainer.append(...sectionStyleContainer.childNodes);
      renderedSections.push({ definition: previewDocuments[index], pages: sectionPages });
      sectionContainer.remove();
      sectionStyleContainer.remove();
    }

    const pages = [...container.querySelectorAll("section.docx-page")];
    if (!pages.length) throw new Error("The Word document has no pages.");

    const layouts = [];
    const sectionCounts = previewDocuments.map(() => 0);
    const sectionPageIndices = previewDocuments.map(() => 0);
    for (const rendered of renderedSections) {
      for (const page of rendered.pages) {
        const layout = buildPageLayout(page, rendered.definition.sectionIndex);
        layout.sectionPageIndex = sectionPageIndices[layout.sectionIndex - 1]++;
        sectionCounts[layout.sectionIndex - 1] += layout.sheets.length;
        layouts.push(layout);
      }
    }

    let nextPageNumber = 1;
    for (let index = 0; index < previewDocuments.length; index += 1) {
      const definition = previewDocuments[index];
      definition.pageNumberStart = definition.explicitPageStart ?? nextPageNumber;
      nextPageNumber = definition.pageNumberStart + Math.max(1, sectionCounts[index]);
    }

    const total = layouts.reduce((count, layout) => count + layout.sheets.length, 0);
    const pdf = await libraries.PDFLib.PDFDocument.create();
    let outputIndex = 0;

    for (const layout of layouts) {
      const { page, width, height, sheetHeight, padding, header, footer } = layout;
      const scale = Math.max(0.75, Math.min(2, 3600 / Math.max(width, height)));
      const originalHeaderVisibility = header?.style.visibility;
      const originalFooterVisibility = footer?.style.visibility;
      if (header) header.style.visibility = "hidden";
      if (footer) footer.style.visibility = "hidden";
      const canvas = await libraries.html2canvas(page, {
        backgroundColor: "#ffffff",
        scale,
        logging: false,
        useCORS: true,
      });
      if (header) header.style.visibility = originalHeaderVisibility || "";
      if (footer) footer.style.visibility = originalFooterVisibility || "";
      const definition = previewDocuments[layout.sectionIndex - 1];

      for (let sheetIndex = 0; sheetIndex < layout.sheets.length; sheetIndex += 1) {
        const sheet = layout.sheets[sheetIndex];
        const values = {
          page: definition.pageNumberStart + layout.sectionPageIndex + sheetIndex,
          totalPages: total,
          sectionPages: sectionCounts[layout.sectionIndex - 1],
          section: layout.sectionIndex,
        };
        report(20 + Math.round((outputIndex / total) * 70), `Rendering page ${outputIndex + 1} of ${total}`);

        const slice = document.createElement("canvas");
        slice.width = canvas.width;
        slice.height = Math.max(1, Math.round(sheetHeight * scale));
        const context = slice.getContext("2d");
        context.fillStyle = "#ffffff";
        context.fillRect(0, 0, slice.width, slice.height);
        drawRegion(context, canvas, { y: 0, height: padding.top }, { y: 0 }, scale);

        let bodyDestination = padding.top;
        if (sheet.repeatHeader && sheet.table?.header) {
          drawRegion(context, canvas, {
            y: sheet.table.header.top,
            height: sheet.table.header.height,
          }, { y: bodyDestination }, scale);
          bodyDestination += sheet.headerHeight;
        }
        drawRegion(context, canvas, {
          y: sheet.start,
          height: Math.max(0, sheet.end - sheet.start),
        }, { y: bodyDestination }, scale);

        if (header) {
          header.innerHTML = layout.headerTemplate;
          applyFieldValues(header, values);
          const headerCanvas = await libraries.html2canvas(header, {
            backgroundColor: null,
            scale,
            logging: false,
            useCORS: true,
          });
          context.drawImage(
            headerCanvas,
            Math.round(layout.headerRect.left * scale),
            Math.round((padding.top + layout.headerMarginTop) * scale),
          );
          headerCanvas.width = 1;
          headerCanvas.height = 1;
        }
        if (footer) {
          footer.innerHTML = layout.footerTemplate;
          applyFieldValues(footer, values);
          const footerCanvas = await libraries.html2canvas(footer, {
            backgroundColor: null,
            scale,
            logging: false,
            useCORS: true,
          });
          const footerTop = layout.footerMarginBottom
            ? sheetHeight + layout.footerMarginBottom - layout.footerRect.height
            : sheetHeight - padding.bottom - layout.footerRect.height;
          context.drawImage(
            footerCanvas,
            Math.round(layout.footerRect.left * scale),
            Math.round(footerTop * scale),
          );
          footerCanvas.width = 1;
          footerCanvas.height = 1;
        }

        const imageBlob = await new Promise((resolve) => slice.toBlob(resolve, "image/jpeg", 0.95));
        slice.width = 1;
        slice.height = 1;
        if (!imageBlob) throw new Error(`Page ${outputIndex + 1} could not be rendered.`);
        const imageBytes = new Uint8Array(await imageBlob.arrayBuffer());
        const image = await pdf.embedJpg(imageBytes);
        const pageWidth = width * 72 / 96;
        const pageHeight = sheetHeight * 72 / 96;
        const pdfPage = pdf.addPage([pageWidth, pageHeight]);
        pdfPage.drawImage(image, { x: 0, y: 0, width: pageWidth, height: pageHeight });
        outputIndex += 1;
      }
      canvas.width = 1;
      canvas.height = 1;
    }

    report(96, "Saving the PDF");
    const output = await pdf.save();
    report(100, "The PDF is ready");
    return new Uint8Array(output);
  } catch (error) {
    if (/Word file is empty|has no pages|no printable area|could not be rendered/.test(error.message)) {
      throw new Error(error.message);
    }
    throw new Error("This Word file could not be converted. Check that it is a valid .docx and try again.");
  } finally {
    for (const image of container.querySelectorAll("img[src^='blob:']")) {
      URL.revokeObjectURL(image.src);
    }
    container.remove();
    styleContainer.remove();
    for (const temporary of temporaryContainers) temporary.remove();
  }
}
