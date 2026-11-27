import { unzipSync, zipSync, strToU8, strFromU8 } from "fflate";

const WORD_NS = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
]);
const REL_OFFICE_NS = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);
const REL_PACKAGE_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPE_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const OLD_OFFICE_NS = "urn:schemas-microsoft-com:office:office";
const DRAWING_NS = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
]);
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const MAX_PART_SIZE = 128 * 1024 * 1024;
const MAX_PACKAGE_SIZE = 512 * 1024 * 1024;
const SKIPPED_DOCUMENT_PARTS = new Set([
  "styles", "numbering", "fontTable", "theme", "settings", "webSettings",
  "footnotes", "endnotes",
]);

class MergeError extends Error {
  constructor(message) {
    super(message);
    this.name = "Error";
    this.stack = "";
  }
}
const invalid = () => { throw new MergeError("This Word file could not be read."); };
const isElement = node => node?.kind === "element";
const localName = name => name.slice(name.indexOf(":") + 1);
const prefixName = name => name.includes(":") ? name.slice(0, name.indexOf(":")) : "";
const isWord = (node, name) => isElement(node) && WORD_NS.has(node.uri) && node.local === name;
const elements = node => node.children.filter(isElement);
const attribute = (node, name) => node.attrs.find(attr => attr.name === name)?.value;
const wordAttribute = (node, name) => node.attrs.find(attr => WORD_NS.has(attr.uri) && attr.local === name)?.value;
const makeAttribute = (name, value) => ({ name, value: String(value), uri: null, local: localName(name) });
const makeElement = (uri, local, prefix, attrs = [], children = []) => ({
  kind: "element",
  name: prefix ? `${prefix}:${local}` : local,
  uri,
  local,
  attrs: attrs.map(([name, value]) => makeAttribute(name, value)),
  children,
});
const makeRoot = (uri, local, prefix = "") => {
  const root = makeElement(uri, local, prefix);
  root.attrs.unshift(makeAttribute(prefix ? `xmlns:${prefix}` : "xmlns", uri));
  return root;
};
const setWordAttribute = (node, name, value) => {
  const attr = node.attrs.find(item => WORD_NS.has(item.uri) && item.local === name);
  if (!attr) invalid();
  attr.value = String(value);
};
const setWordValue = (node, value) => {
  if (wordAttribute(node, "val") === undefined) invalid();
  setWordAttribute(node, "val", value);
};

function decodeXml(value) {
  if (/&(?!(?:amp|lt|gt|apos|quot|#\d+|#x[\da-fA-F]+);)/.test(value)) invalid();
  return value.replace(/&(amp|lt|gt|apos|quot|#\d+|#x[\da-fA-F]+);/g, (entity, name) => {
    if (name[0] !== "#") return { amp: "&", lt: "<", gt: ">", apos: "'", quot: '"' }[name];
    const code = name[1] === "x" ? parseInt(name.slice(2), 16) : Number(name.slice(1));
    if (!Number.isSafeInteger(code) || code <= 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) invalid();
    return String.fromCodePoint(code);
  });
}

function parseXml(data) {
  const source = strFromU8(data).replace(/^\uFEFF/, "");
  if (!source) invalid();
  let offset = 0;
  let root = null;
  const stack = [];
  const namespaces = [new Map([["xml", XML_NS]])];
  const namePattern = /^[A-Za-z_][A-Za-z0-9_.-]*$/;
  const append = node => {
    if (stack.length) stack[stack.length - 1].children.push(node);
    else if (node.kind === "comment" || (node.kind === "text" && !node.text.trim())) return;
    else if (root) invalid();
    else root = node;
  };

  while (offset < source.length) {
    if (source[offset] !== "<") {
      const end = source.indexOf("<", offset);
      const next = end < 0 ? source.length : end;
      append({ kind: "text", text: decodeXml(source.slice(offset, next)) });
      offset = next;
      continue;
    }
    if (source.startsWith("<!--", offset)) {
      const end = source.indexOf("-->", offset + 4);
      if (end < 0) invalid();
      append({ kind: "comment", text: source.slice(offset + 4, end) });
      offset = end + 3;
      continue;
    }
    if (source.startsWith("<![CDATA[", offset)) {
      const end = source.indexOf("]]>", offset + 9);
      if (end < 0 || !stack.length) invalid();
      stack[stack.length - 1].children.push({ kind: "text", text: source.slice(offset + 9, end) });
      offset = end + 3;
      continue;
    }
    if (source.startsWith("<?", offset)) {
      const end = source.indexOf("?>", offset + 2);
      if (end < 0) invalid();
      offset = end + 2;
      continue;
    }
    if (source.startsWith("<!", offset)) invalid();

    let end = offset + 1;
    let quote = "";
    for (; end < source.length; end++) {
      const char = source[end];
      if (quote) {
        if (char === quote) quote = "";
      } else if (char === '"' || char === "'") quote = char;
      else if (char === ">") break;
    }
    if (end === source.length || quote) invalid();
    let raw = source.slice(offset + 1, end).trim();
    offset = end + 1;
    if (raw.startsWith("/")) {
      const name = raw.slice(1).trim();
      if (!stack.length || stack[stack.length - 1].name !== name) invalid();
      stack.pop();
      namespaces.pop();
      continue;
    }

    const selfClosing = raw.endsWith("/");
    if (selfClosing) raw = raw.slice(0, -1).trimEnd();
    const match = /^([^\s/>]+)([\s\S]*)$/.exec(raw);
    if (!match) invalid();
    const name = match[1];
    const nameParts = name.split(":");
    if (nameParts.length > 2 || nameParts.some(part => !namePattern.test(part))) invalid();
    const rawAttrs = [];
    const seen = new Set();
    let rest = match[2];
    while (rest.trim()) {
      const attrMatch = /^\s*([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/.exec(rest);
      if (!attrMatch || seen.has(attrMatch[1])) invalid();
      const attrName = attrMatch[1];
      const attrParts = attrName.split(":");
      if (attrParts.length > 2 || attrParts.some(part => !namePattern.test(part))) invalid();
      const value = attrMatch[2] ?? attrMatch[3] ?? "";
      if (value.includes("<")) invalid();
      rawAttrs.push({ name: attrName, value: decodeXml(value) });
      seen.add(attrName);
      rest = rest.slice(attrMatch[0].length);
    }

    const scope = new Map(namespaces[namespaces.length - 1]);
    for (const attr of rawAttrs) {
      if (attr.name !== "xmlns" && !attr.name.startsWith("xmlns:")) continue;
      const prefix = attr.name === "xmlns" ? "" : attr.name.slice(6);
      if (prefix === "xmlns" || (prefix === "xml" && attr.value !== XML_NS)) invalid();
      scope.set(prefix, attr.value);
    }
    const prefix = nameParts.length === 2 ? nameParts[0] : "";
    const uri = prefix ? scope.get(prefix) : scope.get("") || null;
    if (prefix && !uri) invalid();
    const node = { kind: "element", name, uri, local: nameParts[1] ?? name, attrs: [], children: [] };
    for (const attr of rawAttrs) {
      const parts = attr.name.split(":");
      const attrPrefix = parts.length === 2 ? parts[0] : "";
      const attrUri = attr.name === "xmlns" || attr.name.startsWith("xmlns:")
        ? null
        : attrPrefix ? scope.get(attrPrefix) : null;
      if (attr.name !== "xmlns" && !attr.name.startsWith("xmlns:") && attrPrefix && !attrUri) invalid();
      node.attrs.push({ ...attr, uri: attrUri, local: parts[1] ?? attr.name });
    }
    append(node);
    if (stack.length > 1024) invalid();
    if (!selfClosing) {
      stack.push(node);
      namespaces.push(scope);
    }
  }
  if (stack.length || !isElement(root)) invalid();
  return root;
}

const escapeText = value => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r/g, "&#13;");
const escapeAttribute = value => escapeText(value).replace(/"/g, "&quot;").replace(/\n/g, "&#10;").replace(/\t/g, "&#9;");
function serializeXml(root) {
  const output = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'];
  const write = node => {
    if (node.kind === "text") output.push(escapeText(node.text));
    else if (node.kind === "comment") output.push(`<!--${node.text}-->`);
    else {
      const attrs = node.attrs.map(attr => ` ${attr.name}="${escapeAttribute(attr.value)}"`).join("");
      if (!node.children.length) output.push(`<${node.name}${attrs}/>`);
      else {
        output.push(`<${node.name}${attrs}>`);
        node.children.forEach(write);
        output.push(`</${node.name}>`);
      }
    }
  };
  write(root);
  return strToU8(output.join(""));
}

function namespaceMap(root) {
  const map = new Map([["xml", XML_NS]]);
  for (const attr of root.attrs) {
    if (attr.name === "xmlns") map.set("", attr.value);
    else if (attr.name.startsWith("xmlns:")) map.set(attr.name.slice(6), attr.value);
  }
  return map;
}

function sourceNamespaces(root) {
  const byPrefix = new Map([["xml", XML_NS]]);
  const byUri = new Map([[XML_NS, "xml"]]);
  const pending = [root];
  while (pending.length) {
    const node = pending.pop();
    if (!isElement(node)) continue;
    for (const attr of node.attrs) {
      if (attr.name !== "xmlns" && !attr.name.startsWith("xmlns:")) continue;
      const prefix = attr.name === "xmlns" ? "" : attr.name.slice(6);
      if (!byPrefix.has(prefix)) byPrefix.set(prefix, attr.value);
      if (attr.value && !byUri.has(attr.value)) byUri.set(attr.value, prefix);
    }
    for (const child of node.children) pending.push(child);
  }
  return { byPrefix, byUri };
}

function adoptNodes(nodes, sourceRoot, targetRoot) {
  const source = sourceNamespaces(sourceRoot);
  const target = namespaceMap(targetRoot);
  const targetByUri = new Map();
  for (const [prefix, uri] of target) if (uri && prefix && !targetByUri.has(uri)) targetByUri.set(uri, prefix);
  let serial = 0;
  for (const [prefix, uri] of source.byPrefix) {
    if (!uri || uri === XML_NS || targetByUri.has(uri)) continue;
    let adopted = prefix;
    if (!prefix || target.has(prefix)) {
      do adopted = `m${++serial}`; while (target.has(adopted));
    }
    target.set(adopted, uri);
    targetByUri.set(uri, adopted);
    targetRoot.attrs.push(makeAttribute(`xmlns:${adopted}`, uri));
  }
  const rename = (name, uri) => {
    if (!uri || uri === XML_NS) return name;
    const prefix = targetByUri.get(uri);
    return prefix === undefined ? name : `${prefix}:${localName(name)}`;
  };
  const renameReferences = value => value.split(/(\s+)/).map(token => {
    if (!token || /^\s+$/.test(token)) return token;
    const colon = token.indexOf(":");
    const prefix = colon < 0 ? "" : token.slice(0, colon);
    const adopted = targetByUri.get(source.byPrefix.get(prefix));
    if (adopted === undefined) return token;
    const local = colon < 0 ? token : token.slice(colon + 1);
    return adopted ? `${adopted}:${local}` : local;
  }).join("");
  // mc:Ignorable contains namespace prefixes, so its value must follow adopted prefix names.
  const ignorable = sourceRoot.attrs.find(attr => attr.uri === MC_NS && attr.local === "Ignorable");
  if (ignorable) {
    const value = renameReferences(ignorable.value);
    const existing = targetRoot.attrs.find(attr => attr.uri === MC_NS && attr.local === "Ignorable");
    if (existing) existing.value = [...new Set(`${existing.value} ${value}`.trim().split(/\s+/))].join(" ");
    else targetRoot.attrs.push({ ...ignorable, name: rename(ignorable.name, ignorable.uri), value });
  }
  const clone = node => {
    if (!isElement(node)) return { ...node };
    return {
      ...node,
      name: rename(node.name, node.uri),
      attrs: node.attrs.filter(attr => attr.name !== "xmlns" && !attr.name.startsWith("xmlns:")).map(attr => ({
        ...attr,
        name: rename(attr.name, attr.uri),
        value: attr.uri === MC_NS && (attr.local === "Ignorable" || attr.local === "Requires")
          ? renameReferences(attr.value)
          : attr.value,
      })),
      children: node.children.map(clone),
    };
  };
  return nodes.map(clone);
}

function* walk(nodes) {
  const pending = [...nodes].reverse();
  while (pending.length) {
    const node = pending.pop();
    if (!isElement(node)) continue;
    yield node;
    for (let index = node.children.length - 1; index >= 0; index--) pending.push(node.children[index]);
  }
}

function remapDrawingIds(output) {
  let nextId = 1;
  for (const [part, data] of Object.entries(output)) {
    if (!part.endsWith(".xml") || !strFromU8(data).includes("docPr")) continue;
    const root = parseXml(data);
    let changed = false;
    for (const node of walk([root])) {
      if (!DRAWING_NS.has(node.uri) || node.local !== "docPr") continue;
      const id = node.attrs.find(attr => attr.name === "id" && !attr.uri);
      if (!id || !/^\d+$/.test(id.value) || Number(id.value) > 0x7fffffff) invalid();
      id.value = String(nextId++);
      changed = true;
    }
    if (changed) output[part] = serializeXml(root);
  }
}

function safePartName(name) {
  if (typeof name !== "string" || !name || name.startsWith("/") || /[\\\0]/.test(name)) return false;
  const clean = name.endsWith("/") ? name.slice(0, -1) : name;
  return Boolean(clean) && clean.split("/").every(part => part && part !== "." && part !== "..");
}

function relationshipsName(part) {
  const slash = part.lastIndexOf("/");
  return `${slash < 0 ? "" : part.slice(0, slash + 1)}_rels/${slash < 0 ? part : part.slice(slash + 1)}.rels`;
}

function resolvePart(source, target) {
  const hash = target.indexOf("#");
  const path = hash < 0 ? target : target.slice(0, hash);
  if (!path || /[\\\0]/.test(path)) invalid();
  const parts = path.startsWith("/") ? [] : source ? source.split("/").slice(0, -1) : [];
  for (const part of path.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") {
      if (!parts.length) invalid();
      parts.pop();
    } else parts.push(part);
  }
  const result = parts.join("/");
  if (!safePartName(result)) invalid();
  return result;
}

function relativePart(source, target) {
  const from = source ? source.split("/").slice(0, -1) : [];
  const to = target.split("/");
  let common = 0;
  while (common < from.length && common < to.length && from[common] === to[common]) common++;
  return `${"../".repeat(from.length - common)}${to.slice(common).join("/")}`;
}

function splitTarget(target) {
  const hash = target.indexOf("#");
  return hash < 0 ? { path: target, fragment: "" } : { path: target.slice(0, hash), fragment: target.slice(hash) };
}

function parseRelationships(files, part) {
  const data = files[relationshipsName(part)];
  if (!data) return null;
  const root = parseXml(data);
  if (root.uri !== REL_PACKAGE_NS || root.local !== "Relationships") invalid();
  const list = [];
  const byId = new Map();
  let nextId = 1;
  for (const node of elements(root)) {
    if (node.uri !== REL_PACKAGE_NS || node.local !== "Relationship") invalid();
    const id = attribute(node, "Id");
    const type = attribute(node, "Type");
    const target = attribute(node, "Target");
    const mode = attribute(node, "TargetMode");
    if (!id || !type || !target || byId.has(id) || (mode !== undefined && mode !== "External")) invalid();
    const numeric = /^rId(\d+)$/.exec(id);
    if (numeric) nextId = Math.max(nextId, Number(numeric[1]) + 1);
    const rel = { id, type, target, external: mode === "External", node };
    list.push(rel);
    byId.set(id, rel);
  }
  if (!Number.isSafeInteger(nextId)) invalid();
  return { root, list, byId, nextId };
}

function emptyRelationships() {
  return { root: makeRoot(REL_PACKAGE_NS, "Relationships"), list: [], byId: new Map(), nextId: 1 };
}

function addRelationship(state, type, target, external, requestedId) {
  let id = requestedId;
  if (!id) {
    do id = `rId${state.nextId++}`; while (state.byId.has(id));
  }
  if (state.byId.has(id) || !Number.isSafeInteger(state.nextId)) invalid();
  const attrs = [["Id", id], ["Type", type], ["Target", target]];
  if (external) attrs.push(["TargetMode", "External"]);
  const node = makeElement(REL_PACKAGE_NS, "Relationship", prefixName(state.root.name), attrs);
  state.root.children.push(node);
  const rel = { id, type, target, external, node };
  state.list.push(rel);
  state.byId.set(id, rel);
  return id;
}

function parseContentTypes(files) {
  const root = parseXml(files["[Content_Types].xml"]);
  if (root.uri !== CONTENT_TYPE_NS || root.local !== "Types") invalid();
  const overrides = new Map();
  const defaults = new Map();
  for (const node of elements(root)) {
    if (node.uri !== CONTENT_TYPE_NS) invalid();
    if (node.local === "Override") {
      const part = attribute(node, "PartName");
      const type = attribute(node, "ContentType");
      if (!part?.startsWith("/") || !safePartName(part.slice(1)) || !type || overrides.has(part)) invalid();
      overrides.set(part, type);
    } else if (node.local === "Default") {
      const extension = attribute(node, "Extension")?.toLowerCase();
      const type = attribute(node, "ContentType");
      if (!extension || !type || defaults.has(extension)) invalid();
      defaults.set(extension, type);
    } else invalid();
  }
  return { root, overrides, defaults };
}

function contentType(state, part) {
  return state.overrides.get(`/${part}`) ?? state.defaults.get(part.split(".").pop()?.toLowerCase());
}

function insertChild(root, node, before) {
  const index = root.children.findIndex(child => isElement(child) && before.includes(child.local));
  root.children.splice(index < 0 ? root.children.length : index, 0, node);
}

function addContentType(state, part, type) {
  const name = `/${part}`;
  if (state.overrides.has(name)) {
    if (state.overrides.get(name) !== type) invalid();
    return;
  }
  const node = makeElement(CONTENT_TYPE_NS, "Override", prefixName(state.root.name), [["PartName", name], ["ContentType", type]]);
  state.root.children.push(node);
  state.overrides.set(name, type);
}

function ensureRelationshipsContentType(state) {
  const type = "application/vnd.openxmlformats-package.relationships+xml";
  if (state.defaults.has("rels")) {
    if (state.defaults.get("rels") !== type) invalid();
    return;
  }
  const node = makeElement(CONTENT_TYPE_NS, "Default", prefixName(state.root.name), [["Extension", "rels"], ["ContentType", type]]);
  insertChild(state.root, node, ["Override"]);
  state.defaults.set("rels", type);
}

function parseWordRoot(files, part, local) {
  const data = files[part];
  if (!data) invalid();
  const root = parseXml(data);
  if (!isWord(root, local)) invalid();
  return root;
}

function relatedPart(pkg, name) {
  const rel = pkg.mainRels.list.find(item => !item.external && item.type.endsWith(`/${name}`));
  return rel ? resolvePart(pkg.mainPart, rel.target) : null;
}

function readPackage(input) {
  const data = input?.data;
  const bytes = data instanceof Uint8Array ? data : data instanceof ArrayBuffer ? new Uint8Array(data) : null;
  if (!bytes?.byteLength || bytes.byteLength > MAX_PACKAGE_SIZE) invalid();
  let total = 0;
  let rejected = false;
  const files = unzipSync(bytes, {
    filter: file => {
      const directory = file.name.endsWith("/");
      if (!safePartName(file.name) || !Number.isSafeInteger(file.originalSize) || file.originalSize < 0 || (!directory && file.originalSize > MAX_PART_SIZE)) {
        rejected = true;
        return false;
      }
      total += file.originalSize;
      if (total > MAX_PACKAGE_SIZE) rejected = true;
      return !rejected;
    },
  });
  if (rejected || total > MAX_PACKAGE_SIZE) invalid();
  for (const name of Object.keys(files)) {
    if (!safePartName(name)) invalid();
    if (name.endsWith("/")) delete files[name];
  }
  const contentTypes = parseContentTypes(files);
  for (const part of Object.keys(files)) if (part !== "[Content_Types].xml" && !contentType(contentTypes, part)) invalid();
  const packageRels = parseRelationships(files, "");
  const officeRels = packageRels?.list.filter(rel => !rel.external && rel.type.endsWith("/officeDocument")) ?? [];
  if (officeRels.length !== 1) invalid();
  const mainPart = resolvePart("", officeRels[0].target);
  const root = parseWordRoot(files, mainPart, "document");
  const bodies = elements(root).filter(node => isWord(node, "body"));
  if (bodies.length !== 1) invalid();
  const mainRels = parseRelationships(files, mainPart) ?? emptyRelationships();
  return {
    files,
    contentTypes,
    mainPart,
    root,
    body: bodies[0],
    mainRels,
    officeBase: officeRels[0].type.slice(0, -"officeDocument".length),
  };
}

function copyPart(env, ctx, sourcePart) {
  if (ctx.copied.has(sourcePart)) return ctx.copied.get(sourcePart);
  const data = ctx.source.files[sourcePart];
  if (!data) invalid();
  const destination = `${ctx.prefix}/${sourcePart}`;
  const type = sourcePart.endsWith(".rels") ? null : contentType(ctx.source.contentTypes, sourcePart);
  if (!sourcePart.endsWith(".rels") && !type) invalid();
  ctx.copied.set(sourcePart, destination);
  if (ctx.styles && (type?.endsWith("header+xml") || type?.endsWith("footer+xml"))) {
    const root = parseXml(data);
    remapImportedNodes([root], ctx.styles, ctx.numbering);
    env.output[destination] = serializeXml(root);
  } else env.output[destination] = data;
  const sourceRels = parseRelationships(ctx.source.files, sourcePart);
  if (sourceRels) {
    const copiedRels = emptyRelationships();
    for (const rel of sourceRels.list) {
      let target = rel.target;
      if (!rel.external) {
        const { path, fragment } = splitTarget(rel.target);
        const copied = path ? copyPart(env, ctx, resolvePart(sourcePart, path)) : destination;
        target = `${relativePart(destination, copied)}${fragment}`;
      }
      addRelationship(copiedRels, rel.type, target, rel.external, rel.id);
    }
    env.output[relationshipsName(destination)] = serializeXml(copiedRels.root);
    ensureRelationshipsContentType(env.contentTypes);
  }
  if (!destination.endsWith(".rels")) {
    addContentType(env.contentTypes, destination, type);
  }
  return destination;
}

function targetRelationships(env, part) {
  if (env.partRels.has(part)) return env.partRels.get(part);
  const state = parseRelationships(env.output, part) ?? emptyRelationships();
  state.part = part;
  env.partRels.set(part, state);
  ensureRelationshipsContentType(env.contentTypes);
  return state;
}

function mergePartRelationships(env, ctx, sourcePart, targetPart, nodes) {
  const sourceRels = parseRelationships(ctx.source.files, sourcePart);
  if (!sourceRels) return;
  const target = targetRelationships(env, targetPart);
  const ids = new Map();
  for (const rel of sourceRels.list) {
    let destination = rel.target;
    if (!rel.external) {
      const { path, fragment } = splitTarget(rel.target);
      const copied = path ? copyPart(env, ctx, resolvePart(sourcePart, path)) : targetPart;
      destination = `${relativePart(targetPart, copied)}${fragment}`;
    }
    ids.set(rel.id, addRelationship(target, rel.type, destination, rel.external));
  }
  remapRelationshipAttributes(nodes, ids);
}

function importDocumentRelationships(env, ctx) {
  const ids = new Map();
  for (const rel of ctx.source.mainRels.list) {
    const kind = rel.type.slice(rel.type.lastIndexOf("/") + 1);
    if (SKIPPED_DOCUMENT_PARTS.has(kind)) continue;
    let target = rel.target;
    if (!rel.external) {
      const { path, fragment } = splitTarget(rel.target);
      const copied = path ? copyPart(env, ctx, resolvePart(ctx.source.mainPart, path)) : env.base.mainPart;
      target = `${relativePart(env.base.mainPart, copied)}${fragment}`;
    }
    ids.set(rel.id, addRelationship(env.mainRels, rel.type, target, rel.external));
  }
  return ids;
}

function remapRelationshipAttributes(nodes, ids) {
  for (const node of walk(nodes)) {
    for (const attr of node.attrs) {
      const relationshipAttribute = (REL_OFFICE_NS.has(attr.uri) && ["id", "embed", "link"].includes(attr.local))
        || (attr.uri === OLD_OFFICE_NS && attr.local === "relid");
      if (!relationshipAttribute || !attr.value) continue;
      const id = ids.get(attr.value);
      if (!id) invalid();
      attr.value = id;
    }
  }
}

function remapImportedNodes(nodes, styles, numbering) {
  for (const node of walk(nodes)) {
    if (!WORD_NS.has(node.uri)) continue;
    if (["pStyle", "rStyle", "tblStyle", "basedOn", "next", "link", "styleLink", "numStyleLink"].includes(node.local)) {
      const value = wordAttribute(node, "val");
      if (value !== undefined && styles.has(value)) setWordValue(node, styles.get(value));
    } else if (node.local === "numId") {
      const value = numericAttribute(node, "val");
      if (value !== 0) {
        const key = String(value);
        if (!numbering.has(key)) invalid();
        setWordValue(node, numbering.get(key));
      }
    }
  }
}

function mergeNotes(env, ctx, styles, numbering) {
  const maps = { footnotes: new Map(), endnotes: new Map() };
  for (const kind of ["footnotes", "endnotes"]) {
    const sourcePart = relatedPart(ctx.source, kind);
    if (!sourcePart) continue;
    const sourceRoot = parseWordRoot(ctx.source.files, sourcePart, kind);
    const sourceNotes = elements(sourceRoot).filter(node => isWord(node, kind === "footnotes" ? "footnote" : "endnote"));
    if (!sourceNotes.length) continue;
    const target = ensureTargetPart(env, kind, ctx.source, sourcePart);
    const targetNotes = elements(target.root).filter(node => isWord(node, kind === "footnotes" ? "footnote" : "endnote"));
    const usedIds = new Set(targetNotes.map(node => String(signedNumericAttribute(node, "id"))));
    const specialTypes = new Set(targetNotes.map(node => wordAttribute(node, "type")).filter(Boolean));
    const idMap = maps[kind];
    const imported = [];
    let nextId = 1;
    for (const value of usedIds) {
      const numeric = Number(value);
      if (numeric >= nextId) nextId = numeric + 1;
    }

    for (const note of sourceNotes) {
      const type = wordAttribute(note, "type");
      const oldId = signedNumericAttribute(note, "id");
      if (type === "separator" || type === "continuationSeparator") {
        if (specialTypes.has(type)) continue;
        const clone = adoptNodes([note], sourceRoot, target.root)[0];
        const specialId = String(oldId);
        if (usedIds.has(specialId)) invalid();
        usedIds.add(specialId);
        specialTypes.add(type);
        imported.push(clone);
        continue;
      }
      if (oldId <= 0 || idMap.has(String(oldId)) || !Number.isSafeInteger(nextId)) invalid();
      const freshId = String(nextId++);
      usedIds.add(freshId);
      idMap.set(String(oldId), freshId);
      const clone = adoptNodes([note], sourceRoot, target.root)[0];
      setWordAttribute(clone, "id", freshId);
      imported.push(clone);
    }
    remapImportedNodes(imported, styles, numbering);
    mergePartRelationships(env, ctx, sourcePart, target.part, imported);
    target.root.children.push(...imported);
  }
  return maps;
}

function remapNoteReferences(nodes, maps) {
  for (const node of walk(nodes)) {
    if (!WORD_NS.has(node.uri)) continue;
    const kind = node.local === "footnoteReference" ? "footnotes"
      : node.local === "endnoteReference" ? "endnotes"
        : null;
    if (!kind) continue;
    const oldId = String(numericAttribute(node, "id"));
    const freshId = maps[kind].get(oldId);
    if (!freshId) invalid();
    setWordAttribute(node, "id", freshId);
  }
}

function bookmarkState(body) {
  const names = new Set();
  const ids = new Set();
  for (const node of walk(elements(body))) {
    if (!isWord(node, "bookmarkStart")) continue;
    const name = wordAttribute(node, "name");
    const id = String(numericAttribute(node, "id"));
    if (!name || names.has(name) || ids.has(id)) invalid();
    names.add(name);
    ids.add(id);
  }
  return { names, ids };
}

function remapBookmarks(nodes, state, mergeIndex) {
  const starts = [...walk(nodes)].filter(node => isWord(node, "bookmarkStart"));
  const names = new Map();
  const ids = new Map();
  let maximum = 0;
  for (const value of state.ids) maximum = Math.max(maximum, Number(value));
  let nextId = maximum + 1;
  for (const start of starts) {
    const oldName = wordAttribute(start, "name");
    const oldId = String(numericAttribute(start, "id"));
    if (!oldName || names.has(oldName) || ids.has(oldId) || !Number.isSafeInteger(nextId)) invalid();
    let freshName = oldName;
    if (state.names.has(freshName)) {
      freshName = `${oldName}_M${mergeIndex + 1}`;
      let suffix = 2;
      while (state.names.has(freshName)) freshName = `${oldName}_M${mergeIndex + 1}_${suffix++}`;
    }
    const freshId = String(nextId++);
    names.set(oldName, freshName);
    ids.set(oldId, freshId);
    state.names.add(freshName);
    state.ids.add(freshId);
  }
  for (const node of walk(nodes)) {
    if (!WORD_NS.has(node.uri)) continue;
    if (isWord(node, "bookmarkStart")) {
      setWordAttribute(node, "name", names.get(wordAttribute(node, "name")));
      setWordAttribute(node, "id", ids.get(String(numericAttribute(node, "id"))));
    } else if (isWord(node, "bookmarkEnd")) {
      const freshId = ids.get(String(numericAttribute(node, "id")));
      if (!freshId) invalid();
      setWordAttribute(node, "id", freshId);
    } else if (isWord(node, "hyperlink")) {
      const anchor = wordAttribute(node, "anchor");
      if (anchor && names.has(anchor)) setWordAttribute(node, "anchor", names.get(anchor));
    }
  }
}

function ensureTargetPart(env, kind, source, sourcePart) {
  if (env.targets.has(kind)) return env.targets.get(kind);
  const part = `word/${kind}.xml`;
  let root;
  let created = false;
  if (env.output[part]) root = parseWordRoot(env.output, part, kind);
  else {
    root = makeRoot(env.base.body.uri, kind, "w");
    created = true;
  }
  if (created) {
    const type = contentType(source.contentTypes, sourcePart);
    if (!type) invalid();
    addContentType(env.contentTypes, part, type);
    env.output[part] = serializeXml(root);
  }
  if (!env.mainRels.list.some(rel => !rel.external && rel.type.endsWith(`/${kind}`))) {
    addRelationship(env.mainRels, `${env.officeBase}${kind}`, relativePart(env.base.mainPart, part), false);
  }
  const state = { part, root };
  env.targets.set(kind, state);
  return state;
}

function numericAttribute(node, name) {
  const value = wordAttribute(node, name);
  if (value === undefined || !/^\d+$/.test(value)) invalid();
  const number = Number(value);
  if (!Number.isSafeInteger(number)) invalid();
  return number;
}

function signedNumericAttribute(node, name) {
  const value = wordAttribute(node, name);
  if (value === undefined || !/^-?\d+$/.test(value)) invalid();
  const number = Number(value);
  if (!Number.isSafeInteger(number)) invalid();
  return number;
}

function allocateIds(nodes, attributeName, used) {
  const ids = new Map();
  let maximum = 0;
  for (const value of used) maximum = Math.max(maximum, Number(value));
  let next = maximum + 1;
  for (const node of nodes) {
    const old = numericAttribute(node, attributeName);
    if (ids.has(String(old))) invalid();
    if (!Number.isSafeInteger(next)) invalid();
    const fresh = String(next++);
    used.add(fresh);
    ids.set(String(old), fresh);
  }
  return ids;
}

function usedNumberingIds(root, local, attributeName) {
  const used = new Set();
  for (const node of elements(root).filter(item => isWord(item, local))) used.add(String(numericAttribute(node, attributeName)));
  return used;
}

function insertNumbered(root, node, before) {
  insertChild(root, node, before);
}

function mergeStyles(env, ctx, sourceRoot, sourcePart, numbering, styles) {
  const sourceStyles = elements(sourceRoot).filter(node => isWord(node, "style"));
  if (!sourceStyles.length) return;
  const target = ensureTargetPart(env, "styles", ctx.source, sourcePart);
  const defaultTypes = new Set(elements(target.root)
    .filter(node => isWord(node, "style") && wordAttribute(node, "default") !== undefined)
    .map(node => wordAttribute(node, "type")));
  const imported = [];
  for (const node of sourceStyles) {
    const old = wordAttribute(node, "styleId");
    const fresh = styles.get(old);
    if (!fresh) invalid();
    const clone = adoptNodes([node], sourceRoot, target.root)[0];
    setWordAttribute(clone, "styleId", fresh);
    const name = elements(clone).find(child => isWord(child, "name"));
    if (name) setWordValue(name, fresh);
    const styleType = wordAttribute(clone, "type");
    if (!styleType) invalid();
    if (wordAttribute(clone, "default") !== undefined) {
      if (defaultTypes.has(styleType)) {
        clone.attrs = clone.attrs.filter(attr => !(WORD_NS.has(attr.uri) && attr.local === "default"));
      } else defaultTypes.add(styleType);
    }
    imported.push(clone);
  }
  remapImportedNodes(imported, styles, numbering);
  target.root.children.push(...imported);
  mergePartRelationships(env, ctx, sourcePart, target.part, imported);
}

function mergeNumbering(env, ctx, sourceRoot, sourcePart, styles) {
  const sourceChildren = elements(sourceRoot);
  const pictures = sourceChildren.filter(node => isWord(node, "numPicBullet"));
  const abstracts = sourceChildren.filter(node => isWord(node, "abstractNum"));
  const numbers = sourceChildren.filter(node => isWord(node, "num"));
  if (!pictures.length && !abstracts.length && !numbers.length) return new Map();
  const target = ensureTargetPart(env, "numbering", ctx.source, sourcePart);
  const pictureIds = allocateIds(pictures, "numPicBulletId", usedNumberingIds(target.root, "numPicBullet", "numPicBulletId"));
  const abstractIds = allocateIds(abstracts, "abstractNumId", usedNumberingIds(target.root, "abstractNum", "abstractNumId"));
  const numberIds = allocateIds(numbers, "numId", usedNumberingIds(target.root, "num", "numId"));
  const imported = [];
  const add = (nodes, ids, attributeName, before) => {
    for (const node of nodes) {
      const clone = adoptNodes([node], sourceRoot, target.root)[0];
      setWordAttribute(clone, attributeName, ids.get(String(numericAttribute(node, attributeName))));
      insertNumbered(target.root, clone, before);
      imported.push(clone);
    }
  };
  add(pictures, pictureIds, "numPicBulletId", ["abstractNum", "num", "numIdMacAtCleanup"]);
  add(abstracts, abstractIds, "abstractNumId", ["num", "numIdMacAtCleanup"]);
  add(numbers, numberIds, "numId", ["numIdMacAtCleanup"]);
  for (const node of walk(imported)) {
    if (!WORD_NS.has(node.uri)) continue;
    if (node.local === "abstractNumId" || node.local === "lvlPicBulletId") {
      const value = String(numericAttribute(node, "val"));
      const ids = node.local === "abstractNumId" ? abstractIds : pictureIds;
      if (!ids.has(value)) invalid();
      setWordValue(node, ids.get(value));
    }
  }
  remapImportedNodes(imported, styles, numberIds);
  mergePartRelationships(env, ctx, sourcePart, target.part, imported);
  return numberIds;
}

function pageBreak(body, documentRoot) {
  const namespaces = namespaceMap(documentRoot);
  let wordPrefix;
  for (const [prefix, uri] of namespaces) {
    if (prefix && WORD_NS.has(uri)) wordPrefix = prefix;
  }
  if (!wordPrefix) {
    wordPrefix = "w";
    while (namespaces.has(wordPrefix)) wordPrefix += "_";
    documentRoot.attrs.push(makeAttribute(`xmlns:${wordPrefix}`, body.uri));
  }
  const br = makeElement(body.uri, "br", wordPrefix);
  br.attrs.push({ name: `${wordPrefix}:type`, value: "page", uri: body.uri, local: "type" });
  return makeElement(body.uri, "p", wordPrefix, [], [makeElement(body.uri, "r", wordPrefix, [], [br])]);
}

function trailingSectionIndex(body) {
  for (let index = body.children.length - 1; index >= 0; index--) {
    const node = body.children[index];
    if (node.kind === "text" && !node.text.trim()) continue;
    return isWord(node, "sectPr") ? index : -1;
  }
  return -1;
}

function appendSectionToParagraph(body, section) {
  const prefix = prefixName(body.name);
  let paragraph = body.children[body.children.length - 1];
  let properties = isWord(paragraph, "p") ? elements(paragraph).find(node => isWord(node, "pPr")) : null;
  if (!properties || elements(properties).some(node => isWord(node, "sectPr"))) {
    paragraph = makeElement(body.uri, "p", prefix);
    properties = makeElement(body.uri, "pPr", prefix);
    paragraph.children.push(properties);
    body.children.push(paragraph);
  }
  const changeIndex = properties.children.findIndex(node => isWord(node, "pPrChange"));
  properties.children.splice(changeIndex < 0 ? properties.children.length : changeIndex, 0, section);
}

function startSectionOnNewPage(section) {
  let type = elements(section).find(node => isWord(node, "type"));
  if (type) {
    setWordValue(type, "nextPage");
    return;
  }
  let prefix = prefixName(section.name);
  if (!prefix) {
    prefix = "w";
    while (section.attrs.some(attr => attr.name === `xmlns:${prefix}`)) prefix += "_";
    section.attrs.push(makeAttribute(`xmlns:${prefix}`, section.uri));
  }
  type = makeElement(section.uri, "type", prefixName(section.name));
  type.attrs.push({ name: `${prefix}:val`, value: "nextPage", uri: section.uri, local: "val" });
  insertChild(section, type, [
    "footnotePr", "endnotePr", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
    "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter",
    "docGrid", "printerSettings", "sectPrChange",
  ]);
}

function compatibilityMode(pkg) {
  const part = relatedPart(pkg, "settings");
  if (!part) return null;
  const root = parseWordRoot(pkg.files, part, "settings");
  for (const node of walk([root])) {
    if (!isWord(node, "compatSetting") || attribute(node, "name") !== "compatibilityMode") continue;
    const value = wordAttribute(node, "val");
    return value !== undefined && /^\d+$/.test(value) ? Number(value) : null;
  }
  return null;
}

function alignCompatibilityMode(output, packages) {
  const part = relatedPart(packages[0], "settings");
  if (!part) return;
  const root = parseXml(output[part]);
  if (!isWord(root, "settings")) invalid();
  let setting = null;
  for (const node of walk([root])) {
    if (isWord(node, "compatSetting") && attribute(node, "name") === "compatibilityMode") {
      setting = node;
      break;
    }
  }
  if (!setting) return;
  // A merged document has one compatibility mode, so retain the oldest source layout rules.
  const modes = packages.map(compatibilityMode).filter(value => value !== null);
  if (!modes.length) return;
  setWordValue(setting, Math.min(...modes));
  output[part] = serializeXml(root);
}

function uniquePrefix(output, index) {
  let prefix = `merge${index + 1}`;
  while (Object.keys(output).some(name => name === prefix || name.startsWith(`${prefix}/`))) prefix += "_";
  return prefix;
}

export async function mergeDocxDocuments(inputs, report = () => {}) {
  if (!Array.isArray(inputs) || inputs.length < 2) throw new MergeError("Choose at least two Word files to merge.");
  const say = message => { try { report(message); } catch {} };
  try {
    say("Reading Word files...");
    const packages = inputs.map(readPackage);
    const base = packages[0];
    const output = { ...base.files };
    const env = {
      base,
      output,
      contentTypes: base.contentTypes,
      mainRels: base.mainRels,
      officeBase: base.officeBase,
      targets: new Map(),
      partRels: new Map(),
    };
    for (const kind of ["styles", "numbering", "footnotes", "endnotes"]) {
      const part = relatedPart(base, kind);
      if (part) env.targets.set(kind, { part, root: parseWordRoot(base.files, part, kind) });
    }
    const baseSectionIndex = trailingSectionIndex(base.body);
    const baseSection = baseSectionIndex < 0 ? null : base.body.children[baseSectionIndex];
    const fallbackSection = baseSection ? adoptNodes([baseSection], base.root, base.root)[0] : null;
    if (baseSection) {
      base.body.children.splice(baseSectionIndex, 1);
      appendSectionToParagraph(base.body, baseSection);
    }
    const bookmarks = bookmarkState(base.body);
    let finalSection = null;

    for (let index = 1; index < packages.length; index++) {
      const source = packages[index];
      say(`Merging ${inputs[index]?.name || `Word file ${index + 1}`}...`);
      const ctx = { source, index, prefix: uniquePrefix(output, index), copied: new Map() };
      const sourceStylesPart = relatedPart(source, "styles");
      const sourceStyles = sourceStylesPart ? parseWordRoot(source.files, sourceStylesPart, "styles") : null;
      const sourceNumberingPart = relatedPart(source, "numbering");
      const sourceNumbering = sourceNumberingPart ? parseWordRoot(source.files, sourceNumberingPart, "numbering") : null;

      let styles = new Map();
      if (sourceStyles) {
        const target = ensureTargetPart(env, "styles", source, sourceStylesPart);
        const used = new Set(elements(target.root).filter(node => isWord(node, "style")).map(node => wordAttribute(node, "styleId")));
        if ([...used].some(id => !id)) invalid();
        let prefix = `M${index + 1}_`;
        while ([...used].some(id => id?.startsWith(prefix))) prefix += "_";
        for (const node of elements(sourceStyles).filter(node => isWord(node, "style"))) {
          const old = wordAttribute(node, "styleId");
          if (!old || styles.has(old)) invalid();
          let fresh = `${prefix}${old}`;
          let suffix = 2;
          while (used.has(fresh)) fresh = `${prefix}${old}_${suffix++}`;
          used.add(fresh);
          styles.set(old, fresh);
        }
      }
      const numbering = sourceNumbering ? mergeNumbering(env, ctx, sourceNumbering, sourceNumberingPart, styles) : new Map();
      if (sourceStyles) mergeStyles(env, ctx, sourceStyles, sourceStylesPart, numbering, styles);
      const notes = mergeNotes(env, ctx, styles, numbering);
      ctx.styles = styles;
      ctx.numbering = numbering;
      const relationshipIds = importDocumentRelationships(env, ctx);

      const imported = [...source.body.children];
      const sourceSectionIndex = trailingSectionIndex({ children: imported });
      const sourceSection = sourceSectionIndex < 0 ? null : imported.splice(sourceSectionIndex, 1)[0];
      const firstSourceSection = [...walk(imported)].find(node => isWord(node, "sectPr"));
      const adopted = adoptNodes(imported, source.root, base.root);
      if (firstSourceSection) {
        const adoptedFirstSection = [...walk(adopted)].find(node => isWord(node, "sectPr"));
        startSectionOnNewPage(adoptedFirstSection);
      }
      remapImportedNodes(adopted, styles, numbering);
      remapRelationshipAttributes(adopted, relationshipIds);
      remapNoteReferences(adopted, notes);
      remapBookmarks(adopted, bookmarks, index);
      base.body.children.push(...adopted);
      if (sourceSection) {
        const adoptedSection = adoptNodes([sourceSection], source.root, base.root)[0];
        if (!firstSourceSection) startSectionOnNewPage(adoptedSection);
        remapImportedNodes([adoptedSection], styles, numbering);
        remapRelationshipAttributes([adoptedSection], relationshipIds);
        if (index === packages.length - 1) {
          finalSection = adoptNodes([adoptedSection], base.root, base.root)[0];
        } else {
          appendSectionToParagraph(base.body, adoptedSection);
          finalSection = null;
        }
      } else {
        if (index < packages.length - 1) base.body.children.push(pageBreak(base.body, base.root));
        finalSection = null;
      }
    }

    base.body.children.push(finalSection || fallbackSection || makeElement(base.body.uri, "sectPr", prefixName(base.root.name)));
    alignCompatibilityMode(output, packages);
    output[base.mainPart] = serializeXml(base.root);
    output[relationshipsName(base.mainPart)] = serializeXml(env.mainRels.root);
    for (const target of env.targets.values()) output[target.part] = serializeXml(target.root);
    for (const state of env.partRels.values()) output[relationshipsName(state.part)] = serializeXml(state.root);
    ensureRelationshipsContentType(env.contentTypes);
    output["[Content_Types].xml"] = serializeXml(env.contentTypes.root);
    remapDrawingIds(output);
    say("Word files merged.");
    return zipSync(output, { level: 6 });
  } catch (error) {
    if (error instanceof MergeError) throw error;
    throw new MergeError("This Word file could not be read.");
  }
}
