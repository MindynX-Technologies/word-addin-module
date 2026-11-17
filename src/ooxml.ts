import { BlobReader, BlobWriter, ZipReader, ZipWriter } from '@zip.js/zip.js';
import { DocxEditorError } from './types';

export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
export const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
export const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
export const MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
export const DOCUMENT_TYPE = MIME + '.main+xml';
export const TEMPLATE_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml';
const MAX_FILE = 32 * 1024 * 1024;
const MAX_PART = 48 * 1024 * 1024;
const MAX_EXPANSION = 128 * 1024 * 1024;
const MAX_ENTRIES = 4096;

export type Parts = Map<string, Blob>;

export function invalid(message: string): never {
  throw new DocxEditorError('INVALID_FILE', message);
}

function tooLarge(): never {
  throw new DocxEditorError('FILE_TOO_LARGE', 'This document exceeds the supported archive size.');
}

function validatePath(path: string) {
  if (!path || path.includes('\\') || /[\u0000-\u001f:]/.test(path) ||
    path.split('/').some(segment => !segment || segment === '.' || segment === '..' || /%2f|%5c|%2e/i.test(segment))) {
    invalid('The document contains an invalid package path.');
  }
}

export async function unzip(file: File, templates = false): Promise<Parts> {
  if (/\.(?:doc|dot)$/i.test(file.name)) {
    throw new DocxEditorError('UNSUPPORTED_FILE', 'Legacy Word (.doc/.dot) files require conversion. Save the file as Word Document (.docx) in Word, then open it here. Renaming the extension does not convert it.');
  }
  if (!(templates ? /\.do[ct]x$/i : /\.docx$/i).test(file.name)) {
    invalid(templates ? 'Choose a .docx or .dotx file.' : 'Choose a .docx file.');
  }
  if (file.size > MAX_FILE) tooLarge();
  const header = new Uint8Array(await file.slice(0, 8).arrayBuffer());
  if ([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((byte, index) => header[index] === byte)) {
    throw new DocxEditorError('UNSUPPORTED_FILE', 'This is a legacy binary or encrypted Word file. Save an unencrypted Word Document (.docx) in Word before opening it here.');
  }
  const reader = new ZipReader(new BlobReader(file), { useWebWorkers: false });
  const parts: Parts = new Map();
  try {
    const entries = [];
    const paths = new Set<string>();
    let expansion = 0;
    for await (const entry of reader.getEntriesGenerator()) {
      if (paths.size >= MAX_ENTRIES) tooLarge();
      const path = entry.directory ? entry.filename.replace(/\/$/, '') : entry.filename;
      validatePath(path);
      const key = path.toLowerCase();
      if (paths.has(key) || entry.encrypted || entry.symlink) invalid('The document contains an unsupported archive entry.');
      paths.add(key);
      if (entry.directory) continue;
      if (entry.uncompressedSize > MAX_PART || (expansion += entry.uncompressedSize) > MAX_EXPANSION) tooLarge();
      entries.push(entry);
    }
    for (const entry of entries) {
      const chunks: Blob[] = [];
      let size = 0;
      const sink = new WritableStream<Uint8Array<ArrayBuffer>>({
        write(chunk) {
          size += chunk.byteLength;
          if (size > entry.uncompressedSize || size > MAX_PART) tooLarge();
          chunks.push(new Blob([chunk]));
        },
      });
      await entry.getData(sink, { checkCrc32: true, checkOverlappingEntry: true });
      if (size !== entry.uncompressedSize) invalid('A document part has an invalid size.');
      parts.set(entry.filename, new Blob(chunks));
    }
    return parts;
  } catch (cause) {
    if (cause instanceof DocxEditorError) throw cause;
    throw new DocxEditorError('INVALID_FILE', 'This file could not be opened as a Word document.', { cause });
  } finally {
    await reader.close();
  }
}

export async function zip(parts: Parts, name: string) {
  const writer = new ZipWriter(new BlobWriter(MIME), { useWebWorkers: false });
  for (const [path, blob] of parts) {
    await writer.add(path, new BlobReader(blob), { level: 6 });
  }
  const archive = await writer.close();
  return new File([archive], name, { type: MIME });
}

export function parseXml(text: string) {
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) invalid('The document contains an unsupported XML declaration.');
  const document = new DOMParser().parseFromString(text, 'application/xml');
  if (document.querySelector('parsererror')) invalid('The document contains malformed XML.');
  return document;
}

export async function partXml(parts: Parts, path: string) {
  const blob = parts.get(path);
  if (!blob) invalid('A required Word document part is missing.');
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const encoding = bytes[0] === 255 && bytes[1] === 254 || bytes[0] === 60 && bytes[1] === 0 ? 'utf-16le'
    : bytes[0] === 254 && bytes[1] === 255 || bytes[0] === 0 && bytes[1] === 60 ? 'utf-16be' : 'utf-8';
  return parseXml(new TextDecoder(encoding, { fatal: true }).decode(bytes));
}

export function putXml(parts: Parts, path: string, xml: XMLDocument) {
  parts.set(path, new Blob([new XMLSerializer().serializeToString(xml)], { type: 'application/xml' }));
}

export function resolvePart(from: string, target: string) {
  let decoded: string;
  try {
    decoded = decodeURIComponent(target);
  } catch {
    return invalid('A document relationship has an invalid path.');
  }
  const path = decoded.startsWith('/') ? [] : from.split('/').slice(0, -1);
  for (const segment of decoded.replace(/^\//, '').split('/')) {
    if (segment === '..') {
      if (path.length === 0) invalid('A document relationship leaves the package.');
      path.pop();
    } else if (segment !== '.') {
      path.push(segment);
    }
  }
  const result = path.join('/');
  validatePath(result);
  return result;
}

export function relationshipPath(path: string) {
  const slash = path.lastIndexOf('/') + 1;
  return path.slice(0, slash) + '_rels/' + path.slice(slash) + '.rels';
}

export type Relationship = { id: string; type: string; target: string; external: boolean };

export function relationships(xml: XMLDocument) {
  if (xml.documentElement.namespaceURI !== RELS || xml.documentElement.localName !== 'Relationships') {
    invalid('The document has invalid package relationships.');
  }
  const result = new Map<string, Relationship>();
  for (const node of Array.from(xml.documentElement.children)) {
    if (node.namespaceURI !== RELS || node.localName !== 'Relationship') continue;
    const id = node.getAttribute('Id');
    const target = node.getAttribute('Target');
    const type = node.getAttribute('Type');
    if (!id || !target || !type || result.has(id)) invalid('The document has an invalid relationship.');
    result.set(id, { id, target, type, external: node.getAttribute('TargetMode') === 'External' });
  }
  return result;
}

export function safeHref(href: string) {
  try {
    const url = new URL(href);
    return ['http:', 'https:', 'mailto:'].includes(url.protocol) ? url.href : undefined;
  } catch {
    return undefined;
  }
}

export function child(node: Element | undefined, name: string, ns = W) {
  return node && Array.from(node.children).find(element => element.namespaceURI === ns && element.localName === name);
}

export function children(node: Element | undefined, name: string, ns = W) {
  return node ? Array.from(node.children).filter(element => element.namespaceURI === ns && element.localName === name) : [];
}

export function attr(node: Element | undefined, name = 'val', ns = W) {
  return node?.getAttributeNS(ns, name) ?? undefined;
}

export function containsOnly(node: Element, names: string[]) {
  return Array.from(node.children).every(element => element.namespaceURI === W && names.includes(element.localName));
}

export function element(parent: Element | XMLDocument, name: string, attributes: Record<string, string | number> = {}, ns = W, prefix = 'w') {
  const doc = parent.nodeType === 9 ? parent as XMLDocument : parent.ownerDocument!;
  const node = doc.createElementNS(ns, prefix ? `${prefix}:${name}` : name);
  for (const [key, value] of Object.entries(attributes)) {
    if (ns === W) node.setAttributeNS(W, `w:${key}`, String(value));
    else node.setAttribute(key, String(value));
  }
  return node;
}

export function add(parent: Element, name: string, attributes?: Record<string, string | number>) {
  const node = element(parent, name, attributes);
  parent.append(node);
  return node;
}

export function childProperty(parent: Element, name: string, order: string[], attributes?: Record<string, string | number>) {
  const old = child(parent, name);
  if (!attributes) {
    old?.remove();
    return;
  }
  const node = element(parent, name, attributes);
  if (old) {
    old.replaceWith(node);
  } else {
    const position = order.indexOf(name);
    parent.insertBefore(node, Array.from(parent.children).find(candidate =>
      candidate.namespaceURI === W && order.indexOf(candidate.localName) > position) ?? null);
  }
}
