// A small, allocation-light XML scanner for OOXML parts. It does not validate; it reports tags and
// text in document order. Element names keep their prefix only for extension namespaces, so `x:c`
// and `c` both read as `c` while `x14:cfRule` stays distinct.

export type Attrs = Record<string, string>;

export interface XmlHandler {
  open?(name: string, attrs: Attrs, selfClosing: boolean): void;
  close?(name: string): void;
  text?(text: string): void;
}

const KEEP_PREFIX = /^(x14|x15|x14ac|xr\d*|mc|x16r\d*|xcalcf|xm|a14|a16|c14|c16r\d*)$/;

function localName(raw: string): string {
  const colon = raw.indexOf(":");
  if (colon < 0) return raw;
  const prefix = raw.slice(0, colon);
  return KEEP_PREFIX.test(prefix) ? raw : raw.slice(colon + 1);
}

const ENTITY = /&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g;

export function decodeEntities(text: string): string {
  if (text.indexOf("&") < 0) return text;
  return text.replace(ENTITY, (_, entity: string) => {
    switch (entity) {
      case "lt": return "<";
      case "gt": return ">";
      case "amp": return "&";
      case "quot": return "\"";
      case "apos": return "'";
      default: {
        const code = entity[1] === "x" ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10);
        return Number.isFinite(code) ? String.fromCodePoint(code) : "";
      }
    }
  });
}

const CR_LF = /\r\n?/g;

/** Excel shows CR LF and LF alike; the renderer only needs LF. */
export function normalizeNewlines(text: string): string {
  return text.indexOf("\r") < 0 ? text : text.replace(CR_LF, "\n");
}

/** OOXML escapes control characters in strings as _xHHHH_. */
export function decodeOoxmlEscapes(text: string): string {
  const decoded = text.indexOf("_x") < 0 ? text : text.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)));
  return normalizeNewlines(decoded);
}

const ATTR = /([^\s=/]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

function parseAttrs(source: string): Attrs {
  const attrs: Attrs = {};
  if (source.length === 0) return attrs;
  ATTR.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = ATTR.exec(source))) {
    attrs[localName(match[1]!)] = decodeEntities(match[3] ?? match[4] ?? "");
  }
  return attrs;
}

export function scanXml(xml: string, handler: XmlHandler): void {
  const { open, close, text } = handler;
  let index = 0;
  const length = xml.length;
  while (index < length) {
    const lt = xml.indexOf("<", index);
    if (lt < 0) {
      if (text && index < length) text(decodeEntities(xml.slice(index)));
      break;
    }
    if (text && lt > index) text(decodeEntities(xml.slice(index, lt)));
    const next = xml.charCodeAt(lt + 1);
    if (next === 63 /* ? */) {
      index = xml.indexOf("?>", lt) + 2;
      if (index < 2) break;
      continue;
    }
    if (next === 33 /* ! */) {
      if (xml.startsWith("<!--", lt)) {
        const end = xml.indexOf("-->", lt);
        index = end < 0 ? length : end + 3;
      } else if (xml.startsWith("<![CDATA[", lt)) {
        const end = xml.indexOf("]]>", lt);
        if (text) text(xml.slice(lt + 9, end < 0 ? length : end));
        index = end < 0 ? length : end + 3;
      } else {
        const end = xml.indexOf(">", lt);
        index = end < 0 ? length : end + 1;
      }
      continue;
    }
    const gt = findTagEnd(xml, lt + 1);
    if (gt < 0) break;
    if (next === 47 /* / */) {
      if (close) close(localName(xml.slice(lt + 2, gt).trim()));
      index = gt + 1;
      continue;
    }
    const selfClosing = xml.charCodeAt(gt - 1) === 47;
    const body = xml.slice(lt + 1, selfClosing ? gt - 1 : gt);
    let nameEnd = 0;
    while (nameEnd < body.length) {
      const code = body.charCodeAt(nameEnd);
      if (code === 32 || code === 9 || code === 10 || code === 13) break;
      nameEnd++;
    }
    const name = localName(body.slice(0, nameEnd));
    if (open) open(name, parseAttrs(body.slice(nameEnd)), selfClosing);
    if (selfClosing && close) close(name);
    index = gt + 1;
  }
}

// Attribute values may contain ">", so skip quoted runs.
function findTagEnd(xml: string, from: number): number {
  let quote = 0;
  for (let i = from; i < xml.length; i++) {
    const code = xml.charCodeAt(i);
    if (quote) {
      if (code === quote) quote = 0;
    } else if (code === 34 || code === 39) {
      quote = code;
    } else if (code === 62) {
      return i;
    }
  }
  return -1;
}

/** A tiny element tree for small parts (styles, theme, workbook, drawings). */
export interface XmlNode {
  name: string;
  attrs: Attrs;
  children: XmlNode[];
  text: string;
}

export function parseTree(xml: string): XmlNode {
  const root: XmlNode = { name: "#root", attrs: {}, children: [], text: "" };
  const stack: XmlNode[] = [root];
  scanXml(xml, {
    open(name, attrs) {
      const node: XmlNode = { name, attrs, children: [], text: "" };
      stack[stack.length - 1]!.children.push(node);
      stack.push(node);
    },
    close() {
      if (stack.length > 1) stack.pop();
    },
    text(value) {
      stack[stack.length - 1]!.text += value;
    },
  });
  return root;
}

export function child(node: XmlNode | undefined, name: string): XmlNode | undefined {
  return node?.children.find((entry) => entry.name === name);
}

export function children(node: XmlNode | undefined, name: string): XmlNode[] {
  return node ? node.children.filter((entry) => entry.name === name) : [];
}

export function descendants(node: XmlNode | undefined, name: string, out: XmlNode[] = []): XmlNode[] {
  if (!node) return out;
  for (const entry of node.children) {
    if (entry.name === name) out.push(entry);
    descendants(entry, name, out);
  }
  return out;
}

export function textOf(node: XmlNode | undefined): string {
  if (!node) return "";
  let result = node.text;
  for (const entry of node.children) result += textOf(entry);
  return result;
}
