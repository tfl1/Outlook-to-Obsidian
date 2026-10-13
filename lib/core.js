// Pure helpers for Outlook -> Obsidian. No Office.js or network code here,
// so these can be tested outside Outlook.

const REMOTE_RE = /^(https?:)?\/\//i;
const BAD_URL_RE = /^\s*(javascript|vbscript|data:(?!image\/))/i;

export function pad(n) { return String(n).padStart(2, "0"); }

/** 2026-10-13 */
export function ymd(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }

/** 2026-10-13T09:42 (local time, what Obsidian's date-time property expects) */
export function isoLocal(d) { return `${ymd(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`; }

/** Tue, Oct 13, 2026, 9:42 AM */
export function displayDate(d) {
  return d.toLocaleString("en-US", {
    weekday: "short", year: "numeric", month: "short", day: "numeric",
    hour: "numeric", minute: "2-digit",
  });
}

/** Make a string safe as an Obsidian file name (also safe inside [[wikilinks]]). */
export function sanitizeFilename(s, maxLen = 100) {
  let out = String(s || "")
    .replace(/[\\/:*?"<>|#^\[\]\u0000-\u001f]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^[.\s]+|[.\s]+$/g, "");
  if (out.length > maxLen) out = out.slice(0, maxLen).trim();
  return out;
}

/** Split "report.final.pdf" -> ["report.final", ".pdf"] */
export function splitExt(name) {
  const i = name.lastIndexOf(".");
  return i > 0 ? [name.slice(0, i), name.slice(i)] : [name, ""];
}

/** Return a name not present in `taken` (a Set of lowercase names), adding " (2)", " (3)"... */
export function uniqueName(base, ext, taken) {
  let name = base + ext;
  for (let i = 2; taken.has(name.toLowerCase()); i++) name = `${base} (${i})${ext}`;
  taken.add(name.toLowerCase());
  return name;
}

export function joinPath(folder, name) { return folder ? `${folder}/${name}` : name; }

export function formatAddr(a) {
  if (!a) return "";
  const name = (a.displayName || a.name || "").trim();
  const email = (a.emailAddress || a.address || "").trim();
  if (name && email && name.toLowerCase() !== email.toLowerCase()) return `${name} <${email}>`;
  return name || email;
}

export function normalizeCid(cid) {
  return String(cid || "").trim().replace(/^cid:/i, "").replace(/^<|>$/g, "").toLowerCase();
}

const EXT_BY_MIME = {
  "message/rfc822": ".eml", "application/pdf": ".pdf", "image/png": ".png", "image/jpeg": ".jpg",
  "image/gif": ".gif", "text/plain": ".txt", "text/calendar": ".ics", "text/html": ".html",
};
export function extForMime(mime) { return EXT_BY_MIME[String(mime || "").toLowerCase()] || ".bin"; }

export function bytesToBase64(bytes) {
  let s = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) s += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  return btoa(s);
}

export function base64ToBytes(b64) {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export function parseTags(input) {
  const seen = new Set();
  return String(input || "")
    .split(/[,;\n]/)
    .map(t => t.trim().replace(/^#+/, "").replace(/\s+/g, "-"))
    .filter(t => t && !seen.has(t.toLowerCase()) && seen.add(t.toLowerCase()));
}

function yamlStr(s) { return JSON.stringify(String(s)); } // JSON strings are valid YAML

/** Build the Markdown note. meta: {from, to[], cc[], date(Date), subject}; tags[]; htmlPath; attachments[{path,name}] */
export function buildNote(meta, tags, htmlPath, attachments) {
  const lines = ["---"];
  lines.push(`from: ${yamlStr(meta.from)}`);
  if (meta.to.length) { lines.push("to:"); meta.to.forEach(a => lines.push(`  - ${yamlStr(a)}`)); }
  if (meta.cc.length) { lines.push("cc:"); meta.cc.forEach(a => lines.push(`  - ${yamlStr(a)}`)); }
  lines.push(`date: ${isoLocal(meta.date)}`);
  lines.push(`subject: ${yamlStr(meta.subject)}`);
  if (tags.length) { lines.push("tags:"); tags.forEach(t => lines.push(`  - ${yamlStr(t)}`)); }
  lines.push("---", "", `![[${htmlPath}]]`, "");
  if (attachments.length) {
    lines.push("## Attachments", "");
    attachments.forEach(a => lines.push(`- [[${a.path}|${a.name}]]`));
    lines.push("");
  }
  return lines.join("\n");
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function isTrackingPixel(img) {
  const src = img.getAttribute("src") || "";
  if (!REMOTE_RE.test(src)) return false;
  const w = parseInt(img.getAttribute("width"), 10);
  const h = parseInt(img.getAttribute("height"), 10);
  if (w <= 3 && h <= 3) return true;
  const style = (img.getAttribute("style") || "").toLowerCase().replace(/\s+/g, "");
  if (/display:none|visibility:hidden/.test(style)) return true;
  if (/(^|;)width:[0-3]px/.test(style) && /(^|;)height:[0-3]px/.test(style)) return true;
  return false;
}

/**
 * Turn the email body into a safe, self-contained HTML document with a header block.
 * opts: { DOMParser, meta, html, text, inline: Map(normalizedCid|lowerName -> {dataUri, id}), keepRemote }
 * Returns { html, usedInlineIds: Set }
 */
export function buildHtml(opts) {
  const { meta, inline, keepRemote } = opts;
  const source = opts.html && opts.html.trim()
    ? opts.html
    : `<!doctype html><html><head></head><body><pre style="white-space:pre-wrap;font-family:Calibri,Arial,sans-serif;font-size:11pt">${escapeHtml(opts.text || "")}</pre></body></html>`;
  const doc = new opts.DOMParser().parseFromString(source, "text/html");
  const usedInlineIds = new Set();

  // 1. Remove active / dangerous content.
  doc.querySelectorAll("script,noscript,iframe,frame,frameset,object,embed,applet,base,link,meta[http-equiv],input,button,select,textarea")
    .forEach(el => el.remove());
  doc.querySelectorAll("form").forEach(f => f.replaceWith(...f.childNodes));
  doc.querySelectorAll("*").forEach(el => {
    for (const attr of Array.from(el.attributes)) {
      const n = attr.name.toLowerCase();
      if (n.startsWith("on")) el.removeAttribute(attr.name);
      else if (["href", "src", "action", "xlink:href", "formaction", "background"].includes(n) && BAD_URL_RE.test(attr.value)) el.removeAttribute(attr.name);
    }
  });
  doc.querySelectorAll("a[href]").forEach(a => { a.setAttribute("target", "_blank"); a.setAttribute("rel", "noopener noreferrer"); });

  // 2. Inline (cid:) images -> data URIs.
  doc.querySelectorAll("img[src]").forEach(img => {
    const src = img.getAttribute("src");
    if (!/^cid:/i.test(src)) return;
    const cid = normalizeCid(src);
    const hit = inline.get(cid) || inline.get(cid.split("@")[0]);
    if (hit) { img.setAttribute("src", hit.dataUri); usedInlineIds.add(hit.id); }
    else img.replaceWith(doc.createTextNode(`[inline image not found: ${cid}]`));
  });

  // 3. Tracking pixels always go; other remote images only if keepRemote is off.
  doc.querySelectorAll("img").forEach(img => { if (isTrackingPixel(img)) img.remove(); });
  if (!keepRemote) {
    doc.querySelectorAll("img").forEach(img => {
      const src = img.getAttribute("src") || "";
      if (REMOTE_RE.test(src)) {
        const alt = img.getAttribute("alt");
        const span = doc.createElement("span");
        span.setAttribute("style", "color:#888;font-size:smaller");
        span.textContent = alt ? `[image: ${alt}]` : "[image]";
        img.replaceWith(span);
      } else {
        img.removeAttribute("srcset");
      }
    });
    doc.querySelectorAll("[background]").forEach(el => {
      if (REMOTE_RE.test(el.getAttribute("background"))) el.removeAttribute("background");
    });
    doc.querySelectorAll("[style]").forEach(el => {
      el.setAttribute("style", el.getAttribute("style").replace(/url\(\s*['"]?(https?:)?\/\/[^)]*\)/gi, "none"));
    });
    doc.querySelectorAll("style").forEach(st => {
      st.textContent = st.textContent
        .replace(/@import[^;]+;/gi, "")
        .replace(/url\(\s*['"]?(https?:)?\/\/[^)]*\)/gi, "none");
    });
  }

  // 4. Head: charset + links open outside the viewer.
  const head = doc.head || doc.documentElement.insertBefore(doc.createElement("head"), doc.body);
  const charset = doc.createElement("meta");
  charset.setAttribute("charset", "utf-8");
  head.insertBefore(charset, head.firstChild);
  if (!doc.title) { const t = doc.createElement("title"); t.textContent = meta.subject; head.appendChild(t); }

  // 5. Header block.
  const rows = [["From", meta.from], ["To", meta.to.join("; ")]];
  if (meta.cc.length) rows.push(["Cc", meta.cc.join("; ")]);
  rows.push(["Date", displayDate(meta.date)], ["Subject", meta.subject]);
  const header = doc.createElement("div");
  header.setAttribute("style", "font-family:'Segoe UI',Arial,sans-serif;font-size:13px;line-height:1.45;color:#1f1f1f;background:#f6f6f6;border:0;border-bottom:1px solid #ddd;padding:10px 14px;margin:0 0 14px 0;text-align:left;");
  header.innerHTML = '<table style="border-collapse:collapse;border:0;font-size:13px;width:auto;">' +
    rows.map(([k, v]) =>
      `<tr><td style="color:#666;padding:1px 12px 1px 0;border:0;vertical-align:top;white-space:nowrap;">${k}:</td>` +
      `<td style="padding:1px 0;border:0;color:#1f1f1f;${k === "Subject" ? "font-weight:600;" : ""}">${escapeHtml(v)}</td></tr>`
    ).join("") + "</table>";
  doc.body.insertBefore(header, doc.body.firstChild);

  return { html: "<!doctype html>\n" + doc.documentElement.outerHTML, usedInlineIds };
}
