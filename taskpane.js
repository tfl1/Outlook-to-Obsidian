// Outlook -> Obsidian task pane (v1.0)
import PostalMime from "./lib/postal-mime.js";
import {
  ymd, sanitizeFilename, splitExt, uniqueName, joinPath, formatAddr, normalizeCid,
  extForMime, bytesToBase64, base64ToBytes, parseTags, buildNote, buildHtml,
} from "./lib/core.js";

const API = "https://127.0.0.1:27124";
const ATT = "_Attachments";
const K = { key: "obsidianApiKey", keepRemote: "keepRemoteImages", openAfter: "openAfterSave", recent: "recentFolders" };
const MAX_RECENT = 8;

const $ = id => document.getElementById(id);
let allFolders = null;      // array of folder paths ("" = vault root)
let selectedFolder = null;  // string or null
let busy = false;

// ---------- settings ----------
function getSetting(name, dflt) {
  try { const v = Office.context.roamingSettings.get(name); return v === undefined || v === null ? dflt : v; }
  catch { return dflt; }
}
function setSetting(name, value) {
  Office.context.roamingSettings.set(name, value);
  return new Promise((res, rej) => Office.context.roamingSettings.saveAsync(r =>
    r.status === Office.AsyncResultStatus.Succeeded ? res() : rej(new Error(r.error?.message || "Could not save setting"))));
}

// ---------- status ----------
function status(kind, text) { const s = $("status"); s.className = kind; s.textContent = text; }
function clearStatus() { const s = $("status"); s.className = ""; s.textContent = ""; }

// ---------- Local REST API ----------
const enc = p => p.split("/").map(encodeURIComponent).join("/");

async function api(path, { method = "GET", body, type } = {}) {
  const key = getSetting(K.key, "");
  if (!key) throw new Error("No API key saved. Open Settings at the bottom of this pane.");
  const headers = { Authorization: "Bearer " + key };
  if (type) headers["Content-Type"] = type;
  const opts = { method, headers, body };
  opts.targetAddressSpace = "loopback";
  let res;
  try { res = await fetch(API + path, opts); }
  catch (e) {
    throw new Error("Can't reach Obsidian at 127.0.0.1:27124. Is Obsidian running with the vault open (VeraCrypt drive mounted)?\n(" + e.message + ")");
  }
  if (res.status === 401) throw new Error("Obsidian rejected the API key (401). Re-paste it in Settings.");
  return res;
}

async function listDir(dir) {
  const res = await api("/vault/" + (dir ? enc(dir) + "/" : ""));
  if (res.status === 404) return [];
  if (!res.ok) throw new Error(`Listing "${dir || "vault root"}" failed: HTTP ${res.status}`);
  const j = await res.json();
  return j.files || [];
}

async function putFile(path, body, type) {
  const res = await api("/vault/" + enc(path), { method: "PUT", body, type });
  if (!res.ok) {
    let detail = "";
    try { detail = (await res.text()).slice(0, 300); } catch { /* ignore */ }
    throw new Error(`Writing "${path}" failed: HTTP ${res.status} ${detail}`);
  }
}

async function openInObsidian(path) {
  await api("/open/" + enc(path), { method: "POST" });
}

// ---------- folder list ----------
function skipDir(name) { return name.startsWith(".") || name.toLowerCase() === ATT.toLowerCase(); }

async function crawlFolders() {
  const found = [""];
  const queue = [""];
  let active = 0;
  await new Promise((resolve, reject) => {
    const pump = () => {
      if (!queue.length && active === 0) return resolve();
      while (queue.length && active < 6) {
        const dir = queue.shift();
        active++;
        listDir(dir).then(entries => {
          for (const e of entries) {
            if (!e.endsWith("/")) continue;
            const name = e.slice(0, -1).split("/").pop();
            if (skipDir(name)) continue;
            const full = joinPath(dir, name);
            found.push(full);
            queue.push(full);
          }
          active--;
          pump();
        }, reject);
      }
    };
    pump();
  });
  return found.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}

async function loadFolders() {
  $("folders").innerHTML = '<div class="empty">Loading folders…</div>';
  try {
    allFolders = await crawlFolders();
    renderFolders();
  } catch (e) {
    allFolders = null;
    $("folders").innerHTML = "";
    const d = document.createElement("div"); d.className = "empty"; d.textContent = e.message;
    $("folders").appendChild(d);
  }
  updateSaveButton();
}

function folderLabel(path) {
  if (path === "") return "(vault root)";
  return path;
}

function folderRow(path) {
  const div = document.createElement("div");
  div.className = "folder" + (path === selectedFolder ? " sel" : "");
  div.title = folderLabel(path);
  const i = path.lastIndexOf("/");
  if (i >= 0) {
    const p = document.createElement("span"); p.className = "parent"; p.textContent = path.slice(0, i + 1);
    div.appendChild(p);
    div.appendChild(document.createTextNode(path.slice(i + 1)));
  } else {
    div.textContent = folderLabel(path);
  }
  div.onclick = () => selectFolder(path);
  return div;
}

function renderFolders() {
  const box = $("folders");
  box.innerHTML = "";
  if (!allFolders) return;
  const q = $("filter").value.trim().toLowerCase();
  const words = q.split(/\s+/).filter(Boolean);
  const matches = allFolders.filter(f => words.every(w => folderLabel(f).toLowerCase().includes(w)));

  const addGroup = (title, items) => {
    if (!items.length) return;
    const g = document.createElement("div"); g.className = "group"; g.textContent = title; box.appendChild(g);
    items.slice(0, 200).forEach(f => box.appendChild(folderRow(f)));
  };

  if (!words.length) {
    const recent = getSetting(K.recent, []).filter(f => allFolders.includes(f));
    addGroup("Recent", recent);
    addGroup("All folders", allFolders);
  } else {
    addGroup(`${matches.length} match${matches.length === 1 ? "" : "es"}`, matches);
  }
  if (!box.children.length) box.innerHTML = '<div class="empty">No matching folders.</div>';
}

function selectFolder(path) {
  selectedFolder = path;
  $("chosen").innerHTML = "Save to: <b></b>";
  $("chosen").querySelector("b").textContent = folderLabel(path);
  renderFolders();
  updateSaveButton();
}

function updateSaveButton() {
  $("save").disabled = busy || selectedFolder === null || !Office.context.mailbox.item || !allFolders;
}

async function rememberFolder(path) {
  const recent = [path, ...getSetting(K.recent, []).filter(f => f !== path)].slice(0, MAX_RECENT);
  await setSetting(K.recent, recent);
}

// ---------- reading the email ----------
const officeCall = fn => new Promise((res, rej) => fn(r =>
  r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(new Error(r.error?.message || "Outlook call failed"))));

function readMeta(item, parsedDate) {
  let date = parsedDate ? new Date(parsedDate) : null;
  if (!date || isNaN(date)) date = new Date(item.dateTimeCreated || Date.now());
  return {
    subject: item.subject || "No subject",
    from: formatAddr(item.from),
    to: (item.to || []).map(formatAddr).filter(Boolean),
    cc: (item.cc || []).map(formatAddr).filter(Boolean),
    date,
  };
}

/** Preferred path: whole message as .eml, parsed locally. Returns {html, text, date, attachments:[{id,filename,mimeType,contentId,inlineish,bytes}]} */
async function readViaEml(item) {
  const b64 = await officeCall(cb => item.getAsFileAsync(cb));
  const email = await PostalMime.parse(base64ToBytes(b64), { rfc822Attachments: true, attachmentEncoding: "arraybuffer" });
  return {
    html: email.html || "",
    text: email.text || "",
    date: email.date,
    attachments: (email.attachments || []).map((a, i) => ({
      id: "eml" + i,
      filename: a.filename,
      mimeType: a.mimeType,
      contentId: a.contentId ? normalizeCid(a.contentId) : "",
      inlineish: !!(a.related || a.disposition === "inline"),
      bytes: new Uint8Array(a.content),
    })),
  };
}

/** Fallback path: body + attachments through Office.js (older Outlook builds). */
async function readViaOfficeJs(item) {
  const html = await officeCall(cb => item.body.getAsync(Office.CoercionType.Html, cb));
  const attachments = [];
  const canFetch = Office.context.requirements.isSetSupported("Mailbox", "1.8");
  for (const a of item.attachments || []) {
    if (!canFetch || a.attachmentType === "cloud") continue;
    const c = await officeCall(cb => item.getAttachmentContentAsync(a.id, cb));
    let bytes;
    if (c.format === "base64") bytes = base64ToBytes(c.content);
    else if (c.format === "eml" || c.format === "iCalendar") bytes = new TextEncoder().encode(c.content);
    else continue; // url format: nothing to save
    attachments.push({
      id: a.id,
      filename: a.attachmentType === "item" ? (a.name || "Attached email") + ".eml" : a.name,
      mimeType: a.contentType || "",
      contentId: "",
      inlineish: !!a.isInline,
      bytes,
    });
  }
  return { html, text: "", date: null, attachments };
}

// ---------- save ----------
async function save() {
  const item = Office.context.mailbox.item;
  if (!item || selectedFolder === null) return;
  busy = true; updateSaveButton();
  try {
    status("busy", "Reading email…");
    const useEml = typeof item.getAsFileAsync === "function" && Office.context.requirements.isSetSupported("Mailbox", "1.14");
    const msg = useEml ? await readViaEml(item) : await readViaOfficeJs(item);
    const meta = readMeta(item, msg.date);

    // Inline image lookup: by Content-ID, and by file name (for the fallback path).
    const inline = new Map();
    for (const a of msg.attachments) {
      if (!(a.mimeType || "").toLowerCase().startsWith("image/")) continue;
      const entry = { id: a.id, dataUri: `data:${a.mimeType};base64,${bytesToBase64(a.bytes)}` };
      if (a.contentId) { inline.set(a.contentId, entry); inline.set(a.contentId.split("@")[0], entry); }
      if (a.filename) inline.set(a.filename.toLowerCase(), entry);
    }

    status("busy", "Building HTML copy…");
    const keepRemote = $("keepRemote").checked;
    const { html, usedInlineIds } = buildHtml({ DOMParser, meta, html: msg.html, text: msg.text, inline, keepRemote });

    // Real attachments = not embedded in the HTML and not inline signature images.
    const files = msg.attachments.filter(a =>
      !usedInlineIds.has(a.id) && !(a.inlineish && (a.mimeType || "").toLowerCase().startsWith("image/")));

    status("busy", "Checking for existing files…");
    const folder = selectedFolder;
    const attFolder = joinPath(folder, ATT);
    const [folderEntries, attEntries] = await Promise.all([listDir(folder), listDir(attFolder)]);
    const takenNotes = new Set(folderEntries.filter(e => !e.endsWith("/")).map(e => e.toLowerCase()));
    const takenAtt = new Set(attEntries.filter(e => !e.endsWith("/")).map(e => e.toLowerCase()));

    let base = sanitizeFilename(`${ymd(meta.date)} ${meta.subject}`) || ymd(meta.date);
    // Pick a base name free for both the note and its HTML copy.
    let candidate = base;
    for (let i = 2; takenNotes.has((candidate + ".md").toLowerCase()) || takenAtt.has((candidate + ".html").toLowerCase()); i++) {
      candidate = `${base} (${i})`;
    }
    base = candidate;
    takenAtt.add((base + ".html").toLowerCase());

    const notePath = joinPath(folder, base + ".md");
    const htmlPath = joinPath(attFolder, base + ".html");

    const saved = [];
    for (const [i, a] of files.entries()) {
      status("busy", `Saving attachment ${i + 1} of ${files.length}…`);
      let name = sanitizeFilename(a.filename || "") || `${base} attachment ${i + 1}${extForMime(a.mimeType)}`;
      let [stem, ext] = splitExt(name);
      if (!ext) ext = extForMime(a.mimeType);
      name = uniqueName(stem, ext, takenAtt);
      const path = joinPath(attFolder, name);
      await putFile(path, a.bytes, "application/octet-stream");
      saved.push({ path, name });
    }

    status("busy", "Saving HTML copy…");
    await putFile(htmlPath, new TextEncoder().encode(html), "application/octet-stream");

    status("busy", "Saving note…");
    const note = buildNote(meta, parseTags($("tags").value), htmlPath, saved);
    await putFile(notePath, note, "text/markdown; charset=utf-8");

    await rememberFolder(folder).catch(() => {});
    if (getSetting(K.openAfter, false)) await openInObsidian(notePath).catch(() => {});

    status("ok", `Saved: ${notePath}` + (saved.length ? `\n+ ${saved.length} attachment${saved.length === 1 ? "" : "s"}` : "")
      + (useEml ? "" : "\n(Older Outlook API used; inline images may be missing.)"));
    renderFolders();
  } catch (e) {
    status("err", e.message || String(e));
  } finally {
    busy = false; updateSaveButton();
  }
}

// ---------- current email ----------
function showItem() {
  const item = Office.context.mailbox.item;
  clearStatus();
  $("tags").value = "";
  $("keepRemote").checked = getSetting(K.keepRemote, true);
  if (!item) {
    $("subject").textContent = "Select an email…";
    $("meta").textContent = "";
  } else {
    $("subject").textContent = item.subject || "(no subject)";
    $("meta").textContent = formatAddr(item.from) || "";
  }
  updateSaveButton();
}

// ---------- init ----------
Office.onReady(() => {
  $("defKeepRemote").checked = getSetting(K.keepRemote, true);
  $("defOpenAfter").checked = getSetting(K.openAfter, false);
  if (getSetting(K.key, "")) $("apikey").placeholder = "Key saved (paste a new one to replace)";
  else $("settings").open = true;

  $("defKeepRemote").onchange = async e => {
    await setSetting(K.keepRemote, e.target.checked).catch(err => status("err", err.message));
    $("keepRemote").checked = e.target.checked;
  };
  $("defOpenAfter").onchange = e => setSetting(K.openAfter, e.target.checked).catch(err => status("err", err.message));
  $("saveKey").onclick = async () => {
    const key = $("apikey").value.trim();
    if (!key) return status("err", "Paste the API key first.");
    try {
      await setSetting(K.key, key);
      $("apikey").value = "";
      $("apikey").placeholder = "Key saved (paste a new one to replace)";
      status("ok", "API key saved.");
      loadFolders();
    } catch (e) { status("err", e.message); }
  };

  $("filter").oninput = renderFolders;
  $("refresh").onclick = loadFolders;
  $("save").onclick = save;

  showItem();
  if (Office.context.requirements.isSetSupported("Mailbox", "1.5")) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, showItem);
  }
  if (getSetting(K.key, "")) loadFolders();
  else $("folders").innerHTML = '<div class="empty">Save your API key in Settings below.</div>';
});
