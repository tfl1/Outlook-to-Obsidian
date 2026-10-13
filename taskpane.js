// Outlook -> Obsidian: connection test (v0.1)
// Proves three things: Office.js can read the selected email, Outlook's
// add-in browser can reach the Local REST API at 127.0.0.1, and we can list vault folders.

const API_BASE = "https://127.0.0.1:27124";
const KEY_SETTING = "obsidianApiKey";

function $(id) { return document.getElementById(id); }

function show(html) { $("result").innerHTML = html; }

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function getKey() {
  try { return Office.context.roamingSettings.get(KEY_SETTING) || ""; } catch (e) { return ""; }
}

function saveKey(key) {
  return new Promise((resolve, reject) => {
    Office.context.roamingSettings.set(KEY_SETTING, key);
    Office.context.roamingSettings.saveAsync(r =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error));
  });
}

async function apiGet(path, key) {
  const opts = { method: "GET", headers: { "Authorization": "Bearer " + key } };
  // Hint for browsers that enforce Local Network Access; ignored where unsupported.
  try { opts.targetAddressSpace = "loopback"; } catch (e) { /* ignore */ }
  const res = await fetch(API_BASE + path, opts);
  const text = await res.text();
  let body;
  try { body = JSON.parse(text); } catch (e) { body = text; }
  return { status: res.status, body };
}

async function runTest() {
  const key = $("apikey").value.trim() || getKey();
  if (!key) { show('<span class="err">Enter and save your API key first.</span>'); return; }

  show("Testing…");
  const lines = [];

  // 1. Reach the server and authenticate.
  try {
    const r = await apiGet("/", key);
    if (r.status === 401) {
      show('<span class="err">Reached Obsidian, but the API key was rejected (401). Re-copy the key from Obsidian settings.</span>');
      return;
    }
    if (r.status !== 200) {
      show('<span class="err">Reached Obsidian but got HTTP ' + r.status + ':\n' + escapeHtml(JSON.stringify(r.body)) + '</span>');
      return;
    }
    const authed = r.body && r.body.authenticated;
    lines.push('<div class="ok">✔ Connected to Obsidian' + (authed === false ? " (but not authenticated?)" : "") + '</div>');
  } catch (e) {
    show('<span class="err">✘ Could not reach https://127.0.0.1:27124\n' + escapeHtml(e && e.message ? e.message : e) +
      '\n\nLikely causes: Obsidian not running, vault not mounted, certificate not trusted, or Outlook blocking local network requests.</span>');
    return;
  }

  // 2. List top-level vault folders.
  try {
    const r = await apiGet("/vault/", key);
    if (r.status !== 200) {
      lines.push('<span class="err">Folder listing failed: HTTP ' + r.status + '</span>');
    } else {
      const files = (r.body && r.body.files) || [];
      const folders = files.filter(f => f.endsWith("/"));
      lines.push('<div class="ok">✔ Read vault: ' + folders.length + " top-level folders, " +
        (files.length - folders.length) + " top-level files</div>");
      if (folders.length) {
        lines.push("<ul>" + folders.map(f => "<li>" + escapeHtml(f) + "</li>").join("") + "</ul>");
      }
    }
  } catch (e) {
    lines.push('<span class="err">Folder listing error: ' + escapeHtml(e.message || e) + '</span>');
  }

  show(lines.join(""));
}

Office.onReady(() => {
  try {
    const item = Office.context.mailbox.item;
    $("subject").textContent = item && item.subject ? item.subject : "(no subject)";
  } catch (e) {
    $("subject").textContent = "(could not read email: " + e.message + ")";
  }

  if (getKey()) $("apikey").placeholder = "Key saved (paste a new one to replace)";

  $("save").onclick = async () => {
    const key = $("apikey").value.trim();
    if (!key) { show('<span class="err">Paste a key first.</span>'); return; }
    try {
      await saveKey(key);
      $("apikey").value = "";
      $("apikey").placeholder = "Key saved (paste a new one to replace)";
      show('<span class="ok">Key saved.</span>');
    } catch (e) {
      show('<span class="err">Could not save key: ' + escapeHtml(e.message || e) + '</span>');
    }
  };

  $("test").onclick = runTest;
});
