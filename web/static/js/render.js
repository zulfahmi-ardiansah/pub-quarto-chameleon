const form = document.getElementById("render-form");
const panels = [...document.querySelectorAll(".step-panel")];
const steps = [...document.querySelectorAll("[data-step-nav]")];
const connectors = [...document.querySelectorAll(".connector")];
const stepCount = document.getElementById("step-count");
const mobileProgress = document.getElementById("mobile-progress");
const submitBtn = document.getElementById("submit-btn");
const statusEl = document.getElementById("status");
const loadingOverlay = document.getElementById("loading-overlay");
const loadingSub = document.getElementById("loading-sub");

const STEP_KEYS = { 1: "step.content", 2: "step.info", 3: "step.render", 4: "step.done" };
// Inline SVG check — font-independent, so the "done" badge renders everywhere.
const CHECK_ICON = '<svg class="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"/></svg>';
let current = 1;
let mode = "paste"; // "paste" | "upload"
let lastDownload = null; // { url, filename } — kept for "Download again"

/* ---------- Language: re-render wizard content on switch (chrome lives in app.js) ---------- */

document.addEventListener("langchange", () => {
  refreshFileList();
  renderPreview();
  setStepCount();
  if (current === 3) buildSummary();
});

/* ---------- Wizard navigation ---------- */

function setStepCount() {
  if (stepCount) stepCount.textContent = t("step.count", { n: current, label: t(STEP_KEYS[current]) });
}

function showStep(n) {
  current = n;
  clearErrors();
  panels.forEach((p) => p.classList.toggle("hidden", Number(p.dataset.step) !== n));
  steps.forEach((el) => {
    const step = Number(el.dataset.stepNav);
    const done = step < n, active = step === n;
    el.classList.toggle("step-done", done);
    el.classList.toggle("step-active", active);
    el.classList.toggle("step-clickable", done);
    el.setAttribute("aria-current", active ? "step" : "false");
    const badge = el.querySelector(".badge");
    if (badge) badge.innerHTML = done ? CHECK_ICON : step;
  });
  connectors.forEach((c, i) => c.classList.toggle("connector-active", i + 1 < n));
  if (mobileProgress) mobileProgress.style.width = `${(n / panels.length) * 100}%`;
  setStepCount();
  if (n === 3) buildSummary();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

steps.forEach((el) =>
  el.addEventListener("click", () => {
    const step = Number(el.dataset.stepNav);
    if (step < current) showStep(step);
  })
);

function hasContent() {
  const pasted = document.getElementById("markdown").value.trim();
  const files = form.querySelector('input[name="files"]').files;
  return mode === "paste" ? pasted.length > 0 : files.length > 0;
}

function validateStep(n) {
  clearErrors();
  if (n === 1 && !hasContent()) { showError('[data-error-for="1"]', t("validate.content")); return false; }
  if (n === 2) {
    const title = form.querySelector('input[name="cover_title"]');
    if (!title.value.trim()) {
      showError("#err-cover-title", t("validate.title"));
      title.classList.add("field-invalid");
      title.focus();
      return false;
    }
  }
  return true;
}

function showError(sel, msg) {
  const el = document.querySelector(sel);
  if (el) { el.textContent = msg; el.classList.remove("hidden"); }
}
function clearErrors() {
  document.querySelectorAll(".step-error").forEach((e) => { e.textContent = ""; e.classList.add("hidden"); });
  form.querySelectorAll(".field-invalid").forEach((e) => e.classList.remove("field-invalid"));
}

form.querySelectorAll("[data-next]").forEach((b) =>
  b.addEventListener("click", () => { if (validateStep(current)) showStep(current + 1); })
);
form.querySelectorAll("[data-back]").forEach((b) =>
  b.addEventListener("click", () => showStep(current - 1))
);
document.getElementById("markdown").addEventListener("input", clearErrors);
form.querySelector('input[name="files"]').addEventListener("change", clearErrors);
form.querySelector('input[name="cover_title"]').addEventListener("input", clearErrors);

/* ---------- Content mode tabs ---------- */

const tabs = [...document.querySelectorAll(".mode-tab")];
function setMode(m) {
  mode = m;
  tabs.forEach((tb) => tb.classList.toggle("mode-tab-active", tb.dataset.mode === m));
  document.querySelectorAll("[data-mode-panel]").forEach((p) =>
    p.classList.toggle("hidden", p.dataset.modePanel !== m)
  );
}
tabs.forEach((tb) => tb.addEventListener("click", () => setMode(tb.dataset.mode)));

/* ---------- Markdown preview (inline sheet + modal, kept in sync) ---------- */

const mdInput = document.getElementById("markdown");
const previewTargets = [...document.querySelectorAll(".preview-target")];
function renderPreview() {
  const src = mdInput.value;
  const html = src.trim()
    ? DOMPurify.sanitize(marked.parse(src))
    : `<p class="doc-empty">${escapeHtml(t("preview.empty"))}</p>`;
  previewTargets.forEach((el) => { el.innerHTML = html; });
}
mdInput.addEventListener("input", renderPreview);

const previewModal = document.getElementById("preview-modal");
function openPreview() { previewModal.classList.remove("hidden"); document.body.classList.add("overflow-hidden"); }
function closePreview() { previewModal.classList.add("hidden"); document.body.classList.remove("overflow-hidden"); }
["open-preview", "open-preview-lg"].forEach((id) => {
  const el = document.getElementById(id);
  if (el) el.addEventListener("click", openPreview);
});
document.getElementById("close-preview").addEventListener("click", closePreview);
document.getElementById("preview-backdrop").addEventListener("click", closePreview);

/* ---------- Paste from clipboard ---------- */

const clipboardNote = document.getElementById("clipboard-note");
function showClipboardNote(msg) { clipboardNote.textContent = msg; clipboardNote.classList.remove("hidden"); }
document.getElementById("paste-clipboard").addEventListener("click", async () => {
  clipboardNote.classList.add("hidden");
  try {
    const text = await navigator.clipboard.readText();
    if (!text) { showClipboardNote(t("clipboard.empty")); return; }
    mdInput.value = mdInput.value.trim() ? `${mdInput.value.replace(/\s+$/, "")}\n\n${text}` : text;
    mdInput.dispatchEvent(new Event("input", { bubbles: true }));
    mdInput.focus();
  } catch (_) {
    showClipboardNote(t("clipboard.fail"));
  }
});

/* ---------- Upload: drag & drop ---------- */

const fileInput = form.querySelector('input[name="files"]');
const dropzone = document.getElementById("dropzone");
const fileList = document.getElementById("file-list");
const ALLOWED = /\.(md|qmd|zip)$/i;
const isZip = (f) => /\.zip$/i.test(f.name);

function syncInputFiles(files) {
  const dt = new DataTransfer();
  files.forEach((f) => dt.items.add(f));
  fileInput.files = dt.files;
}

function addFiles(incoming, current = [...fileInput.files]) {
  const accepted = [...incoming].filter((f) => ALLOWED.test(f.name));
  const zips = accepted.filter(isZip);
  const fail = (key) => { syncInputFiles(current); refreshFileList(); showError('[data-error-for="1"]', t(key)); };

  // A project zip is a whole bundle: exactly one, never mixed with loose md/qmd.
  if (zips.length > 1) return fail("upload.oneZip");
  if (zips.length === 1) {
    if (accepted.length > 1 || current.length) return fail("upload.zipAlone");
    syncInputFiles([zips[0]]); refreshFileList(); return;
  }

  // Loose md/qmd: many allowed, but not alongside an already-selected zip.
  if (current.some(isZip)) return fail("upload.zipAlone");
  const existing = current.filter((f) => !isZip(f));
  const names = new Set(existing.map((f) => f.name));
  accepted.forEach((f) => {
    if (!names.has(f.name)) { existing.push(f); names.add(f.name); }
  });

  syncInputFiles(existing.sort((a, b) => a.name.localeCompare(b.name)));
  refreshFileList();
}

function removeFile(name) {
  syncInputFiles([...fileInput.files].filter((f) => f.name !== name));
  refreshFileList();
}

function refreshFileList() {
  const files = [...fileInput.files];
  fileList.innerHTML = files.map((f, i) => `
    <li class="flex items-center justify-between rounded-lg border border-gray-200 bg-white px-3 py-2 dark:border-ink-800 dark:bg-ink-900">
      <span class="flex min-w-0 items-center gap-2">
        ${isZip(f)
          ? `<span class="inline-flex h-5 w-5 shrink-0 items-center justify-center rounded bg-primary-100 text-primary-700 dark:bg-primary-900/40 dark:text-primary-300"><svg class="h-3.5 w-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M3 7a2 2 0 012-2h4l2 2h8a2 2 0 012 2v8a2 2 0 01-2 2H5a2 2 0 01-2-2V7z"/></svg></span>`
          : `<span class="inline-flex h-5 w-5 shrink-0 items-center justify-center rounded bg-primary-100 text-xs font-semibold text-primary-700 dark:bg-primary-900/40 dark:text-primary-300">${i + 1}</span>`}
        <span class="truncate text-gray-700 dark:text-gray-200">${escapeHtml(f.name)}${isZip(f) ? ` <span class="text-xs text-gray-400 dark:text-gray-500">${escapeHtml(t("upload.zipTag"))}</span>` : ""}</span>
      </span>
      <button type="button" class="file-remove ml-3 shrink-0 text-gray-400 transition hover:text-red-600" data-name="${escapeHtml(f.name)}" aria-label="${escapeHtml(t("upload.remove"))}">
        <svg class="h-4 w-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12"/></svg>
      </button>
    </li>`).join("");
  fileList.querySelectorAll(".file-remove").forEach((b) =>
    b.addEventListener("click", () => removeFile(b.dataset.name))
  );
}

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInput.click(); } });
fileInput.addEventListener("change", () => addFiles([...fileInput.files], []));
["dragenter", "dragover"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add("dropzone-active"); })
);
["dragleave", "drop"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.remove("dropzone-active"); })
);
dropzone.addEventListener("drop", (e) => { if (e.dataTransfer?.files) addFiles(e.dataTransfer.files); });

/* ---------- Help modal ---------- */

const helpModal = document.getElementById("help-modal");
const helpImg = document.getElementById("help-img");
const helpTabs = [...document.querySelectorAll(".help-tab")];
const HELP_SRC = helpImg.getAttribute("src");                          // /static/help-chat-gpt.png?v=123
const HELP_QS = HELP_SRC.includes("?") ? HELP_SRC.slice(HELP_SRC.indexOf("?")) : "";
const HELP_DIR = HELP_SRC.replace(/help-[\w-]+\.png.*$/, "");          // /static/

function selectHelp(provider) {
  helpImg.src = `${HELP_DIR}help-${provider}.png${HELP_QS}`;
  helpImg.alt = `Copying a response in ${provider.replace("-", " ")}`;
  helpTabs.forEach((tb) => tb.classList.toggle("help-tab-active", tb.dataset.help === provider));
}
function openHelp() { helpModal.classList.remove("hidden"); document.body.classList.add("overflow-hidden"); }
function closeHelp() { helpModal.classList.add("hidden"); document.body.classList.remove("overflow-hidden"); }
helpTabs.forEach((tb) => tb.addEventListener("click", () => selectHelp(tb.dataset.help)));
document.getElementById("open-help").addEventListener("click", openHelp);
document.getElementById("close-help").addEventListener("click", closeHelp);
document.getElementById("help-backdrop").addEventListener("click", closeHelp);

/* Esc closes whichever modal is open. */
document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  if (!previewModal.classList.contains("hidden")) closePreview();
  if (!helpModal.classList.contains("hidden")) closeHelp();
});

/* ---------- Step 3 summary ---------- */

function buildSummary() {
  const get = (name) => form.querySelector(`[name="${name}"]`)?.value.trim() || "—";
  const upFiles = [...fileInput.files];
  const contentDesc = mode === "paste"
    ? t("sum.contentPaste")
    : (upFiles.length === 1 && isZip(upFiles[0])
        ? t("sum.contentZip", { name: upFiles[0].name })
        : t("sum.contentFiles", { n: upFiles.length }));
  const custom = form.querySelector('input[name="custom_template"]').files[0];
  const preset = get("preset");
  const tplDesc = custom ? t("sum.templateCustom", { name: custom.name })
    : (preset === "—" ? t("sum.templateDefault") : preset);

  const rows = [
    [t("sum.content"), contentDesc],
    [t("sum.title"), get("cover_title")],
    [t("sum.author"), get("cover_author")],
    [t("sum.date"), get("cover_date")],
    [t("sum.template"), tplDesc],
  ];
  document.getElementById("summary").innerHTML = rows.map(([k, v]) =>
    `<div class="flex items-center justify-between gap-4 px-4 py-3"><dt class="text-gray-400 dark:text-gray-500">${escapeHtml(k)}</dt><dd class="text-right font-medium">${escapeHtml(v)}</dd></div>`
  ).join("");
}

/* ---------- Status + submit ---------- */

function setStatus(kind, html) {
  const styles = {
    ok: "border-green-300 bg-green-50 text-green-800 dark:border-green-800 dark:bg-green-900/30 dark:text-green-200",
    error: "border-red-300 bg-red-50 text-red-800 dark:border-red-800 dark:bg-red-900/30 dark:text-red-200",
  };
  statusEl.className = `mt-4 rounded-lg border p-4 text-sm ${styles[kind]}`;
  statusEl.innerHTML = html;
  statusEl.classList.remove("hidden");
}
function clearStatus() { statusEl.classList.add("hidden"); statusEl.innerHTML = ""; }

function showLoading() { loadingSub.textContent = t("loading.sub"); loadingOverlay.classList.remove("hidden"); document.body.classList.add("overflow-hidden"); }
function hideLoading() { loadingOverlay.classList.add("hidden"); document.body.classList.remove("overflow-hidden"); }
/* Keep the overlay up at least `ms` so a fast render reads as deliberate, not a flicker. */
function holdLoading(startedAt, ms = 650) {
  const left = ms - (Date.now() - startedAt);
  return left > 0 ? new Promise((r) => setTimeout(r, left)) : Promise.resolve();
}

function triggerDownload(url, filename) {
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

function showSuccess(filename) {
  document.getElementById("success-body").textContent = t("success.body", { file: filename });
  showStep(4);
}

/* Download again — reuses the blob from the last successful render. */
document.getElementById("download-again").addEventListener("click", () => {
  if (lastDownload) triggerDownload(lastDownload.url, lastDownload.filename);
});

/* Create another — reset the form and return to step 1. */
document.getElementById("create-another").addEventListener("click", () => {
  if (lastDownload) { URL.revokeObjectURL(lastDownload.url); lastDownload = null; }
  form.reset();
  syncInputFiles([]);
  refreshFileList();
  renderPreview();
  clearStatus();
  clearErrors();
  setMode("paste");
  showStep(1);
});

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  if (!validateStep(1) || !validateStep(2)) return;

  clearStatus();
  submitBtn.disabled = true;
  const started = Date.now();
  showLoading();

  try {
    const res = await fetch("/render", { method: "POST", body: new FormData(form) });
    if (!res.ok) {
      let msg = `Request failed (${res.status}).`;
      try {
        const data = await res.json();
        msg = data.error || msg;
        if (data.detail) msg += `\n\n${data.detail}`;
      } catch (_) { /* non-JSON body */ }
      await holdLoading(started);
      hideLoading();
      setStatus("error", `<strong>${escapeHtml(t("status.failed"))}</strong><pre class="mt-2 whitespace-pre-wrap break-words">${escapeHtml(msg)}</pre>`);
      return;
    }

    const disp = res.headers.get("Content-Disposition") || "";
    const match = disp.match(/filename="?([^"]+)"?/);
    const filename = match ? match[1] : "qlon-output";

    if (lastDownload) URL.revokeObjectURL(lastDownload.url);
    const url = URL.createObjectURL(await res.blob());
    lastDownload = { url, filename };
    triggerDownload(url, filename);
    await holdLoading(started);
    hideLoading();
    showSuccess(filename);
  } catch (err) {
    await holdLoading(started);
    hideLoading();
    setStatus("error", `<strong>${escapeHtml(t("status.network"))}</strong><pre class="mt-2 whitespace-pre-wrap break-words">${escapeHtml(String(err))}</pre>`);
  } finally {
    submitBtn.disabled = false;
  }
});

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => (
    { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]
  ));
}

/* ---------- Init (page chrome + i18n boot handled by app.js) ---------- */
setMode("paste");
selectHelp("chat-gpt");
showStep(1);
renderPreview();
