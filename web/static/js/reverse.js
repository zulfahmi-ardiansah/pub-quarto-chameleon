/* Reverse wizard (docx → Markdown). Self-contained module mirroring render.js, scoped
   to #reverse-form so it never touches the forward wizard's DOM or globals. Exposes
   window.ReverseWizard = { reset } for the mode controller to call on entry. */
(function () {
  const form = document.getElementById("reverse-form");
  if (!form) return;

  const panels = [...form.querySelectorAll("[data-rev-step]")];
  // The stepper lives in the header (#reverse-progress), like the forward wizard's.
  const nav = document.getElementById("reverse-progress");
  const steps = [...nav.querySelectorAll("[data-rev-step-nav]")];
  const connectors = [...nav.querySelectorAll(".connector")];
  const stepCount = document.getElementById("rev-step-count");
  const mobileProgress = document.getElementById("rev-mobile-progress");
  const submitBtn = document.getElementById("rev-submit-btn");
  const statusEl = document.getElementById("rev-status");
  const loadingOverlay = document.getElementById("loading-overlay");

  const CHECK_ICON = '<svg class="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"/></svg>';
  const STEP_KEYS = { 1: "rev.step.upload", 2: "rev.step.options", 3: "rev.step.review", 4: "rev.step.done" };
  let current = 1;
  let lastDownload = null; // { url, filename }

  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => (
    { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]
  ));

  /* ---------- Navigation ---------- */

  function showStep(n) {
    current = n;
    clearErrors();
    panels.forEach((p) => p.classList.toggle("hidden", Number(p.dataset.revStep) !== n));
    steps.forEach((el) => {
      const step = Number(el.dataset.revStepNav);
      const done = step < n, active = step === n;
      el.classList.toggle("step-done", done);
      el.classList.toggle("step-active", active);
      el.classList.toggle("step-clickable", done);
      el.setAttribute("aria-current", active ? "step" : "false");
      const badge = el.querySelector(".badge");
      if (badge) badge.innerHTML = done ? CHECK_ICON : step;
    });
    connectors.forEach((c, i) => c.classList.toggle("connector-active", i + 1 < n));
    setStepCount();
    if (mobileProgress) mobileProgress.style.width = `${(n / panels.length) * 100}%`;
    if (n === 3) buildSummary();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function setStepCount() {
    if (stepCount) stepCount.textContent = t("step.count", { n: current, label: t(STEP_KEYS[current]) });
  }

  steps.forEach((el) =>
    el.addEventListener("click", () => {
      const step = Number(el.dataset.revStepNav);
      if (step < current) showStep(step);
    })
  );

  function validateStep(n) {
    clearErrors();
    if (n === 1 && fileInput.files.length === 0) {
      showError('[data-rev-error-for="1"]', t("rev.validate.file"));
      return false;
    }
    return true;
  }

  function showError(sel, msg) {
    const el = form.querySelector(sel);
    if (el) { el.textContent = msg; el.classList.remove("hidden"); }
  }
  function clearErrors() {
    form.querySelectorAll(".step-error").forEach((e) => { e.textContent = ""; e.classList.add("hidden"); });
  }

  form.querySelectorAll("[data-rev-next]").forEach((b) =>
    b.addEventListener("click", () => { if (validateStep(current)) showStep(current + 1); })
  );
  form.querySelectorAll("[data-rev-back]").forEach((b) =>
    b.addEventListener("click", () => showStep(current - 1))
  );

  /* ---------- Upload: single-file dropzone ---------- */

  const fileInput = document.getElementById("rev-file-input");
  const dropzone = document.getElementById("rev-dropzone");
  const fileList = document.getElementById("rev-file-list");
  const isDocx = (f) => /\.docx$/i.test(f.name);

  function setFile(file) {
    const dt = new DataTransfer();
    if (file) dt.items.add(file);
    fileInput.files = dt.files;
    refreshFileList();
  }
  function refreshFileList() {
    const files = [...fileInput.files];
    fileList.innerHTML = files.map((f) => `
      <li class="flex items-center justify-between rounded-lg border border-gray-200 bg-white px-3 py-2 dark:border-ink-800 dark:bg-ink-900">
        <span class="flex min-w-0 items-center gap-2">
          <span class="inline-flex h-5 w-5 shrink-0 items-center justify-center rounded bg-primary-100 text-primary-700 dark:bg-primary-900/40 dark:text-primary-300"><svg class="h-3.5 w-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg></span>
          <span class="truncate text-gray-700 dark:text-gray-200">${esc(f.name)}</span>
        </span>
        <button type="button" class="rev-file-remove ml-3 shrink-0 text-gray-400 transition hover:text-red-600" aria-label="${esc(t("upload.remove"))}">
          <svg class="h-4 w-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12"/></svg>
        </button>
      </li>`).join("");
    fileList.querySelectorAll(".rev-file-remove").forEach((b) =>
      b.addEventListener("click", () => setFile(null))
    );
  }
  // A dropped/selected .docx replaces any previous one.
  function acceptFiles(incoming) {
    const docx = [...incoming].find(isDocx);
    if (docx) { setFile(docx); clearErrors(); }
  }

  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInput.click(); } });
  fileInput.addEventListener("change", () => { if (fileInput.files.length) acceptFiles(fileInput.files); });
  ["dragenter", "dragover"].forEach((ev) =>
    dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add("dropzone-active"); })
  );
  ["dragleave", "drop"].forEach((ev) =>
    dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.remove("dropzone-active"); })
  );
  dropzone.addEventListener("drop", (e) => { if (e.dataTransfer?.files) acceptFiles(e.dataTransfer.files); });

  /* ---------- LLM reveal + fuma lock ---------- */

  const layoutSel = document.getElementById("rev-layout");
  const useLlm = document.getElementById("rev-use-llm");
  const llmFields = document.getElementById("rev-llm-fields");
  const fumaNote = document.getElementById("rev-fuma-note");

  function syncLlm() {
    const fuma = layoutSel.value === "fuma";
    if (fuma) { useLlm.checked = true; useLlm.disabled = true; }
    else { useLlm.disabled = false; }
    fumaNote.classList.toggle("hidden", !fuma);
    llmFields.classList.toggle("hidden", !useLlm.checked);
  }
  layoutSel.addEventListener("change", syncLlm);
  useLlm.addEventListener("change", syncLlm);

  /* ---------- Step 3 summary ---------- */

  function buildSummary() {
    const file = fileInput.files[0];
    const fuma = layoutSel.value === "fuma";
    const llm = useLlm.checked;
    const skipToc = form.querySelector('[name="skip_toc"]').value.trim();
    const noMerge = form.querySelector('[name="no_merge"]').checked;
    const reorder = form.querySelector('[name="allow_reorder"]').checked;
    const model = form.querySelector('[name="model_name"]').value.trim();

    const rows = [
      [t("rev.sum.file"), file ? file.name : "—"],
      [t("rev.sum.layout"), fuma ? t("rev.layout.fuma") : t("rev.layout.flat")],
      [t("rev.sum.llm"), llm ? t("rev.sum.on") : t("rev.sum.off")],
    ];
    if (llm && model) rows.push([t("rev.field.modelName"), model]);
    if (llm && reorder) rows.push([t("rev.field.allowReorder"), t("rev.sum.on")]);
    if (skipToc) rows.push([t("rev.field.skipToc"), skipToc]);
    if (noMerge) rows.push([t("rev.field.noMerge"), t("rev.sum.on")]);

    document.getElementById("rev-summary").innerHTML = rows.map(([k, v]) =>
      `<div class="flex items-center justify-between gap-4 px-4 py-3"><dt class="text-gray-400 dark:text-gray-500">${esc(k)}</dt><dd class="text-right font-medium">${esc(v)}</dd></div>`
    ).join("");
  }
  document.addEventListener("langchange", () => { setStepCount(); if (current === 3) buildSummary(); });

  /* ---------- Status + submit ---------- */

  function setStatus(kind, html) {
    const styles = {
      error: "border-red-300 bg-red-50 text-red-800 dark:border-red-800 dark:bg-red-900/30 dark:text-red-200",
    };
    statusEl.className = `mt-4 rounded-lg border p-4 text-sm ${styles[kind]}`;
    statusEl.innerHTML = html;
    statusEl.classList.remove("hidden");
  }
  function clearStatus() { statusEl.classList.add("hidden"); statusEl.innerHTML = ""; }

  function showLoading() { loadingOverlay.classList.remove("hidden"); document.body.classList.add("overflow-hidden"); }
  function hideLoading() { loadingOverlay.classList.add("hidden"); document.body.classList.remove("overflow-hidden"); }
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

  document.getElementById("rev-download-again").addEventListener("click", () => {
    if (lastDownload) triggerDownload(lastDownload.url, lastDownload.filename);
  });
  document.getElementById("rev-create-another").addEventListener("click", () => reset());

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (!validateStep(1)) { showStep(1); return; }

    clearStatus();
    submitBtn.disabled = true;
    const started = Date.now();
    showLoading();

    try {
      const res = await fetch("/reverse", { method: "POST", body: new FormData(form) });
      if (!res.ok) {
        let msg = `Request failed (${res.status}).`;
        try {
          const data = await res.json();
          msg = data.error || msg;
          if (data.detail) msg += `\n\n${data.detail}`;
        } catch (_) { /* non-JSON body */ }
        await holdLoading(started);
        hideLoading();
        setStatus("error", `<strong>${esc(t("status.failed"))}</strong><pre class="mt-2 whitespace-pre-wrap break-words">${esc(msg)}</pre>`);
        return;
      }

      const disp = res.headers.get("Content-Disposition") || "";
      const match = disp.match(/filename="?([^"]+)"?/);
      const filename = match ? match[1] : "qlon-markdown.zip";

      if (lastDownload) URL.revokeObjectURL(lastDownload.url);
      const url = URL.createObjectURL(await res.blob());
      lastDownload = { url, filename };
      triggerDownload(url, filename);
      await holdLoading(started);
      hideLoading();
      document.getElementById("rev-success-body").textContent = t("rev.success.body", { file: filename });
      showStep(4);
    } catch (err) {
      await holdLoading(started);
      hideLoading();
      setStatus("error", `<strong>${esc(t("status.network"))}</strong><pre class="mt-2 whitespace-pre-wrap break-words">${esc(String(err))}</pre>`);
    } finally {
      submitBtn.disabled = false;
    }
  });

  /* ---------- Reset (called on wizard entry / "convert another") ---------- */

  function reset() {
    if (lastDownload) { URL.revokeObjectURL(lastDownload.url); lastDownload = null; }
    form.reset();
    setFile(null);
    syncLlm();
    clearStatus();
    clearErrors();
    showStep(1);
  }

  /* ---------- Init ---------- */
  syncLlm();
  showStep(1);
  window.ReverseWizard = { reset };
})();
