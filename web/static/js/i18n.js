/* Lightweight i18n: text keyed by [data-i18n], placeholders by [data-i18n-placeholder],
   aria-labels by [data-i18n-aria]. Values are trusted constants (set via innerHTML so a
   few keys can carry markup). Language persists in localStorage. */

const I18N = {
  en: {
    "brand.sub": "Quarto Chameleon",
    "step.content": "Content",
    "step.content.desc": "Paste or upload your Markdown",
    "step.info": "Details",
    "step.info.desc": "Title, author, template",
    "step.render": "Review",
    "step.render.desc": "Check and download .docx",
    "step.done": "Done",
    "step.count": "Step {n} of 4: {label}",

    "eyebrow.step1": "<b class='font-bold'>Step 1</b> — Content",
    "eyebrow.step2": "<b class='font-bold'>Step 2</b> — Details",
    "eyebrow.step3": "<b class='font-bold'>Step 3</b> — Review",

    "s1.title": "Bring in your Markdown",
    "s1.desc": "Paste an AI answer or your own Markdown, or upload .md/.qmd files. Each file becomes a chapter.",
    "s2.title": "Describe the document",
    "s2.desc": "These appear on the cover and in page headers. Only the title is required.",
    "s3.title": "Review and generate",
    "s3.desc": "Everything look right? Generate your styled Word document.",

    "loading.title": "Generating your document",
    "loading.sub": "Rendering with Quarto. This can take up to a minute.",
    "success.title": "Your document is ready",
    "success.body": "Downloaded {file}. Check your downloads folder, or grab it again below.",
    "success.downloadAgain": "Download again",
    "success.createAnother": "Create another",

    "tab.paste": "Paste Markdown",
    "tab.upload": "Upload files",
    "label.markdown": "Markdown",
    "label.preview": "Preview",
    "btn.preview": "Preview",
    "btn.expand": "Expand",
    "btn.pasteClipboard": "Paste from clipboard",
    "modal.title": "Preview",
    "modal.note": "Approximate (Final styling comes from the template)",
    "placeholder.markdown": "# My Document\n\nPaste your AI response here!",
    "preview.empty": "Your document preview appears here as you type.",
    "clipboard.empty": "Clipboard is empty.",
    "clipboard.fail": "Couldn't read the clipboard. Paste manually with Ctrl/Cmd+V.",

    "help.open": "How do I copy from ChatGPT, Claude, Gemini, or others?",
    "help.title": "How to Copy from Your AI",
    "help.steps": "Hover the AI's reply and click the <strong>Copy</strong> button, then paste it into the Markdown box.",
    "help.other": "Using a different tool? It works the same way!",

    "upload.dropTitle": "Drag & drop files here",
    "upload.dropHint": "or click to browse .md / .qmd, each file becomes a chapter",
    "upload.remove": "Remove",

    "legend.cover": "Cover",
    "legend.header": "Header & TOC",
    "legend.template": "Template",
    "field.title": "Title <span class=\"text-primary-600\">*</span>",
    "field.subtitle": "Subtitle",
    "field.author": "Author",
    "field.date": "Date",
    "field.headerTitle": "Header title",
    "field.headerSubtitle": "Header subtitle",
    "field.tocTitle": "Table of contents title",
    "field.preset": "Built-in preset",
    "field.custom": "Custom template",
    "option.default": "Default (basic)",
    "template.hint.title": "Want to customise the look?",
    "template.hint.desc": "Download the base template, edit styles in Word, then upload it above.",

    "btn.next": "Next",
    "btn.back": "Back",
    "btn.render": "Generate .docx",

    "sum.content": "Content",
    "sum.title": "Title",
    "sum.author": "Author",
    "sum.date": "Date",
    "sum.template": "Template",
    "sum.contentPaste": "Pasted Markdown (single document)",
    "sum.contentFiles": "{n} uploaded file(s)",
    "sum.templateDefault": "Default (basic)",
    "sum.templateCustom": "Custom: {name}",

    "status.rendering": "Rendering. This can take a minute (Quarto + diagrams).",
    "status.done": "Done. Downloaded {file}.",
    "status.failed": "Failed.",
    "status.network": "Network error.",
    "validate.content": "Add some Markdown or upload at least one file before continuing.",
    "validate.title": "A cover title is required.",
  },
  id: {
    "brand.sub": "Quarto Chameleon",
    "step.content": "Konten",
    "step.content.desc": "Tempel atau unggah Markdown",
    "step.info": "Detail",
    "step.info.desc": "Judul, penulis, templat",
    "step.render": "Tinjau",
    "step.render.desc": "Periksa dan unduh .docx",
    "step.done": "Selesai",
    "step.count": "Langkah {n} dari 4: {label}",

    "eyebrow.step1": "<b class='font-bold'>Langkah 1</b> — Konten",
    "eyebrow.step2": "<b class='font-bold'>Langkah 2</b> — Detail",
    "eyebrow.step3": "<b class='font-bold'>Langkah 3</b> — Tinjau",

    "s1.title": "Masukkan Markdown Anda",
    "s1.desc": "Tempel jawaban AI atau Markdown Anda sendiri, atau unggah berkas .md/.qmd. Tiap berkas menjadi satu bab.",
    "s2.title": "Jelaskan dokumennya",
    "s2.desc": "Ini muncul di sampul dan header halaman. Hanya judul yang wajib diisi.",
    "s3.title": "Tinjau dan hasilkan",
    "s3.desc": "Semua sudah benar? Hasilkan dokumen Word Anda yang rapi.",

    "loading.title": "Membuat dokumen Anda",
    "loading.sub": "Merender dengan Quarto. Bisa memakan waktu hingga satu menit.",
    "success.title": "Dokumen Anda sudah siap",
    "success.body": "Berkas {file} terunduh. Periksa folder unduhan Anda, atau ambil lagi di bawah.",
    "success.downloadAgain": "Unduh lagi",
    "success.createAnother": "Buat lagi",

    "tab.paste": "Tempel Markdown",
    "tab.upload": "Unggah berkas",
    "label.markdown": "Markdown",
    "label.preview": "Pratinjau",
    "btn.preview": "Pratinjau",
    "btn.expand": "Perbesar",
    "btn.pasteClipboard": "Tempel dari papan klip",
    "modal.title": "Pratinjau",
    "modal.note": "Perkiraan (Gaya akhir mengikuti templat)",
    "placeholder.markdown": "# Dokumen Saya\n\nTempel respons AI Anda di sini!",
    "preview.empty": "Pratinjau dokumen muncul di sini saat Anda mengetik.",
    "clipboard.empty": "Papan klip kosong.",
    "clipboard.fail": "Tidak dapat membaca papan klip. Tempel manual dengan Ctrl/Cmd+V.",

    "help.open": "Bagaimana cara menyalin dari ChatGPT, Claude, Gemini, atau lainnya?",
    "help.title": "Cara Menyalin dari AI Anda",
    "help.steps": "Arahkan kursor ke balasan AI lalu klik tombol <strong>Copy</strong>, kemudian tempel ke kotak Markdown.",
    "help.other": "Pakai alat lain? Caranya sama!",

    "upload.dropTitle": "Seret & lepas berkas di sini",
    "upload.dropHint": "atau klik untuk memilih .md / .qmd, tiap berkas menjadi satu bab",
    "upload.remove": "Hapus",

    "legend.cover": "Sampul",
    "legend.header": "Header & Daftar isi",
    "legend.template": "Templat",
    "field.title": "Judul <span class=\"text-primary-600\">*</span>",
    "field.subtitle": "Subjudul",
    "field.author": "Penulis",
    "field.date": "Tanggal",
    "field.headerTitle": "Judul header",
    "field.headerSubtitle": "Subjudul header",
    "field.tocTitle": "Judul daftar isi",
    "field.preset": "Templat bawaan",
    "field.custom": "Kustom templat",
    "option.default": "Bawaan (basic)",
    "template.hint.title": "Ingin mengubah tampilan?",
    "template.hint.desc": "Unduh templat dasar, edit gaya di Word, lalu unggah di atas.",

    "btn.next": "Lanjut",
    "btn.back": "Kembali",
    "btn.render": "Hasilkan .docx",

    "sum.content": "Konten",
    "sum.title": "Judul",
    "sum.author": "Penulis",
    "sum.date": "Tanggal",
    "sum.template": "Templat",
    "sum.contentPaste": "Markdown ditempel (satu dokumen)",
    "sum.contentFiles": "{n} berkas diunggah",
    "sum.templateDefault": "Bawaan (basic)",
    "sum.templateCustom": "Kustom: {name}",

    "status.rendering": "Merender — proses ini bisa memakan waktu satu menit (Quarto + diagram).",
    "status.done": "Selesai. Berkas {file} terunduh.",
    "status.failed": "Gagal.",
    "status.network": "Kesalahan jaringan.",
    "validate.content": "Tambahkan Markdown atau unggah minimal satu berkas sebelum melanjutkan.",
    "validate.title": "Judul sampul wajib diisi.",
  },
};

let currentLang = localStorage.getItem("qlon-lang") || (navigator.language || "en").slice(0, 2);
if (!I18N[currentLang]) currentLang = "en";

function t(key, vars) {
  let s = (I18N[currentLang] && I18N[currentLang][key]) || I18N.en[key] || key;
  if (vars) for (const [k, v] of Object.entries(vars)) s = s.replaceAll(`{${k}}`, v);
  return s;
}

function applyI18n() {
  document.documentElement.lang = currentLang;
  document.querySelectorAll("[data-i18n]").forEach((el) => { el.innerHTML = t(el.dataset.i18n); });
  document.querySelectorAll("[data-i18n-placeholder]").forEach((el) => { el.placeholder = t(el.dataset.i18nPlaceholder); });
  document.querySelectorAll("[data-i18n-aria]").forEach((el) => { el.setAttribute("aria-label", t(el.dataset.i18nAria)); });
}

function setLang(lang) {
  if (!I18N[lang]) return;
  currentLang = lang;
  localStorage.setItem("qlon-lang", lang);
  applyI18n();
  document.dispatchEvent(new CustomEvent("langchange"));
}
