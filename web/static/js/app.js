/* Page shell shared by both wizards: theme toggle, language switch, and the one-time
   i18n boot. Wizard behaviour lives in render.js / reverse.js — this file owns only
   what belongs to the page chrome, and must load before either wizard. */

/* ---------- Theme toggle (multiple instances: rail + mobile bar) ---------- */

const themeToggles = [...document.querySelectorAll(".theme-toggle")];
function syncThemeIcons() {
  const dark = document.documentElement.classList.contains("dark");
  themeToggles.forEach((tg) => {
    tg.querySelector(".icon-sun").classList.toggle("hidden", dark);
    tg.querySelector(".icon-moon").classList.toggle("hidden", !dark);
  });
}
themeToggles.forEach((tg) =>
  tg.addEventListener("click", () => {
    const dark = document.documentElement.classList.toggle("dark");
    localStorage.setItem("qlon-theme", dark ? "dark" : "light");
    syncThemeIcons();
  })
);

/* ---------- Language toggle ---------- */

const langBtns = [...document.querySelectorAll(".lang-btn")];
function syncLangButtons() {
  langBtns.forEach((b) => b.classList.toggle("lang-active", b.dataset.lang === currentLang));
}
langBtns.forEach((b) => b.addEventListener("click", () => setLang(b.dataset.lang)));
// Each wizard listens for `langchange` itself to re-render its own content.
document.addEventListener("langchange", syncLangButtons);

/* ---------- Init ---------- */
applyI18n();
syncLangButtons();
syncThemeIcons();
