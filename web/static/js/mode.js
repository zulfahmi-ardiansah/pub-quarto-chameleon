/* Wizard mode controller: toggles the whole page between the forward wizard
   (#render-form, Markdown → Word) and the reverse wizard (#reverse-form,
   Word → Markdown). Flips `body.mode-reverse`, which the CSS uses to swap the two
   header steppers ([data-forward-only] / [data-reverse-only]) and to hide the
   inactive form. */
(function () {
  const toggle = document.getElementById("mode-toggle");
  const label = document.getElementById("mode-toggle-label");
  const forwardForm = document.getElementById("render-form");
  const reverseForm = document.getElementById("reverse-form");
  if (!toggle || !forwardForm || !reverseForm) return;

  let mode = "forward"; // "forward" | "reverse"

  function apply() {
    const reverse = mode === "reverse";
    document.body.classList.toggle("mode-reverse", reverse);
    forwardForm.classList.toggle("hidden", reverse);
    reverseForm.classList.toggle("hidden", !reverse);
    // Label advertises the *destination* wizard, mirroring the toggle's action.
    label.dataset.i18n = reverse ? "mode.toForward" : "mode.toReverse";
    label.textContent = t(label.dataset.i18n);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  toggle.addEventListener("click", () => {
    mode = mode === "forward" ? "reverse" : "forward";
    if (mode === "reverse") window.ReverseWizard?.reset();
    apply();
  });

  // Keep the label translated when the language changes.
  document.addEventListener("langchange", () => { label.textContent = t(label.dataset.i18n); });

  apply();
})();
