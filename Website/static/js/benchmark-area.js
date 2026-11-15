// Switches between the runs of each benchmark workload. Every run is rendered at build time from the
// comparison view (site.json BenchmarkEvidence), so without this script all runs stay listed.
(function () {
  "use strict";
  function enhance(section) {
    var buttons = Array.prototype.slice.call(section.querySelectorAll(".imo-bench-run[data-run]"));
    var panels = Array.prototype.slice.call(section.querySelectorAll("[data-run-panel]"));
    if (!buttons.length || !panels.length) return;
    function select(id) {
      buttons.forEach(function (b) { b.setAttribute("aria-pressed", b.getAttribute("data-run") === id ? "true" : "false"); });
      panels.forEach(function (p) { p.hidden = p.getAttribute("data-run-panel") !== id; });
    }
    buttons.forEach(function (b) { b.addEventListener("click", function () { select(b.getAttribute("data-run")); }); });
    var preferred = section.querySelector("[data-run-panel][data-preferred]") || panels[0];
    select(preferred.getAttribute("data-run-panel"));
  }
  function init() { Array.prototype.forEach.call(document.querySelectorAll("[data-benchmark-workload]"), enhance); }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
