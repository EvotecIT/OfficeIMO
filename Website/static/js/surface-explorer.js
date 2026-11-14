// Turns each [data-surface-explorer] list of groups into a tablist with one visible panel.
// Without this script every group stays listed, so the page is complete at rest.
(function () {
  "use strict";
  function enhance(root) {
    var tablist = root.querySelector('[role="tablist"]');
    if (!tablist) return;
    var tabs = Array.prototype.slice.call(tablist.querySelectorAll('[role="tab"]'));
    var panels = tabs.map(function (tab) { return document.getElementById(tab.getAttribute("aria-controls")); });
    if (!tabs.length || panels.some(function (p) { return !p; })) return;
    function select(index, focus) {
      tabs.forEach(function (tab, i) {
        var on = i === index;
        tab.setAttribute("aria-selected", on ? "true" : "false");
        tab.tabIndex = on ? 0 : -1;
        panels[i].hidden = !on;
      });
      if (focus) tabs[index].focus();
    }
    tabs.forEach(function (tab, i) {
      tab.addEventListener("click", function () { select(i, false); });
      tab.addEventListener("keydown", function (e) {
        var next = null;
        if (e.key === "ArrowDown" || e.key === "ArrowRight") next = (i + 1) % tabs.length;
        else if (e.key === "ArrowUp" || e.key === "ArrowLeft") next = (i - 1 + tabs.length) % tabs.length;
        else if (e.key === "Home") next = 0;
        else if (e.key === "End") next = tabs.length - 1;
        if (next !== null) { e.preventDefault(); select(next, true); }
      });
    });
    var start = 0;
    var hash = location.hash.replace("#", "");
    if (hash) { panels.forEach(function (p, i) { if (p.id === hash) start = i; }); }
    tablist.hidden = false;
    root.classList.add("is-enhanced");
    select(start, false);
  }
  function init() { Array.prototype.forEach.call(document.querySelectorAll("[data-surface-explorer]"), enhance); }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
