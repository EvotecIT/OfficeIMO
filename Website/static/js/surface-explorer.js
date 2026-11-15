// Turns each [data-surface-explorer] list of groups into a tablist with one visible panel,
// and trims the [data-surface-proof] script to the height of its document.
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
  // The proof script is complete at rest. With script it is trimmed to the height of the document
  // beside it, and the button expands it in place rather than scrolling inside a box.
  function enhanceProof(root) {
    var code = root.querySelector("[data-surface-proof-code]");
    var toggle = root.querySelector("[data-surface-proof-toggle]");
    if (!code || !toggle) return;
    root.classList.add("is-clamped");
    toggle.hidden = false;
    toggle.addEventListener("click", function () {
      var expanded = root.classList.toggle("is-expanded");
      toggle.setAttribute("aria-expanded", expanded ? "true" : "false");
      toggle.textContent = toggle.getAttribute(expanded ? "data-label-less" : "data-label-more");
      if (!expanded) {
        var top = root.getBoundingClientRect().top;
        if (top < 0) root.scrollIntoView({ block: "start" });
      }
    });
  }
  // Long CLI option lists show the first eight flags until the reader asks for the rest.
  function enhanceSignature(sig) {
    var toggle = sig.querySelector(".imo-cli-sig__more");
    if (!toggle) return;
    sig.classList.add("is-collapsed");
    toggle.hidden = false;
    toggle.addEventListener("click", function () {
      var collapsed = sig.classList.toggle("is-collapsed");
      toggle.setAttribute("aria-expanded", collapsed ? "false" : "true");
      toggle.textContent = toggle.getAttribute(collapsed ? "data-label-more" : "data-label-less");
    });
  }
  function init() {
    Array.prototype.forEach.call(document.querySelectorAll("[data-surface-explorer]"), enhance);
    Array.prototype.forEach.call(document.querySelectorAll("[data-cli-sig-long]"), enhanceSignature);
    Array.prototype.forEach.call(document.querySelectorAll("[data-surface-proof]"), enhanceProof);
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
