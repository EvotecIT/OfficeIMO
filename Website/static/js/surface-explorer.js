// Turns each [data-surface-explorer] list of groups into a tablist with one visible panel,
// adds a command finder over the PowerShell command lists,
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
    enhanceFilter(root, tabs, panels, select);
  }
  // The finder narrows every family to the commands whose name, aliases or synopsis contain all typed words.
  // Clearing it restores the tab that was selected before.
  function enhanceFilter(root, tabs, panels, select) {
    var box = root.querySelector("[data-ps-filter]");
    var input = box && box.querySelector("input");
    var status = box && box.querySelector(".imo-ps-filter__status");
    if (!input) return;
    var idle = status.textContent;
    var rows = Array.prototype.slice.call(root.querySelectorAll("[data-ps-command]")).map(function (row) {
      return { el: row, text: row.textContent.replace(/\s+/g, " ").toLowerCase() };
    });
    var groups = Array.prototype.slice.call(root.querySelectorAll("[data-ps-group]"));
    var lists = Array.prototype.slice.call(root.querySelectorAll("[data-ps-commands]"));
    var counts = panels.map(function (panel) {
      var small = panel.querySelector("h3 small");
      return { el: small, text: small ? small.textContent : "" };
    });
    var opened = null;
    var selected = 0;
    function apply() {
      var words = input.value.toLowerCase().trim().split(/\s+/).filter(Boolean);
      var filtering = words.length > 0;
      if (filtering && !opened) {
        opened = lists.map(function (list) { return list.open; });
        tabs.forEach(function (tab, i) { if (tab.getAttribute("aria-selected") === "true") selected = i; });
        root.classList.add("is-filtering");
      }
      if (!filtering) {
        if (!opened) return;
        rows.forEach(function (row) { row.el.hidden = false; });
        groups.forEach(function (group) { group.hidden = false; });
        lists.forEach(function (list, i) { list.open = opened[i]; });
        counts.forEach(function (count) { if (count.el) count.el.textContent = count.text; });
        opened = null;
        root.classList.remove("is-filtering");
        select(selected, false);
        status.textContent = idle;
        return;
      }
      rows.forEach(function (row) {
        row.el.hidden = !words.every(function (word) { return row.text.indexOf(word) !== -1; });
      });
      groups.forEach(function (group) { group.hidden = !group.querySelector("[data-ps-command]:not([hidden])"); });
      var total = 0;
      var families = 0;
      panels.forEach(function (panel, i) {
        var found = panel.querySelectorAll("[data-ps-command]:not([hidden])").length;
        panel.hidden = found === 0;
        var list = panel.querySelector("[data-ps-commands]");
        if (list) list.open = true;
        if (counts[i].el) counts[i].el.textContent = found + " of " + counts[i].text;
        total += found;
        if (found) families++;
      });
      status.textContent = total === 0 ? "No command matches “" + input.value.trim() + "”." :
        total + (total === 1 ? " command" : " commands") + " in " + families + (families === 1 ? " family" : " families") + " match.";
    }
    input.addEventListener("input", apply);
    input.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && input.value) { e.preventDefault(); input.value = ""; apply(); }
    });
    box.hidden = false;
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
