// Turns runs of "### C#", "### PowerShell" and "### CLI" sections in a docs article into one tab group.
// The reader's choice is remembered and applied to every group on every page; without script the sections stay listed.
(function () {
  "use strict";
  var LANGS = { "c#": "csharp", "powershell": "powershell", "cli": "cli" };
  var KEY = "imo-docs-language";
  function stored() { try { return localStorage.getItem(KEY); } catch (e) { return null; } }
  function store(v) { try { localStorage.setItem(KEY, v); } catch (e) { /* storage unavailable */ } }
  function langOf(node) {
    if (!node || node.tagName !== "H3") return null;
    return LANGS[node.textContent.trim().toLowerCase()] || null;
  }
  var groups = [];
  function build(run) {
    if (run.length < 2) return;
    var group = document.createElement("div");
    group.className = "imo-lang-tabs";
    var list = document.createElement("div");
    list.className = "imo-lang-tabs__list";
    list.setAttribute("role", "tablist");
    list.setAttribute("aria-label", "Language");
    group.appendChild(list);
    run[0].heading.parentNode.insertBefore(group, run[0].heading);
    var entries = run.map(function (section, i) {
      var id = (section.heading.id || ("lang-" + groups.length + "-" + i));
      var tab = document.createElement("button");
      tab.type = "button";
      tab.setAttribute("role", "tab");
      tab.id = id + "-tab";
      tab.textContent = section.heading.textContent.trim();
      tab.dataset.lang = section.lang;
      var panel = document.createElement("div");
      panel.className = "imo-lang-tabs__panel";
      panel.setAttribute("role", "tabpanel");
      panel.setAttribute("aria-labelledby", tab.id);
      panel.id = id + "-panel";
      tab.setAttribute("aria-controls", panel.id);
      section.nodes.forEach(function (n) { panel.appendChild(n); });
      section.heading.classList.add("imo-lang-tabs__heading");
      panel.insertBefore(section.heading, panel.firstChild);
      list.appendChild(tab);
      group.appendChild(panel);
      return { tab: tab, panel: panel, lang: section.lang };
    });
    function select(lang, focus) {
      var hit = entries.some(function (e) { return e.lang === lang; }) ? lang : entries[0].lang;
      entries.forEach(function (e) {
        var on = e.lang === hit;
        e.tab.setAttribute("aria-selected", on ? "true" : "false");
        e.tab.tabIndex = on ? 0 : -1;
        e.panel.hidden = !on;
        if (on && focus) e.tab.focus();
      });
    }
    entries.forEach(function (e, i) {
      e.tab.addEventListener("click", function () { store(e.lang); groups.forEach(function (g) { g(e.lang); }); });
      e.tab.addEventListener("keydown", function (ev) {
        var step = ev.key === "ArrowRight" ? 1 : ev.key === "ArrowLeft" ? -1 : 0;
        if (!step) return;
        ev.preventDefault();
        var next = entries[(i + step + entries.length) % entries.length];
        store(next.lang);
        groups.forEach(function (g) { g(next.lang); });
        next.tab.focus();
      });
    });
    groups.push(select);
  }
  function init() {
    var root = document.querySelector(".imo-docs__article, .imo-docs__main article, main article");
    if (!root) return;
    var headings = Array.prototype.slice.call(root.querySelectorAll("h3"));
    var used = new Set();
    headings.forEach(function (h) {
      if (used.has(h) || !langOf(h)) return;
      var run = [];
      var node = h;
      while (node && langOf(node) && node.parentNode === h.parentNode) {
        var section = { heading: node, lang: langOf(node), nodes: [] };
        used.add(node);
        var next = node.nextElementSibling;
        while (next && !/^H[1-3]$/.test(next.tagName)) { section.nodes.push(next); next = next.nextElementSibling; }
        run.push(section);
        node = next;
      }
      build(run);
    });
    var pick = stored() || "csharp";
    groups.forEach(function (g) { g(pick); });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
