(function () {
  "use strict";

  function initShowcase() {
    var form = document.getElementById("showcase-controls");
    if (!form) return;
    var query = document.getElementById("showcase-query");
    var format = document.getElementById("showcase-format");
    var capability = document.getElementById("showcase-capability");
    var count = document.getElementById("showcase-result-count");
    var empty = document.getElementById("showcase-empty");
    var cards = Array.prototype.slice.call(document.querySelectorAll("[data-showcase-format]"));

    function normalize(value) {
      return value.normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
    }

    var examples = cards.map(function (card) {
      return {
        element: card,
        format: card.getAttribute("data-showcase-format"),
        capabilities: (card.getAttribute("data-showcase-capabilities") || "").split(/\s+/),
        text: normalize(card.getAttribute("data-showcase-search") || "")
      };
    });

    function selectKnown(select, value) {
      var known = Array.prototype.some.call(select.options, function (option) { return option.value === value; });
      select.value = known ? value : "all";
    }

    function applyFilters() {
      var terms = normalize(query.value).split(/\s+/).filter(Boolean);
      var visible = 0;
      examples.forEach(function (example) {
        var matches = (format.value === "all" || format.value === example.format) &&
          (capability.value === "all" || example.capabilities.indexOf(capability.value) !== -1) &&
          terms.every(function (term) { return example.text.indexOf(term) !== -1; });
        example.element.hidden = !matches;
        if (matches) visible += 1;
      });
      count.textContent = visible === examples.length ? visible + " examples" :
        visible + " of " + examples.length + " examples";
      empty.hidden = visible !== 0;
    }

    function writeLocation(push, preserveHash) {
      var url = new URL(window.location.href);
      [["q", query.value.trim()], ["format", format.value], ["capability", capability.value]].forEach(function (entry) {
        if (entry[1] && entry[1] !== "all") url.searchParams.set(entry[0], entry[1]);
        else url.searchParams.delete(entry[0]);
      });
      if (!preserveHash) url.hash = "";
      if (url.href !== window.location.href) {
        window.history[push ? "pushState" : "replaceState"](window.history.state, "", url.pathname + url.search + url.hash);
      }
    }

    function openLinkedExample() {
      var id;
      try { id = decodeURIComponent(window.location.hash.slice(1)); } catch (_) { return; }
      var example = examples.find(function (entry) { return entry.element.id === id; });
      if (!example) return;
      if (example.element.hidden) {
        query.value = "";
        format.value = example.format;
        capability.value = "all";
        applyFilters();
        writeLocation(false, true);
      }
      var details = example.element.querySelector(".imo-example__details");
      if (details) details.open = true;
      example.element.scrollIntoView({ block: "start", behavior: "instant" });
    }

    function readLocation() {
      var params = new URLSearchParams(window.location.search);
      query.value = (params.get("q") || "").slice(0, 160);
      selectKnown(format, params.get("format"));
      selectKnown(capability, params.get("capability"));
      applyFilters();
      openLinkedExample();
    }

    form.addEventListener("submit", function (event) { event.preventDefault(); });
    form.addEventListener("reset", function (event) {
      event.preventDefault();
      query.value = "";
      format.value = "all";
      capability.value = "all";
      applyFilters();
      writeLocation(true, false);
    });
    query.addEventListener("input", function () {
      applyFilters();
      writeLocation(false, false);
    });
    [format, capability].forEach(function (select) {
      select.addEventListener("change", function () {
        applyFilters();
        writeLocation(true, false);
      });
    });
    window.addEventListener("popstate", readLocation);
    window.addEventListener("hashchange", openLinkedExample);
    readLocation();
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", initShowcase);
  else initShowcase();
})();
