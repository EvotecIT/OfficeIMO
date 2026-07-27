(function () {
  "use strict";

  function directContentRoot(container) {
    try {
      return container.querySelector(":scope > article") || container;
    } catch (_) {
      return container.firstElementChild && container.firstElementChild.tagName === "ARTICLE"
        ? container.firstElementChild
        : container;
    }
  }

  function markIntroduction(source) {
    var paragraphs = Array.prototype.filter.call(source.children, function (element) {
      return element.tagName === "P";
    });

    if (paragraphs[0]) {
      paragraphs[0].classList.add("imo-comparison-detail-intro");
    }
    if (paragraphs[1]) {
      paragraphs[1].classList.add("imo-comparison-detail-evidence");
    }
  }

  function enhanceTables(source) {
    var tables = Array.prototype.slice.call(source.querySelectorAll(":scope > table"));

    tables.forEach(function (table, index) {
      var heading = table.previousElementSibling;
      var shell = document.createElement("section");
      var toolbar = document.createElement("div");
      var title = document.createElement("span");
      var cue = document.createElement("span");
      var scroll = document.createElement("div");
      var label = heading && /^H[2-4]$/.test(heading.tagName)
        ? heading.textContent.trim()
        : "Comparison matrix";

      shell.className = "imo-comparison-matrix";
      toolbar.className = "imo-comparison-matrix__toolbar";
      title.className = "imo-comparison-matrix__title";
      cue.className = "imo-comparison-matrix__cue";
      scroll.className = "imo-comparison-matrix__scroll";
      title.textContent = label;
      cue.innerHTML = "<span aria-hidden=\"true\">↔</span> Scroll to compare";
      scroll.tabIndex = 0;
      scroll.setAttribute("role", "region");
      scroll.setAttribute("aria-label", label + " table");

      table.parentNode.insertBefore(shell, table);
      toolbar.appendChild(title);
      toolbar.appendChild(cue);
      shell.appendChild(toolbar);
      shell.appendChild(scroll);
      scroll.appendChild(table);

      function updateScrollState() {
        var scrollable = scroll.scrollWidth > scroll.clientWidth + 2;
        shell.classList.toggle("is-scrollable", scrollable);
        shell.classList.toggle("is-scrolled", scroll.scrollLeft > 2);
      }

      scroll.addEventListener("scroll", updateScrollState, { passive: true });
      window.addEventListener("resize", updateScrollState, { passive: true });
      window.requestAnimationFrame(updateScrollState);

      table.setAttribute("data-comparison-table", String(index + 1));
    });
  }

  function decisionRecord(heading) {
    var firstContent = heading.nextElementSibling;
    var list = firstContent && (firstContent.tagName === "UL" || firstContent.tagName === "OL")
      ? firstContent
      : null;
    var note = list ? list.nextElementSibling : firstContent;

    if (!note || note.tagName !== "P") {
      note = null;
    }
    if (!list && !note) {
      return null;
    }

    var text = heading.textContent.trim();
    var isCombined = /\b(use both|use together|clear owner|combine|hybrid|at the edge)\b/i.test(text);
    var isOfficeImo = /\bofficeimo\b/i.test(text);
    var isChoice = /^(choose|choosing|when|where)\b/i.test(text);

    return {
      heading: heading,
      list: list,
      note: note,
      label: isCombined
        ? "Combined architecture"
        : (isOfficeImo ? "OfficeIMO fit" : (isChoice ? "Best fit" : "Validation gate")),
      variant: isCombined
        ? "combined"
        : (isOfficeImo ? "officeimo" : (isChoice ? "alternative" : "validation"))
    };
  }

  function enhanceChoices(source) {
    var headings = Array.prototype.filter.call(source.children, function (element) {
      if (element.tagName !== "H2") {
        return false;
      }
      var content = element.nextElementSibling;
      return content && (content.tagName === "UL" || content.tagName === "OL" || content.tagName === "P");
    });
    var records = headings.map(decisionRecord).filter(Boolean);

    if (!records.length) {
      return;
    }

    var grid = document.createElement("section");
    grid.id = "comparison-recommendations";
    grid.className = "imo-comparison-choice-grid";
    grid.setAttribute("aria-label", "Decision recommendations");
    records[0].heading.parentNode.insertBefore(grid, records[0].heading);

    records.forEach(function (record) {
      var card = document.createElement("section");
      var label = document.createElement("span");

      card.className = "imo-comparison-choice-card imo-comparison-choice-card--" + record.variant;
      label.className = "imo-comparison-choice-card__label";
      label.textContent = record.label;
      card.appendChild(label);
      card.appendChild(record.heading);
      if (record.list) {
        card.appendChild(record.list);
      }
      if (record.note) {
        card.appendChild(record.note);
      }
      grid.appendChild(card);
    });
  }

  function initialize() {
    var page = document.querySelector("[data-comparison-page]");
    var container = page && page.querySelector(".imo-comparison-detail-content");
    if (!container) {
      return;
    }

    var source = directContentRoot(container);
    markIntroduction(source);
    enhanceTables(source);
    enhanceChoices(source);
    container.classList.add("is-enhanced");
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }
}());
