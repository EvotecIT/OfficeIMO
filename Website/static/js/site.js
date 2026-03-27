/* OfficeIMO – Site JS – Liquid Glass Premium */
(function () {
  "use strict";

  var THEME_KEY = "imo-theme";

  /* === Theme Toggle === */
  function applyTheme(mode) {
    var resolved = mode === "auto" ? (matchMedia("(prefers-color-scheme:dark)").matches ? "dark" : "light") : mode;
    document.documentElement.setAttribute("data-theme", resolved);
    document.querySelectorAll(".imo-theme-toggle").forEach(function (btn) {
      var sun = btn.querySelector(".icon-sun");
      var moon = btn.querySelector(".icon-moon");
      if (sun) sun.style.display = resolved === "dark" ? "none" : "block";
      if (moon) moon.style.display = resolved === "dark" ? "block" : "none";
    });
  }

  function initTheme() {
    var stored = localStorage.getItem(THEME_KEY) || "dark";
    applyTheme(stored);
    document.querySelectorAll(".imo-theme-toggle").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var cur = localStorage.getItem(THEME_KEY) || "dark";
        var next = cur === "dark" ? "light" : cur === "light" ? "auto" : "dark";
        localStorage.setItem(THEME_KEY, next);
        applyTheme(next);
      });
    });
  }

  /* === Mobile Nav === */
  function initMobileNav() {
    var hamburger = document.querySelector(".imo-hamburger");
    var nav = document.querySelector(".imo-nav");
    if (!hamburger || !nav) return;
    hamburger.addEventListener("click", function () {
      hamburger.classList.toggle("is-active");
      nav.classList.toggle("is-open");
      document.body.style.overflow = nav.classList.contains("is-open") ? "hidden" : "";
    });
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && nav.classList.contains("is-open")) {
        hamburger.classList.remove("is-active");
        nav.classList.remove("is-open");
        document.body.style.overflow = "";
      }
    });
  }

  /* === Dropdown Menus === */
  function initDropdowns() {
    var items = document.querySelectorAll(".imo-nav__item");
    items.forEach(function (item) {
      var btn = item.querySelector("button.imo-nav__link");
      if (!btn) return;

      // On mobile: toggle on click
      btn.addEventListener("click", function (e) {
        if (window.innerWidth < 1024) {
          e.preventDefault();
          e.stopPropagation();
          items.forEach(function (other) { if (other !== item) other.classList.remove("is-open"); });
          item.classList.toggle("is-open");
        }
      });
    });

    document.addEventListener("click", function (e) {
      if (window.innerWidth < 1024 && !e.target.closest(".imo-nav__item")) {
        items.forEach(function (item) { item.classList.remove("is-open"); });
      }
    });
  }

  /* === Code Copy === */
  function initCodeCopy() {
    document.addEventListener("click", function (e) {
      var btn = e.target.closest(".imo-install__copy, [data-copy]");
      if (!btn) return;
      var text = btn.getAttribute("data-copy");
      if (!text) {
        var code = btn.closest(".imo-install");
        if (code) text = code.querySelector(".imo-install__code").textContent.trim();
      }
      if (!text) return;
      if (navigator.clipboard) {
        navigator.clipboard.writeText(text).then(function () {
          btn.classList.add("is-copied");
          setTimeout(function () { btn.classList.remove("is-copied"); }, 2000);
        });
      }
    });
  }

  /* === Tab Switching (Code Examples) === */
  function initTabs() {
    // Find tab containers by role="tablist"
    document.querySelectorAll('[role="tablist"]').forEach(function (tablist) {
      var tabs = tablist.querySelectorAll('[role="tab"]');
      tabs.forEach(function (tab) {
        tab.addEventListener("click", function () {
          var panelId = tab.getAttribute("aria-controls");
          var panel = panelId ? document.getElementById(panelId) : null;

          // Deactivate all tabs and panels in this group
          tabs.forEach(function (t) {
            t.classList.remove("is-active");
            t.setAttribute("aria-selected", "false");
            t.setAttribute("tabindex", "-1");
            var p = document.getElementById(t.getAttribute("aria-controls"));
            if (p) { p.classList.remove("is-active"); p.hidden = true; }
          });

          // Activate clicked tab and its panel
          tab.classList.add("is-active");
          tab.setAttribute("aria-selected", "true");
          tab.removeAttribute("tabindex");
          if (panel) { panel.classList.add("is-active"); panel.hidden = false; }
        });
      });
    });
  }

  /* === Header Scroll Shadow === */
  function initHeaderScroll() {
    var header = document.querySelector(".imo-header");
    if (!header) return;
    window.addEventListener("scroll", function () {
      header.classList.toggle("is-scrolled", window.scrollY > 10);
    }, { passive: true });
  }

  /* === Docs Sidebar === */
  function initDocsSidebar() {
    var toggle = document.querySelector(".imo-docs__sidebar-toggle");
    var sidebar = document.querySelector(".imo-docs__sidebar");
    if (toggle && sidebar) {
      toggle.addEventListener("click", function () {
        sidebar.classList.toggle("is-open");
      });
    }

    // Auto-open the group containing the current page
    var currentPath = window.location.pathname;
    document.querySelectorAll(".imo-docs__group").forEach(function (group) {
      var links = group.querySelectorAll(".imo-docs__link");
      var hasActive = false;
      links.forEach(function (link) {
        var href = link.getAttribute("href");
        if (href && currentPath === href) {
          hasActive = true;
          link.classList.add("active");
          link.setAttribute("aria-current", "page");
        }
      });
      if (hasActive) {
        group.setAttribute("open", "");
      }
    });

    // Also check top-level links
    document.querySelectorAll(".imo-docs__link--top").forEach(function (link) {
      var href = link.getAttribute("href");
      if (href && currentPath === href) {
        link.classList.add("active");
      }
    });
  }

  /* === Prism Manual Trigger === */
  function initPrism() {
    // Prism autoloader handles language detection automatically
    // Just ensure it runs after DOM is ready
    if (typeof Prism !== "undefined" && Prism.highlightAll) {
      setTimeout(function () { Prism.highlightAll(); }, 100);
    }
  }

  /* === Init === */
  function init() {
    initTheme();
    initMobileNav();
    initDropdowns();
    initCodeCopy();
    initTabs();
    initHeaderScroll();
    initDocsSidebar();
    initPrism();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
