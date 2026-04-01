/* OfficeIMO – Site JS – Liquid Glass Premium */
(function () {
  "use strict";

  var THEME_KEY = "imo-theme";
  var PRISM_LANGUAGES_PATH = "/assets/prism/components/";

  window.Prism = window.Prism || {};
  window.Prism.manual = true;

  function getPreferredTheme() {
    return matchMedia("(prefers-color-scheme: light)").matches ? "light" : "dark";
  }

  function applyTheme(mode) {
    var resolved = mode === "light" || mode === "dark" ? mode : getPreferredTheme();
    document.documentElement.setAttribute("data-theme", resolved);
    document.documentElement.style.colorScheme = resolved;
    document.querySelectorAll(".imo-theme-toggle").forEach(function (btn) {
      var sun = btn.querySelector(".icon-sun");
      var moon = btn.querySelector(".icon-moon");
      if (sun) sun.style.display = resolved === "dark" ? "none" : "block";
      if (moon) moon.style.display = resolved === "dark" ? "block" : "none";
    });
  }

  function initTheme() {
    var stored = localStorage.getItem(THEME_KEY);
    if (stored !== "light" && stored !== "dark") {
      stored = getPreferredTheme();
    }
    applyTheme(stored);
    document.querySelectorAll(".imo-theme-toggle").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var cur = document.documentElement.getAttribute("data-theme") || getPreferredTheme();
        var next = cur === "dark" ? "light" : "dark";
        localStorage.setItem(THEME_KEY, next);
        applyTheme(next);
      });
    });
  }

  function initHeaderLinks() {
    function normalizePath(pathname) {
      return pathname.replace(/\/+$/, "") || "/";
    }

    function getLocalTarget(link) {
      var href = link.getAttribute("href");
      if (!href || href.charAt(0) === "#") return null;

      try {
        var target = new URL(href, window.location.origin);
        if (target.origin !== window.location.origin) return null;
        return target;
      } catch (error) {
        return null;
      }
    }

    var current = normalizePath(window.location.pathname);
    var navItems = document.querySelectorAll(".imo-header .imo-nav__item");

    navItems.forEach(function (item) {
      var topLink = item.querySelector(":scope > a.imo-nav__link");
      var topButton = item.querySelector(":scope > button.imo-nav__link");
      var dropdownLinks = item.querySelectorAll(".imo-dropdown a[href]");
      var isActive = false;

      if (topLink) {
        var topTarget = getLocalTarget(topLink);
        if (topTarget) {
          var topPath = normalizePath(topTarget.pathname);
          if (topPath === "/") {
            isActive = current === "/";
          } else {
            isActive = current === topPath || current.indexOf(topPath + "/") === 0;
          }

          topLink.addEventListener("click", function (event) {
            var samePath = topPath === current;
            var sameSearch = topTarget.search === window.location.search;
            var noHash = !topTarget.hash;
            if (samePath && sameSearch && noHash) {
              event.preventDefault();
              window.scrollTo({ top: 0, behavior: "smooth" });
            }
          });
        }
      }

      if (!isActive && dropdownLinks.length) {
        dropdownLinks.forEach(function (link) {
          var target = getLocalTarget(link);
          if (!target) return;

          var path = normalizePath(target.pathname);
          if (current === path || current.indexOf(path + "/") === 0) {
            isActive = true;
            link.classList.add("is-active");
            if (!link.hasAttribute("aria-current")) {
              link.setAttribute("aria-current", "page");
            }
          }

          link.addEventListener("click", function (event) {
            var samePath = path === current;
            var sameSearch = target.search === window.location.search;
            var noHash = !target.hash;
            if (samePath && sameSearch && noHash) {
              event.preventDefault();
              window.scrollTo({ top: 0, behavior: "smooth" });
            }
          });
        });
      }

      if (isActive) {
        item.classList.add("is-current");
        if (topLink) {
          topLink.classList.add("is-active");
          topLink.setAttribute("aria-current", "page");
        }
        if (topButton) {
          topButton.classList.add("is-active");
        }
      }
    });
  }

  function initMobileNav() {
    var hamburger = document.querySelector(".imo-hamburger");
    var nav = document.querySelector(".imo-nav");
    if (!hamburger || !nav) return;

    function closeNav() {
      hamburger.classList.remove("is-active");
      hamburger.setAttribute("aria-expanded", "false");
      nav.classList.remove("is-open");
      document.body.style.overflow = "";
    }

    function openNav() {
      hamburger.classList.add("is-active");
      hamburger.setAttribute("aria-expanded", "true");
      nav.classList.add("is-open");
      document.body.style.overflow = "hidden";
    }

    hamburger.setAttribute("aria-expanded", "false");

    hamburger.addEventListener("click", function () {
      if (nav.classList.contains("is-open")) {
        closeNav();
      } else {
        openNav();
      }
    });

    nav.querySelectorAll("a[href]").forEach(function (link) {
      link.addEventListener("click", function () {
        closeNav();
      });
    });

    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && nav.classList.contains("is-open")) {
        closeNav();
      }
    });

    window.addEventListener("resize", function () {
      if (window.innerWidth >= 1024 && nav.classList.contains("is-open")) {
        closeNav();
      }
    });
  }

  function initDropdowns() {
    var items = document.querySelectorAll(".imo-nav__item");
    items.forEach(function (item) {
      var btn = item.querySelector("button.imo-nav__link");
      if (!btn) return;

      btn.addEventListener("click", function (e) {
        if (window.innerWidth < 1024) {
          e.preventDefault();
          e.stopPropagation();
          items.forEach(function (other) {
            if (other !== item) other.classList.remove("is-open");
          });
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
        }).catch(function () {
          btn.classList.remove("is-copied");
        });
      }
    });
  }

  function initTabs() {
    document.querySelectorAll('[role="tablist"]').forEach(function (tablist) {
      var tabs = tablist.querySelectorAll('[role="tab"]');
      tabs.forEach(function (tab) {
        tab.addEventListener("click", function () {
          var panelId = tab.getAttribute("aria-controls");
          var panel = panelId ? document.getElementById(panelId) : null;

          tabs.forEach(function (t) {
            t.classList.remove("is-active");
            t.setAttribute("aria-selected", "false");
            t.setAttribute("tabindex", "-1");
            var p = document.getElementById(t.getAttribute("aria-controls"));
            if (p) {
              p.classList.remove("is-active");
              p.hidden = true;
            }
          });

          tab.classList.add("is-active");
          tab.setAttribute("aria-selected", "true");
          tab.removeAttribute("tabindex");
          if (panel) {
            panel.classList.add("is-active");
            panel.hidden = false;
          }
        });
      });
    });
  }

  function initHeaderScroll() {
    var header = document.querySelector(".imo-header");
    if (!header) return;
    window.addEventListener("scroll", function () {
      header.classList.toggle("is-scrolled", window.scrollY > 10);
    }, { passive: true });
  }

  function initDocsSidebar() {
    var toggle = document.querySelector(".imo-docs__sidebar-toggle");
    var sidebar = document.querySelector(".imo-docs__sidebar");
    var overlay = document.querySelector(".imo-docs__sidebar-overlay");

    if (toggle && sidebar) {
      function closeSidebar() {
        sidebar.classList.remove("is-open");
        toggle.setAttribute("aria-expanded", "false");
        if (overlay) overlay.hidden = true;
        document.body.style.overflow = "";
      }

      function openSidebar() {
        sidebar.classList.add("is-open");
        toggle.setAttribute("aria-expanded", "true");
        if (overlay) overlay.hidden = false;
        document.body.style.overflow = "hidden";
      }

      toggle.setAttribute("aria-expanded", "false");
      toggle.addEventListener("click", function () {
        if (sidebar.classList.contains("is-open")) {
          closeSidebar();
        } else {
          openSidebar();
        }
      });

      if (overlay) {
        overlay.hidden = true;
        overlay.addEventListener("click", closeSidebar);
      }

      sidebar.querySelectorAll("a[href]").forEach(function (link) {
        link.addEventListener("click", function () {
          if (window.innerWidth < 1024) {
            closeSidebar();
          }
        });
      });

      document.addEventListener("keydown", function (e) {
        if (e.key === "Escape" && sidebar.classList.contains("is-open")) {
          closeSidebar();
        }
      });

      window.addEventListener("resize", function () {
        if (window.innerWidth >= 1024) {
          closeSidebar();
        }
      });
    }

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

    document.querySelectorAll(".imo-docs__link--top").forEach(function (link) {
      var href = link.getAttribute("href");
      if (href && currentPath === href) {
        link.classList.add("active");
      }
    });
  }

  function initPrism() {
    var attempts = 0;

    function configurePrism() {
      if (typeof Prism === "undefined") return false;
      Prism.manual = true;
      Prism.plugins = Prism.plugins || {};
      if (Prism.plugins.autoloader && Prism.plugins.autoloader.languages_path !== PRISM_LANGUAGES_PATH) {
        Prism.plugins.autoloader.languages_path = PRISM_LANGUAGES_PATH;
      }
      return typeof Prism.highlightAll === "function" || typeof Prism.highlightAllUnder === "function";
    }

    function highlight() {
      attempts += 1;
      if (!configurePrism()) {
        if (attempts < 10) {
          setTimeout(highlight, 60);
        }
        return;
      }

      if (typeof Prism.highlightAllUnder === "function") {
        Prism.highlightAllUnder(document);
      } else if (typeof Prism.highlightAll === "function") {
        Prism.highlightAll();
      }
    }

    highlight();
  }

  function init() {
    initTheme();
    initMobileNav();
    initDropdowns();
    initHeaderLinks();
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
