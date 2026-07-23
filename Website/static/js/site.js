/* OfficeIMO – Site JS – Liquid Glass Premium */
(function () {
  "use strict";

  var THEME_KEY = "imo-theme";
  var PRISM_LANGUAGES_PATH = "/assets/prism/components/";

  window.Prism = window.Prism || {};
  window.Prism.manual = true;

  function getDefaultTheme() {
    return "light";
  }

  function applyTheme(mode) {
    var resolved = mode === "light" || mode === "dark" ? mode : getDefaultTheme();
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
      stored = getDefaultTheme();
    }
    applyTheme(stored);
    document.querySelectorAll(".imo-theme-toggle").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var cur = document.documentElement.getAttribute("data-theme") || getDefaultTheme();
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
    var navLinks = Array.prototype.slice.call(document.querySelectorAll(".imo-header .imo-nav a[href]"));
    var matchingLinks = [];

    navLinks.forEach(function (link) {
      var target = getLocalTarget(link);
      if (!target) return;

      var path = normalizePath(target.pathname);
      var matches = path === "/"
        ? current === "/"
        : current === path || current.indexOf(path + "/") === 0;
      if (matches) {
        var isTopLevel = link.classList.contains("imo-nav__link") &&
          link.parentElement && link.parentElement.classList.contains("imo-nav__item");
        matchingLinks.push({ link: link, path: path, exact: current === path, topLevel: isTopLevel });
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

    matchingLinks.sort(function (left, right) {
      if (left.exact !== right.exact) return left.exact ? -1 : 1;
      if (left.topLevel !== right.topLevel) return left.topLevel ? -1 : 1;
      return right.path.length - left.path.length;
    });
    var activeLink = matchingLinks.length ? matchingLinks[0].link : null;
    var navItems = document.querySelectorAll(".imo-header .imo-nav__item");

    navItems.forEach(function (item) {
      var topLink = item.querySelector(":scope > a.imo-nav__link");
      var topButton = item.querySelector(":scope > button.imo-nav__link");
      var dropdownLinks = item.querySelectorAll(".imo-dropdown a[href]");
      var isActive = activeLink && item.contains(activeLink);

      dropdownLinks.forEach(function (link) {
        if (link !== activeLink) return;
        link.classList.add("is-active");
        link.setAttribute("aria-current", "page");
      });

      if (isActive) {
        item.classList.add("is-current");
        if (topLink === activeLink) {
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
    var items = Array.prototype.slice.call(document.querySelectorAll(".imo-nav__item"));

    function positionDropdown(item) {
      var menu = item.querySelector(":scope > .imo-dropdown");
      if (!menu) return;

      menu.style.removeProperty("--imo-menu-shift");
      var rect = menu.getBoundingClientRect();
      var gutter = 12;
      var shift = 0;
      if (rect.left < gutter) shift += gutter - rect.left;
      if (rect.right + shift > window.innerWidth - gutter) {
        shift -= rect.right + shift - (window.innerWidth - gutter);
      }
      if (shift) menu.style.setProperty("--imo-menu-shift", shift + "px");
    }

    function setOpen(item, open) {
      var btn = item.querySelector("button.imo-nav__link");
      if (!btn) return;
      item.classList.toggle("is-open", open);
      btn.setAttribute("aria-expanded", open ? "true" : "false");
      if (!open) delete item.dataset.openedByHover;
      var menu = item.querySelector(":scope > .imo-dropdown");
      if (!open || window.innerWidth < 1024) {
        if (menu) menu.style.removeProperty("--imo-menu-shift");
        return;
      }
      window.requestAnimationFrame(function () {
        if (item.classList.contains("is-open")) positionDropdown(item);
      });
    }

    function closeAll(except) {
      items.forEach(function (item) {
        if (item !== except) setOpen(item, false);
      });
    }

    items.forEach(function (item) {
      var btn = item.querySelector("button.imo-nav__link");
      if (!btn) return;

      btn.addEventListener("click", function (e) {
        e.preventDefault();
        e.stopPropagation();
        var wasOpenedByHover = item.dataset.openedByHover === "true";
        delete item.dataset.openedByHover;
        var willOpen = wasOpenedByHover || !item.classList.contains("is-open");
        closeAll(item);
        setOpen(item, willOpen);
      });

      item.addEventListener("mouseenter", function () {
        if (window.innerWidth >= 1024 && window.matchMedia("(hover: hover)").matches) {
          if (item.classList.contains("is-open")) return;
          closeAll(item);
          setOpen(item, true);
          item.dataset.openedByHover = "true";
        }
      });

      item.addEventListener("mouseleave", function () {
        if (window.innerWidth >= 1024 && window.matchMedia("(hover: hover)").matches && !item.contains(document.activeElement)) {
          setOpen(item, false);
        }
      });

      item.addEventListener("focusout", function () {
        if (window.innerWidth < 1024) return;
        window.requestAnimationFrame(function () {
          if (!item.contains(document.activeElement)) setOpen(item, false);
        });
      });
    });

    document.addEventListener("click", function (e) {
      if (!e.target.closest(".imo-nav__item")) closeAll();
    });

    document.addEventListener("keydown", function (e) {
      if (e.key !== "Escape") return;
      var openItem = document.querySelector(".imo-nav__item.is-open");
      if (!openItem) return;
      var trigger = openItem.querySelector("button.imo-nav__link");
      closeAll();
      if (trigger) trigger.focus();
    });

    window.addEventListener("resize", function () {
      closeAll();
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

  function initConverterFrame() {
    var frame = document.querySelector('.imo-converter-launch__frame');
    if (!frame) return;

    var observer = null;

    function syncTheme() {
      try {
        var frameRoot = frame.contentDocument && frame.contentDocument.documentElement;
        if (!frameRoot) return;
        var theme = document.documentElement.getAttribute('data-theme') || 'light';
        frameRoot.setAttribute('data-theme', theme);
        frameRoot.style.colorScheme = theme;
      } catch (error) {
        // Same-origin in production; keep the converter's saved preference otherwise.
      }
    }

    function syncHeight() {
      try {
        var frameDocument = frame.contentDocument;
        if (!frameDocument) return;

        var documentHeight = frameDocument.documentElement ? frameDocument.documentElement.scrollHeight : 0;
        var bodyHeight = frameDocument.body ? frameDocument.body.scrollHeight : 0;
        var height = Math.max(documentHeight, bodyHeight);
        if (height > 0) frame.style.height = Math.ceil(height) + 'px';
      } catch (error) {
        // The converter is same-origin in production. Keep the CSS fallback if a preview host changes that.
      }
    }

    function observeFrame() {
      if (observer) observer.disconnect();
      syncTheme();
      syncHeight();

      try {
        var frameWindow = frame.contentWindow;
        var frameDocument = frame.contentDocument;
        if (frameWindow && frameWindow.ResizeObserver && frameDocument) {
          observer = new frameWindow.ResizeObserver(syncHeight);
          if (frameDocument.documentElement) observer.observe(frameDocument.documentElement);
          if (frameDocument.body) observer.observe(frameDocument.body);
        }
      } catch (error) {
        // The fixed minimum height remains usable if same-origin observation is unavailable.
      }

      window.setTimeout(syncHeight, 250);
      window.setTimeout(syncHeight, 1000);
    }

    frame.addEventListener('load', observeFrame);
    window.addEventListener('resize', syncHeight);
    new MutationObserver(syncTheme).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });
    observeFrame();
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

  function initShowcaseFilters() {
    var controls = Array.prototype.slice.call(document.querySelectorAll("[data-showcase-filter]"));
    var cards = Array.prototype.slice.call(document.querySelectorAll("[data-showcase-tags]"));
    var resultCount = document.getElementById("showcase-result-count");
    if (!controls.length || !cards.length) return;

    function applyFilter(filter) {
      var visible = 0;

      cards.forEach(function (card) {
        var tags = (card.getAttribute("data-showcase-tags") || "").split(/\s+/);
        var matches = filter === "all" || tags.indexOf(filter) !== -1;
        card.hidden = !matches;
        if (matches) visible += 1;
      });

      controls.forEach(function (control) {
        var active = control.getAttribute("data-showcase-filter") === filter;
        control.classList.toggle("is-active", active);
        control.setAttribute("aria-pressed", active ? "true" : "false");
      });

      if (resultCount) {
        resultCount.textContent = visible + (visible === 1 ? " workflow" : " workflows");
      }
    }

    controls.forEach(function (control) {
      control.addEventListener("click", function () {
        applyFilter(control.getAttribute("data-showcase-filter") || "all");
      });
    });
  }

  function init() {
    initTheme();
    initMobileNav();
    initDropdowns();
    initHeaderLinks();
    initCodeCopy();
    initTabs();
    initHeaderScroll();
    initConverterFrame();
    initDocsSidebar();
    initPrism();
    initShowcaseFilters();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
