/* OfficeIMO – Site JS – Liquid Glass Premium */
(function () {
  "use strict";

  var THEME_KEY = "imo-theme";
  var PRISM_LANGUAGES_PATH = "/assets/prism/components/";

  window.Prism = window.Prism || {};
  window.Prism.manual = true;

  var BACKGROUND_KEY = "imo-bg";
  var BACKGROUNDS = ["blueprint", "paper", "aurora", "plain"];
  var colorSchemeQuery = window.matchMedia ? window.matchMedia("(prefers-color-scheme: dark)") : null;

  function readPreference(key) {
    try { return localStorage.getItem(key); } catch (error) { return null; }
  }

  function writePreference(key, value) {
    try {
      if (value === null) localStorage.removeItem(key);
      else localStorage.setItem(key, value);
    } catch (error) { /* Storage can be unavailable; the page still works. */ }
  }

  function getDefaultTheme() {
    return colorSchemeQuery && colorSchemeQuery.matches ? "dark" : "light";
  }

  function getThemePreference() {
    var stored = readPreference(THEME_KEY);
    return stored === "light" || stored === "dark" ? stored : "system";
  }

  function setPressed(selector, attribute, value) {
    document.querySelectorAll(selector).forEach(function (btn) {
      btn.setAttribute("aria-pressed", btn.getAttribute(attribute) === value ? "true" : "false");
    });
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
    setPressed("[data-imo-theme]", "data-imo-theme", getThemePreference());
  }

  function applyBackground(background) {
    var resolved = BACKGROUNDS.indexOf(background) >= 0 ? background : BACKGROUNDS[0];
    document.documentElement.setAttribute("data-bg", resolved);
    setPressed("[data-imo-bg]", "data-imo-bg", resolved);
  }

  function initTheme() {
    applyTheme(getThemePreference());
    applyBackground(readPreference(BACKGROUND_KEY));

    if (colorSchemeQuery) {
      var onSystemChange = function () {
        if (getThemePreference() === "system") applyTheme("system");
      };
      if (colorSchemeQuery.addEventListener) colorSchemeQuery.addEventListener("change", onSystemChange);
      else if (colorSchemeQuery.addListener) colorSchemeQuery.addListener(onSystemChange);
    }

    document.querySelectorAll(".imo-theme-toggle").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var cur = document.documentElement.getAttribute("data-theme") || getDefaultTheme();
        var next = cur === "dark" ? "light" : "dark";
        writePreference(THEME_KEY, next);
        applyTheme(next);
      });
    });

    document.querySelectorAll("[data-imo-theme]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var choice = btn.getAttribute("data-imo-theme");
        writePreference(THEME_KEY, choice === "light" || choice === "dark" ? choice : null);
        applyTheme(choice);
      });
    });

    document.querySelectorAll("[data-imo-bg]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var choice = btn.getAttribute("data-imo-bg");
        writePreference(BACKGROUND_KEY, choice === BACKGROUNDS[0] ? null : choice);
        applyBackground(choice);
      });
    });
  }

  function initHeaderMenus() {
    var menus = Array.prototype.slice.call(document.querySelectorAll("details.imo-menu"));
    if (!menus.length) return;

    menus.forEach(function (menu) {
      menu.addEventListener("toggle", function () {
        if (!menu.open) return;
        menus.forEach(function (other) { if (other !== menu) other.open = false; });
      });
    });

    document.addEventListener("click", function (event) {
      menus.forEach(function (menu) {
        if (menu.open && !menu.contains(event.target)) menu.open = false;
      });
    });

    document.addEventListener("keydown", function (event) {
      if (event.key !== "Escape") return;
      menus.forEach(function (menu) {
        if (!menu.open) return;
        menu.open = false;
        var summary = menu.querySelector("summary");
        if (summary) summary.focus();
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

    var actualCurrent = normalizePath(window.location.pathname);
    var current = actualCurrent;
    if (current.indexOf('/products/') === 0 && current !== '/products/pswriteoffice') current = '/libraries';
    if (document.body.classList.contains('imo-body--docs')) current = '/docs';
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
        var samePath = path === actualCurrent;
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
      nav.dispatchEvent(new Event("navigationclose"));
    }

    function openNav() {
      hamburger.classList.add("is-active");
      hamburger.setAttribute("aria-expanded", "true");
      nav.classList.add("is-open");
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
        if (nav.querySelector(".imo-nav__item.is-open")) return;
        closeNav();
        hamburger.focus();
        e.preventDefault();
      }
    });

    document.addEventListener("click", function (e) {
      if (nav.classList.contains("is-open") && !e.target.closest(".imo-header")) closeNav();
    });

    nav.addEventListener("focusout", function () {
      window.requestAnimationFrame(function () {
        if (nav.classList.contains("is-open") && !document.activeElement.closest(".imo-header")) closeNav();
      });
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
      var menu = item.querySelector(":scope > .imo-dropdown");
      if (menu) menu.hidden = !open;
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

    var nav = document.querySelector(".imo-nav");
    if (nav) nav.addEventListener("navigationclose", function () { closeAll(); });

    items.forEach(function (item) {
      var btn = item.querySelector("button.imo-nav__link");
      if (!btn) return;

      btn.addEventListener("click", function (e) {
        e.preventDefault();
        e.stopPropagation();
        var willOpen = !item.classList.contains("is-open");
        closeAll(item);
        setOpen(item, willOpen);
      });

      btn.addEventListener("keydown", function (e) {
        if (e.key !== "ArrowDown" && e.key !== "ArrowUp") return;
        e.preventDefault();
        closeAll(item);
        setOpen(item, true);
        var links = item.querySelectorAll(".imo-dropdown a[href]");
        var target = e.key === "ArrowUp" ? links[links.length - 1] : links[0];
        window.requestAnimationFrame(function () {
          if (target && item.classList.contains("is-open") && document.activeElement === btn) target.focus();
        });
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
      e.preventDefault();
      if (trigger) trigger.focus();
    });

    window.addEventListener("resize", function () {
      closeAll();
    });
  }

  function initCodeCopy() {
    var status = document.createElement("span");
    status.className = "imo-copy-status";
    status.setAttribute("role", "status");
    document.body.appendChild(status);
    var statusTimer;

    function reportCopy(message) {
      clearTimeout(statusTimer);
      status.textContent = message;
      statusTimer = setTimeout(function () { status.textContent = ""; }, 4000);
    }

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
          reportCopy("Copied to clipboard.");
          setTimeout(function () { btn.classList.remove("is-copied"); }, 2000);
        }).catch(function () {
          btn.classList.remove("is-copied");
          reportCopy("Clipboard unavailable. Select the command and copy it manually.");
        });
      } else {
        reportCopy("Clipboard unavailable. Select the command and copy it manually.");
      }
    });
  }

  function initTabs() {
    document.querySelectorAll('[role="tablist"]').forEach(function (tablist) {
      var tabs = tablist.querySelectorAll('[role="tab"]');
      tabs.forEach(function (tab, index) {
        tab.addEventListener("keydown", function (event) {
          var next;
          if (event.key === "ArrowRight") next = (index + 1) % tabs.length;
          else if (event.key === "ArrowLeft") next = (index + tabs.length - 1) % tabs.length;
          else if (event.key === "Home") next = 0;
          else if (event.key === "End") next = tabs.length - 1;
          else return;
          event.preventDefault();
          tabs[next].click();
          tabs[next].focus();
        });
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
    var template = document.getElementById('browser-workspace-template');
    if (!template) return;
    var directory = document.querySelector('.imo-browser-tools');
    var search = directory.querySelector('.imo-browser-tools__search');
    var input = search.querySelector('input');
    var countLabel = search.querySelector('[data-tool-count]');
    var cards = Array.from(directory.querySelectorAll('.imo-browser-tools__card'));
    var frameShell = document.querySelector('.imo-converter-launch__frame-shell');
    var frame = null;
    var workspaceUrl = null;
    var pendingFile = null;
    search.hidden = false;

    function fileExtension(name) {
      var dot = name.lastIndexOf('.');
      return dot > 0 ? name.slice(dot).toLowerCase() : '';
    }

    function acceptsPendingFile(card) {
      if (!pendingFile) return true;
      return (card.getAttribute('data-accept') || '').split(',').indexOf(fileExtension(pendingFile.name)) >= 0;
    }

    function filterTools() {
      var query = input.value.trim().toLowerCase();
      var matches = 0;
      cards.forEach(function (card) {
        var terms = card.textContent.replace(/\s+/g, ' ').trim();
        card.hidden = !terms.toLowerCase().includes(query) || !acceptsPendingFile(card);
        if (!card.hidden) matches++;
      });
      directory.querySelectorAll('[data-tool-group]').forEach(function (group) {
        group.hidden = !group.querySelector('.imo-browser-tools__card:not([hidden])');
      });
      countLabel.textContent = matches ? matches + (matches === 1 ? ' tool' : ' tools') : 'No tools match your search.';
      return matches;
    }
    input.addEventListener('input', filterTools);
    directory.querySelectorAll('.imo-browser-tools__toolbar nav a').forEach(function (link) {
      link.addEventListener('click', function () { input.value = ''; filterTools(); });
    });
    filterTools();

    function hasWorkspaceSelection(parameters) {
      var workspace = (parameters.get('workspace') || '').toLowerCase();
      return !!(parameters.get('route') || '').trim() || workspace === 'pdf' || workspace === 'provenance';
    }

    function syncTheme() {
      try {
        var frameRoot = frame && frame.contentDocument && frame.contentDocument.documentElement;
        if (!frameRoot) return;
        var theme = document.documentElement.getAttribute('data-theme') || 'light';
        frameRoot.setAttribute('data-theme', theme);
        frameRoot.style.colorScheme = theme;
      } catch (error) {
        // The standalone application retains its saved theme if hosted elsewhere.
      }
    }

    function createFrame() {
      frameShell.appendChild(template.content.cloneNode(true));
      frame = frameShell.querySelector('iframe');
      workspaceUrl = new URL(frame.getAttribute('data-workspace-src'), window.location.href);

      window.addEventListener('message', function (event) {
        if (event.source !== frame.contentWindow || event.origin !== workspaceUrl.origin ||
            !event.data || event.data.type !== 'officeimo:workspace-selection') return;
        var selection = event.data;
        if (['workspace', 'route', 'tool'].some(function (key) {
          return selection[key] != null && (typeof selection[key] !== 'string' || selection[key].length > 100);
        })) return;
        if (selection.title != null && (typeof selection.title !== 'string' ||
            selection.title.length > 160 || selection.title.trim().length === 0)) return;
        var target = new URL(window.location.href);
        ['workspace', 'route', 'tool'].forEach(function (key) {
          var value = selection[key];
          if (key === 'workspace' && value === 'convert') value = null;
          if (value) target.searchParams.set(key, value);
          else target.searchParams.delete(key);
        });
        if (target.href !== window.location.href) {
          window.history[selection.replace === true ? 'replaceState' : 'pushState'](null, '', target);
        }
        if (selection.title) document.title = selection.title;
      });
      frame.addEventListener('load', syncTheme);
      new MutationObserver(syncTheme).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });
    }

    // The app owns route validation; the website forwards only its public selection keys.
    function workspaceAddress(parameters) {
      var address = new URL(workspaceUrl.href);
      ['workspace', 'route', 'tool'].forEach(function (key) {
        if (parameters.has(key)) address.searchParams.set(key, parameters.get(key));
      });
      return address.href;
    }

    // The app listens for selection messages only after its workspace has rendered.
    function workspaceReady() {
      try { return !!(frame.contentDocument && frame.contentDocument.querySelector('.ocx-workspace-content')); } catch (error) { return false; }
    }

    // Browsing the directory never creates a frame or starts the WebAssembly runtime.
    function openWorkspace(parameters) {
      directory.hidden = true;
      frameShell.hidden = false;
      showNotice('');
      if (!frame) {
        createFrame();
        frame.src = workspaceAddress(parameters);
        syncTheme();
      } else if (!workspaceReady()) {
        frame.src = workspaceAddress(parameters);
      } else {
        frame.contentWindow.postMessage({ type: 'officeimo:restore-selection',
          workspace: parameters.get('workspace'), route: parameters.get('route'), tool: parameters.get('tool')
        }, workspaceUrl.origin);
      }
      window.scrollTo(0, 0);
    }

    function showDirectory() {
      frameShell.hidden = true;
      directory.hidden = false;
    }

    var notice = null;
    function showNotice(message) {
      if (!notice) {
        if (!message) return;
        notice = document.createElement('p');
        notice.className = 'imo-handoff-notice';
        notice.setAttribute('role', 'status');
        frameShell.insertBefore(notice, frameShell.firstChild);
      }
      notice.textContent = message;
      notice.hidden = !message;
    }

    // Hands a file chosen on the directory to the tool's own input once the workspace renders it.
    function deliverFile(file, card) {
      var route = card.getAttribute('data-route');
      var tool = card.getAttribute('data-pdf-tool');
      var isText = card.getAttribute('data-input') === 'text';
      var selector = route
        ? (isText ? '.ocx-conversion-workspace[data-active-route="' + route + '"] .ocx-textarea' : '#conversion-file-input-' + route)
        : tool ? '#pdf-file-input-' + tool : '#provenance-file-input';
      var started = Date.now();
      var tooLarge = file.name + ' is too long for this text tool. Paste a shorter section, or convert the file with another tool.';
      (function attempt() {
        var element = null;
        try { element = frame.contentDocument && frame.contentDocument.querySelector(selector); } catch (error) { return; }
        if (!element) {
          if (frameShell.hidden) return;
          if (Date.now() - started < 120000) setTimeout(attempt, 200);
          else showNotice('The tool did not finish loading, so ' + file.name + ' was not opened. Choose the file again inside the tool.');
          return;
        }
        var view = frame.contentWindow;
        if (isText) {
          // UTF-8 uses at most four bytes per character, so larger files cannot fit.
          if (element.maxLength > 0 && file.size > element.maxLength * 4) { showNotice(tooLarge); return; }
          file.text().then(function (text) {
            if (element.maxLength > 0 && text.length > element.maxLength) { showNotice(tooLarge); return; }
            element.value = text;
            element.dispatchEvent(new view.Event('input', { bubbles: true }));
          }, function () {
            showNotice(file.name + ' could not be read. Choose it again inside the tool.');
          });
          return;
        }
        var transfer = new view.DataTransfer();
        transfer.items.add(new view.File([file], file.name, { type: file.type, lastModified: file.lastModified }));
        element.files = transfer.files;
        element.dispatchEvent(new view.Event('change', { bubbles: true }));
      })();
    }

    function initFileDrop(drop) {
      var fileInput = drop.querySelector('#browser-drop-input');
      var target = drop.querySelector('[data-browser-drop-target]');
      var selected = drop.querySelector('[data-browser-drop-selected]');
      var status = drop.querySelector('[data-browser-drop-status]');
      drop.hidden = false;

      function formatSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1048576) return Math.round(bytes / 1024) + ' KB';
        return (bytes / 1048576).toFixed(1) + ' MB';
      }

      function choose(file) {
        pendingFile = file || null;
        input.value = '';
        var matches = filterTools();
        target.hidden = !!pendingFile;
        selected.hidden = !pendingFile;
        drop.classList.toggle('has-file', !!pendingFile);
        if (!pendingFile) return;
        drop.querySelector('[data-browser-drop-name]').textContent = pendingFile.name;
        drop.querySelector('[data-browser-drop-size]').textContent = formatSize(pendingFile.size);
        status.textContent = matches === 0
          ? 'No browser tool opens this file type yet.'
          : matches === 1 ? 'Open the tool below to continue.' : 'Pick one of the ' + matches + ' tools below.';
      }

      function carriesFiles(event) {
        return event.dataTransfer && Array.prototype.indexOf.call(event.dataTransfer.types, 'Files') >= 0;
      }

      fileInput.addEventListener('change', function () { choose(fileInput.files[0]); });
      drop.querySelector('[data-browser-drop-clear]').addEventListener('click', function () {
        fileInput.value = '';
        choose(null);
        target.focus();
      });

      var dragDepth = 0;
      document.addEventListener('dragenter', function (event) {
        if (!carriesFiles(event) || directory.hidden) return;
        dragDepth++;
        drop.classList.add('is-dragging');
      });
      document.addEventListener('dragleave', function (event) {
        if (!carriesFiles(event)) return;
        dragDepth = Math.max(0, dragDepth - 1);
        if (!dragDepth) drop.classList.remove('is-dragging');
      });
      document.addEventListener('dragover', function (event) {
        if (!carriesFiles(event) || directory.hidden) return;
        event.preventDefault();
        event.dataTransfer.dropEffect = 'copy';
      });
      document.addEventListener('drop', function (event) {
        if (!carriesFiles(event) || directory.hidden) return;
        event.preventDefault();
        dragDepth = 0;
        drop.classList.remove('is-dragging');
        choose(event.dataTransfer.files[0]);
      });

      cards.forEach(function (card) {
        card.addEventListener('click', function (event) {
          if (!pendingFile || event.button !== 0 || event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
          event.preventDefault();
          var destination = new URL(card.href, window.location.href);
          window.history.pushState(null, '', destination);
          openWorkspace(destination.searchParams);
          deliverFile(pendingFile, card);
        });
      });
    }

    var drop = directory.querySelector('[data-browser-drop]');
    if (drop) initFileDrop(drop);

    window.addEventListener('popstate', function () {
      var selection = new URLSearchParams(window.location.search);
      if (!hasWorkspaceSelection(selection)) {
        showDirectory();
        return;
      }
      openWorkspace(selection);
    });

    var pageParameters = new URLSearchParams(window.location.search);
    if (hasWorkspaceSelection(pageParameters)) openWorkspace(pageParameters);
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
        if (sidebar.contains(document.activeElement)) toggle.focus();
      }

      function openSidebar() {
        sidebar.classList.add("is-open");
        toggle.setAttribute("aria-expanded", "true");
        if (overlay) overlay.hidden = false;
        document.body.style.overflow = "hidden";
        var firstLink = sidebar.querySelector('a[aria-current], a[href]');
        if (firstLink) firstLink.focus();
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
        if (e.key === "Tab" && sidebar.classList.contains("is-open")) {
          var focusable = Array.from(sidebar.querySelectorAll('a[href], summary')).filter(function (el) { return el.getClientRects().length > 0; });
          var first = focusable[0], last = focusable[focusable.length - 1];
          if (e.shiftKey && document.activeElement === first) { e.preventDefault(); toggle.focus(); }
          else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); toggle.focus(); }
          else if (document.activeElement === toggle) { e.preventDefault(); (e.shiftKey ? last : first).focus(); }
        }
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
    document.querySelectorAll('.imo-docs__nav-intro a').forEach(function (link) {
      if (link.getAttribute('href') === currentPath) link.setAttribute('aria-current', 'page');
    });
    var referenceDetails = document.querySelector('.imo-reference-browser > details');
    if (referenceDetails && window.innerWidth < 1024) referenceDetails.open = false;
    document.querySelectorAll(".imo-docs__group").forEach(function (group) {
      var links = group.querySelectorAll(".imo-docs__link");
      var hasActive = false;
      links.forEach(function (link) {
        var href = link.getAttribute("href");
        if (href && (currentPath === href || (link.hasAttribute('data-reference-root') && currentPath.indexOf(href) === 0))) {
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
      if (href && (currentPath === href || (link.hasAttribute('data-reference-root') && currentPath.indexOf(href) === 0))) {
        link.classList.add("active");
      }
    });
  }

  var prismLoading = null;

  function prismReady() {
    return typeof Prism !== "undefined" && (typeof Prism.highlightAllUnder === "function" || typeof Prism.highlightAll === "function");
  }

  function loadScript(src) {
    return new Promise(function (resolve, reject) {
      var existing = document.querySelector('script[src="' + src + '"]');
      if (existing && existing.getAttribute("data-loaded") === "true") { resolve(); return; }
      var script = existing || document.createElement("script");
      script.addEventListener("load", function () { script.setAttribute("data-loaded", "true"); resolve(); });
      script.addEventListener("error", reject);
      if (!existing) { script.src = src; document.head.appendChild(script); }
    });
  }

  // Loads the highlighter only for pages that show code, so other pages skip it entirely.
  function ensurePrism() {
    if (prismReady()) return Promise.resolve();
    if (!prismLoading) {
      prismLoading = loadScript("/assets/prism/prism-core.min.js")
        .then(function () { return loadScript("/assets/prism/prism-autoloader.min.js"); })
        .then(function () {
          Prism.manual = true;
          Prism.plugins = Prism.plugins || {};
          if (Prism.plugins.autoloader) Prism.plugins.autoloader.languages_path = PRISM_LANGUAGES_PATH;
        });
    }
    return prismLoading;
  }
  window.OfficeIMOEnsurePrism = ensurePrism;

  function initPrism() {
    if (!document.querySelector('code[class*="language-"], pre[class*="language-"]')) return;
    ensurePrism().then(function () {
      if (typeof Prism.highlightAllUnder === "function") Prism.highlightAllUnder(document);
      else Prism.highlightAll();
    }).catch(function () { /* Code stays readable without highlighting. */ });
  }

  function init() {
    initTheme();
    initHeaderMenus();
    initMobileNav();
    initDropdowns();
    initHeaderLinks();
    initCodeCopy();
    initTabs();
    initHeaderScroll();
    initConverterFrame();
    initDocsSidebar();
    initPrism();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
