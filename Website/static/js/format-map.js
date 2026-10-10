/*
 * Format map (partials/sections/format-map.html). One data set and one side panel, two views:
 *   radial - home page: formats around a circle, conversions as lines, paths between two formats.
 *   list   - browser tools: formats as chips grouped by family, curves to what the chosen one becomes.
 * Phones get the chip list in both views. ES5 for the site minifier.
 * The radial view also runs a tour (starts when the circle is in view; Play tour button): it spotlights formats around the
 * circle, finds paths between formats and steps through the surfaces, with a caption for each scene.
 */
(function () {
  'use strict';

  var NS = 'http://www.w3.org/2000/svg';
  var SURFACES = ['dotnet', 'browser', 'cli', 'studio', 'powershell'];
  var FIDELITY = { Editable: 'editable output', FixedLayout: 'fixed layout', Semantic: 'structure and text' };
  var still = !!(window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches);

  function el(tag, cls, text) {
    var node = document.createElement(tag);
    if (cls) node.className = cls;
    if (text != null) node.textContent = text;
    return node;
  }
  function sv(tag, attrs, text) {
    var node = document.createElementNS(NS, tag);
    for (var key in attrs) if (Object.prototype.hasOwnProperty.call(attrs, key)) node.setAttribute(key, attrs[key]);
    if (text != null) node.textContent = text;
    return node;
  }
  function plural(n, word) { return n + ' ' + word + (n === 1 ? '' : 's'); }
  function each(list, fn) { Array.prototype.forEach.call(list, fn); }

  function init(root) {
    var dataNode = root.querySelector('.imo-fmap__data');
    var data;
    try { data = JSON.parse(dataNode.textContent); } catch (e) { return; }

    var view = root.getAttribute('data-view') === 'list' ? 'list' : 'radial';
    var families = data.families;
    var familyOf = {};
    var order = [];
    data.formats.forEach(function (f) { familyOf[f[0]] = f[1]; order.push(f[0]); });
    var routes = data.routes.map(function (r) {
      var tool = data.tools[r[6]];
      return { source: r[0], target: r[1], pkg: r[2], surfaces: r[3].split(' '), fidelity: r[4], support: r[5], id: r[6], api: r[7], cmdlet: r[8] || '', example: r[9] || '', tool: tool ? tool[0] : null, toolTitle: tool ? tool[1] : null };
    });
    var packages = data.packages || {};
    var surfaceLinks = data.surfaces || {};
    var surfaceNames = {};
    each(root.querySelectorAll('[data-surface][data-name]'), function (b) { surfaceNames[b.getAttribute('data-surface')] = b.getAttribute('data-name'); });

    var state = { surface: root.getAttribute('data-surface') || 'all', from: null, to: null, hover: null, focus: null, hidden: {} };
    var panel = root.querySelector('.imo-fmap__panel');
    var chips = Array.prototype.slice.call(root.querySelectorAll('.imo-fmap__chip'));
    var surfaceButtons = Array.prototype.slice.call(root.querySelectorAll('button.imo-fmap__surface'));

    // ---------------------------------------------------------------- shared rules
    function onSurface(r) { return state.surface === 'all' || r.surfaces.indexOf(state.surface) >= 0; }
    function shownFamily(name) { return !state.hidden[familyOf[name]]; }
    function visible(r) { return onSurface(r) && shownFamily(r.source) && shownFamily(r.target); }
    function outOf(n) { return routes.filter(function (r) { return r.source === n && visible(r); }); }
    function inTo(n) { return routes.filter(function (r) { return r.target === n && visible(r); }); }
    function where() { return state.surface === 'all' ? '' : ' in ' + surfaceNames[state.surface]; }
    function highlight(r) { return state.surface === 'all' ? r.surfaces.indexOf('browser') >= 0 : true; }

    function shortest(a, b) {
      var depth = {}, prev = {}, queue = [a], found = -1;
      depth[a] = 0; prev[a] = [];
      while (queue.length) {
        var n = queue.shift();
        if (found >= 0 && depth[n] >= found) break;
        outOf(n).forEach(function (r) {
          var m = r.target;
          if (depth[m] == null) { depth[m] = depth[n] + 1; prev[m] = [r]; queue.push(m); }
          else if (depth[m] === depth[n] + 1) prev[m].push(r);
          if (m === b) found = depth[m];
        });
      }
      if (depth[b] == null || a === b) return [];
      var paths = [];
      (function walk(n, acc) {
        if (paths.length >= 3) return;
        if (n === a) { paths.push(acc); return; }
        prev[n].forEach(function (r) { walk(r.source, [r].concat(acc)); });
      })(b, []);
      return paths;
    }

    // ---------------------------------------------------------------- shared panel pieces
    function badges(r) {
      var wrap = el('span', 'imo-fmap__badges');
      SURFACES.forEach(function (s) {
        if (r.surfaces.indexOf(s) >= 0) wrap.appendChild(el('span', 'imo-fmap__badge imo-fmap__badge--' + s, surfaceNames[s] || s));
      });
      return wrap;
    }
    function row(r, label) {
      var li = el('li', 'imo-fmap__row' + (highlight(r) ? ' is-hl' : ''));
      var body = el('span', 'imo-fmap__body');
      body.appendChild(el('strong', null, label));
      // On the PowerShell filter the useful name is the cmdlet, not the .NET package behind it.
      var what = state.surface === 'powershell' && r.cmdlet ? r.cmdlet : r.pkg;
      body.appendChild(el('span', 'imo-fmap__pkg', what + (FIDELITY[r.fidelity] ? ' · ' + FIDELITY[r.fidelity] : '')));
      li.appendChild(body);
      li.appendChild(badges(r));
      if (r.tool) {
        var go = el('a', 'imo-fmap__try', 'Try it');
        go.href = r.tool;
        go.setAttribute('aria-label', 'Try ' + r.source + ' to ' + r.target + ' in your browser');
        li.appendChild(go);
      }
      return li;
    }
    function chipButton(name) {
      var b = el('button', 'imo-fmap__mini imo-fam--' + familyOf[name], name);
      b.type = 'button';
      b.setAttribute('data-go', name);
      return b;
    }
    function heading(text, small) {
      var h = el('h3', 'imo-fmap__title', text);
      if (small) h.appendChild(el('small', null, small));
      return h;
    }

    // ---------------------------------------------------------------- "Go further": where to take a format or route next
    function link(href, text, sub, surface) {
      var a = el('a', 'imo-fmap__link' + (surface ? ' imo-fmap__link--' + surface : ''));
      a.href = href;
      a.appendChild(el('i', null));
      var body = el('span', null);
      // Package names may wrap only after a dot (OfficeIMO.PowerPoint.<wbr>IWork), never mid-word.
      var label = el('b', null);
      text.split('.').forEach(function (part, i, parts) {
        label.appendChild(document.createTextNode(part + (i < parts.length - 1 ? '.' : '')));
        if (i < parts.length - 1) label.appendChild(document.createElement('wbr'));
      });
      body.appendChild(label);
      if (sub) body.appendChild(el('small', null, sub));
      a.appendChild(body);
      return a;
    }
    function packageLink(pkg) {
      var p = packages[pkg] || [], product = (data.products || {})[pkg];
      if (product) return link(product, pkg, 'Product page', 'dotnet');
      if (p[0]) return link(p[0], pkg, 'API reference', 'dotnet');
      if (p[1]) return link(p[1], pkg, 'Documentation', 'dotnet');
      return p[2] ? link(p[2], pkg, 'NuGet', 'dotnet') : null;
    }
    function surfaceLink(id, sub) {
      var s = surfaceLinks[id];
      return s ? link(s[1], s[2], sub || s[0], id) : null;
    }
    function toolLinks(list) {
      var seen = {}, wrap = el('div', 'imo-fmap__tools');
      list.forEach(function (r) {
        if (!r.tool || seen[r.tool]) return;
        seen[r.tool] = true;
        var a = el('a', 'imo-fmap__tool', r.toolTitle || (r.source + ' to ' + r.target));
        a.href = r.tool;
        wrap.appendChild(a);
      });
      return wrap.childNodes.length ? wrap : null;
    }
    function next(title, items) {
      items = items.filter(function (x) { return !!x; });
      if (!items.length) return;
      var box = el('div', 'imo-fmap__next');
      box.appendChild(el('span', 'imo-fmap__label', title));
      items.forEach(function (x) { box.appendChild(x); });
      panel.appendChild(box);
    }
    function linkList(links) {
      links = links.filter(function (x) { return !!x; });
      if (!links.length) return null;
      var ul = el('ul', 'imo-fmap__links');
      links.forEach(function (a) { var li = el('li', null); li.appendChild(a); ul.appendChild(li); });
      return ul;
    }
    function count(list, surface) { return list.filter(function (r) { return r.surfaces.indexOf(surface) >= 0; }).length; }

    function goFurtherIdle() {
      var live = routes.filter(function (r) { return r.tool && visible(r); });
      next('Try in your browser', [toolLinks(live)]);
      next('Go further', [linkList([
        surfaceLink('browser'), surfaceLink('dotnet'), surfaceLink('cli'), surfaceLink('studio'), surfaceLink('powershell'),
        link('/convert/guides/', 'Read the conversion guides', 'Fidelity, limits and examples per format')
      ])]);
    }
    function goFurtherFormat(name, outs) {
      var pkgs = [];
      outs.forEach(function (r) { if (pkgs.indexOf(r.pkg) < 0) pkgs.push(r.pkg); });
      var shown = pkgs.slice(0, 4).map(packageLink);
      var cli = count(outs, 'cli'), studio = count(outs, 'studio'), ps = count(outs, 'powershell');
      var of = function (n) { return n === outs.length ? 'Runs all of these' : 'Runs ' + n + ' of these'; };
      next('Go further with ' + name, [
        toolLinks(outs),
        linkList(shown.concat([
          pkgs.length > 4 ? link('/libraries/', 'And ' + plural(pkgs.length - 4, 'more package'), 'Every package on one page', 'dotnet') : null,
          cli ? surfaceLink('cli', of(cli)) : null,
          studio ? surfaceLink('studio', of(studio)) : null,
          ps ? surfaceLink('powershell', of(ps)) : null
        ]))
      ]);
    }
    function goFurtherPath(path) {
      var everywhere = function (s) { return path.every(function (r) { return r.surfaces.indexOf(s) >= 0; }); };
      var code = el('pre', 'imo-fmap__code');
      var lines = [];
      // On the PowerShell filter (or when it is the only way to run every step) show the cmdlets instead of the .NET calls.
      var shell = everywhere('powershell') && (state.surface === 'powershell' || !everywhere('dotnet'));
      if (shell) {
        lines.push('Install-Module PSWriteOffice');
        var previous = null;
        path.forEach(function (r, i) {
          var example = r.example || r.cmdlet;
          if (previous) example = example.replace(/\.\/in\.[a-z0-9]+\b/gi, previous);
          var output = example.match(/\.\/out\.[a-z0-9]+\b/i);
          if (output && i < path.length - 1) {
            previous = output[0].replace('./out.', './step-' + (i + 1) + '.');
            example = example.replace(output[0], previous);
          } else previous = null;
          lines.push('# ' + r.source + ' → ' + r.target, example);
        });
      } else {
        path.forEach(function (r) { if (lines.indexOf('dotnet add package ' + r.pkg) < 0) lines.push('dotnet add package ' + r.pkg); });
        path.forEach(function (r) { lines.push('// ' + r.source + ' → ' + r.target, r.api); });
      }
      code.appendChild(el('code', null, lines.join('\n')));
      var refs = [];
      path.forEach(function (r) { var a = packageLink(r.pkg); if (a && refs.every(function (x) { return x.href !== a.href; })) refs.push(a); });
      next('Use it in your code', [code, linkList(refs.concat([
        everywhere('cli') ? surfaceLink('cli', 'Runs this from the command line') : null,
        everywhere('studio') ? surfaceLink('studio', 'Runs this in the desktop app') : null,
        everywhere('powershell') ? surfaceLink('powershell', 'Runs this from PowerShell') : null
      ]))]);
    }

    var side = root.querySelector('.imo-fmap__side'), shownKey = '';
    function renderPanel() {
      var key = [state.surface, state.from, state.to, state.from ? '' : (state.hover || state.focus)].join('|');
      buildPanel();
      if (key === shownKey) return;
      shownKey = key;
      if (side) side.scrollTop = 0;
      if (still) return;
      panel.classList.remove('is-swap');
      void panel.offsetWidth;
      panel.classList.add('is-swap');
    }
    function buildPanel() {
      panel.textContent = '';
      var focus = state.from || state.hover || state.focus;
      if (!focus) {
        panel.appendChild(heading('Pick a format'));
        var lines = state.surface === 'all'
          ? 'Highlighted lines run in the browser tools today.'
          : 'Showing the conversions that run in ' + surfaceNames[state.surface] + '.';
        panel.appendChild(el('p', 'imo-fmap__hint', (view === 'radial' ? 'Hover a dot, or choose one above. ' : 'Choose a format to see what it converts to. ') + lines));
        var top = order.map(function (n) { return [n, outOf(n).length]; }).filter(function (p) { return p[1] > 0; })
          .sort(function (a, b) { return b[1] - a[1]; }).slice(0, 6);
        if (top.length) {
          panel.appendChild(el('span', 'imo-fmap__label', 'Most connected' + where()));
          var wrap = el('div', 'imo-fmap__minis');
          top.forEach(function (p) { wrap.appendChild(chipButton(p[0])); });
          panel.appendChild(wrap);
        }
        goFurtherIdle();
        return;
      }
      if (state.from && state.to) {
        var paths = shortest(state.from, state.to);
        panel.appendChild(heading(state.from + ' → ' + state.to, paths.length ? (paths[0].length === 1 ? 'direct' : paths[0].length + ' steps') : 'no route yet'));
        if (!paths.length) {
          panel.appendChild(el('p', 'imo-fmap__hint', 'No mapped route' + where() + ' from ' + state.from + ' to ' + state.to + ' yet. Try another format, or clear the choice.'));
          return;
        }
        paths.forEach(function (path) {
          var list = el('ol', 'imo-fmap__list imo-fmap__path');
          path.forEach(function (r) { list.appendChild(row(r, r.source + ' → ' + r.target)); });
          panel.appendChild(list);
        });
        if (paths[0].length > 1) panel.appendChild(el('p', 'imo-fmap__hint', 'Each step is a separate conversion with its own report.'));
        goFurtherPath(paths[0]);
        return;
      }
      var outs = outOf(focus).sort(function (a, b) { return a.target.localeCompare(b.target); });
      var ins = inTo(focus);
      panel.appendChild(heading(focus, families[familyOf[focus]]));
      panel.appendChild(el('span', 'imo-fmap__label', outs.length ? 'Converts to ' + plural(outs.length, 'format') + where() : 'No mapped conversion from ' + focus + where() + ' yet'));
      var list = el('ul', 'imo-fmap__list imo-fmap__outs');
      outs.forEach(function (r) { list.appendChild(row(r, r.target)); });
      if (view === 'list') list.addEventListener('scroll', function () { relayout(false); });
      panel.appendChild(list);
      if (ins.length) {
        panel.appendChild(el('span', 'imo-fmap__label', 'Made from ' + plural(ins.length, 'format') + where()));
        var made = el('div', 'imo-fmap__minis');
        ins.forEach(function (r) { made.appendChild(chipButton(r.source)); });
        panel.appendChild(made);
      }
      if (view === 'radial' && state.from) panel.appendChild(el('p', 'imo-fmap__hint', 'Pick a second format to find the way from ' + focus + '.'));
      if (outs.length) goFurtherFormat(focus, outs);
    }

    // ---------------------------------------------------------------- shared controls
    function syncSurfaceButtons() {
      surfaceButtons.forEach(function (b) {
        var on = b.getAttribute('data-surface') === state.surface;
        b.classList.toggle('is-active', on);
        b.setAttribute('aria-pressed', on ? 'true' : 'false');
      });
      root.setAttribute('data-active-surface', state.surface);
    }
    function syncChips() {
      chips.forEach(function (chip) {
        var name = chip.getAttribute('data-format');
        var n = outOf(name).length + inTo(name).length;
        chip.disabled = n === 0;
        var on = name === state.from;
        chip.classList.toggle('is-active', on);
        chip.setAttribute('aria-pressed', on ? 'true' : 'false');
      });
    }

    var counts = { all: routes.length };
    SURFACES.forEach(function (s) { counts[s] = routes.filter(function (r) { return r.surfaces.indexOf(s) >= 0; }).length; });
    each(root.querySelectorAll('[data-count]'), function (node) { node.textContent = String(counts[node.getAttribute('data-count')] || 0); });

    // Picking: a first format focuses it; in the radial view a second one asks for the way between them.
    function pick(name, second) {
      if (second && state.from && state.from !== name && !state.to) state.to = name;
      else if (state.from === name && !state.to) state.from = null;
      else { state.from = name; state.to = null; }
      update(false, true);
    }

    surfaceButtons.forEach(function (b) {
      b.addEventListener('click', function () {
        state.surface = b.getAttribute('data-surface');
        if (state.from && !outOf(state.from).length && !inTo(state.from).length) state.from = state.to = null;
        update(true);
      });
    });
    chips.forEach(function (chip) {
      chip.addEventListener('click', function () {
        pick(chip.getAttribute('data-format'), view === 'radial');
        // Stacked on a phone, the answer sits below the chips: bring it into view.
        if (panel.getBoundingClientRect().top > window.innerHeight * 0.65) {
          panel.scrollIntoView({ block: 'start', behavior: still ? 'auto' : 'smooth' });
        }
      });
    });
    panel.addEventListener('click', function (e) {
      var b = e.target.closest ? e.target.closest('[data-go]') : null;
      if (b) pick(b.getAttribute('data-go'), view === 'radial');
    });

    // ---------------------------------------------------------------- list view: curves from the chip to its results
    var wires = root.querySelector('.imo-fmap__wires');
    function drawWires(animate) {
      if (!wires) return;
      while (wires.firstChild) wires.removeChild(wires.firstChild);
      var chip = state.from && root.querySelector('.imo-fmap__chip.is-active');
      var rows = Array.prototype.slice.call(panel.querySelectorAll('.imo-fmap__outs .imo-fmap__row'));
      // The list scrolls when it is long: only draw to the rows that are in view.
      var outsBox = panel.querySelector('.imo-fmap__outs'), clip = outsBox && outsBox.getBoundingClientRect();
      if (clip) rows = rows.filter(function (li) { var r = li.getBoundingClientRect(); return r.top >= clip.top - 2 && r.bottom <= clip.bottom + 2; });
      if (!chip || !rows.length || getComputedStyle(wires).display === 'none') return;
      // Leave from the end of the chip's row, so curves don't cut through its neighbours.
      var box = wires.getBoundingClientRect(), start = chip.getBoundingClientRect(), rowEnd = start.right;
      each(chip.parentNode.children, function (other) {
        var r = other.getBoundingClientRect();
        if (Math.abs(r.top - start.top) < 2) rowEnd = Math.max(rowEnd, r.right);
      });
      var x0 = rowEnd - box.left + 6, y0 = start.top + start.height / 2 - box.top;
      wires.setAttribute('viewBox', '0 0 ' + Math.max(1, box.width) + ' ' + Math.max(1, box.height));
      rows.forEach(function (li, index) {
        var r = li.getBoundingClientRect();
        var x1 = r.left - box.left - 2, y1 = r.top + r.height / 2 - box.top, bend = Math.max(24, (x1 - x0) / 2);
        var path = sv('path', {
          d: 'M' + x0 + ' ' + y0 + ' C' + (x0 + bend) + ' ' + y0 + ' ' + (x1 - bend) + ' ' + y1 + ' ' + x1 + ' ' + y1,
          'class': 'imo-fmap__edge is-out' + (li.classList.contains('is-hl') ? ' is-hl' : '')
        });
        wires.appendChild(path);
        if (still || animate === false) return;
        var length = Math.ceil(path.getTotalLength());
        path.style.strokeDasharray = length + ' ' + length;
        path.style.strokeDashoffset = String(length);
        path.style.animationDelay = Math.min(index * 18, 360) + 'ms';
        path.classList.add('is-drawing');
      });
    }

    // ---------------------------------------------------------------- radial view (ported from the review prototype)
    var svg = root.querySelector('.imo-fmap__svg');
    var C = 500, R = 318, edgeEls = [], nodeEls = {}, pulseLayer = null;
    var pickers = root.querySelectorAll('[data-pick]');
    var legend = Array.prototype.slice.call(root.querySelectorAll('.imo-fmap__legend button'));

    function drawRadial() {
      state.hover = state.focus = null;
      if (!svg) return;
      while (svg.firstChild) svg.removeChild(svg.firstChild);
      edgeEls = []; nodeEls = {};
      var fams = data.families ? Object.keys(families).filter(function (f) { return !state.hidden[f] && order.some(function (n) { return familyOf[n] === f; }); }) : [];
      var nodes = [], labels = [], gap = 0.9, k = 0, pos = {};
      fams.forEach(function (f) { order.forEach(function (n) { if (familyOf[n] === f) nodes.push(n); }); });
      var slots = nodes.length + fams.length * gap;
      fams.forEach(function (f) {
        var start = k;
        order.forEach(function (n) {
          if (familyOf[n] !== f) return;
          var a = -Math.PI / 2 + (k / slots) * Math.PI * 2;
          pos[n] = { a: a, x: C + R * Math.cos(a), y: C + R * Math.sin(a) };
          k++;
        });
        labels.push([f, (start + k - 1) / 2, k - start]);
        k += gap;
      });
      var gF = sv('g'), gE = sv('g', { 'class': 'imo-fmap__edges' }), gN = sv('g');
      pulseLayer = sv('g');
      labels.forEach(function (l) {
        if (l[2] < 2) return;
        var a = -Math.PI / 2 + (l[1] / slots) * Math.PI * 2, rr = R - 30, x = C + rr * Math.cos(a), y = C + rr * Math.sin(a);
        var deg = (Math.sin(a) > 0.05 ? a - Math.PI / 2 : a + Math.PI / 2) * 180 / Math.PI;
        gF.appendChild(sv('text', { 'class': 'imo-fmap__fam imo-fam--' + l[0], x: x, y: y, 'text-anchor': 'middle', 'dominant-baseline': 'middle', transform: 'rotate(' + deg.toFixed(1) + ' ' + x.toFixed(1) + ' ' + y.toFixed(1) + ')' }, families[l[0]]));
      });
      routes.forEach(function (r) {
        var a = pos[r.source], b = pos[r.target];
        if (!a || !b || !onSurface(r)) return;
        var mx = (a.x + b.x) / 2, my = (a.y + b.y) / 2, t = 0.18;
        var p = sv('path', { 'class': 'imo-fmap__edge' + (highlight(r) ? ' is-hl' : ''), d: 'M' + a.x.toFixed(1) + ',' + a.y.toFixed(1) + ' Q' + (C + (mx - C) * t).toFixed(1) + ',' + (C + (my - C) * t).toFixed(1) + ' ' + b.x.toFixed(1) + ',' + b.y.toFixed(1) });
        p._route = r;
        edgeEls.push(p);
        gE.appendChild(p);
      });
      nodes.forEach(function (n) {
        var at = pos[n], c = Math.cos(at.a), s = Math.sin(at.a);
        var degree = outOf(n).length + inTo(n).length;
        var g = sv('g', { 'class': 'imo-fmap__node imo-fam--' + familyOf[n] + (degree ? '' : ' is-idle'), tabindex: degree ? '0' : '-1', role: 'button', 'aria-label': n + ': converts to ' + plural(outOf(n).length, 'format') });
        g.appendChild(sv('circle', { cx: at.x, cy: at.y, r: Math.min(11, 4.5 + degree * 0.28) }));
        var lx = C + (R + 18) * c, ly = C + (R + 18) * s;
        g.appendChild(sv('text', { x: lx, y: ly, 'text-anchor': c >= 0 ? 'start' : 'end', 'dominant-baseline': 'middle', transform: 'rotate(' + ((c >= 0 ? at.a : at.a + Math.PI) * 180 / Math.PI) + ' ' + lx + ' ' + ly + ')' }, n));
        if (degree) {
          g.addEventListener('mouseenter', function () { if (moved()) setHover(n); });
          g.addEventListener('mousemove', function () { if (moved()) setHover(n); });
          g.addEventListener('mouseleave', function () { setHover(null); });
          g.addEventListener('focus', function () { tour.touched = true; stopTour(true); state.focus = n; refreshPreview(true); });
          g.addEventListener('blur', function () { state.focus = null; refreshPreview(true); });
          g.addEventListener('click', function (e) { e.stopPropagation(); pick(n, true); });
          g.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick(n, true); } });
        }
        nodeEls[n] = g;
        gN.appendChild(g);
      });
      var hub = sv('g', { 'pointer-events': 'none' });
      hub.appendChild(sv('text', { 'class': 'imo-fmap__hub', x: C, y: C - 6, 'text-anchor': 'middle' }, 'OfficeIMO'));
      hub.appendChild(sv('text', { 'class': 'imo-fmap__hubsub', x: C, y: C + 24, 'text-anchor': 'middle' }, plural(edgeEls.length, 'conversion') + where()));
      svg.appendChild(gF); svg.appendChild(gE); svg.appendChild(hub); svg.appendChild(gN); svg.appendChild(pulseLayer);
      fillPickers(nodes.filter(function (n) { return outOf(n).length + inTo(n).length > 0; }));
    }

    // Lines grow out of their source dot. Used when a format is chosen or the tour moves on, not for plain hovering.
    function drawOn(p, delay, ms) {
      if (still) return;
      var len = p._len || (p._len = Math.ceil(p.getTotalLength()));
      p.style.transition = 'none';
      p.style.strokeDasharray = len + ' ' + len;
      p.style.strokeDashoffset = String(len);
      void p.getBoundingClientRect();
      p.style.transition = 'stroke-dashoffset ' + ms + 'ms cubic-bezier(.3,.6,.2,1) ' + delay + 'ms, opacity .3s, stroke-width .3s, stroke .3s';
      p.style.strokeDashoffset = '0';
      window.setTimeout(function () { p.style.transition = p.style.strokeDasharray = p.style.strokeDashoffset = ''; }, delay + ms + 60);
    }

    function paint(animate) {
      if (!svg) return;
      var focus = state.from || state.hover || state.focus;
      svg.classList.toggle('is-focus', !!focus);
      var paths = state.from && state.to ? shortest(state.from, state.to) : [];
      var onPath = [];
      paths.forEach(function (p) { onPath = onPath.concat(p); });
      var hot = {};
      if (focus) hot[focus] = true;
      var grow = 0;
      edgeEls.forEach(function (p) {
        var r = p._route;
        var out = !!focus && !state.to && r.source === focus, inn = !!focus && !state.to && r.target === focus, on = onPath.indexOf(r) >= 0;
        var wasOut = p.classList.contains('is-out'), wasOn = p.classList.contains('is-on');
        p.classList.toggle('is-out', out);
        p.classList.toggle('is-in', inn);
        p.classList.toggle('is-on', on);
        if (out || inn || on) { hot[r.source] = true; hot[r.target] = true; }
        if (animate && ((out && !wasOut) || (on && !wasOn))) drawOn(p, on ? onPath.indexOf(r) * 420 : Math.min(grow++ * 28, 420), on ? 700 : 650);
      });
      for (var n in nodeEls) {
        nodeEls[n].classList.toggle('is-hot', !!hot[n]);
        nodeEls[n].classList.toggle('is-picked', n === state.from || n === state.to);
      }
    }

    function fillPickers(names) {
      if (!pickers.length) return;
      var sorted = names.slice().sort(function (a, b) { return a.localeCompare(b); });
      each(pickers, function (select) {
        var which = select.getAttribute('data-pick');
        select.textContent = '';
        var none = el('option', null, which === 'from' ? 'Any format' : 'Anything');
        none.value = '';
        select.appendChild(none);
        sorted.forEach(function (n) { var o = el('option', null, n); o.value = n; select.appendChild(o); });
      });
      syncPickers();
    }
    function syncPickers() {
      each(pickers, function (select) {
        var value = select.getAttribute('data-pick') === 'from' ? state.from : state.to;
        select.value = value || '';
      });
    }
    each(pickers, function (select) {
      select.addEventListener('change', function () {
        if (select.getAttribute('data-pick') === 'from') { state.from = select.value || null; if (!state.from) state.to = null; }
        else { state.to = select.value || null; if (state.to && !state.from) { state.from = state.to; state.to = null; } }
        update(false, true);
      });
    });
    legend.forEach(function (b) {
      b.addEventListener('click', function () {
        var f = b.getAttribute('data-family');
        var shown = legend.filter(function (x) { return x.getAttribute('aria-pressed') === 'true'; }).length;
        if (!state.hidden[f] && shown <= 1) return;
        state.hidden[f] = !state.hidden[f];
        b.setAttribute('aria-pressed', state.hidden[f] ? 'false' : 'true');
        if (state.from && state.hidden[familyOf[state.from]]) state.from = state.to = null;
        if (state.to && state.hidden[familyOf[state.to]]) state.to = null;
        update(true);
      });
    });
    if (svg) svg.addEventListener('click', function () { state.from = state.to = null; update(); });

    // Ambient pulses along highlighted routes while nothing is chosen and the map is on screen.
    var onScreen = true;
    if (svg && !still) {
      if (window.IntersectionObserver) new window.IntersectionObserver(function (entries) { onScreen = entries[0].isIntersecting; }).observe(svg);
      var pulses = [];
      window.setInterval(function () {
        if (state.hover || state.focus || document.hidden || !onScreen || !pulseLayer) return;
        var pool;
        if (state.from) {
          // A chosen format or path: carry data along its own lines.
          pool = edgeEls.filter(function (p) { return p.classList.contains(state.to ? 'is-on' : 'is-out'); });
        } else {
          pool = edgeEls.filter(function (p) { return p.classList.contains('is-hl'); });
          if (!pool.length || Math.random() < 0.35) pool = edgeEls;
        }
        if (!pool.length) return;
        var p = pool[Math.floor(Math.random() * pool.length)];
        var dot = sv('circle', { r: 4.5, 'class': 'imo-fmap__pulse', opacity: 0 });
        pulseLayer.appendChild(dot);
        pulses.push({ p: p, dot: dot, t0: performance.now(), len: p.getTotalLength() });
        if (pulses.length === 1) window.requestAnimationFrame(step);
      }, 650);
      var step = function (now) {
        for (var i = pulses.length - 1; i >= 0; i--) {
          var u = (now - pulses[i].t0) / 1700;
          if (u >= 1 || state.hover || state.focus || !pulses[i].p.isConnected) { pulses[i].dot.remove(); pulses.splice(i, 1); continue; }
          var pt = pulses[i].p.getPointAtLength(pulses[i].len * u);
          pulses[i].dot.setAttribute('cx', pt.x);
          pulses[i].dot.setAttribute('cy', pt.y);
          pulses[i].dot.setAttribute('opacity', Math.sin(u * Math.PI).toFixed(2));
        }
        if (pulses.length) window.requestAnimationFrame(step);
      };
    }

    // ---------------------------------------------------------------- hover smoothing
    // Sweeping the pointer across the circle should read as one smooth motion: the lines repaint at once (CSS eases them),
    // but the panel is only rebuilt once the pointer settles, and leaving a dot keeps its routes lit for a moment so that
    // crossing the gap to the next dot never flashes back to the idle state.
    var hoverLeave = 0, panelTimer = 0;
    // Scrolling or a redraw slides dots under a parked pointer and browsers answer with mouseenter. That is not the visitor
    // reaching for a dot, so hover only counts while the pointer has actually moved a moment ago.
    var pointer = { x: -1, y: -1, at: 0 };
    document.addEventListener('mousemove', function (e) {
      if (pointer.x >= 0 && (e.screenX !== pointer.x || e.screenY !== pointer.y)) pointer.at = Date.now();
      pointer.x = e.screenX;
      pointer.y = e.screenY;
    }, true);
    function moved() { return Date.now() - pointer.at < 400; }
    function setHover(n) {
      window.clearTimeout(hoverLeave);
      if (tour.on) return;
      if (n) {
        if (state.hover === n) return;
        state.hover = n;
        refreshPreview();
      } else if (state.hover) {
        hoverLeave = window.setTimeout(function () { state.hover = null; refreshPreview(); }, 170);
      }
    }
    function schedulePanel(now) {
      window.clearTimeout(panelTimer);
      if (now) { renderPanel(); return; }
      panelTimer = window.setTimeout(renderPanel, 90);
    }

    // ---------------------------------------------------------------- update
    function refreshPreview(now) {
      paint();
      schedulePanel(now);
    }
    function update(redraw, animate) {
      window.clearTimeout(panelTimer);
      syncSurfaceButtons();
      syncChips();
      if (redraw) drawRadial();
      paint(animate);
      syncPickers();
      renderPanel();
      drawWires();
    }

    // ---------------------------------------------------------------- tour
    // Scenes: an intro, a spotlight on some formats around the circle, routes between formats, then each surface in turn.
    // It starts once the circle is in view and stops at the first touch; the Play tour button starts it again.
    //   ?tour=0   no autoplay        ?speed=1.5   faster or slower (0.4 to 3)
    var query = {};
    (window.location.search || '').replace(/^\?/, '').split('&').forEach(function (kv) {
      var p = kv.split('=');
      try {
        if (p[0]) query[decodeURIComponent(p[0])] = decodeURIComponent((p[1] || '').replace(/\+/g, ' '));
      } catch (e) { /* Ignore a malformed tracking parameter without disabling the map. */ }
    });
    var speed = Math.max(0.4, Math.min(3, parseFloat(query.speed) || 1));
    var tour = { on: false, touched: false, timer: 0, scrollTimer: 0, drift: 0, index: -1, scenes: [], before: 'all', button: null, caption: null };
    var surfaceNotes = {};
    each(root.querySelectorAll('[data-surface][data-note]'), function (b) { surfaceNotes[b.getAttribute('data-surface')] = b.getAttribute('data-note'); });

    // ---- Scenes. Each builder returns null when the map has nothing to show for it (an unknown format, no route on that surface).
    var SPOTS = ['DOCX', 'XLSX', 'PPTX', 'PDF', 'OneNote', 'MSG', 'BibTeX', 'Markdown', 'HTML'];
    var PATHS = [['DOCX', 'Markdown'], ['XLSX', 'PDF'], ['Markdown', 'DOCX'], ['PPTX', 'PDF'], ['OneNote', 'PDF'], ['HTML', 'DOCX'], ['Pages', 'Markdown'], ['EPUB', 'DOCX']];
    function norm(s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
    function findFormat(name) {
      var key = norm(name);
      for (var i = 0; i < order.length; i++) if (norm(order[i]) === key) return order[i];
      return null;
    }
    function introScene() {
      return { kind: 'OfficeIMO', title: order.length + ' formats, ' + routes.length + ' conversions', note: 'Every dot is a format. Every line is a conversion.', surface: 'all', dwell: 3600 };
    }
    function spotScene(n, surface) {
      surface = surface || 'all';
      var outs = routes.filter(function (r) { return r.source === n && (surface === 'all' || r.surfaces.indexOf(surface) >= 0); });
      if (!outs.length) return null;
      var inBrowser = outs.filter(function (r) { return r.surfaces.indexOf('browser') >= 0; }).length;
      var where = surface === 'all' ? (inBrowser ? ' · ' + inBrowser + ' in your browser' : '') : ' in ' + surfaceNames[surface];
      return { kind: families[familyOf[n]], family: familyOf[n], title: n, from: n, surface: surface, dwell: 4200, note: 'Converts to ' + plural(outs.length, 'format') + where };
    }
    function pathScene(a, b, surface) {
      surface = surface || 'all';
      var keep = state.surface;
      state.surface = surface;
      var paths = shortest(a, b);
      state.surface = keep;
      if (!paths.length) return null;
      var steps = paths[0].length;
      var first = paths[0][0];
      return { kind: 'Find the way', title: a + ' → ' + b, from: a, to: b, surface: surface, dwell: 3600 + steps * 900,
        note: steps === 1 ? 'One step · ' + (surface === 'powershell' && first.cmdlet ? first.cmdlet : first.pkg) : steps + ' steps · each one is its own conversion with its own report' };
    }
    function surfaceScene(id) {
      if (!counts[id]) return null;
      return { kind: 'Where it runs', title: surfaceNames[id] + ' · ' + plural(counts[id], 'conversion'), surface: id, dwell: 3800, note: surfaceNotes[id] || '' };
    }
    function compact(list) { return list.filter(function (s) { return !!s; }); }
    function defaultSpots(surface) { return compact(SPOTS.map(function (n) { return findFormat(n) && spotScene(findFormat(n), surface); })); }
    function defaultPaths(surface) { return compact(PATHS.map(function (p) { var a = findFormat(p[0]), b = findFormat(p[1]); return a && b && pathScene(a, b, surface); })); }
    function defaultSurfaces() { return compact(['dotnet', 'browser', 'cli', 'studio', 'powershell'].map(surfaceScene)); }

    function buildScenes() {
      return [introScene()].concat(defaultSpots('all'), defaultPaths('all'), defaultSurfaces());
    }

    function showCaption(s, ms) {
      var c = tour.caption;
      if (!c) return;
      c.className = 'imo-fmap__caption' + (s.family ? ' imo-fam--' + s.family : s.surface !== 'all' ? ' imo-fmap__caption--' + s.surface : '');
      c.querySelector('.imo-fmap__capkind').textContent = s.kind;
      c.querySelector('.imo-fmap__captitle').textContent = s.title;
      c.querySelector('.imo-fmap__capnote').textContent = s.note || '';
      var bar = c.querySelector('i');
      bar.style.animation = 'none';
      void bar.offsetWidth;
      bar.style.animation = 'imo-fmap-bar ' + ms + 'ms linear forwards';
      c.classList.remove('is-swap');
      void c.offsetWidth;
      c.classList.add('is-swap');
    }

    function runScene(s) {
      window.clearTimeout(hoverLeave);
      window.clearTimeout(tour.scrollTimer);
      state.hover = state.focus = null;
      var changed = s.surface !== state.surface;
      state.surface = s.surface;
      state.from = s.from || null;
      state.to = s.to || null;
      update(changed, true);
      var ms = s.dwell / speed;
      showCaption(s, ms);
      // A long answer drifts upward so the whole list shows during the scene.
      if (side && !still) {
        var token = ++tour.drift;
        tour.scrollTimer = window.setTimeout(function () {
          var room = side.scrollHeight - side.clientHeight;
          if (room <= 8) return;
          var from = side.scrollTop, began = performance.now(), length = Math.min(1200, ms * 0.3);
          var frame = function (now) {
            if (token !== tour.drift) return;
            var u = Math.min(1, (now - began) / length);
            side.scrollTop = from + (room - from) * (u < 0.5 ? 2 * u * u : 1 - Math.pow(-2 * u + 2, 2) / 2);
            if (u < 1) window.requestAnimationFrame(frame);
          };
          window.requestAnimationFrame(frame);
        }, ms * 0.45);
      }
    }

    // Moves to the next scene. Waits while the tab is hidden or the circle is out of view.
    function advance() {
      window.clearTimeout(tour.timer);
      if (!tour.on) return;
      if (document.hidden || !onScreen) { tour.timer = window.setTimeout(advance, 400); return; }
      tour.index = (tour.index + 1) % tour.scenes.length;
      var scene = tour.scenes[tour.index];
      runScene(scene);
      tour.timer = window.setTimeout(advance, scene.dwell / speed);
    }

    function syncTourButton() {
      if (!tour.button) return;
      tour.button.setAttribute('aria-pressed', tour.on ? 'true' : 'false');
      tour.button.lastChild.textContent = tour.on ? 'Stop tour' : 'Play tour';
    }
    function startTour() {
      if (tour.on || !svg) return;
      state.hidden = {};
      legend.forEach(function (b) { b.setAttribute('aria-pressed', 'true'); });
      update(true);
      tour.scenes = buildScenes();
      tour.on = true;
      panel.setAttribute('aria-live', 'off');
      tour.index = -1;
      tour.before = state.surface;
      root.classList.add('is-touring');
      syncTourButton();
      advance();
    }
    function stopTour(keepView) {
      if (!tour.on) return;
      tour.on = false;
      ++tour.drift;
      window.clearTimeout(tour.timer);
      window.clearTimeout(tour.scrollTimer);
      root.classList.remove('is-touring');
      panel.setAttribute('aria-live', 'polite');
      syncTourButton();
      // Keep the current DOM intact until a user's click or keyboard activation completes.
      if (keepView) return;
      state.from = state.to = null;
      var changed = state.surface !== tour.before;
      state.surface = tour.before;
      update(changed);
      syncTourButton();
    }

    if (svg && order.length) {
      var canvas = root.querySelector('.imo-fmap__canvas');
      tour.caption = el('div', 'imo-fmap__caption');
      tour.caption.setAttribute('aria-hidden', 'true');
      tour.caption.appendChild(el('span', 'imo-fmap__capkind'));
      tour.caption.appendChild(el('b', 'imo-fmap__captitle'));
      tour.caption.appendChild(el('span', 'imo-fmap__capnote'));
      tour.caption.appendChild(el('i'));
      canvas.appendChild(tour.caption);
      // The button sits in the circle's empty top-left corner, so it never wraps or moves with the chip row above.
      tour.button = el('button', 'imo-fmap__tourbtn');
      tour.button.type = 'button';
      tour.button.setAttribute('aria-pressed', 'false');
      tour.button.appendChild(el('i'));
      tour.button.appendChild(document.createTextNode('Play tour'));
      tour.button.addEventListener('click', function () { tour.touched = true; if (tour.on) stopTour(); else startTour(); });
      canvas.appendChild(tour.button);
      // Touching the map ends the tour.
      root.addEventListener('click', function (e) { if (e.target !== tour.button && !tour.button.contains(e.target)) { tour.touched = true; stopTour(true); } }, true);
      root.addEventListener('change', function () { tour.touched = true; stopTour(true); }, true);
      root.addEventListener('keydown', function (e) { if (e.target !== tour.button && e.key !== 'Tab') { tour.touched = true; stopTour(true); } }, true);
    }

    var pending = 0;
    function relayout(animate) {
      if (pending) return;
      pending = window.requestAnimationFrame(function () { pending = 0; drawWires(animate); });
    }
    if (wires) {
      window.addEventListener('resize', relayout);
      if (window.ResizeObserver) new window.ResizeObserver(relayout).observe(root.querySelector('.imo-fmap__board'));
    }

    // Choosing a format must not move what is below the map. Where the answer panel is not pinned to the height of the circle
    // (the list view, and the stacked layouts), reserve the height of the tallest answer, measured once per layout.
    var reserveTimer = 0;
    function reservePanelHeight() {
      if (!side) return;
      side.style.minHeight = '';
      if (/size/.test(getComputedStyle(side).contain || '')) return;
      var keep = { from: state.from, to: state.to, hover: state.hover, focus: state.focus, surface: state.surface, top: side.scrollTop };
      var tallest = 0;
      state.hover = state.focus = state.to = null;
      state.surface = 'all';
      order.forEach(function (n) {
        if (!outOf(n).length) return;
        state.from = n;
        buildPanel();
        tallest = Math.max(tallest, side.offsetHeight);
      });
      state.from = keep.from; state.to = keep.to; state.hover = keep.hover; state.focus = keep.focus; state.surface = keep.surface;
      buildPanel();
      side.scrollTop = keep.top;
      if (tallest) side.style.minHeight = tallest + 'px';
      drawWires();
    }
    function scheduleReserve() {
      window.clearTimeout(reserveTimer);
      reserveTimer = window.setTimeout(reservePanelHeight, 150);
    }
    window.addEventListener('resize', scheduleReserve);
    if (document.fonts && document.fonts.ready) document.fonts.ready.then(scheduleReserve);

    if (view === 'list') state.from = 'DOCX';
    update(true);
    scheduleReserve();
    // The tour starts by itself as soon as the circle is in view, unless the visitor got there first (hovered or picked
    // something) or asked for no tour (?tour=0). Reduced motion keeps the button only.
    if (svg && view === 'radial' && !still && query.tour !== '0') {
      if (window.IntersectionObserver) {
        new window.IntersectionObserver(function (entries, observer) {
          if (!entries[0].isIntersecting) return;
          observer.disconnect();
          if (!tour.on && !tour.touched && !state.from && !state.hover && !state.focus) startTour();
        }, { threshold: 0.4 }).observe(svg);
      }
    }
  }

  function start() { each(document.querySelectorAll('[data-format-map]'), init); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
