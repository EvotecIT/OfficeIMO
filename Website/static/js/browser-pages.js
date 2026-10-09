/* PDF page selection, ordering and thumbnail lifetime. */
(function () {
  'use strict';
  function create(options) {
    var host = options.host, mode = options.mode, name = options.name, count = options.count;
    var el = options.element, engine = options.engine, form = options.form;
    var plural = options.plural;
    var thumbnailUrls = [];
    var generation = options.generation;
    var disposed = false;
    var selected = {};
    var selectionProblem = '';
    var order = [];
    for (var p = 1; p <= count; p++) order.push(p);
    var wrap = el('div', 'bt-grid-picker');
    var bar = el('div', 'bt-grid-picker__bar');
    var grid = el('div', 'bt-thumbs');
    grid.setAttribute('role', mode === 'order' ? 'list' : 'group');
    grid.setAttribute('aria-label', 'Pages');
    var typed = el('input', 'bt-grid-picker__text');
    typed.type = 'text';
    typed.setAttribute('aria-label', mode === 'order' ? 'Page order' : 'Pages, for example 1-3,5,last');
    typed.placeholder = mode === 'order' ? '3,1,2' : 'e.g. 1-3, 5, last';
    var thumbs = {};
    var observer = 'IntersectionObserver' in window ? new IntersectionObserver(function (entries) {
      if (disposed) return;
      entries.forEach(function (entry) {
        if (!entry.isIntersecting || !grid.contains(entry.target)) return;
        observer.unobserve(entry.target);
        drawThumb(parseInt(entry.target.getAttribute('data-page'), 10));
      });
    }, { rootMargin: '200px' }) : null;

    function drawThumb(page) {
      var thumb = thumbs[page];
      if (!thumb || thumb.drawn) return;
      thumb.drawn = true;
      engine.call('render', { source: 'input', index: 0, page: page, size: 220 }).then(function (buffer) {
        if (disposed || generation !== options.currentGeneration()) return;
        if (!buffer || !buffer.byteLength) return;
        var img = el('img');
        img.alt = '';
        img.src = URL.createObjectURL(new Blob([buffer], { type: 'image/png' }));
        thumbnailUrls.push(img.src);
        thumb.sheet.innerHTML = '';
        thumb.sheet.appendChild(img);
      }).catch(function () { if (!disposed) thumb.sheet.textContent = 'Preview unavailable'; });
    }

    function button(text, handler) {
      var b = el('button', 'bt-chip', text);
      b.type = 'button';
      b.addEventListener('click', handler);
      bar.appendChild(b);
      return b;
    }
    if (mode === 'order') {
      button('Reverse order', function () { order.reverse(); render(); changed(); });
      button('Reset', function () { order.sort(function (a, b) { return a - b; }); render(); changed(); });
    } else {
      button('All', function () { order.forEach(function (n) { selected[n] = true; }); render(); changed(); });
      button('None', function () { selected = {}; render(); changed(); });
      button('Odd', function () { selected = {}; order.forEach(function (n) { if (n % 2) selected[n] = true; }); render(); changed(); });
      button('Even', function () { selected = {}; order.forEach(function (n) { if (!(n % 2)) selected[n] = true; }); render(); changed(); });
    }

    function rotation() {
      var checked = form && form.querySelector('input[name="rotation"]:checked');
      return checked ? parseInt(checked.value, 10) : 90;
    }

    function render() {
      if (observer) observer.disconnect();
      grid.innerHTML = '';
      order.forEach(function (page, position) {
        var item = el('button', 'bt-thumb');
        item.type = 'button';
        item.setAttribute('data-page', String(page));
        var isOn = !!selected[page];
        if (mode !== 'order') item.setAttribute('aria-pressed', isOn ? 'true' : 'false');
        item.setAttribute('aria-label', 'Page ' + page + (mode === 'order' ? ', position ' + (position + 1) + '. Use Alt and the arrow keys to move it.' : ''));
        if (isOn) item.classList.add(mode === 'remove' ? 'is-removed' : 'is-selected');
        var sheet = thumbs[page] && thumbs[page].drawn ? thumbs[page].sheet : el('span', 'bt-thumb__sheet');
        if (mode === 'rotate' && isOn) sheet.style.transform = 'rotate(' + rotation() + 'deg)';
        else sheet.style.transform = '';
        thumbs[page] = thumbs[page] || { sheet: sheet, drawn: false };
        thumbs[page].sheet = sheet;
        item.appendChild(sheet);
        item.appendChild(el('span', 'bt-thumb__tick', mode === 'remove' ? '×' : '✓'));
        item.appendChild(el('small', null, mode === 'order' ? (position + 1) + ' · p.' + page : String(page)));
        if (mode === 'order') {
          item.draggable = true;
          item.addEventListener('dragstart', function (event) { event.dataTransfer.setData('text/plain', String(page)); item.classList.add('is-dragging'); });
          item.addEventListener('dragend', function () { item.classList.remove('is-dragging'); });
          item.addEventListener('dragover', function (event) { event.preventDefault(); item.classList.add('is-over'); });
          item.addEventListener('dragleave', function () { item.classList.remove('is-over'); });
          item.addEventListener('drop', function (event) {
            event.preventDefault();
            var from = parseInt(event.dataTransfer.getData('text/plain'), 10);
            if (!from || from === page) return;
            order.splice(order.indexOf(from), 1);
            order.splice(order.indexOf(page), 0, from);
            render(); changed();
          });
          item.addEventListener('keydown', function (event) {
            if (!event.altKey || (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight' && event.key !== 'ArrowUp' && event.key !== 'ArrowDown')) return;
            event.preventDefault();
            var index = order.indexOf(page);
            var target = index + (event.key === 'ArrowLeft' || event.key === 'ArrowUp' ? -1 : 1);
            if (target < 0 || target >= order.length) return;
            order.splice(index, 1);
            order.splice(target, 0, page);
            render(); changed();
            var moved = grid.querySelector('[data-page="' + page + '"]');
            if (moved) moved.focus();
          });
        } else {
          item.addEventListener('click', function () {
            if (selected[page]) delete selected[page]; else selected[page] = true;
            render(); changed();
          });
        }
        grid.appendChild(item);
        // The first rows draw straight away; long documents fill in as they scroll into view.
        if (!thumbs[page].drawn) {
          if (position < 12 || !observer) drawThumb(page); else observer.observe(item);
        }
      });
      if (document.activeElement !== typed) { typed.value = text(); selectionProblem = ''; typed.removeAttribute('aria-invalid'); }
    }

    function selectedList() { return order.filter(function (n) { return selected[n]; }); }
    function ranges(list) {
      var out = [];
      for (var i = 0; i < list.length; i++) {
        var j = i;
        while (j + 1 < list.length && list[j + 1] === list[j] + 1) j++;
        out.push(j > i ? list[i] + '-' + list[j] : String(list[i]));
        i = j;
      }
      return out.join(',');
    }
    function text() { return mode === 'order' ? ranges(order) : ranges(selectedList()); }
    function parse(value) {
      var out = [];
      var parts = value.toLowerCase().replace(/\s+/g, '').split(',');
      for (var i = 0; i < parts.length; i++) {
        if (!parts[i]) continue;
        var match = parts[i].replace(/last/g, String(count)).match(/^(\d+)(?:-(\d+))?$/);
        if (!match) return null;
        var a = +match[1], b = match[2] ? +match[2] : a;
        if (a < 1 || b < 1 || a > count || b > count) return null;
        var step = a <= b ? 1 : -1;
        for (var n = a; step > 0 ? n <= b : n >= b; n += step) out.push(n);
      }
      return out;
    }
    typed.addEventListener('input', function () {
      var list = parse(typed.value);
      selectionProblem = '';
      if (!list) { selectionProblem = 'Enter valid page numbers or ranges within this PDF.'; typed.setAttribute('aria-invalid', 'true'); changed(); return; }
      if (mode === 'order') {
        var unique = {};
        list.forEach(function (n) { unique[n] = true; });
        if (list.length !== count || Object.keys(unique).length !== count) {
          selectionProblem = 'Enter every page exactly once in the new order.';
          typed.setAttribute('aria-invalid', 'true'); changed(); return;
        }
        order = list;
      } else {
        selected = {};
        list.forEach(function (n) { selected[n] = true; });
      }
      typed.setAttribute('aria-invalid', 'false'); render(); changed();
    });
    typed.addEventListener('blur', function () { if (!selectionProblem) { typed.value = text(); typed.removeAttribute('aria-invalid'); } });

    function changed() { options.changed(); }

    var rotationChanged = function (event) { if (event.target.name === 'rotation') render(); };
    if (form) form.addEventListener('change', rotationChanged);

    var field = el('label', 'bt-grid-picker__field');
    field.appendChild(el('span', null, mode === 'order' ? 'Or type the order' : 'Or type pages'));
    field.appendChild(typed);
    wrap.appendChild(bar);
    wrap.appendChild(grid);
    wrap.appendChild(field);
    host.appendChild(wrap);
    render();

    return {
      count: count,
      generation: generation,
      dispose: function () {
        disposed = true;
        if (observer) observer.disconnect();
        if (form) form.removeEventListener('change', rotationChanged);
        thumbnailUrls.forEach(function (url) { URL.revokeObjectURL(url); });
        thumbnailUrls = [];
      },
      name: name,
      value: function () {
        if (selectionProblem) return { value: '', problem: selectionProblem };
        if (mode === 'order') {
          var moved = order.some(function (n, i) { return n !== i + 1; });
          return { value: order.join(','), problem: moved ? '' : 'Drag the pages into a new order first.' };
        }
        var list = selectedList();
        if (!list.length) {
          var verb = window.matchMedia && window.matchMedia('(pointer: coarse)').matches ? 'Tap' : 'Click';
          return { value: '', problem: verb + (mode === 'remove' ? ' the pages you want to delete.' : mode === 'rotate' ? ' the pages you want to rotate.' : ' the pages you want to keep.') };
        }
        if (mode === 'remove' && list.length >= count) return { value: ranges(list), problem: 'You can’t delete every page.' };
        return { value: ranges(list), problem: '' };
      },
      actionLabel: function (fallback) {
        if (mode === 'order') return fallback;
        var n = selectedList().length;
        if (!n) return fallback;
        if (mode === 'remove') return 'Delete ' + plural(n, 'page') + ' and save a copy';
        if (mode === 'rotate') return 'Rotate ' + plural(n, 'page') + ' and save';
        return 'Save ' + plural(n, 'page') + ' as a new PDF';
      }
    };
  }

  window.OfficeIMOBrowserPages = { create: create };
})();
