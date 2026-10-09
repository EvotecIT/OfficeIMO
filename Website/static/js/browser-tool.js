/*
 * OfficeIMO browser tools. Drives one statically rendered tool page (layouts/browser-tool.html).
 * The page works from data-* attributes produced from data/browser_tools.json; the OfficeIMO engine
 * runs in a Web Worker (engine-worker.js) so the page never freezes while a file is processed.
 * Written in ES5 style to match the site's minifier.
 */
(function () {
  'use strict';

  var root = document.querySelector('[data-browser-tool]');
  if (!root) return;

  var MAX_FILE_BYTES = 25 * 1024 * 1024;
  var MAX_TOTAL_BYTES = 75 * 1024 * 1024;
  var HANDOFF_DB = 'officeimo-browser-tools';
  var HANDOFF_TTL_MS = 10 * 60 * 1000;

  var cfg = {
    id: root.getAttribute('data-browser-tool'),
    base: root.getAttribute('data-engine-base') || '/apps/officeimo-converter/',
    kind: root.getAttribute('data-kind'),
    target: root.getAttribute('data-target') || '',
    action: root.getAttribute('data-action'),
    commit: root.getAttribute('data-commit') || '',
    find: root.getAttribute('data-find') || '',
    confirm: root.getAttribute('data-confirm') === 'true',
    assemblies: split(root.getAttribute('data-assemblies'), ','),
    input: root.getAttribute('data-input') || 'file',
    accept: split((root.getAttribute('data-accept') || '').toLowerCase(), ','),
    label: root.getAttribute('data-label') || 'file',
    min: parseInt(root.getAttribute('data-min'), 10) || 0,
    max: parseInt(root.getAttribute('data-max'), 10) || 0,
    auto: root.getAttribute('data-auto') === 'true',
    live: root.getAttribute('data-live') === 'true',
    sample: root.getAttribute('data-sample') || '',
    sampleName: root.getAttribute('data-sample-name') || '',
    sampleSecond: root.getAttribute('data-sample-second') || '',
    sampleSecondName: root.getAttribute('data-sample-second-name') || '',
    sampleNote: root.getAttribute('data-sample-note') || '',
    sampleOptions: split(root.getAttribute('data-sample-options'), '|'),
    slots: split(root.getAttribute('data-slots'), '|'),
    runLabel: root.getAttribute('data-run-label') || 'Run',
    confirmLabel: root.getAttribute('data-confirm-label') || '',
    busy: split(root.getAttribute('data-busy'), '|')
  };
  if (cfg.input === 'pair') { cfg.min = 2; cfg.max = 2; }
  if (cfg.input === 'file') { cfg.min = 1; cfg.max = 1; }

  var ui = {
    steps: document.querySelectorAll('[data-bt-steps] li'),
    status: document.querySelector('[data-bt-engine-text]'),
    statusBox: document.querySelector('[data-bt-engine-status]'),
    fileInput: root.querySelector('[data-bt-file-input]'),
    drop: root.querySelector('[data-bt-drop]'),
    files: root.querySelector('[data-bt-files]'),
    sample: root.querySelector('[data-bt-sample]'),
    text: root.querySelector('[data-bt-text]'),
    resetText: root.querySelector('[data-bt-reset-text]'),
    form: root.querySelector('[data-bt-options]'),
    pages: root.querySelector('[data-bt-pages]'),
    run: root.querySelector('[data-bt-run]'),
    hint: root.querySelector('[data-bt-hint]'),
    output: root.querySelector('[data-bt-output]'),
    next: document.querySelector('[data-bt-next]')
  };

  var state = {
    files: [],          // { name, size, ext, buffer }
    text: ui.text ? ui.text.value : '',
    textFromFile: false,
    result: null,
    resultAction: '',
    artifacts: [],      // { info, blob, url }
    urls: [],
    busy: false,
    stale: false,
    probe: null,
    grid: null,
    generation: 0
  };

  // ---------------------------------------------------------------- helpers
  function split(value, separator) {
    if (!value) return [];
    return value.split(separator).map(function (part) { return part.trim(); }).filter(function (part) { return part.length > 0; });
  }
  function el(tag, className, text) {
    var node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined && text !== null) node.textContent = text;
    return node;
  }
  function extOf(name) {
    var dot = name.lastIndexOf('.');
    return dot > 0 ? name.slice(dot).toLowerCase() : '';
  }
  function bytes(size) {
    if (size < 1024) return size + ' B';
    var units = ['KB', 'MB', 'GB'];
    var value = size / 1024, unit = 0;
    while (value >= 1024 && unit < units.length - 1) { value /= 1024; unit++; }
    return (value >= 100 ? Math.round(value) : Math.round(value * 10) / 10) + ' ' + units[unit];
  }
  function debounce(fn, wait) {
    var timer = 0;
    return function () { window.clearTimeout(timer); timer = window.setTimeout(fn, wait); };
  }
  function plural(count, one, many) { return count + ' ' + (count === 1 ? one : (many || one + 's')); }
  function setHint(text, tone) {
    if (!ui.hint) return;
    ui.hint.textContent = text || '';
    ui.hint.setAttribute('data-tone', tone || '');
  }
  function setStep(step) {
    for (var i = 0; i < ui.steps.length; i++) {
      var n = i + 1;
      ui.steps[i].className = n < step ? 'is-done' : (n === step ? 'is-now' : '');
    }
  }
  function revokeUrls() {
    state.urls.forEach(function (url) { URL.revokeObjectURL(url); });
    state.urls = [];
  }
  function objectUrl(blob) {
    var url = URL.createObjectURL(blob);
    state.urls.push(url);
    return url;
  }

  // ---------------------------------------------------------------- engine
  var engine = (function () {
    var worker = null, seq = 0, waiting = {}, progress = [], broken = null;
    function fail(message) {
      broken = new Error(message);
      Object.keys(waiting).forEach(function (key) { waiting[key].reject(broken); delete waiting[key]; });
    }
    function start() {
      if (worker || broken) return worker;
      try {
        worker = new Worker(cfg.base + 'engine-worker.js', { type: 'module' });
      } catch (error) {
        fail('This browser can’t run the tool. Try a current version of Edge, Chrome, Firefox or Safari.');
        return null;
      }
      worker.onmessage = function (event) {
        var data = event.data;
        if (data.type === 'progress') { progress.forEach(function (fn) { fn(data.bytes); }); return; }
        var entry = waiting[data.id];
        if (!entry) return;
        delete waiting[data.id];
        if (data.ok) entry.resolve(data.value); else entry.reject(new Error(data.error));
      };
      worker.onerror = function (event) {
        if (event && event.preventDefault) event.preventDefault();
        fail('The tool couldn’t start. Check your connection, then reload the page.');
      };
      return worker;
    }
    function call(type, payload, transfer) {
      return new Promise(function (resolve, reject) {
        var w = start();
        if (!w) { reject(broken); return; }
        var id = ++seq;
        waiting[id] = { resolve: resolve, reject: reject };
        var message = payload || {};
        message.id = id;
        message.type = type;
        w.postMessage(message, transfer || []);
      });
    }
    return { call: call, onProgress: function (fn) { progress.push(fn); } };
  })();

  var ready = engine.call('prepare', { kind: cfg.kind, assemblies: cfg.assemblies });
  var engineReady = false;
  engine.onProgress(function (loaded) {
    if (engineReady || !ui.status) return;
    ui.status.textContent = 'Getting the tool ready · ' + bytes(loaded);
    if (ui.statusBox) ui.statusBox.setAttribute('data-state', 'loading');
  });
  ready.then(function () {
    engineReady = true;
    if (ui.status) ui.status.textContent = 'Ready · runs in this tab, nothing is uploaded';
    if (ui.statusBox) ui.statusBox.setAttribute('data-state', 'ready');
  }, function (error) {
    if (ui.status) ui.status.textContent = error.message;
    if (ui.statusBox) ui.statusBox.setAttribute('data-state', 'error');
  });

  // ---------------------------------------------------------------- input: files
  function acceptFile(name) {
    return cfg.accept.length === 0 || cfg.accept.indexOf(extOf(name)) >= 0;
  }
  function describeAccept() {
    return cfg.accept.map(function (ext) { return ext.replace('.', '').toUpperCase(); }).join(', ');
  }
  function readFile(file) {
    return new Promise(function (resolve, reject) {
      if (file.arrayBuffer) { file.arrayBuffer().then(resolve, reject); return; }
      var reader = new FileReader();
      reader.onload = function () { resolve(reader.result); };
      reader.onerror = function () { reject(reader.error); };
      reader.readAsArrayBuffer(file);
    });
  }

  function addFiles(list) {
    var incoming = Array.prototype.slice.call(list || []);
    if (incoming.length === 0) return Promise.resolve();
    var bad = incoming.filter(function (file) { return !acceptFile(file.name); });
    if (bad.length) { setHint(bad[0].name + ' isn’t supported here. Choose ' + describeAccept() + '.', 'bad'); return Promise.resolve(); }
    var tooBig = incoming.filter(function (file) { return file.size > MAX_FILE_BYTES; });
    if (tooBig.length) { setHint(tooBig[0].name + ' is ' + bytes(tooBig[0].size) + '. The limit here is 25 MB per file.', 'bad'); return Promise.resolve(); }
    var keep = cfg.max > 1 ? state.files.slice() : [];
    var room = cfg.max > 1 ? cfg.max - keep.length : 1;
    if (room <= 0) { setHint('This tool takes up to ' + cfg.max + ' files. Remove one first.', 'bad'); return Promise.resolve(); }
    incoming = incoming.slice(0, room);
    return Promise.all(incoming.map(function (file) {
      return readFile(file).then(function (buffer) { return { name: file.name, size: file.size, ext: extOf(file.name), buffer: buffer }; });
    })).then(function (loaded) {
      var all = keep.concat(loaded);
      var total = all.reduce(function (sum, file) { return sum + file.size; }, 0);
      if (total > MAX_TOTAL_BYTES) { setHint('Together these files are ' + bytes(total) + '. The limit is 75 MB.', 'bad'); return; }
      return setFiles(all);
    }, function () { setHint('That file couldn’t be read. It may be open in another program.', 'bad'); });
  }

  function setFiles(files) {
    state.files = files;
    state.generation++;
    state.probe = null;
    clearResult();
    renderFiles();
    // The visitor can press the button straight away; the run waits for staging, which waits for the engine.
    state.staging = stageFiles();
    setStep(inputComplete() ? 2 : 1);
    updateRunButton();
    return state.staging.then(afterInputChanged);
  }

  function stageFiles() {
    var generation = state.generation;
    var chain = engine.call('clear');
    state.files.forEach(function (file, index) {
      chain = chain.then(function () {
        if (generation !== state.generation) return;
        var copy = file.buffer.slice(0);
        return engine.call('stage', { slot: index, bytes: copy, name: file.name }, [copy]);
      });
    });
    return chain.then(function () {
      if (generation !== state.generation || cfg.kind !== 'pdf' || state.files.length !== 1) return;
      return engine.call('probe', { slot: 0 }).then(function (probe) {
        if (generation === state.generation) state.probe = probe;
      });
    }).catch(function (error) { setHint(error.message, 'bad'); });
  }

  function renderFiles() {
    if (!ui.files) return;
    ui.files.innerHTML = '';
    ui.files.hidden = state.files.length === 0;
    if (ui.drop) ui.drop.classList.toggle('has-files', state.files.length > 0);
    state.files.forEach(function (file, index) {
      var item = el('li', 'bt-file');
      var badge = el('span', 'bt-file__badge', file.ext.replace('.', '').toUpperCase());
      badge.setAttribute('data-format', file.ext.replace('.', ''));
      var meta = el('span', 'bt-file__meta');
      if (cfg.slots[index]) meta.appendChild(el('small', 'bt-file__slot', cfg.slots[index]));
      meta.appendChild(el('b', null, file.name));
      meta.appendChild(el('small', null, bytes(file.size)));
      item.appendChild(badge);
      item.appendChild(meta);
      var actions = el('span', 'bt-file__actions');
      if (cfg.input === 'files' || cfg.input === 'pair') {
        if (index > 0) actions.appendChild(fileButton('↑', 'Move ' + file.name + ' up', function () { move(index, -1); }));
        if (index < state.files.length - 1) actions.appendChild(fileButton('↓', 'Move ' + file.name + ' down', function () { move(index, 1); }));
        actions.appendChild(fileButton('×', 'Remove ' + file.name, function () { remove(index); }));
      } else {
        var change = el('button', 'bt-link', 'Change');
        change.type = 'button';
        change.addEventListener('click', function () { ui.fileInput.click(); });
        actions.appendChild(change);
      }
      item.appendChild(actions);
      ui.files.appendChild(item);
    });
    if (ui.drop) {
      // A single-file tool shows the chosen file with its own Change link, so the drop zone steps aside.
      ui.drop.hidden = state.files.length > 0 && state.files.length >= Math.max(cfg.max, 1);
      if (ui.sample) ui.sample.hidden = state.files.length > 0;
      var choose = ui.drop.querySelector('[data-bt-choose]');
      if (choose) {
        if (state.files.length === 0) choose.textContent = cfg.max > 1 ? 'Choose files' : 'Choose a file';
        else choose.textContent = cfg.max > 1 && state.files.length < cfg.max ? 'Add more' : 'Choose a different file';
        choose.className = state.files.length === 0 ? 'bt-btn bt-btn--primary' : 'bt-btn';
      }
    }
  }
  function fileButton(text, label, handler) {
    var button = el('button', 'bt-icon-btn', text);
    button.type = 'button';
    button.setAttribute('aria-label', label);
    button.addEventListener('click', handler);
    return button;
  }
  function move(index, offset) {
    var files = state.files.slice();
    var file = files.splice(index, 1)[0];
    files.splice(index + offset, 0, file);
    setFiles(files);
  }
  function remove(index) {
    var files = state.files.slice();
    files.splice(index, 1);
    setFiles(files);
  }

  function loadSample() {
    if (!cfg.sample) return Promise.resolve();
    setHint('Loading the sample…');
    function get(path) {
      return fetch(cfg.base + path).then(function (response) {
        if (!response.ok) throw new Error('The sample couldn’t be downloaded.');
        return response.arrayBuffer();
      });
    }
    var wantsTwo = cfg.input === 'pair' || cfg.input === 'files';
    return Promise.all([get(cfg.sample), wantsTwo && cfg.sampleSecond ? get(cfg.sampleSecond) : Promise.resolve(null)]).then(function (buffers) {
      var buffer = buffers[0];
      var files = [{ name: cfg.sampleName || cfg.sample.split('/').pop(), size: buffer.byteLength, ext: extOf(cfg.sample), buffer: buffer }];
      if (wantsTwo) {
        var second = buffers[1] || buffer.slice(0);
        files.push({ name: buffers[1] ? (cfg.sampleSecondName || cfg.sampleSecond.split('/').pop()) : 'copy-of-' + files[0].name, size: second.byteLength, ext: files[0].ext, buffer: second });
      }
      // Samples can prefill fields, such as the published password of the protected sample.
      cfg.sampleOptions.forEach(function (pair) {
        var at = pair.indexOf('=');
        var field = ui.form && at > 0 ? ui.form.querySelector('[name="' + pair.slice(0, at) + '"]') : null;
        if (field) field.value = pair.slice(at + 1);
      });
      return setFiles(files).then(function () { if (cfg.sampleNote) setHint(cfg.sampleNote); });
    }).catch(function (error) { setHint(error.message, 'bad'); });
  }

  // ---------------------------------------------------------------- input: text
  function setText(value, fromFile) {
    state.text = value;
    state.textFromFile = !!fromFile;
    if (ui.text && ui.text.value !== value) ui.text.value = value;
    afterInputChanged();
  }

  // ---------------------------------------------------------------- options
  function collectOptions() {
    var options = {};
    var problem = '';
    if (ui.form) {
      var fields = ui.form.querySelectorAll('input, select, textarea');
      for (var i = 0; i < fields.length; i++) {
        var field = fields[i];
        if (!field.name || field.name.indexOf('__confirm') > 0) continue;
        if ((field.type === 'radio' || field.type === 'checkbox') && !field.checked) continue;
        options[field.name] = field.value;
        if (field.hasAttribute('data-required') && !field.value.trim()) problem = problem || 'Fill in “' + labelFor(field) + '”.';
      }
      var confirms = ui.form.querySelectorAll('[data-confirms]');
      for (var c = 0; c < confirms.length; c++) {
        var original = ui.form.querySelector('[name="' + confirms[c].getAttribute('data-confirms') + '"]');
        if (original && original.value && original.value !== confirms[c].value) problem = problem || 'The two passwords don’t match.';
      }
    }
    if (state.grid) {
      var pages = state.grid.value();
      if (!pages.value) problem = problem || pages.problem;
      else if (pages.problem) problem = problem || pages.problem;
      options[state.grid.name] = pages.value;
    }
    if (cfg.id === 'protect-pdf' && !options.ownerPassword) options.ownerPassword = options.userPassword || '';
    return { values: options, problem: problem };
  }
  function labelFor(field) {
    var label = field.closest('label');
    var span = label && label.querySelector('span');
    return span ? span.textContent.trim() : field.name;
  }

  function inputComplete() {
    if (cfg.input === 'text') return (state.text || '').trim().length > 0 || state.textFromFile;
    return state.files.length >= Math.max(cfg.min, 1) && (cfg.max === 0 || state.files.length <= cfg.max);
  }

  function updateRunButton() {
    if (!ui.run) return;
    var complete = inputComplete();
    var options = collectOptions();
    var label = state.grid ? state.grid.actionLabel(cfg.runLabel) : cfg.runLabel;
    if (state.result && state.result.ok && !state.stale && (cfg.auto || cfg.live)) label = cfg.runLabel;
    ui.run.textContent = state.busy ? 'Working…' : label;
    // Once there is a fresh result, the download is the next step, so the run button steps back.
    var settled = state.result && state.result.ok && !state.stale && !state.busy;
    ui.run.className = 'bt-btn bt-btn--block' + (settled ? '' : ' bt-btn--primary');
    ui.run.disabled = state.busy || !complete || !!options.problem || !!(state.probe && state.probe.needsPassword && cfg.target !== 'unlock');
    // Tools that run as soon as a file arrives have nothing to "check again" before the first file.
    var waitingForFirstFile = !!cfg.auto && cfg.input !== 'text' && !complete;
    ui.run.hidden = waitingForFirstFile;
    if (state.busy) return;
    if (waitingForFirstFile) { setHint(''); return; }
    if (!complete) {
      if (cfg.input === 'pair') setHint(state.files.length === 1 ? 'Add the second PDF to compare.' : 'Add two PDFs to compare.');
      else if (cfg.input === 'files') setHint(state.files.length === 1 ? 'Add at least one more PDF.' : 'Add two or more PDFs.');
      else if (cfg.input === 'text') setHint('Type or paste some ' + cfg.label + ' first.');
      else setHint('Add a file to start.');
    } else if (state.probe && state.probe.needsPassword && cfg.target !== 'unlock') {
      setHint('This PDF needs a password before it can be changed. Unlock it first.', 'bad');
    } else if (options.problem) {
      setHint(options.problem, 'warn');
    } else if (state.stale) {
      setHint('Settings changed. Run it again to update the result.', 'warn');
    } else if (!state.result) {
      setHint('Ready.');
    }
  }

  function afterInputChanged() {
    setupPages();
    if (cfg.target === 'unlock' && state.probe && !state.probe.encrypted) setHint('This PDF doesn’t have a password, so there’s nothing to remove.', 'warn');
    setStep(inputComplete() ? 2 : 1);
    updateRunButton();
    if (!inputComplete()) return;
    if (cfg.auto && !(state.probe && state.probe.needsPassword)) run(cfg.action);
  }

  // ---------------------------------------------------------------- page picker
  function setupPages() {
    if (!ui.pages) return;
    var count = state.probe && state.probe.ok ? state.probe.pageCount : 0;
    if (state.grid && state.grid.count === count) return;
    state.grid = null;
    var empty = ui.pages.querySelector('[data-bt-pages-empty]');
    var old = ui.pages.querySelector('.bt-grid-picker');
    if (old) old.parentNode.removeChild(old);
    if (!count) { if (empty) empty.hidden = false; return; }
    if (empty) empty.hidden = true;
    state.grid = pageGrid(ui.pages, ui.pages.getAttribute('data-bt-pages'), ui.pages.getAttribute('data-option'), count);
  }

  function pageGrid(host, mode, name, count) {
    var selected = {};
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
      entries.forEach(function (entry) {
        if (!entry.isIntersecting) return;
        observer.unobserve(entry.target);
        drawThumb(parseInt(entry.target.getAttribute('data-page'), 10));
      });
    }, { rootMargin: '200px' }) : null;

    function drawThumb(page) {
      var thumb = thumbs[page];
      if (!thumb || thumb.drawn) return;
      thumb.drawn = true;
      engine.call('render', { source: 'input', index: 0, page: page, size: 220 }).then(function (buffer) {
        if (!buffer || !buffer.byteLength) return;
        var img = el('img');
        img.alt = '';
        img.src = objectUrl(new Blob([buffer], { type: 'image/png' }));
        thumb.sheet.innerHTML = '';
        thumb.sheet.appendChild(img);
      });
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
      var checked = ui.form && ui.form.querySelector('input[name="rotation"]:checked');
      return checked ? parseInt(checked.value, 10) : 90;
    }

    function render() {
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
      if (document.activeElement !== typed) typed.value = text();
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
      typed.setAttribute('aria-invalid', list ? 'false' : 'true');
      if (!list) return;
      if (mode === 'order') {
        var unique = {};
        list.forEach(function (n) { unique[n] = true; });
        if (list.length !== count || Object.keys(unique).length !== count) return;
        order = list;
      } else {
        selected = {};
        list.forEach(function (n) { selected[n] = true; });
      }
      render(); changed();
    });
    typed.addEventListener('blur', function () { typed.value = text(); typed.removeAttribute('aria-invalid'); });

    function changed() { markStale(); updateRunButton(); }

    if (ui.form) ui.form.addEventListener('change', function (event) { if (event.target.name === 'rotation') render(); });

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
      name: name,
      value: function () {
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

  // ---------------------------------------------------------------- run
  var busyTimer = 0;
  function run(action, extra) {
    if (state.busy) return;
    action = action || cfg.action;
    var collected = collectOptions();
    if (collected.problem && (action === cfg.action || action === cfg.find)) { setHint(collected.problem, 'warn'); return; }
    var options = collected.values;
    if (extra) Object.keys(extra).forEach(function (key) { options[key] = extra[key]; });
    if (cfg.input === 'text') {
      if (cfg.kind === 'text' && state.textFromFile && action === cfg.action) options.source = 'file';
      else options.text = state.text;
    }
    if (cfg.confirm && action === cfg.action) options.confirm = 'true';
    var generation = state.generation;
    state.busy = true;
    state.stale = false;
    updateRunButton();
    renderBusy(action);
    var started = Date.now();
    Promise.all([ready, state.staging || Promise.resolve()]).then(function () {
      return engine.call('run', { kind: cfg.kind, target: cfg.target, action: action, options: options });
    }).then(function (result) {
      if (generation !== state.generation) return null;
      return fetchArtifacts(result).then(function () { return result; });
    }).then(function (result) {
      state.busy = false;
      window.clearInterval(busyTimer);
      if (!result) { updateRunButton(); return; }
      result.clientMs = Date.now() - started;
      showResult(result, action);
      if (webMcpWaiter) { webMcpWaiter(result); webMcpWaiter = null; }
    }).catch(function (error) {
      state.busy = false;
      window.clearInterval(busyTimer);
      showResult({ ok: false, verdict: { tone: 'bad', title: 'The tool couldn’t finish', detail: error.message }, facts: [], items: [], artifacts: [] }, action);
      if (webMcpWaiter) { webMcpWaiter(null, error); webMcpWaiter = null; }
    });
  }

  function renderBusy(action) {
    var light = cfg.live || (cfg.auto && state.result);
    if (light && state.result) { ui.output.classList.add('is-updating'); return; }
    ui.output.innerHTML = '';
    ui.output.classList.remove('is-updating');
    var box = el('div', 'bt-busy');
    box.appendChild(el('span', 'bt-spinner'));
    var title = el('h2', null, engineReady ? 'Working on it' : 'Getting the tool ready');
    box.appendChild(title);
    var steps = el('ol', 'bt-busy__steps');
    var labels = action === cfg.find ? ['Searching the PDF'] : cfg.busy.length ? cfg.busy : ['Processing'];
    labels.forEach(function (label) { steps.appendChild(el('li', null, label)); });
    box.appendChild(steps);
    var clock = el('p', 'bt-hint', 'This runs on your device. Bigger files take longer.');
    box.appendChild(clock);
    ui.output.appendChild(box);
    var started = Date.now();
    window.clearInterval(busyTimer);
    busyTimer = window.setInterval(function () {
      var seconds = Math.round((Date.now() - started) / 1000);
      title.textContent = engineReady ? 'Working on it · ' + seconds + ' s' : 'Getting the tool ready · ' + seconds + ' s';
    }, 1000);
  }

  function fetchArtifacts(result) {
    revokeUrls();
    state.artifacts = [];
    var list = (result && result.artifacts) || [];
    return Promise.all(list.map(function (info) {
      return engine.call('artifact', { index: info.index }).then(function (buffer) {
        var blob = new Blob([buffer], { type: info.contentType.split(';')[0] });
        return { info: info, blob: blob, url: objectUrl(blob), buffer: buffer };
      });
    })).then(function (artifacts) { state.artifacts = artifacts; });
  }

  // ---------------------------------------------------------------- results
  function clearResult() {
    state.result = null;
    state.stale = false;
    revokeUrls();
    state.artifacts = [];
    if (ui.output && !ui.output.querySelector('[data-bt-empty]')) {
      ui.output.innerHTML = '';
      ui.output.appendChild(emptyState());
    }
  }
  var emptyTemplate = ui.output ? ui.output.innerHTML : '';
  function emptyState() {
    var wrap = el('div');
    wrap.innerHTML = emptyTemplate;
    return wrap.firstElementChild || wrap;
  }

  function markStale() {
    if (!state.result || state.busy) return;
    state.stale = true;
    ui.output.classList.add('is-stale');
    if (cfg.live) { liveRun(); return; }
  }

  function showResult(result, action) {
    state.result = result;
    state.resultAction = action;
    if (cfg.kind === 'text' && state.textFromFile && result.ok && result.preview && ui.text && action === cfg.action) {
      ui.text.value = result.preview.text; // show the decoded file; typing afterwards checks the edited text
      state.text = result.preview.text;
    }
    ui.output.classList.remove('is-updating', 'is-stale');
    ui.output.innerHTML = '';
    // Machine-readable outcome for tests and monitoring; the visible verdict is the human answer.
    var primaryInfo = (result.artifacts || []).filter(function (a) { return a.role === 'primary'; })[0];
    ui.output.setAttribute('data-result-state', result.ok ? 'ok' : 'error');
    ui.output.setAttribute('data-result-action', action);
    ui.output.setAttribute('data-conversion-ms', String(result.elapsedMilliseconds || 0));
    ui.output.setAttribute('data-peak-retained-bytes', String(result.peakRetainedBytes || 0));
    ui.output.setAttribute('data-result-bytes', String(primaryInfo ? primaryInfo.bytes : 0));
    var awaitingCommit = result.ok && ((cfg.commit && action === cfg.action && hasSelectable(result)) || (cfg.find && action === cfg.find && result.items && result.items.length));
    setStep(result.ok ? (awaitingCommit ? 2 : 3) : 2);

    ui.output.appendChild(verdictBox(result.verdict, result.ok));
    if (result.facts && result.facts.length) ui.output.appendChild(factList(result.facts));

    if (!result.ok) {
      var retry = el('button', 'bt-btn', 'Try again');
      retry.type = 'button';
      retry.addEventListener('click', function () { run(action); });
      ui.output.appendChild(retry);
      updateRunButton();
      return;
    }

    if (awaitingCommit) {
      ui.output.appendChild(commitPanel(result, action));
    } else if (result.items && result.items.length) {
      ui.output.appendChild(itemList(result.items));
    }

    var primary = state.artifacts.filter(function (a) { return a.info.role === 'primary'; })[0];
    var preview = previewBox(result, primary);
    if (preview) ui.output.appendChild(preview);
    if (!awaitingCommit) ui.output.appendChild(downloads(primary));
    updateRunButton();
    if (cfg.auto && cfg.input !== 'text') setHint(awaitingCommit ? 'Choose what to remove, then save a copy.' : 'Done. Your original file is unchanged.');
    else if (!cfg.live) setHint(awaitingCommit ? 'Review the matches before you continue.' : 'Done. Your original file is unchanged.');
  }

  function hasSelectable(result) {
    return (result.items || []).some(function (item) { return item.selectable; });
  }

  function verdictBox(verdict, ok) {
    var box = el('div', 'bt-verdict bt-verdict--' + ((verdict && verdict.tone) || (ok ? 'good' : 'bad')));
    var icon = el('span', 'bt-verdict__icon', verdict && verdict.tone === 'good' ? '✓' : verdict && verdict.tone === 'bad' ? '!' : verdict && verdict.tone === 'warn' ? '!' : 'i');
    icon.setAttribute('aria-hidden', 'true');
    var text = el('div');
    text.appendChild(el('h2', null, verdict ? verdict.title : ''));
    if (verdict && verdict.detail) text.appendChild(el('p', null, verdict.detail));
    box.appendChild(icon);
    box.appendChild(text);
    return box;
  }

  function factList(facts) {
    var list = el('dl', 'bt-facts');
    facts.forEach(function (fact) {
      var row = el('div', fact.tone ? 'is-' + fact.tone : null);
      row.appendChild(el('dt', null, fact.label));
      row.appendChild(el('dd', null, fact.value));
      list.appendChild(row);
    });
    return list;
  }

  var STATE_LABEL = { found: 'Found', removed: 'Removed', kept: 'Kept', warning: 'Check', info: 'Note', good: 'OK', bad: 'Problem' };
  function itemRow(item, selectable) {
    var row = el('li', 'bt-item bt-item--' + item.state);
    if (selectable) {
      var box = el('input');
      box.type = 'checkbox';
      box.value = item.id;
      box.checked = !!item.selected;
      box.id = 'bt-item-' + item.id;
      row.appendChild(box);
    }
    var body = el(selectable ? 'label' : 'div', 'bt-item__body');
    if (selectable) body.htmlFor = 'bt-item-' + item.id;
    body.appendChild(el('b', null, item.title));
    if (item.detail) body.appendChild(el('span', null, item.detail));
    row.appendChild(body);
    if (item.state !== 'detail') row.appendChild(el('span', 'bt-state bt-state--' + item.state, STATE_LABEL[item.state] || item.state));
    return row;
  }

  function itemList(items) {
    var wrap = el('div', 'bt-items');
    var groups = [];
    var byGroup = {};
    items.forEach(function (item) {
      var key = item.group || '';
      if (!byGroup[key]) { byGroup[key] = []; groups.push(key); }
      byGroup[key].push(item);
    });
    var shown = 0;
    var overflow = el('details', 'bt-items__more');
    var hiddenCount = 0;
    groups.forEach(function (key) {
      var host = shown < 6 ? wrap : overflow;
      if (key) host.appendChild(el('h3', 'bt-items__group', key));
      var list = el('ul', 'bt-item-list');
      byGroup[key].forEach(function (item) { list.appendChild(itemRow(item, false)); shown++; if (host === overflow) hiddenCount++; });
      host.appendChild(list);
    });
    if (hiddenCount) {
      var summary = el('summary', null, 'Show ' + plural(hiddenCount, 'more item'));
      overflow.insertBefore(summary, overflow.firstChild);
      wrap.appendChild(overflow);
    }
    return wrap;
  }

  function commitPanel(result, action) {
    var panel = el('div', 'bt-commit');
    var selectable = result.items.filter(function (item) { return item.selectable; });
    var others = result.items.filter(function (item) { return !item.selectable; });
    var list = el('ul', 'bt-item-list');
    (selectable.length ? selectable : result.items).forEach(function (item) { list.appendChild(itemRow(item, !!item.selectable)); });
    if (cfg.kind === 'text' && selectable.length > 1) {
      var tools = el('div', 'bt-commit__bulk');
      var risky = el('button', 'bt-chip', 'Select the ones that can change meaning');
      risky.type = 'button';
      risky.addEventListener('click', function () {
        selectable.forEach(function (item) { var box = list.querySelector('#bt-item-' + item.id); if (box) box.checked = item.group === 'Can change meaning'; });
        sync();
      });
      var all = el('button', 'bt-chip', 'Select all');
      all.type = 'button';
      all.addEventListener('click', function () { Array.prototype.forEach.call(list.querySelectorAll('input'), function (box) { box.checked = true; }); sync(); });
      var none = el('button', 'bt-chip', 'Clear');
      none.type = 'button';
      none.addEventListener('click', function () { Array.prototype.forEach.call(list.querySelectorAll('input'), function (box) { box.checked = false; }); sync(); });
      tools.appendChild(risky); tools.appendChild(all); tools.appendChild(none);
      panel.appendChild(tools);
    }
    panel.appendChild(list);
    if (others.length && selectable.length) panel.appendChild(itemList(others));
    var go = el('button', 'bt-btn bt-btn--primary bt-btn--block');
    go.type = 'button';
    function chosen() { return Array.prototype.map.call(list.querySelectorAll('input:checked'), function (box) { return box.value; }); }
    function sync() {
      if (cfg.find && action === cfg.find) { go.textContent = cfg.confirmLabel || 'Continue'; go.disabled = false; return; }
      var n = chosen().length;
      go.disabled = n === 0;
      go.textContent = n === 0 ? 'Select at least one to remove' : (cfg.confirmLabel || 'Continue').replace('selected', n === selectable.length ? (n === 1 ? 'it' : 'all ' + n) : n + ' of ' + selectable.length);
    }
    list.addEventListener('change', sync);
    go.addEventListener('click', function () {
      if (cfg.find && action === cfg.find) run(cfg.action);
      else run(cfg.commit, { remove: chosen().join(',') });
    });
    sync();
    panel.appendChild(go);
    panel.appendChild(el('p', 'bt-hint', 'Your original stays as it is. Changes only go into a new copy.'));
    return panel;
  }

  // ---------------------------------------------------------------- previews
  function previewBox(result, primary) {
    var preview = result.preview;
    if (!preview) {
      if (cfg.kind === 'origin' && state.files[0] && /\.(png|jpe?g|webp)$/.test(state.files[0].ext)) return originalImage();
      return null;
    }
    var box = el('figure', 'bt-preview');
    if (preview.kind === 'pdf' && primary) {
      box.appendChild(pdfPager(preview.pageCount || 0, preview.artifactIndex || 0));
    } else if (preview.kind === 'image' && primary) {
      if (cfg.kind === 'origin' && state.files[0]) {
        box.className = 'bt-preview bt-preview--pair';
        box.appendChild(imageFigure(objectUrl(new Blob([state.files[0].buffer])), 'Original'));
        box.appendChild(imageFigure(primary.url, 'Clean copy'));
        box.appendChild(el('figcaption', null, 'These look the same because only hidden data changed.'));
      } else {
        var img = el('img');
        img.src = primary.url;
        img.alt = 'Result preview';
        box.appendChild(img);
      }
    } else if (preview.kind === 'html' && preview.html) {
      var frame = el('iframe', 'bt-preview__frame');
      frame.setAttribute('sandbox', '');
      frame.setAttribute('title', 'Result preview');
      frame.srcdoc = preview.html;
      box.appendChild(frame);
      if (preview.text) box.appendChild(sourceToggle(preview.text, 'Show the HTML'));
    } else if (preview.kind === 'frame' && primary) {
      var gallery = el('iframe', 'bt-preview__frame bt-preview__frame--tall');
      gallery.setAttribute('sandbox', '');
      gallery.setAttribute('title', 'Comparison');
      gallery.src = primary.url;
      box.appendChild(gallery);
    } else if (preview.kind === 'text' && preview.text !== undefined && preview.text !== null) {
      if (cfg.kind === 'text') box.appendChild(markedText(preview.text, result));
      else { var pre = el('pre', 'bt-preview__text'); pre.textContent = preview.text; box.appendChild(pre); }
    } else {
      return null;
    }
    return box;
  }

  function originalImage() {
    var box = el('figure', 'bt-preview');
    box.appendChild(imageFigure(objectUrl(new Blob([state.files[0].buffer])), state.files[0].name));
    return box;
  }
  function imageFigure(url, caption) {
    var fig = el('figure', 'bt-preview__image');
    var img = el('img');
    img.src = url;
    img.alt = caption;
    fig.appendChild(img);
    fig.appendChild(el('figcaption', null, caption));
    return fig;
  }
  function sourceToggle(text, label) {
    var details = el('details', 'bt-source');
    details.appendChild(el('summary', null, label));
    var pre = el('pre', 'bt-preview__text');
    pre.textContent = text;
    details.appendChild(pre);
    var copy = el('button', 'bt-chip', 'Copy');
    copy.type = 'button';
    copy.addEventListener('click', function () {
      if (navigator.clipboard) navigator.clipboard.writeText(text).then(function () { copy.textContent = 'Copied'; }, function () { copy.textContent = 'Select the text to copy'; });
    });
    details.appendChild(copy);
    return details;
  }

  function markedText(text, result) {
    var pre = el('pre', 'bt-preview__text bt-marked');
    var findings = (result.items || []).filter(function (item) { return item.start !== undefined && item.start !== null; })
      .sort(function (a, b) { return a.start - b.start; });
    var offset = 0;
    findings.forEach(function (item) {
      if (item.start < offset) return;
      pre.appendChild(document.createTextNode(text.slice(offset, item.start)));
      var mark = el('mark', 'bt-mark bt-mark--' + item.state, item.title.replace(/.*\((U\+[0-9A-F]+)\)$/, '$1'));
      mark.title = item.title + ' — ' + item.detail;
      pre.appendChild(mark);
      offset = item.start + (item.length || 1);
    });
    pre.appendChild(document.createTextNode(text.slice(offset)));
    return pre;
  }

  function pdfPager(pageCount, artifactIndex) {
    var wrap = el('div', 'bt-pager');
    var stage = el('div', 'bt-pager__page');
    var nav = el('div', 'bt-pager__nav');
    var prev = el('button', 'bt-icon-btn', '‹');
    var next = el('button', 'bt-icon-btn', '›');
    prev.type = next.type = 'button';
    prev.setAttribute('aria-label', 'Previous page');
    next.setAttribute('aria-label', 'Next page');
    var label = el('span', null, '');
    nav.appendChild(prev); nav.appendChild(label); nav.appendChild(next);
    wrap.appendChild(stage);
    wrap.appendChild(nav);
    var page = 1;
    var cache = {};
    function show() {
      label.textContent = pageCount ? 'Page ' + page + ' of ' + pageCount : 'Preview';
      nav.hidden = pageCount <= 1;
      prev.disabled = page <= 1;
      next.disabled = !pageCount || page >= pageCount;
      if (cache[page]) { stage.innerHTML = ''; stage.appendChild(cache[page]); return; }
      stage.innerHTML = '';
      stage.appendChild(el('span', 'bt-spinner'));
      var requested = page;
      engine.call('render', { source: 'artifact', index: artifactIndex, page: page, size: 1100 }).then(function (buffer) {
        var node;
        if (buffer && buffer.byteLength) {
          node = el('img');
          node.alt = 'Page ' + requested + ' of the result';
          node.src = objectUrl(new Blob([buffer], { type: 'image/png' }));
        } else {
          node = el('p', 'bt-hint', 'This page can’t be previewed here, but it’s in the download.');
        }
        cache[requested] = node;
        if (requested === page) { stage.innerHTML = ''; stage.appendChild(node); }
      });
    }
    prev.addEventListener('click', function () { if (page > 1) { page--; show(); } });
    next.addEventListener('click', function () { if (page < pageCount) { page++; show(); } });
    if (!pageCount) engine.call('pages', { source: 'artifact', index: artifactIndex }).then(function (count) { pageCount = count; show(); });
    else show();
    return wrap;
  }

  // ---------------------------------------------------------------- downloads and next steps
  function downloads(primary) {
    var wrap = el('div', 'bt-downloads');
    if (primary) {
      var link = el('a', 'bt-btn bt-btn--primary bt-btn--block', 'Download ' + primary.info.fileName + ' · ' + bytes(primary.info.bytes));
      link.href = primary.url;
      link.download = primary.info.fileName;
      link.setAttribute('data-bt-download', 'primary');
      wrap.appendChild(link);
    }
    var extras = state.artifacts.filter(function (a) { return a.info.role !== 'primary'; });
    var nextLinks = nextTools(primary);
    var supportable = cfg.kind === 'convert' && primary && /pdf$/i.test(primary.info.contentType);
    if (!extras.length && !nextLinks.length && !supportable) return wrap;
    var more = el('details', 'bt-more');
    more.appendChild(el('summary', null, 'More files and next steps'));
    if (extras.length) {
      var files = el('div', 'bt-more__row');
      extras.forEach(function (artifact) {
        var a = el('a', 'bt-chip', ROLE_LABEL[artifact.info.role] || artifact.info.fileName);
        a.href = artifact.url;
        a.download = artifact.info.fileName;
        a.title = artifact.info.fileName;
        files.appendChild(a);
      });
      more.appendChild(files);
    }
    if (supportable) more.appendChild(supportBundle());
    if (nextLinks.length) {
      more.appendChild(el('p', 'bt-more__label', 'Continue with this file'));
      var row = el('div', 'bt-more__row');
      nextLinks.forEach(function (link) { row.appendChild(link); });
      more.appendChild(row);
    }
    more.open = !!nextLinks.length && !extras.length;
    wrap.appendChild(more);
    return wrap;
  }
  var ROLE_LABEL = { report: 'Report (JSON)', overlay: 'Layout check image', support: 'Bug report files (ZIP)' };

  function supportBundle() {
    var box = el('div', 'bt-support');
    var label = el('label', 'bt-check');
    var include = el('input');
    include.type = 'checkbox';
    label.appendChild(include);
    var text = el('span');
    text.appendChild(el('b', null, 'Include my file in the bug report'));
    text.appendChild(el('small', null, 'Off by default. Without it, the ZIP only holds fingerprints and diagnostics.'));
    label.appendChild(text);
    var button = el('button', 'bt-chip', 'Prepare bug report files');
    button.type = 'button';
    button.addEventListener('click', function () {
      button.disabled = true;
      button.textContent = 'Preparing…';
      engine.call('run', { kind: cfg.kind, target: cfg.target, action: 'support', options: { includeContent: include.checked ? 'true' : 'false' } }).then(function (result) {
        if (!result.ok || !result.artifacts.length) throw new Error(result.verdict ? result.verdict.detail : 'Not available.');
        var info = result.artifacts[0];
        return engine.call('artifact', { index: info.index }).then(function (buffer) {
          var a = el('a', 'bt-chip', 'Download bug report files');
          a.href = objectUrl(new Blob([buffer], { type: 'application/zip' }));
          a.download = info.fileName;
          button.replaceWith(a);
        });
      }).catch(function (error) { button.disabled = false; button.textContent = 'Prepare bug report files'; setHint(error.message, 'bad'); });
    });
    box.appendChild(label);
    box.appendChild(button);
    return box;
  }

  function nextTools(primary) {
    if (!ui.next || !primary) return [];
    var ext = extOf(primary.info.fileName);
    return Array.prototype.slice.call(ui.next.querySelectorAll('a')).filter(function (a) {
      return split(a.getAttribute('data-accept') || '', ',').indexOf(ext) >= 0;
    }).map(function (source) {
      var link = el('a', 'bt-chip bt-chip--next', source.textContent);
      link.href = source.getAttribute('href') + '?from=previous';
      link.addEventListener('click', function (event) {
        event.preventDefault();
        handoff({ name: primary.info.fileName, type: primary.blob.type, buffer: primary.buffer }).then(function () {
          window.location.href = link.href;
        }, function () { window.location.href = source.getAttribute('href'); });
      });
      return link;
    });
  }

  // Hand a result to the next tool page. One entry, deleted as soon as the next page reads it.
  function openHandoff() {
    return new Promise(function (resolve, reject) {
      if (!window.indexedDB) { reject(new Error('unavailable')); return; }
      var request = indexedDB.open(HANDOFF_DB, 1);
      request.onupgradeneeded = function () { request.result.createObjectStore('handoff'); };
      request.onsuccess = function () { resolve(request.result); };
      request.onerror = function () { reject(request.error); };
    });
  }
  function handoff(file) {
    return openHandoff().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction('handoff', 'readwrite');
        tx.objectStore('handoff').put({ name: file.name, type: file.type, buffer: file.buffer, created: Date.now() }, 'file');
        tx.oncomplete = function () { db.close(); resolve(); };
        tx.onerror = function () { db.close(); reject(tx.error); };
      });
    });
  }
  function takeHandoff() {
    return openHandoff().then(function (db) {
      return new Promise(function (resolve) {
        var tx = db.transaction('handoff', 'readwrite');
        var store = tx.objectStore('handoff');
        var get = store.get('file');
        get.onsuccess = function () { store.delete('file'); resolve(get.result || null); };
        get.onerror = function () { resolve(null); };
        tx.oncomplete = function () { db.close(); };
      });
    }).catch(function () { return null; });
  }

  // ---------------------------------------------------------------- live text tools
  var liveRun = debounce(function () { if (inputComplete()) run(cfg.action); }, 450);

  // ---------------------------------------------------------------- WebMCP
  // One Website Tool per tool page. It runs the visible tool on the file the visitor already chose; it never
  // accepts paths or bytes, and its output is bounded so an agent can't be flooded by a long file name.
  var webMcpWaiter = null;
  var webMcpRegistration = null;
  function bounded(value, limit) {
    var text = String(value === undefined || value === null ? '' : value).trim();
    if (text.length <= limit) return text;
    var end = limit;
    var code = text.charCodeAt(end - 1);
    if (code >= 0xD800 && code <= 0xDBFF) end--; // never split a surrogate pair
    return text.slice(0, end);
  }
  function webMcpOutput(result, error) {
    var base = { tool: cfg.id, route: cfg.target || cfg.kind };
    if (error || !result || !result.ok) {
      base.success = false;
      base.message = bounded(error ? error.message : (result && result.verdict ? result.verdict.title + '. ' + result.verdict.detail : 'The tool did not finish.'), 300);
      return base;
    }
    var primary = (result.artifacts || []).filter(function (a) { return a.role === 'primary'; })[0];
    base.success = true;
    base.verdict = bounded(result.verdict.title, 160);
    base.outputFileName = primary ? bounded(primary.fileName, 180) : null;
    base.outputBytes = primary ? primary.bytes : 0;
    base.warningCount = (result.items || []).filter(function (item) { return item.state === 'warning'; }).length;
    base.elapsedMilliseconds = result.elapsedMilliseconds || 0;
    base.message = 'Finished locally. Review the visible result and download button before saving the file.';
    return base;
  }
  function registerWebMcp() {
    var context = document.modelContext || navigator.modelContext;
    if (!context || typeof context.registerTool !== 'function' || webMcpRegistration) return;
    var controller = new AbortController();
    webMcpRegistration = controller;
    try {
      Promise.resolve(context.registerTool({
        name: 'convert_selected_document',
        description: 'Run the open OfficeIMO browser tool (' + bounded(document.title, 120) + ') on the file already chosen on the page, with the settings shown. Processing stays in the browser.',
        inputSchema: { type: 'object', properties: {}, additionalProperties: false },
        annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false, untrustedContentHint: true },
        execute: function (input, callContext) {
          if (callContext && callContext.signal && callContext.signal.aborted) return Promise.resolve({ success: false, message: 'Conversion was cancelled before it started.' });
          if (!inputComplete()) return Promise.resolve({ success: false, tool: cfg.id, message: 'No file is chosen yet. Choose a file or load the sample on the page, then try again.' });
          if (state.busy) return Promise.resolve({ success: false, tool: cfg.id, message: 'The tool is already running in this tab.' });
          return new Promise(function (resolve) {
            webMcpWaiter = function (result, error) { resolve(webMcpOutput(result, error)); };
            run(cfg.find || cfg.action);
          });
        }
      }, { signal: controller.signal })).then(function () {
        document.body.setAttribute('data-webmcp-status', 'registered');
      }, function () {
        document.body.setAttribute('data-webmcp-status', 'failed');
      });
    } catch (error) {
      document.body.setAttribute('data-webmcp-status', 'failed');
    }
  }
  window.addEventListener('pagehide', function (event) {
    if (webMcpRegistration) webMcpRegistration.abort();
    webMcpRegistration = null;
    if (!event.persisted) webMcpWaiter = null;
  });
  window.addEventListener('pageshow', function (event) { if (event.persisted) registerWebMcp(); });

  // Read-only diagnostics used by the CI performance budget (Website/scripts/converter-performance.playwright.js).
  window.OfficeIMOBrowserTool = { id: cfg.id, engineMemory: function () { return engine.call('memory'); } };

  // ---------------------------------------------------------------- wiring
  if (ui.fileInput) {
    ui.fileInput.addEventListener('change', function () {
      var files = ui.fileInput.files;
      if (cfg.input === 'text') {
        var file = files && files[0];
        if (!file) return;
        if (cfg.kind === 'text') {
          addTextFile(file);
        } else {
          file.text().then(function (text) { setText(text, false); });
        }
        ui.fileInput.value = '';
        return;
      }
      addFiles(files);
      ui.fileInput.value = '';
    });
  }
  function addTextFile(file) {
    if (file.size > 1024 * 1024) { setHint('Text files up to 1 MB can be checked here.', 'bad'); return; }
    readFile(file).then(function (buffer) {
      state.files = [{ name: file.name, size: file.size, ext: extOf(file.name), buffer: buffer }];
      var copy = buffer.slice(0);
      return engine.call('clear').then(function () { return engine.call('stage', { slot: 0, bytes: copy, name: file.name }, [copy]); });
    }).then(function () {
      state.textFromFile = true;
      run(cfg.action);
    });
  }

  if (ui.drop) {
    ['dragenter', 'dragover'].forEach(function (type) {
      ui.drop.addEventListener(type, function (event) { event.preventDefault(); ui.drop.classList.add('is-dragging'); });
    });
    ['dragleave', 'drop'].forEach(function (type) {
      ui.drop.addEventListener(type, function () { ui.drop.classList.remove('is-dragging'); });
    });
    ui.drop.addEventListener('drop', function (event) {
      event.preventDefault();
      addFiles(event.dataTransfer && event.dataTransfer.files);
    });
  }
  // Dropping anywhere on the page feeds the tool rather than navigating away.
  document.addEventListener('dragover', function (event) { if (event.dataTransfer && Array.prototype.indexOf.call(event.dataTransfer.types || [], 'Files') >= 0) event.preventDefault(); });
  document.addEventListener('drop', function (event) {
    if (!event.dataTransfer || !event.dataTransfer.files || !event.dataTransfer.files.length) return;
    if (ui.drop && ui.drop.contains(event.target)) return;
    event.preventDefault();
    if (cfg.input === 'text') { var file = event.dataTransfer.files[0]; if (cfg.kind === 'text') addTextFile(file); else file.text().then(function (text) { setText(text, false); }); }
    else addFiles(event.dataTransfer.files);
  });

  if (ui.sample) ui.sample.addEventListener('click', loadSample);
  if (ui.text) {
    ui.text.addEventListener('input', function () {
      state.text = ui.text.value;
      state.textFromFile = false;
      if (state.result) markStale();
      updateRunButton();
      if (cfg.live || cfg.auto) liveRun();
    });
  }
  if (ui.resetText) ui.resetText.addEventListener('click', function () { setText(ui.text.getAttribute('data-sample') || '', false); });
  if (ui.form) {
    ui.form.addEventListener('input', function () { markStale(); updateRunButton(); });
    ui.form.addEventListener('change', function () { markStale(); updateRunButton(); });
    ui.form.addEventListener('submit', function (event) { event.preventDefault(); if (!ui.run.disabled) run(cfg.find || cfg.action); });
    Array.prototype.forEach.call(ui.form.querySelectorAll('[data-bt-reveal]'), function (button) {
      button.addEventListener('click', function () {
        var field = button.parentNode.querySelector('input');
        var show = field.type === 'password';
        field.type = show ? 'text' : 'password';
        button.textContent = show ? 'Hide' : 'Show';
      });
    });
  }
  if (ui.run) ui.run.addEventListener('click', function () { run(cfg.find || cfg.action); });

  window.addEventListener('pagehide', revokeUrls);

  // A result handed over from another tool arrives through IndexedDB, once.
  var params = new URLSearchParams(window.location.search);
  if (params.get('from') === 'previous') {
    if (window.history && window.history.replaceState) window.history.replaceState(null, '', window.location.pathname);
    takeHandoff().then(function (file) {
      if (!file || Date.now() - file.created > HANDOFF_TTL_MS) return;
      if (cfg.input === 'text') { setText(new TextDecoder().decode(file.buffer), false); return; }
      if (!acceptFile(file.name)) return;
      setFiles([{ name: file.name, size: file.buffer.byteLength, ext: extOf(file.name), buffer: file.buffer }]).then(function () {
        setHint('Carried over from the previous tool.');
      });
    });
  } else {
    takeHandoff(); // clear anything left behind by an abandoned hand-off
  }

  registerWebMcp();
  updateRunButton();
  if (cfg.input === 'text' && (cfg.live || cfg.auto) && inputComplete()) run(cfg.action);
})();
