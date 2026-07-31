(function () {
  var root = document.querySelector('[data-excel-benchmarks]');
  if (!root) return;

  var tbody = root.querySelector('[data-benchmark-matrix]');
  if (!tbody) return;

  var rows = Array.prototype.slice.call(tbody.querySelectorAll('[data-benchmark-row]'));
  var filters = root.querySelectorAll('[data-benchmark-filter]');
  var buttons = root.querySelectorAll('[data-benchmark-sort]');
  var sortMetric = root.querySelector('[data-benchmark-sort-mode]');
  var reset = root.querySelector('[data-benchmark-reset]');
  var count = root.querySelector('[data-benchmark-count]');
  var sortState = { key: 'original', direction: 'none', type: 'number' };

  rows.forEach(function (row) {
    row._libraryCells = {};
    Array.prototype.forEach.call(row.querySelectorAll('[data-library]'), function (cell) {
      row._libraryCells[cell.getAttribute('data-library')] = cell;
    });
    row._filterText = (row.textContent + ' ' + row.getAttribute('data-fastest-library') + ' ' + Object.keys(row._libraryCells).join(' ')).toLowerCase();
  });

  function isFiniteNumber(value) {
    return typeof value === 'number' && isFinite(value);
  }

  function numberValue(value) {
    var match = String(value || '').replace(/,/g, '').match(/-?\d+(?:\.\d+)?/);
    var parsed = match ? Number(match[0]) : NaN;
    return isFiniteNumber(parsed) ? parsed : null;
  }

  function durationValue(cell) {
    if (!cell || cell.querySelector('.imo-benchmark-missing')) return null;
    var attributeValue = numberValue(cell.getAttribute('data-mean-ms'));
    if (isFiniteNumber(attributeValue)) return attributeValue;
    var strong = cell.querySelector('strong');
    var text = (strong && strong.textContent ? strong.textContent : '').trim().toLowerCase();
    var parsed = numberValue(text);
    if (!isFiniteNumber(parsed)) return null;
    if (/\ss$/.test(text) && !/\sms$/.test(text)) return parsed * 1000;
    return parsed;
  }

  function ratioValue(cell) {
    if (!cell || cell.querySelector('.imo-benchmark-missing')) return null;
    return numberValue(cell.getAttribute('data-ratio-to-fastest'));
  }

  function activeSortMetric() {
    return sortMetric && sortMetric.value === 'ratio' ? 'ratio' : 'time';
  }

  function librarySortValue(cell) {
    var time = durationValue(cell);
    var ratio = ratioValue(cell);
    return activeSortMetric() === 'ratio' ? [ratio, time] : [time, ratio];
  }

  function rowValue(row, key) {
    if (key === 'original') return numberValue(row.getAttribute('data-original-index')) || 0;
    if (key === 'scenario') return row.getAttribute('data-scenario') || '';
    if (key === 'fastest') return numberValue(row.getAttribute('data-fastest-ms'));
    if (key.indexOf('library:') === 0) {
      return librarySortValue(row._libraryCells[key.substring(8)]);
    }
    return '';
  }

  function missingValue(value) {
    if (Array.isArray(value)) {
      return value.every(function (entry) { return missingValue(entry); });
    }
    return value === null || typeof value === 'undefined' || value === '';
  }

  function compareNumberValues(leftValue, rightValue) {
    var leftValues = Array.isArray(leftValue) ? leftValue : [leftValue];
    var rightValues = Array.isArray(rightValue) ? rightValue : [rightValue];
    var length = Math.max(leftValues.length, rightValues.length);

    for (var index = 0; index < length; index++) {
      var leftEntry = leftValues[index];
      var rightEntry = rightValues[index];
      var leftMissing = missingValue(leftEntry);
      var rightMissing = missingValue(rightEntry);

      if (leftMissing && rightMissing) continue;
      if (leftMissing) return 1;
      if (rightMissing) return -1;

      var result = leftEntry - rightEntry;
      if (result !== 0) return result;
    }

    return 0;
  }

  function compareRows(left, right) {
    if (sortState.direction === 'none') {
      return rowValue(left, 'original') - rowValue(right, 'original');
    }

    var leftValue = rowValue(left, sortState.key);
    var rightValue = rowValue(right, sortState.key);
    var leftMissing = missingValue(leftValue);
    var rightMissing = missingValue(rightValue);

    if (leftMissing && rightMissing) return rowValue(left, 'original') - rowValue(right, 'original');
    if (leftMissing) return 1;
    if (rightMissing) return -1;

    var result = sortState.type === 'number'
      ? compareNumberValues(leftValue, rightValue)
      : String(leftValue).localeCompare(String(rightValue), undefined, { numeric: true, sensitivity: 'base' });

    if (result === 0) result = rowValue(left, 'original') - rowValue(right, 'original');
    return sortState.direction === 'desc' ? -result : result;
  }

  function filterValue(name) {
    var filter = root.querySelector('[data-benchmark-filter="' + name + '"]');
    return filter && filter.value ? filter.value : '';
  }

  function filterRow(row) {
    var search = filterValue('search').toLowerCase();
    var rowCount = filterValue('rowCount');
    var workload = filterValue('workload');
    var category = filterValue('category');
    var library = filterValue('library');

    if (search && row._filterText.indexOf(search) === -1) return false;
    if (rowCount && row.getAttribute('data-row-count') !== rowCount) return false;
    if (workload && row.getAttribute('data-workload') !== workload) return false;
    if (category && row.getAttribute('data-category') !== category) return false;
    if (library) {
      var cell = row._libraryCells[library];
      if (!cell || cell.querySelector('.imo-benchmark-missing')) return false;
    }

    return true;
  }

  function updateHeaders(activeButton) {
    Array.prototype.forEach.call(buttons, function (button) {
      var th = button.closest ? button.closest('th') : button.parentNode;
      if (th) th.setAttribute('aria-sort', 'none');
      button.removeAttribute('data-sort-active');
      button.setAttribute('data-sort-direction', 'none');
    });

    if (activeButton && sortState.direction !== 'none') {
      var activeTh = activeButton.closest ? activeButton.closest('th') : activeButton.parentNode;
      if (activeTh) activeTh.setAttribute('aria-sort', sortState.direction === 'asc' ? 'ascending' : 'descending');
      activeButton.setAttribute('data-sort-active', 'true');
      activeButton.setAttribute('data-sort-direction', sortState.direction);
    }
  }

  function apply() {
    var visible = 0;
    rows.forEach(function (row) {
      var keep = filterRow(row);
      row.hidden = !keep;
      if (keep) visible++;
    });
    rows.sort(compareRows).forEach(function (row) { tbody.appendChild(row); });
    if (count) count.textContent = 'Showing ' + visible + ' of ' + rows.length + ' rows';
  }

  function sortBy(key, direction) {
    var activeButton = null;
    Array.prototype.forEach.call(buttons, function (button) {
      if (button.getAttribute('data-benchmark-sort') === key) activeButton = button;
    });

    sortState = {
      key: direction === 'none' ? 'original' : key,
      direction: direction || 'asc',
      type: activeButton ? (activeButton.getAttribute('data-sort-type') || 'text') : 'text'
    };
    if (sortState.direction === 'none') sortState.key = 'original';

    updateHeaders(activeButton);
    apply();
  }

  function setFilter(name, value) {
    var filter = root.querySelector('[data-benchmark-filter="' + name + '"]');
    if (filter) filter.value = value || '';
    apply();
  }

  function resetMatrix() {
    Array.prototype.forEach.call(filters, function (filter) { filter.value = ''; });
    if (sortMetric) sortMetric.value = 'time';
    sortState = { key: 'original', direction: 'none', type: 'number' };
    updateHeaders(null);
    apply();
  }

  function on(element, eventName, handler) {
    if (!element) return;
    if (element.addEventListener) {
      element.addEventListener(eventName, handler);
    } else {
      element['on' + eventName] = handler;
    }
  }

  Array.prototype.forEach.call(filters, function (filter) {
    on(filter, 'input', apply);
    on(filter, 'change', apply);
  });

  on(sortMetric, 'change', apply);

  Array.prototype.forEach.call(buttons, function (button) {
    on(button, 'click', function () {
      var key = button.getAttribute('data-benchmark-sort');
      if (sortState.key !== key) {
        sortBy(key, 'asc');
      } else if (sortState.direction === 'asc') {
        sortBy(key, 'desc');
      } else if (sortState.direction === 'desc') {
        sortBy('original', 'none');
      } else {
        sortBy(key, 'asc');
      }
    });
  });

  on(reset, 'click', resetMatrix);

  window.OfficeImoBenchmarkMatrix = {
    apply: apply,
    reset: resetMatrix,
    setFilter: setFilter,
    setSortMetric: function (value) {
      if (sortMetric) sortMetric.value = value === 'ratio' ? 'ratio' : 'time';
      apply();
    },
    sortBy: sortBy
  };

  apply();
}());

(function () {
  var root = document.querySelector('[data-library-comparison-benchmarks]');
  if (!root || !window.fetch) return;

  var state = root.querySelector('[data-library-comparison-state]');
  var table = root.querySelector('[data-library-comparison-table]');
  var rows = root.querySelector('[data-library-comparison-rows]');
  var meta = root.querySelector('[data-library-comparison-meta]');
  var workloadButtons = root.querySelectorAll('[data-library-comparison-workload]');
  var platformButtons = root.querySelectorAll('[data-library-comparison-platform]');
  var modeButtons = root.querySelectorAll('[data-library-comparison-mode]');
  var selectedComparison = queryValue(
    'benchmark-workload',
    root.getAttribute('data-comparison-id'));
  var catalog;
  var activeRequestId = 0;

  function queryValue(name, fallback) {
    try {
      return new URL(window.location.href).searchParams.get(name) || fallback;
    } catch (_) {
      return fallback;
    }
  }

  var selectedPlatform = queryValue('benchmark-os', 'windows').toLowerCase();
  var selectedMode = queryValue('benchmark-mode', 'full').toLowerCase();

  function setQuery() {
    try {
      var url = new URL(window.location.href);
      url.searchParams.set('benchmark-workload', selectedComparison);
      url.searchParams.set('benchmark-os', selectedPlatform);
      url.searchParams.set('benchmark-mode', selectedMode);
      window.history.replaceState(null, '', url.toString());
    } catch (_) {
      // The selector still works in hosts without the URL API.
    }
  }

  function activateButtons() {
    Array.prototype.forEach.call(workloadButtons, function (button) {
      var active = button.getAttribute('data-library-comparison-workload') === selectedComparison;
      button.classList.toggle('active', active);
      button.setAttribute('aria-pressed', active ? 'true' : 'false');
    });
    Array.prototype.forEach.call(platformButtons, function (button) {
      var active = button.getAttribute('data-library-comparison-platform') === selectedPlatform;
      button.classList.toggle('active', active);
      button.setAttribute('aria-pressed', active ? 'true' : 'false');
      var availability = catalog && catalog.availability
        ? catalog.availability.find(function (item) {
          return item.comparisonId === selectedComparison &&
            item.runMode === selectedMode &&
            item.platform === button.getAttribute('data-library-comparison-platform');
        })
        : null;
      button.classList.toggle('missing', !!availability && !availability.available);
    });
    Array.prototype.forEach.call(modeButtons, function (button) {
      var active = button.getAttribute('data-library-comparison-mode') === selectedMode;
      button.classList.toggle('active', active);
      button.setAttribute('aria-pressed', active ? 'true' : 'false');
    });
  }

  function formatDuration(value) {
    if (typeof value !== 'number') return '—';
    if (value >= 1000) return (value / 1000).toFixed(2) + ' s';
    return value.toFixed(value >= 100 ? 1 : 2) + ' ms';
  }

  function formatBytes(value) {
    if (typeof value !== 'number') return '—';
    if (value >= 1024 * 1024) return (value / (1024 * 1024)).toFixed(2) + ' MB';
    if (value >= 1024) return (value / 1024).toFixed(1) + ' KB';
    return Math.round(value) + ' B';
  }

  function platformLabel(value) {
    var labels = {
      windows: 'Windows',
      linux: 'Linux',
      macos: 'macOS'
    };
    return labels[value] || value;
  }

  function metric(row, name) {
    if (!row || !row.metrics) return null;
    if (typeof row.metrics[name] === 'number') return row.metrics[name];
    var key = Object.keys(row.metrics).find(function (candidate) {
      return candidate.toLowerCase() === name.toLowerCase();
    });
    return key ? row.metrics[key] : null;
  }

  function workloadName() {
    var names = {
      'markpflug-65k-csv-decoded-net10.0': 'CSV · decoded strings',
      'markpflug-65k-xlsx-typed-net10.0': 'XLSX · typed values',
      'markpflug-65k-xlsb-typed-net10.0': 'XLSB · typed values'
    };
    return names[selectedComparison] || selectedComparison || 'Library comparison';
  }

  function compatibilityValue(entry, name) {
    var compatibility = entry && entry.compatibility;
    if (!compatibility) return null;
    var key = Object.keys(compatibility).find(function (candidate) {
      return candidate.toLowerCase() === name.toLowerCase();
    });
    return key ? compatibility[key] : null;
  }

  function renderResult(entry, result, requestId) {
    if (requestId !== activeRequestId) return;
    var summaries = result.summary || [];
    var groups = {};
    summaries.forEach(function (row) {
      var workload = workloadName();
      if (!groups[workload]) groups[workload] = [];
      groups[workload].push(row);
    });

    rows.innerHTML = '';
    Object.keys(groups).forEach(function (workload) {
      var group = groups[workload];
      var fastest = Math.min.apply(null, group.map(function (row) {
        return typeof row.medianMs === 'number' ? row.medianMs : Number.POSITIVE_INFINITY;
      }));
      group.sort(function (left, right) {
        var leftMedian = typeof left.medianMs === 'number' ? left.medianMs : Number.POSITIVE_INFINITY;
        var rightMedian = typeof right.medianMs === 'number' ? right.medianMs : Number.POSITIVE_INFINITY;
        return leftMedian - rightMedian || String(left.scenario).localeCompare(String(right.scenario));
      });
      group.forEach(function (row, index) {
        var median = row.medianMs;
        var ratio = typeof median === 'number' && isFinite(fastest) && fastest > 0 ? median / fastest : null;
        var tr = document.createElement('tr');
        tr.innerHTML =
          '<td>' + (index === 0 ? workload : '') + '</td>' +
          '<td><strong>' + String(row.scenario || 'Unknown') + '</strong></td>' +
          '<td>' + formatDuration(median) + '</td>' +
          '<td>' + formatBytes(metric(row, 'Allocated')) + '</td>' +
          '<td>' + (ratio === null ? '—' : ratio.toFixed(2) + 'x') + '</td>';
        if (ratio !== null && ratio <= 1.0005) tr.classList.add('imo-library-comparison-fastest');
        rows.appendChild(tr);
      });
    });

    meta.innerHTML = '';
    var sourceCommit = compatibilityValue(entry, 'gitSha');
    [
      workloadName(),
      platformLabel(selectedPlatform),
      selectedMode,
      entry.environment && entry.environment.processorName,
      entry.environment && entry.environment.runtimeVersion,
      entry.generatedUtc && new Date(entry.generatedUtc).toLocaleString(),
      sourceCommit && 'source ' + sourceCommit.substring(0, 12)
    ].filter(Boolean).forEach(function (value) {
      var span = document.createElement('span');
      span.textContent = value;
      meta.appendChild(span);
    });

    state.hidden = true;
    table.hidden = false;
  }

  function renderSelection() {
    var requestId = ++activeRequestId;
    activateButtons();
    setQuery();
    table.hidden = true;
    meta.innerHTML = '';
    var entry = (catalog.entries || []).find(function (candidate) {
      return candidate.comparisonId === selectedComparison &&
        candidate.platform === selectedPlatform &&
        candidate.runMode === selectedMode &&
        (selectedMode !== 'full' || candidate.publish === true);
    });
    if (!entry) {
      state.hidden = false;
      state.className = 'imo-library-comparison-state missing';
      state.textContent = 'No ' + selectedMode + ' evidence has been published for ' + platformLabel(selectedPlatform) + ' yet.';
      return;
    }
    if (entry.comparable === false) {
      state.hidden = false;
      state.className = 'imo-library-comparison-state incompatible';
      state.textContent = 'This lane is not directly comparable: ' + (entry.compatibilityIssues || []).join(' ');
      return;
    }

    state.hidden = false;
    state.className = 'imo-library-comparison-state';
    state.textContent = 'Loading ' + platformLabel(selectedPlatform) + ' ' + selectedMode + ' results…';
    fetch(entry.resultPath, { credentials: 'same-origin' })
      .then(function (response) {
        if (!response.ok) throw new Error('HTTP ' + response.status);
        return response.json();
      })
      .then(function (result) { renderResult(entry, result, requestId); })
      .catch(function (error) {
        if (requestId !== activeRequestId) return;
        state.hidden = false;
        state.className = 'imo-library-comparison-state incompatible';
        state.textContent = 'Benchmark evidence could not be loaded: ' + error.message;
      });
  }

  Array.prototype.forEach.call(workloadButtons, function (button) {
    button.addEventListener('click', function () {
      selectedComparison = button.getAttribute('data-library-comparison-workload');
      renderSelection();
    });
  });
  Array.prototype.forEach.call(platformButtons, function (button) {
    button.addEventListener('click', function () {
      selectedPlatform = button.getAttribute('data-library-comparison-platform');
      renderSelection();
    });
  });
  Array.prototype.forEach.call(modeButtons, function (button) {
    button.addEventListener('click', function () {
      selectedMode = button.getAttribute('data-library-comparison-mode');
      renderSelection();
    });
  });

  fetch(root.getAttribute('data-index-url'), { credentials: 'same-origin' })
    .then(function (response) {
      if (!response.ok) throw new Error('HTTP ' + response.status);
      return response.json();
    })
    .then(function (value) {
      catalog = value;
      renderSelection();
    })
    .catch(function (error) {
      state.className = 'imo-library-comparison-state incompatible';
      state.textContent = 'Benchmark catalog could not be loaded: ' + error.message;
    });
}());
