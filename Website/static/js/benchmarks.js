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
