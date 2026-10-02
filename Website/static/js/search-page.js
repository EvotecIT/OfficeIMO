(function () {
  'use strict';
  var dialog = document.querySelector('[data-site-search-dialog]');
  if (!dialog) return;
  var panels = Array.from(document.querySelectorAll('[data-site-search-panel]'));
  var pagePanel = document.querySelector('[data-site-search-page]');
  var dialogPanel = dialog.querySelector('[data-site-search-panel]');
  var opener = null;
  var api = window.PowerForgeWebMcpSearch = window.PowerForgeWebMcpSearch || {};
  var labels = { api: 'API reference', powershell: 'PowerShell', docs: 'Guide', products: 'Library', 'pdf-workflows': 'PDF workflow', conversions: 'Conversion', blog: 'Article', pages: 'Page', solutions: 'Solution', comparisons: 'Comparison' };

  function openSearch(trigger) {
    if (!dialog.open) {
      opener = trigger || document.activeElement;
      dialog.showModal();
    }
    var input = dialogPanel.querySelector('[data-search-page-input]');
    input.focus();
    input.select();
  }

  function updatePageQuery(panel, query) {
    if (panel !== pagePanel) return;
    var url = new URL(window.location.href);
    if (query) url.searchParams.set('q', query);
    else url.searchParams.delete('q');
    history.replaceState(history.state, '', url);
  }

  function render(panel, response) {
    var results = panel.querySelector('[data-search-page-results]');
    var meta = panel.querySelector('[data-site-search-meta]');
    var items = response.results || [];
    results.replaceChildren();
    items.forEach(function (item) {
      var url;
      try { url = new URL(item.url, window.location.origin); } catch (_) { return; }
      if (url.origin !== window.location.origin || !/^https?:$/.test(url.protocol)) return;
      var card = document.createElement('article');
      card.className = 'imo-search__result';
      var type = document.createElement('span');
      type.className = 'imo-search__type';
      type.textContent = [labels[item.collection] || 'Page', item.kind, item.project].filter(Boolean).join(' · ');
      var link = document.createElement('a');
      link.href = url.pathname + url.search + url.hash;
      link.textContent = item.title || item.url;
      card.append(type, link);
      if (item.meta && item.meta.signature) {
        var signature = document.createElement('code');
        signature.className = 'imo-search__signature';
        signature.textContent = item.meta.signature;
        card.append(signature);
      }
      if (item.meta && item.meta.namespace) {
        var context = document.createElement('p');
        context.className = 'imo-search__context';
        context.textContent = 'using ' + item.meta.namespace + ';' +
          (item.meta.receiverType ? ' · Extension for ' + item.meta.receiverType : '');
        card.append(context);
      }
      var description = item.description || item.snippet;
      if (description) {
        var text = document.createElement('p');
        text.textContent = description;
        card.append(text);
      }
      results.append(card);
    });
    var count = response.totalMatches;
    meta.textContent = count === 0 ? 'No results. Try a shorter topic or an API or command name.'
      : 'Showing ' + items.length + ' of ' + count + ' results for “' + response.query + '”.';
    panel.querySelector('[data-site-search-more]').hidden = items.length >= count || items.length >= 100;
    if (items.length >= 100 && count > 100) meta.textContent += ' Refine your query to narrow the results.';
    updatePageQuery(panel, response.query);
    dialog.querySelector('[data-site-search-page-link]').href = '/search/?q=' + encodeURIComponent(response.query);
  }

  async function search(panel, limit) {
    if (panel.searchController) panel.searchController.abort();
    panel.searchController = new AbortController();
    var query = panel.querySelector('[data-search-page-input]').value.trim();
    var version = panel.searchVersion = (panel.searchVersion || 0) + 1;
    var meta = panel.querySelector('[data-site-search-meta]');
    panel.searchLimit = limit || 20;
    panel.querySelector('[data-site-search-more]').hidden = true;
    if (!query) {
      panel.querySelector('[data-search-page-results]').replaceChildren();
      meta.textContent = 'Enter a topic, API type, or PowerShell command.';
      if (panel === dialogPanel) dialog.querySelector('[data-site-search-page-link]').href = '/search/';
      updatePageQuery(panel, '');
      return;
    }
    if (query.length < 2 && !panel.querySelector('[data-search-package]').value) {
      panel.querySelector('[data-search-page-results]').replaceChildren();
      meta.textContent = 'Type at least two characters or choose a package.';
      return;
    }
    meta.textContent = 'Searching…';
    try {
      var response = await api.search({ query: query, limit: panel.searchLimit,
        project: panel.querySelector('[data-search-package]').value,
        kind: panel.querySelector('[data-search-kind]').value,
        signal: panel.searchController.signal });
      if (version !== panel.searchVersion) return;
      response.query = query;
      render(panel, response);
    } catch (error) {
      if (version !== panel.searchVersion) return;
      panel.querySelector('[data-search-page-results]').replaceChildren();
      if (error.name === 'AbortError') return;
      meta.textContent = error.code === 'SEARCH_QUERY_TOO_BROAD' ? error.message : 'Search is unavailable. Please try again.';
    }
  }

  panels.forEach(function (panel) {
    var filters = document.createElement('div');
    filters.className = 'imo-search__filters';
    [['package', 'All packages', []], ['kind', 'All types and members', ['class', 'interface', 'struct', 'enum', 'method', 'extension', 'constructor', 'property', 'field', 'event']]].forEach(function (spec) {
      var label = document.createElement('label');
      label.textContent = spec[0] === 'package' ? 'Package' : 'API kind';
      var select = document.createElement('select');
      select.setAttribute('data-search-' + spec[0], '');
      var all = document.createElement('option');
      all.value = ''; all.textContent = spec[1]; select.append(all);
      spec[2].forEach(function (value) { var option = document.createElement('option'); option.value = value; option.textContent = value; select.append(option); });
      label.append(select); filters.append(label);
      select.addEventListener('change', function () { search(panel); });
    });
    panel.querySelector('[data-search-page-results]').before(filters);
    panel.querySelector('[data-search-page-input]').placeholder = 'Search methods, types, packages, or tasks…';
    panel.querySelector('[data-search-page-input]').addEventListener('focus', populatePackages, { once: true });
    panel.querySelector('[data-search-page-input]').addEventListener('input', function () {
      if (panel.searchController) panel.searchController.abort();
      panel.searchVersion = (panel.searchVersion || 0) + 1;
      clearTimeout(panel.searchTimer);
      panel.searchTimer = setTimeout(function () { search(panel); }, 150);
    });
    panel.addEventListener('keydown', function (event) {
      if (event.altKey || event.ctrlKey || event.metaKey || event.isComposing || !['ArrowDown', 'ArrowUp'].includes(event.key)) return;
      var links = Array.from(panel.querySelectorAll('[data-search-page-results] a'));
      var index = links.indexOf(document.activeElement);
      if (!links.length || index < 0 && !event.target.matches('[data-search-page-input]')) return;
      event.preventDefault();
      if (event.key === 'ArrowUp' && index <= 0) panel.querySelector('[data-search-page-input]').focus();
      else links[Math.min(links.length - 1, Math.max(0, index + (event.key === 'ArrowDown' ? 1 : -1)))].focus();
    });
    panel.querySelector('[data-site-search-more]').addEventListener('click', function () { search(panel, (panel.searchLimit || 20) + 20); });
  });
  var packagePromise;
  function populatePackages() {
    if (!packagePromise && typeof api.facets === 'function') packagePromise = api.facets().then(function (facets) {
      panels.forEach(function (panel) {
        var select = panel.querySelector('[data-search-package]');
        facets.projects.forEach(function (project) { var option = document.createElement('option'); option.value = project; option.textContent = project; select.append(option); });
      });
    }).catch(function () { packagePromise = null; });
  }
  document.querySelectorAll('[data-site-search-open]').forEach(function (link) {
    link.addEventListener('click', function (event) {
      if (event.button !== 0 || event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
      event.preventDefault();
      openSearch(link);
    });
  });
  dialog.querySelector('[data-site-search-close]').addEventListener('click', function () { dialog.close(); });
  dialog.addEventListener('close', function () {
    if (dialogPanel.searchController) dialogPanel.searchController.abort();
    if (opener && opener.isConnected) opener.focus();
  });
  window.addEventListener('message', function (event) {
    var frame = document.querySelector('iframe[data-workspace-src]');
    if (event.origin === location.origin && frame && event.source === frame.contentWindow && event.data && event.data.type === 'officeimo:open-search') openSearch(frame);
  });
  document.addEventListener('keydown', function (event) {
    if (event.defaultPrevented || event.isComposing) return;
    var typing = event.target.closest && event.target.closest('input, textarea, select, [contenteditable="true"], [role="textbox"]');
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k' ||
        event.key === '/' && !typing && !event.ctrlKey && !event.metaKey && !event.altKey) {
      event.preventDefault();
      openSearch();
    }
  });

  // WebMCP searches use the same visible UI without leaving the current page.
  api.renderVisibleResults = function (response) {
    var panel = pagePanel || dialogPanel;
    panel.searchVersion = (panel.searchVersion || 0) + 1;
    panel.querySelector('[data-search-page-input]').value = response.query;
    panel.querySelector('[data-search-package]').value = '';
    panel.querySelector('[data-search-kind]').value = '';
    if (!pagePanel) openSearch();
    render(panel, response);
  };
  function seedPage() {
    if (!pagePanel) return;
    populatePackages();
    pagePanel.querySelector('[data-search-page-input]').value = new URLSearchParams(location.search).get('q') || '';
    search(pagePanel);
  }
  if (document.readyState !== 'complete') document.addEventListener('DOMContentLoaded', seedPage, { once: true });
  else seedPage();
})();
