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
      type.textContent = labels[item.collection] || 'Page';
      var link = document.createElement('a');
      link.href = url.pathname + url.search + url.hash;
      link.textContent = item.title || item.url;
      card.append(type, link);
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
    meta.textContent = 'Searching…';
    try {
      var response = await api.search({ query: query, limit: panel.searchLimit });
      if (version !== panel.searchVersion) return;
      response.query = query;
      render(panel, response);
    } catch (_) {
      if (version !== panel.searchVersion) return;
      panel.querySelector('[data-search-page-results]').replaceChildren();
      meta.textContent = 'Search is unavailable. Please try again.';
    }
  }

  panels.forEach(function (panel) {
    panel.querySelector('[data-search-page-input]').addEventListener('input', function () { search(panel); });
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
  document.querySelectorAll('[data-site-search-open]').forEach(function (link) {
    link.addEventListener('click', function (event) {
      if (event.button !== 0 || event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
      event.preventDefault();
      openSearch(link);
    });
  });
  dialog.querySelector('[data-site-search-close]').addEventListener('click', function () { dialog.close(); });
  dialog.addEventListener('close', function () { if (opener && opener.isConnected) opener.focus(); });
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
    if (!pagePanel) openSearch();
    render(panel, response);
  };
  function seedPage() {
    if (!pagePanel) return;
    pagePanel.querySelector('[data-search-page-input]').value = new URLSearchParams(location.search).get('q') || '';
    search(pagePanel);
  }
  if (document.readyState !== 'complete') document.addEventListener('DOMContentLoaded', seedPage, { once: true });
  else seedPage();
})();
