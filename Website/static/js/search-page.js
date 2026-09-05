(async function () {
  const input = document.getElementById('imo-search-query');
  const meta = document.getElementById('imo-search-meta');
  const results = document.getElementById('imo-search-results');

  if (!input || !meta || !results) {
    return;
  }

  let entries = [];
  let indexLoadError = null;
  let resolveIndexReady;
  const indexReady = new Promise(function (resolve) {
    resolveIndexReady = resolve;
  });
  const webMcpSearch = window.PowerForgeWebMcpSearch || {};
  window.PowerForgeWebMcpSearch = webMcpSearch;

  const params = new URLSearchParams(window.location.search);
  const seededQuery = (params.get('q') || '').trim();
  if (seededQuery) {
    input.value = seededQuery;
  }

  function escapeHtml(value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function searchText(entry) {
    const tags = Array.isArray(entry.tags) ? entry.tags.join(' ') : '';
    return [
      entry.title,
      entry.description,
      entry.snippet,
      entry.searchText,
      entry.collection,
      entry.kind,
      tags
    ].join(' ').toLowerCase();
  }

  function render(entries, query) {
    if (!entries.length) {
      results.innerHTML = '<p>No results found.</p>';
      meta.textContent = query ? '0 results for "' + query + '"' : 'No search entries found.';
      return;
    }

    meta.textContent = query
      ? entries.length + ' results for "' + query + '"'
      : entries.length + ' pages indexed';

    results.innerHTML = entries.map(function (item) {
      const title = item.title || item.url || '/';
      const description = item.description
        ? '<div class="imo-search__desc">' + escapeHtml(item.description) + '</div>'
        : '';
      const snippet = item.snippet
        ? '<div class="imo-search__snippet">' + escapeHtml(item.snippet) + '</div>'
        : '';
      const tags = Array.isArray(item.tags) && item.tags.length
        ? '<div class="imo-search__tags">' + item.tags.map(function (tag) {
            return '<span class="imo-search__tag">' + escapeHtml(tag) + '</span>';
          }).join('') + '</div>'
        : '';

      return '<article class="imo-search__result"><a href="' + escapeHtml(item.url || '/') + '">' + escapeHtml(title) + '</a>' + description + snippet + tags + '</article>';
    }).join('');
  }

  function awaitIndex(signal) {
    if (!signal) {
      return indexReady;
    }
    if (signal.aborted) {
      return Promise.reject(new DOMException('The OfficeIMO search was cancelled.', 'AbortError'));
    }

    return new Promise(function (resolve, reject) {
      function cleanup() {
        signal.removeEventListener('abort', abort);
      }
      function abort() {
        cleanup();
        reject(new DOMException('The OfficeIMO search was cancelled.', 'AbortError'));
      }
      signal.addEventListener('abort', abort, { once: true });
      indexReady.then(function () {
        cleanup();
        resolve();
      });
    });
  }

  webMcpSearch.adapter = {
    search: async function (request) {
      await awaitIndex(request.signal);
      if (indexLoadError) {
        throw indexLoadError;
      }
      if (typeof webMcpSearch.searchEntries !== 'function') {
        throw new Error('The PowerForge WebMCP search runtime is unavailable.');
      }

      const result = webMcpSearch.searchEntries(entries, request.query, request.limit);
      input.value = request.query;
      render(result.results, request.query);
      return result;
    }
  };

  try {
    let indexPath = '/search/index.json';
    const manifestResponse = await fetch('/search/manifest.json', { cache: 'no-cache' });
    if (manifestResponse.ok) {
      const manifest = await manifestResponse.json();
      if (manifest && typeof manifest.searchIndexPath === 'string' && manifest.searchIndexPath.trim()) {
        indexPath = manifest.searchIndexPath;
      }
    }

    const indexResponse = await fetch(indexPath, { cache: 'no-cache' });
    if (!indexResponse.ok) {
      throw new Error('Failed to load search index: ' + indexResponse.status);
    }

    entries = await indexResponse.json();
  } catch (error) {
    indexLoadError = error instanceof Error ? error : new Error(String(error));
    meta.textContent = 'Search index unavailable.';
    results.innerHTML = '<p>' + escapeHtml(error && error.message ? error.message : error) + '</p>';
    resolveIndexReady();
    return;
  }
  resolveIndexReady();

  function runSearch() {
    const query = input.value.trim().toLowerCase();
    if (!query) {
      render(entries, '');
      return;
    }

    const matches = entries
      .map(function (item) {
        return { item: item, haystack: searchText(item), weight: Number(item.weight || 1) };
      })
      .filter(function (row) {
        return row.haystack.indexOf(query) >= 0;
      })
      .sort(function (left, right) {
        if (right.weight !== left.weight) {
          return right.weight - left.weight;
        }

        return String(left.item.title || '').localeCompare(String(right.item.title || ''));
      })
      .map(function (row) {
        return row.item;
      });

    render(matches, query);
  }

  input.addEventListener('input', runSearch);
  runSearch();
})();
