using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

public static partial class MarkdownRenderer {
    private static string BuildIncrementalUpdateScript(MarkdownRendererOptions options) {
        bool mermaid = options.Mermaid?.Enabled == true;
        bool chart = options.Chart?.Enabled == true;
        var mathOptions = options.Math;
        bool codeCopy = options.EnableCodeCopyButtons;
        bool tableCopy = options.EnableTableCopyButtons;

        // Notes:
        // - We keep <base> in <head> so relative links/images resolve.
        // - We preserve already-rendered Mermaid SVGs by comparing data-mermaid-hash attributes.
        // - We re-run Prism highlighting after updates (if Prism is present).
        var sb = new StringBuilder(8 * 1024);
        sb.Append("""
async function updateContent(newBodyHtml) {
  const root = document.getElementById('omdRoot') || document.body;
  // Extract <base href="..."> from payload and place it in <head>.
  try {
    const baseMatch = newBodyHtml.match(/<base\s+href="([^"]*)"[^>]*>/i);
    if (baseMatch) {
      let baseEl = document.getElementById('omdBase');
      if (!baseEl) {
        baseEl = document.createElement('base');
        baseEl.id = 'omdBase';
        document.head.appendChild(baseEl);
      }
      baseEl.href = baseMatch[1];
      newBodyHtml = newBodyHtml.replace(baseMatch[0], '');
    } else {
      const baseEl = document.getElementById('omdBase');
      if (baseEl) baseEl.href = 'about:blank';
    }
  } catch(e) { /* best-effort */ }
""");

        if (chart) {
            sb.Append("""
  // Destroy existing Chart.js instances before replacing DOM to avoid leaks.
  try {
    if (window.Chart && typeof Chart.getChart === 'function') {
      root.querySelectorAll('canvas.omd-chart').forEach(c => {
        try { const inst = Chart.getChart(c); if (inst) inst.destroy(); } catch(e) { /* ignore */ }
      });
    }
  } catch(e) { /* ignore */ }
""");
        }

        if (mermaid) {
            sb.Append("""
  // Cache existing Mermaid SVGs keyed by data-mermaid-hash.
  const existingSvgs = new Map();
  root.querySelectorAll('[data-mermaid-hash]').forEach(el => {
    const hash = el.getAttribute('data-mermaid-hash');
    const svg = el.querySelector('svg') || (el.nextElementSibling && el.nextElementSibling.tagName === 'svg' ? el.nextElementSibling : null);
    if (hash && svg) existingSvgs.set(hash, svg.cloneNode(true));
  });
""");
        }

        AppendCustomUpdateScripts(sb, options, beforeReplace: true);

        sb.Append("""
  // Replace rendered contents.
  root.innerHTML = newBodyHtml;
""");

        if (codeCopy || tableCopy) {
            sb.Append("""

  // Copy helpers (optional)
  function omdCopyText(text) {
    const s = String(text ?? '');
    try {
      const wv = window.chrome && window.chrome.webview;
      if (wv && typeof wv.postMessage === 'function') {
        // Host can optionally handle this message and place text on clipboard.
        wv.postMessage({ type: 'omd.copy', text: s });
      }
    } catch(_) { /* ignore */ }

    try {
      if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
        return navigator.clipboard.writeText(s);
      }
    } catch(_) { /* ignore */ }

    try {
      const ta = document.createElement('textarea');
      ta.value = s;
      ta.setAttribute('readonly', 'readonly');
      ta.style.position = 'fixed';
      ta.style.left = '-9999px';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy'); } catch(_) { /* ignore */ }
      document.body.removeChild(ta);
    } catch(_) { /* ignore */ }

    return Promise.resolve();
  }

  function omdFlash(btn, label) {
    try {
      const orig = btn.textContent;
      btn.textContent = label;
      btn.setAttribute('data-omd-flash', '1');
      setTimeout(() => { try { btn.textContent = orig; btn.removeAttribute('data-omd-flash'); } catch(_){} }, 900);
    } catch(_) {}
  }
""");
        }

        if (codeCopy) {
            sb.Append("""

  function omdSetupCodeCopyButtons(rootEl) {
    try {
      rootEl.querySelectorAll('pre > code').forEach(code => {
        const pre = code.parentElement;
        if (!pre || pre.getAttribute('data-omd-code-inited') === '1') return;
        pre.setAttribute('data-omd-code-inited', '1');
        pre.classList && pre.classList.add('omd-has-actions');

        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'omd-copy-btn omd-copy-code';
        btn.textContent = 'Copy';
        btn.addEventListener('click', ev => {
          try { ev.preventDefault(); ev.stopPropagation(); } catch(_) {}
          omdCopyText(code.textContent || '');
          omdFlash(btn, 'Copied');
        });

        // Put the button as the first child so it stays visible even if Prism modifies <code>.
        pre.insertBefore(btn, pre.firstChild);
      });
    } catch(_) { /* ignore */ }
  }
""");
        }

        if (tableCopy) {
            sb.Append("""

  function omdCellText(cell) {
    const t = (cell && (cell.innerText || cell.textContent)) ? String(cell.innerText || cell.textContent) : '';
    return t.replace(/\r?\n/g, ' ').trim();
  }

  function omdTableToTsv(table) {
    const rows = [];
    const trs = table.querySelectorAll('tr');
    trs.forEach(tr => {
      const cells = tr.querySelectorAll('th,td');
      if (!cells || cells.length === 0) return;
      const vals = [];
      cells.forEach(c => vals.push(omdCellText(c)));
      rows.push(vals.join('\\t'));
    });
    return rows.join('\\n');
  }

  function omdCsvEscape(value) {
    const s = String(value ?? '');
    if (s.indexOf('\"') >= 0 || s.indexOf(',') >= 0 || s.indexOf('\\n') >= 0 || s.indexOf('\\r') >= 0) {
      return '\"' + s.replace(/\"/g, '\"\"') + '\"';
    }
    return s;
  }

  function omdTableToCsv(table) {
    const rows = [];
    const trs = table.querySelectorAll('tr');
    trs.forEach(tr => {
      const cells = tr.querySelectorAll('th,td');
      if (!cells || cells.length === 0) return;
      const vals = [];
      cells.forEach(c => vals.push(omdCsvEscape(omdCellText(c))));
      rows.push(vals.join(','));
    });
    return rows.join('\\n');
  }

  function omdSetupTableCopyButtons(rootEl) {
    try {
      rootEl.querySelectorAll('table').forEach(table => {
        if (table.getAttribute('data-omd-table-inited') === '1') return;
        table.setAttribute('data-omd-table-inited', '1');

        const actions = document.createElement('div');
        actions.className = 'omd-table-actions';

        const b1 = document.createElement('button');
        b1.type = 'button';
        b1.className = 'omd-copy-btn omd-copy-tsv';
        b1.textContent = 'Copy TSV';
        b1.addEventListener('click', ev => {
          try { ev.preventDefault(); ev.stopPropagation(); } catch(_) {}
          omdCopyText(omdTableToTsv(table));
          omdFlash(b1, 'Copied');
        });

        const b2 = document.createElement('button');
        b2.type = 'button';
        b2.className = 'omd-copy-btn omd-copy-csv';
        b2.textContent = 'Copy CSV';
        b2.addEventListener('click', ev => {
          try { ev.preventDefault(); ev.stopPropagation(); } catch(_) {}
          omdCopyText(omdTableToCsv(table));
          omdFlash(b2, 'Copied');
        });

        actions.appendChild(b1);
        actions.appendChild(b2);

        table.parentElement && table.parentElement.insertBefore(actions, table);
      });
    } catch(_) { /* ignore */ }
  }
""");
        }

        if (mermaid) {
            sb.Append("""
  // Restore cached Mermaid SVGs for unchanged diagrams.
  root.querySelectorAll('[data-mermaid-hash]').forEach(el => {
    const hash = el.getAttribute('data-mermaid-hash');
    if (existingSvgs.has(hash)) {
      const cachedSvg = existingSvgs.get(hash);
      el.innerHTML = '';
      el.appendChild(cachedSvg);
      el.setAttribute('data-mermaid-rendered', 'true');
    }
  });

  // Render only new/changed Mermaid blocks.
  const unrendered = root.querySelectorAll('[data-mermaid-hash]:not([data-mermaid-rendered])');
  if (unrendered.length > 0 && window.mermaid) {
    try { await window.mermaid.run({ nodes: unrendered }); } catch(e) { console.warn('Mermaid render error:', e); }
  }
  // Render plain Mermaid blocks (language-mermaid) when hashes are not present.
  const plainMermaid = root.querySelectorAll('pre > code.language-mermaid:not([data-mermaid-rendered]), .mermaid:not([data-mermaid-rendered]):not(svg)');
  if (plainMermaid.length > 0 && window.mermaid) {
    try { await window.mermaid.run({ nodes: plainMermaid }); } catch(e) { console.warn('Mermaid render error:', e); }
  }
""");
        }

        if (chart) {
            sb.Append("""
  // Chart.js rendering (optional).
  try {
    function b64ToUtf8(b64) {
      try {
        const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
        if (window.TextDecoder) return new TextDecoder('utf-8').decode(bytes);
        // Fallback for older engines.
        return decodeURIComponent(escape(String.fromCharCode.apply(null, Array.from(bytes))));
      } catch(e) { return ''; }
    }
    if (window.Chart) {
      root.querySelectorAll('canvas.omd-chart:not([data-chart-rendered])').forEach(c => {
        const b64 = c.getAttribute('data-chart-config-b64');
        if (!b64) return;
        const jsonText = b64ToUtf8(b64);
        if (!jsonText) return;
        let cfg = null;
        try { cfg = JSON.parse(jsonText); } catch(e) { console.warn('Chart config JSON parse error:', e); return; }
        try {
          const ctx = c.getContext && c.getContext('2d');
          if (!ctx) return;
          new Chart(ctx, cfg);
          c.setAttribute('data-omd-visual-rendered', 'true');
          c.setAttribute('data-chart-rendered', 'true');
        } catch(e) { console.warn('Chart render error:', e); }
      });
    }
  } catch(e) { /* ignore */ }
""");
        }

        AppendCustomUpdateScripts(sb, options, beforeReplace: false);

        if (codeCopy || tableCopy) {
            sb.Append("""

  // Add optional copy buttons after updates (best-effort).
  try {
""");
            if (codeCopy) sb.Append("    omdSetupCodeCopyButtons(root);\n");
            if (tableCopy) sb.Append("    omdSetupTableCopyButtons(root);\n");
            sb.Append("""
  } catch(_) { /* ignore */ }
""");
        }

        sb.Append("""
  // Prism highlighting (optional).
  try {
    if (window.Prism) {
      if (typeof Prism.highlightAllUnder === 'function') Prism.highlightAllUnder(root);
      else if (typeof Prism.highlightAll === 'function') Prism.highlightAll();
    }
  } catch(e) { /* ignore */ }
""");

        if (mathOptions != null && mathOptions.Enabled) {
            sb.Append("""

  // KaTeX auto-render (optional).
  try {
    if (window.renderMathInElement) {
      const delimiters = [];
""");
            if (mathOptions.EnableDollarDisplay) sb.Append("      delimiters.push({ left: \"$$\", right: \"$$\", display: true });\n");
            if (mathOptions.EnableDollarInline) sb.Append("      delimiters.push({ left: \"$\", right: \"$\", display: false });\n");
            if (mathOptions.EnableBracketDisplay) sb.Append("      delimiters.push({ left: \"\\\\[\", right: \"\\\\]\", display: true });\n");
            if (mathOptions.EnableParenInline) sb.Append("      delimiters.push({ left: \"\\\\(\", right: \"\\\\)\", display: false });\n");
            sb.Append("""
      if (delimiters.length > 0) {
        window.renderMathInElement(root, {
          delimiters: delimiters,
          throwOnError: false,
          strict: 'ignore',
          ignoredTags: ['script', 'noscript', 'style', 'textarea', 'pre', 'code']
        });
      }
    }
  } catch(e) { /* ignore */ }
""");
        }

        sb.Append("""
}

// Optional WebView2 integration: allow hosts to push updates without ExecuteScriptAsync.
// - PostWebMessageAsString(bodyHtml)  => e.data is a string
// - PostWebMessageAsJson({ bodyHtml }) => e.data is an object
(function(){
  try {
    const wv = window.chrome && window.chrome.webview;
    if (!wv || typeof wv.addEventListener !== 'function') return;
    wv.addEventListener('message', e => {
      try {
        const d = e && e.data;
        if (d && typeof d === 'object' && d.type === 'omd.update' && typeof d.bodyHtml === 'string') { updateContent(d.bodyHtml); return; }
        if (typeof d === 'string') { updateContent(d); return; }
        if (d && typeof d === 'object' && typeof d.bodyHtml === 'string') { updateContent(d.bodyHtml); return; }
      } catch(_) { /* ignore */ }
    });
  } catch(_) { /* ignore */ }
})();
""");
        return sb.ToString();
    }
}
