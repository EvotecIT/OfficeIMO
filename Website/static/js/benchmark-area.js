// Draws each workload on a benchmark area page from the published result summary
// (static/data/benchmarks/library-comparisons/summary.json, written by
// Build/Generate-LibraryComparisonWebsiteData.ps1). The server-rendered page already lists
// every run with its raw result JSON; this adds, per workload:
//   1. every contender across every run (platform, full and quick), relative to the fastest;
//   2. for the selected run, a verdict and a case-by-library matrix with time and allocations;
//   3. the per-case tables with bars, as the complete detail.
(function () {
  "use strict";
  var SUMMARY_URL = "/data/benchmarks/library-comparisons/summary.json";
  var summaryPromise = null;

  function el(tag, cls, text) {
    var node = document.createElement(tag);
    if (cls) node.className = cls;
    if (text !== undefined && text !== null) node.textContent = text;
    return node;
  }
  function loadSummary() {
    if (!summaryPromise) {
      summaryPromise = fetch(SUMMARY_URL, { credentials: "same-origin" }).then(function (r) {
        if (!r.ok) throw new Error("HTTP " + r.status);
        return r.json();
      });
    }
    return summaryPromise;
  }
  function isNum(v) { return typeof v === "number" && isFinite(v); }
  function formatMs(ms) {
    if (!isNum(ms)) return "—";
    if (ms >= 1000) return (ms / 1000).toFixed(ms >= 10000 ? 1 : 2) + " s";
    if (ms >= 100) return ms.toFixed(0) + " ms";
    if (ms >= 10) return ms.toFixed(1) + " ms";
    if (ms >= 1) return ms.toFixed(2) + " ms";
    return (ms * 1000).toFixed(ms >= 0.1 ? 0 : 1) + " µs";
  }
  function formatBytes(bytes) {
    if (!isNum(bytes)) return "—";
    if (bytes >= 1073741824) return (bytes / 1073741824).toFixed(2) + " GB";
    if (bytes >= 1048576) return (bytes / 1048576).toFixed(bytes >= 104857600 ? 0 : 1) + " MB";
    if (bytes >= 1024) return (bytes / 1024).toFixed(0) + " KB";
    return Math.round(bytes) + " B";
  }
  function formatRatio(r) { return r >= 10 ? r.toFixed(0) : r >= 1.995 ? r.toFixed(1) : r.toFixed(2); }
  function geomean(values) {
    var list = values.filter(function (v) { return isNum(v) && v > 0; });
    if (!list.length) return null;
    return Math.exp(list.reduce(function (s, v) { return s + Math.log(v); }, 0) / list.length);
  }
  function key(s) { return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, ""); }
  function isOfficeImo(name) { return /^officeimo/i.test(name); }
  var PLATFORM = { windows: "Windows", linux: "Linux", macos: "macOS" };
  function runLabel(run) { return (PLATFORM[run.platform] || run.platform) + " · " + (run.runMode === "full" ? "full" : "quick"); }

  // Splits scenario names such as "MagickNETEncode" or "OfficeIMO_WriteDataReader" into the library and,
  // when the same suffix appears for two or more libraries, an operation that becomes part of the case.
  function normalizeRows(rows, libraries) {
    var libs = (libraries || []).map(function (name) { return { name: name, key: key(name) }; })
      .sort(function (a, b) { return b.key.length - a.key.length; });
    var parsed = rows.map(function (row) {
      var raw = String(row.library || "");
      var flat = key(raw);
      var match = null;
      for (var i = 0; i < libs.length; i++) { if (libs[i].key && flat.indexOf(libs[i].key) === 0) { match = libs[i]; break; } }
      var rest = "";
      if (match) {
        // Walk the raw name until the matched letters are consumed, then keep the rest.
        var consumed = 0, j = 0;
        while (j < raw.length && consumed < match.key.length) { if (/[a-z0-9]/i.test(raw[j])) consumed++; j++; }
        rest = raw.slice(j).replace(/^[\s_\-.]+/, "").replace(/_/g, " ");
      }
      return { row: row, library: match ? match.name : raw.replace(/_/g, " "), rest: rest };
    });
    var restLibraries = {};
    parsed.forEach(function (p) { if (p.rest) { (restLibraries[p.rest] = restLibraries[p.rest] || {})[p.library] = true; } });
    var operations = {};
    Object.keys(restLibraries).forEach(function (r) { if (Object.keys(restLibraries[r]).length > 1) operations[r] = true; });
    return parsed.map(function (p) {
      var c = {};
      Object.keys(p.row.case || {}).forEach(function (k) { c[k] = p.row.case[k]; });
      var contender = p.library;
      if (p.rest && operations[p.rest]) c.Operation = p.rest;
      else if (p.rest) contender = p.library + " " + p.rest;
      return { contender: contender, isOfficeImo: isOfficeImo(contender), case: c, timeMs: p.row.timeMs, statistic: p.row.statistic,
        allocatedBytes: p.row.allocatedBytes, failed: !!p.row.status && p.row.status !== "Succeeded" };
    });
  }

  // Groups rows into cases. Parameters that never change are shown once as the workload, not per case.
  function buildModel(run, libraries) {
    var rows = normalizeRows(run.rows || [], libraries);
    var paramValues = {};
    rows.forEach(function (r) { Object.keys(r.case).forEach(function (k) { (paramValues[k] = paramValues[k] || {})[r.case[k]] = true; }); });
    var varying = Object.keys(paramValues).filter(function (k) { return Object.keys(paramValues[k]).length > 1; });
    var constant = Object.keys(paramValues).filter(function (k) { return varying.indexOf(k) < 0; })
      .map(function (k) { return k + ": " + Object.keys(paramValues[k])[0]; });
    var cases = [], caseIndex = {}, contenders = [], seen = {};
    rows.forEach(function (r) {
      var label = varying.map(function (k) { return (varying.length > 1 ? k + " " : "") + (r.case[k] === undefined ? "—" : r.case[k]); }).join(" · ");
      if (!(label in caseIndex)) { caseIndex[label] = cases.length; cases.push({ label: label, cells: {} }); }
      cases[caseIndex[label]].cells[r.contender] = r;
      if (!seen[r.contender]) { seen[r.contender] = true; contenders.push(r.contender); }
    });
    cases.forEach(function (c) {
      var best = null, bestAlloc = null;
      contenders.forEach(function (name) {
        var cell = c.cells[name];
        if (!cell || cell.failed || !isNum(cell.timeMs)) return;
        if (best === null || cell.timeMs < best) best = cell.timeMs;
        if (isNum(cell.allocatedBytes) && (bestAlloc === null || cell.allocatedBytes < bestAlloc)) bestAlloc = cell.allocatedBytes;
      });
      c.best = best; c.bestAlloc = bestAlloc;
    });
    // Geometric mean of each contender's time relative to the fastest in every case it ran.
    var scores = contenders.map(function (name) {
      var ratios = [], wins = 0, measured = 0;
      cases.forEach(function (c) {
        var cell = c.cells[name];
        if (!cell || cell.failed || !isNum(cell.timeMs) || !c.best) return;
        measured++;
        ratios.push(cell.timeMs / c.best);
        if (cell.timeMs / c.best < 1.005) wins++;
      });
      return { name: name, ratio: geomean(ratios), wins: wins, measured: measured };
    });
    contenders.sort(function (a, b) {
      var sa = scores[contenders.indexOf(a)], sb = scores[contenders.indexOf(b)];
      if (isOfficeImo(a) !== isOfficeImo(b)) return isOfficeImo(a) ? -1 : 1;
      return (sa.ratio || Infinity) - (sb.ratio || Infinity);
    });
    var scoreByName = {};
    scores.forEach(function (s) { scoreByName[s.name] = s; });
    var statistic = rows.some(function (r) { return r.statistic === "mean"; }) ? "Mean" : "Median";
    return { cases: cases, contenders: contenders, scores: scoreByName, constant: constant, statistic: statistic };
  }

  // OfficeIMO against each other contender: geometric mean of per-case time and allocation ratios.
  function headToHead(model, a, b) {
    var t = [], m = [];
    model.cases.forEach(function (c) {
      var x = c.cells[a], y = c.cells[b];
      if (!x || !y || x.failed || y.failed || !isNum(x.timeMs) || !isNum(y.timeMs)) return;
      t.push(x.timeMs / y.timeMs);
      if (isNum(x.allocatedBytes) && isNum(y.allocatedBytes) && y.allocatedBytes > 0 && x.allocatedBytes > 0) m.push(x.allocatedBytes / y.allocatedBytes);
    });
    return { time: geomean(t), memory: geomean(m), cases: t.length };
  }
  function compareText(ratio, subject, other, word) {
    if (!isNum(ratio)) return null;
    if (ratio < 0.97) return subject + " is " + formatRatio(1 / ratio) + "× " + word[0] + " than " + other;
    if (ratio > 1.03) return other + " is " + formatRatio(ratio) + "× " + word[0] + " than " + subject;
    return subject + " and " + other + " are " + word[1];
  }

  function renderOverview(target, runs, libraries) {
    target.textContent = "";
    if (runs.length < 2) return;
    var models = runs.map(function (run) { return buildModel(run, libraries); });
    var names = [];
    models.forEach(function (m) { m.contenders.forEach(function (n) { if (names.indexOf(n) < 0) names.push(n); }); });
    names.sort(function (a, b) { return isOfficeImo(a) === isOfficeImo(b) ? 0 : isOfficeImo(a) ? -1 : 1; });
    var figure = el("figure", "imo-bench-glance");
    figure.appendChild(el("figcaption", null, "Every run at a glance: time relative to the fastest library in each run (geometric mean over its cases). 1.00× is the fastest; lower is better."));
    var wrap = el("div", "imo-bench-table-wrap");
    var table = el("table", "imo-bench-table imo-bench-table--glance");
    var head = el("tr");
    var corner = el("th", null, "Library");
    corner.scope = "col";
    head.appendChild(corner);
    runs.forEach(function (run) {
      var th = el("th", "imo-bench-num", runLabel(run));
      th.scope = "col";
      th.appendChild(el("small", null, (run.generatedUtc || "").slice(0, 10)));
      head.appendChild(th);
    });
    var thead = el("thead");
    thead.appendChild(head);
    table.appendChild(thead);
    var body = el("tbody");
    names.forEach(function (name) {
      var tr = el("tr", isOfficeImo(name) ? "is-officeimo" : null);
      var th = el("th", null, name);
      th.scope = "row";
      tr.appendChild(th);
      models.forEach(function (m) {
        var s = m.scores[name];
        var td = el("td", "imo-bench-num");
        if (!s || !isNum(s.ratio)) { td.textContent = "—"; td.title = "Not in this run"; }
        else {
          td.textContent = s.ratio < 1.005 ? "1.00×" : formatRatio(s.ratio) + "×";
          if (s.ratio < 1.005) td.className += " is-best";
        }
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });
    table.appendChild(body);
    wrap.appendChild(table);
    figure.appendChild(wrap);
    target.appendChild(figure);
  }

  function renderRun(target, run, libraries) {
    var model = buildModel(run, libraries);
    target.textContent = "";

    var env = run.environment || {};
    var parts = [];
    if (env.processor) parts.push(env.processor + (env.cores ? " (" + env.cores + " threads)" : ""));
    if (env.os) parts.push(env.os);
    if (env.runtime) parts.push(env.runtime);
    if (env.affinity) parts.push("CPU affinity " + env.affinity);
    var meta = el("p", "imo-bench-meta");
    meta.appendChild(el("strong", null, runLabel(run) + " run"));
    meta.appendChild(document.createTextNode(" · " + (parts.length ? parts.join(" · ") : "Environment not recorded for this run") + " · " + (run.generatedUtc || "").slice(0, 10)));
    target.appendChild(meta);
    if (run.runMode !== "full") {
      target.appendChild(el("p", "imo-bench-note", "Quick run: a short diagnostic measurement, shown so every contender and platform is visible. Full runs carry more weight for comparisons."));
    }
    if (model.constant.length) target.appendChild(el("p", "imo-bench-workload-shape", "Workload: " + model.constant.join(" · ")));

    // Verdict: OfficeIMO against each other contender.
    var ours = model.contenders.filter(isOfficeImo);
    var theirs = model.contenders.filter(function (n) { return !isOfficeImo(n); });
    if (ours.length && theirs.length) {
      var list = el("ul", "imo-bench-verdict");
      ours.forEach(function (a) {
        theirs.forEach(function (b) {
          var h = headToHead(model, a, b);
          if (!h.cases) return;
          var li = el("li");
          var time = compareText(h.time, a, b, ["faster", "about as fast"]);
          var faster = isNum(h.time) && h.time < 0.97, slower = isNum(h.time) && h.time > 1.03;
          li.className = faster ? "is-ahead" : slower ? "is-behind" : "is-even";
          li.appendChild(el("strong", null, time));
          var detail = [];
          if (isNum(h.memory)) detail.push(h.memory < 0.97 ? a + " allocates " + formatRatio(1 / h.memory) + "× less" : h.memory > 1.03 ? a + " allocates " + formatRatio(h.memory) + "× more" : "similar allocations");
          detail.push(h.cases === 1 ? "1 case" : "geometric mean of " + h.cases + " cases");
          li.appendChild(el("span", null, detail.join(" · ")));
          list.appendChild(li);
        });
      });
      target.appendChild(list);
    }

    // Matrix: every case against every contender.
    var figure = el("figure", "imo-bench-figure");
    figure.appendChild(el("figcaption", null, model.statistic + " time per operation, allocations below. The fastest in each row is highlighted; lower is better."));
    var wrap = el("div", "imo-bench-table-wrap");
    var table = el("table", "imo-bench-table imo-bench-table--matrix");
    var hr = el("tr");
    var corner = el("th", null, model.cases.length > 1 ? "Case" : "Workload");
    corner.scope = "col";
    hr.appendChild(corner);
    model.contenders.forEach(function (name) {
      var th = el("th", "imo-bench-num" + (isOfficeImo(name) ? " is-officeimo" : ""), name);
      th.scope = "col";
      var version = (run.versions || {})[key(name.split(" ")[0])];
      if (version) th.appendChild(el("small", null, version));
      hr.appendChild(th);
    });
    var thead = el("thead");
    thead.appendChild(hr);
    table.appendChild(thead);
    var body = el("tbody");
    model.cases.forEach(function (c) {
      var tr = el("tr");
      var th = el("th", null, c.label || "All");
      th.scope = "row";
      tr.appendChild(th);
      model.contenders.forEach(function (name) {
        var cell = c.cells[name];
        var td = el("td", "imo-bench-num" + (isOfficeImo(name) ? " is-officeimo" : ""));
        if (!cell) td.textContent = "—";
        else if (cell.failed) td.textContent = "Failed";
        else {
          var best = c.best && isNum(cell.timeMs) && cell.timeMs / c.best < 1.005;
          if (best) td.className += " is-best";
          td.appendChild(el("span", "imo-bench-time", formatMs(cell.timeMs)));
          var sub = best ? "fastest" : (c.best && isNum(cell.timeMs) ? formatRatio(cell.timeMs / c.best) + "×" : "");
          if (isNum(cell.allocatedBytes)) sub += (sub ? " · " : "") + formatBytes(cell.allocatedBytes);
          if (sub) td.appendChild(el("small", null, sub));
        }
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });
    table.appendChild(body);
    wrap.appendChild(table);
    figure.appendChild(wrap);
    target.appendChild(figure);

    // Complete detail: one bar table per case.
    if (model.cases.length) {
      var details = el("details", "imo-bench-detail");
      details.appendChild(el("summary", null, "Bars for each case"));
      model.cases.forEach(function (c) {
        var rows = model.contenders.filter(function (n) { return c.cells[n]; }).map(function (n) { return c.cells[n]; })
          .sort(function (a, b) { return (isNum(a.timeMs) ? a.timeMs : Infinity) - (isNum(b.timeMs) ? b.timeMs : Infinity); });
        var slowest = rows.reduce(function (max, r) { return isNum(r.timeMs) && r.timeMs > max ? r.timeMs : max; }, 0);
        var block = el("div", "imo-bench-bars");
        if (c.label) block.appendChild(el("h4", null, c.label));
        rows.forEach(function (r) {
          var line = el("div", "imo-bench-bars__row" + (r.isOfficeImo ? " is-officeimo" : ""));
          line.appendChild(el("span", "imo-bench-bars__name", r.contender));
          var track = el("span", "imo-bench-bars__track");
          if (!r.failed && isNum(r.timeMs) && slowest > 0) {
            var bar = el("span", "imo-bench-bar");
            bar.style.width = Math.max(1.5, r.timeMs / slowest * 100).toFixed(1) + "%";
            track.appendChild(bar);
          }
          line.appendChild(track);
          line.appendChild(el("span", "imo-bench-bars__value", r.failed ? "Failed" : formatMs(r.timeMs)));
          block.appendChild(line);
        });
        details.appendChild(block);
      });
      target.appendChild(details);
    }

    var links = el("p", "imo-bench-links");
    if (run.sourceCommit) {
      var commit = el("a", null, "Source " + String(run.sourceCommit).substring(0, 12));
      commit.href = "https://github.com/EvotecIT/OfficeIMO/tree/" + run.sourceCommit;
      links.appendChild(commit);
      links.appendChild(document.createTextNode(" · "));
    }
    var raw = el("a", null, "Raw result JSON with every sample");
    raw.href = run.resultPath;
    links.appendChild(raw);
    var versions = Object.keys(run.versions || {}).map(function (k) { return k + " " + run.versions[k]; });
    if (versions.length) links.appendChild(document.createTextNode(" · Versions: " + versions.join(", ")));
    target.appendChild(links);
  }

  function enhance(section, summary) {
    var id = section.getAttribute("data-comparison");
    var libraries = (section.getAttribute("data-libraries") || "").split("|").filter(Boolean);
    var buttons = Array.prototype.slice.call(section.querySelectorAll(".imo-bench-run"));
    var target = section.querySelector("[data-benchmark-chart]");
    var overview = section.querySelector("[data-benchmark-overview]");
    var byPath = {};
    (summary.runs || []).forEach(function (run) { if (run.comparisonId === id) byPath[run.resultPath] = run; });
    var runs = buttons.map(function (b) { return byPath[b.getAttribute("data-result")]; }).filter(Boolean);
    if (!buttons.length || !target || !runs.length) return;
    if (overview) renderOverview(overview, runs, libraries);
    function select(button) {
      var run = byPath[button.getAttribute("data-result")];
      if (!run) return;
      buttons.forEach(function (b) { b.setAttribute("aria-pressed", b === button ? "true" : "false"); });
      renderRun(target, run, libraries);
    }
    buttons.forEach(function (b) { b.addEventListener("click", function () { select(b); }); });
    function score(b) {
      var s = 0;
      if (b.getAttribute("data-mode") === "full") s += 4;
      if (b.getAttribute("data-published") === "true") s += 2;
      if (b.getAttribute("data-platform") === "windows") s += 1;
      return s;
    }
    select(buttons.slice().sort(function (a, b) { return score(b) - score(a); })[0]);
  }

  function preferredRun(runs) {
    function score(run) { return (run.runMode === "full" ? 4 : 0) + (run.publish ? 2 : 0) + (run.platform === "windows" ? 1 : 0); }
    return runs.slice().sort(function (a, b) { return score(b) - score(a) || String(b.generatedUtc).localeCompare(String(a.generatedUtc)); })[0];
  }

  // Benchmarks hub: one row per measured comparison, using its strongest run, so every contender is visible at once.
  function renderHub(root, summary) {
    var byId = {}, order = [];
    (summary.runs || []).forEach(function (run) {
      if (!byId[run.comparisonId]) { byId[run.comparisonId] = []; order.push(run.comparisonId); }
      byId[run.comparisonId].push(run);
    });
    order.sort(function (a, b) {
      var ra = byId[a][0], rb = byId[b][0];
      return String(ra.area).localeCompare(String(rb.area)) || String(ra.title).localeCompare(String(rb.title)) || a.localeCompare(b);
    });
    var wrap = el("div", "imo-bench-table-wrap");
    var table = el("table", "imo-bench-table imo-bench-table--hub");
    var hr = el("tr");
    ["Workload", "Libraries compared", "Run shown", "Result"].forEach(function (label) {
      var th = el("th", null, label);
      th.scope = "col";
      hr.appendChild(th);
    });
    var thead = el("thead");
    thead.appendChild(hr);
    table.appendChild(thead);
    var body = el("tbody");
    order.forEach(function (id) {
      var runs = byId[id];
      var run = preferredRun(runs);
      var model = buildModel(run, run.libraries);
      var ranked = model.contenders.slice().filter(function (n) { return model.scores[n] && isNum(model.scores[n].ratio); })
        .sort(function (a, b) { return model.scores[a].ratio - model.scores[b].ratio; });
      var tr = el("tr");
      var th = el("th");
      th.scope = "row";
      var link = el("a", null, (run.title || id) + (/-net8\.0/.test(id) ? " (.NET 8)" : ""));
      link.href = "/benchmarks/" + run.area + "/#" + id;
      th.appendChild(link);
      tr.appendChild(th);
      tr.appendChild(el("td", "imo-bench-hub__libs", model.contenders.join(", ")));
      var runCell = el("td", null, runLabel(run));
      if (runs.length > 1) runCell.appendChild(el("small", null, " + " + (runs.length - 1) + " more"));
      tr.appendChild(runCell);
      var result = el("td", "imo-bench-hub__result");
      var ours = ranked.filter(isOfficeImo)[0];
      if (!ours || !ranked.length) result.textContent = "—";
      else {
        var rank = ranked.indexOf(ours) + 1;
        var badge = el("span", "imo-bench-rank" + (rank === 1 ? " is-first" : ""), rank === 1 ? "Fastest" : ordinal(rank) + " of " + ranked.length);
        result.appendChild(badge);
        var text;
        if (rank === 1) {
          var next = ranked.filter(function (n) { return !isOfficeImo(n); })[0];
          var h = next ? headToHead(model, ours, next) : null;
          text = next && h && isNum(h.time) ? next + " is " + formatRatio(1 / h.time) + "× slower" : "";
        } else {
          var h2 = headToHead(model, ours, ranked[0]);
          text = isNum(h2.time) ? ranked[0] + " is " + formatRatio(h2.time) + "× faster" : "";
        }
        if (text) result.appendChild(el("span", null, text));
      }
      tr.appendChild(result);
      body.appendChild(tr);
    });
    table.appendChild(body);
    wrap.appendChild(table);
    root.textContent = "";
    root.appendChild(wrap);
  }
  function ordinal(n) { return n === 2 ? "2nd" : n === 3 ? "3rd" : n + "th"; }

  function init() {
    var sections = Array.prototype.slice.call(document.querySelectorAll("[data-benchmark-workload]"));
    var hub = document.querySelector("[data-benchmark-hub-results]");
    if (!sections.length && !hub) return;
    loadSummary().then(function (summary) {
      if (hub) renderHub(hub, summary);
      sections.forEach(function (section) { enhance(section, summary); });
    }).catch(function (error) {
      sections.forEach(function (section) {
        var target = section.querySelector("[data-benchmark-chart]");
        if (target) target.insertBefore(el("p", "imo-bench-note", "The result summary couldn't be loaded (" + error.message + "). Each run's raw result is linked below."), target.firstChild);
      });
    });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
