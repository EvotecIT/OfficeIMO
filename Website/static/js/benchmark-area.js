// Draws each workload on a benchmark area page from its published result JSON.
// The server-rendered page already lists every run and its raw result; this adds the charts.
(function () {
  "use strict";
  var IGNORED_VARIABLES = { namespace: true, type: true, fullname: true };
  var cache = {};

  function el(tag, cls, text) {
    var node = document.createElement(tag);
    if (cls) node.className = cls;
    if (text !== undefined && text !== null) node.textContent = text;
    return node;
  }
  function load(path) {
    if (!cache[path]) {
      cache[path] = fetch(path, { credentials: "same-origin" }).then(function (r) {
        if (!r.ok) throw new Error("HTTP " + r.status);
        return r.json();
      });
    }
    return cache[path];
  }
  function formatMs(ms) {
    if (typeof ms !== "number" || !isFinite(ms)) return "—";
    if (ms >= 1000) return (ms / 1000).toFixed(ms >= 10000 ? 1 : 2) + " s";
    if (ms >= 100) return ms.toFixed(0) + " ms";
    if (ms >= 10) return ms.toFixed(1) + " ms";
    if (ms >= 1) return ms.toFixed(2) + " ms";
    return (ms * 1000).toFixed(ms >= 0.1 ? 0 : 1) + " µs";
  }
  function formatBytes(bytes) {
    if (typeof bytes !== "number" || !isFinite(bytes)) return "—";
    if (bytes >= 1073741824) return (bytes / 1073741824).toFixed(2) + " GB";
    if (bytes >= 1048576) return (bytes / 1048576).toFixed(bytes >= 104857600 ? 0 : 1) + " MB";
    if (bytes >= 1024) return (bytes / 1024).toFixed(0) + " KB";
    return bytes + " B";
  }
  function allocated(row) {
    var m = row.metrics || {};
    var v = m.BytesAllocatedPerOperation;
    if (typeof v !== "number") v = m.Allocated;
    return typeof v === "number" ? v : null;
  }
  function libraryName(scenario) { return String(scenario || "").replace(/_/g, " "); }
  function libraryVersion(metadata, scenario) {
    var key = String(scenario || "").split("_")[0].toLowerCase();
    return metadata["benchmark.dependency." + key + ".version"] || "";
  }
  function isOfficeImo(scenario) { return /^officeimo/i.test(String(scenario || "")); }
  function groupKey(row) {
    var vars = row.variables || {};
    return Object.keys(vars).filter(function (k) { return !IGNORED_VARIABLES[k.toLowerCase()]; }).sort()
      .map(function (k) { return k + ": " + vars[k]; }).join(" · ");
  }

  function render(target, result, run) {
    var metadata = result.metadata || {};
    var env = result.environment || {};
    target.textContent = "";

    var meta = el("p", "imo-bench-meta");
    var parts = [];
    if (env.processorName) parts.push(env.processorName);
    if (env.osDescription) parts.push(String(env.osDescription).split(" [")[0]);
    if (env.runtimeVersion) parts.push(String(env.runtimeVersion).split(" (")[0]);
    parts.push(run.mode === "full" ? "Full run" : "Quick run");
    var affinity = metadata["benchmark.workload.affinityMask"];
    if (affinity) parts.push("CPU affinity " + affinity);
    var versions = Object.keys(metadata).filter(function (k) { return /^benchmark\.dependency\.[^.]+\.version$/.test(k); })
      .map(function (k) { return k.split(".")[2] + " " + metadata[k]; });
    meta.textContent = parts.join(" · ");
    target.appendChild(meta);
    if (run.mode !== "full") {
      target.appendChild(el("p", "imo-bench-note", "Quick runs are short diagnostic measurements. Use full runs for comparisons."));
    }

    var groups = {};
    var order = [];
    (result.summary || []).forEach(function (row) {
      var key = groupKey(row);
      if (!groups[key]) { groups[key] = []; order.push(key); }
      groups[key].push(row);
    });
    order.forEach(function (key) {
      var rows = groups[key].slice().sort(function (a, b) {
        var am = typeof a.medianMs === "number" ? a.medianMs : Infinity;
        var bm = typeof b.medianMs === "number" ? b.medianMs : Infinity;
        return am - bm;
      });
      var fastest = rows.length && typeof rows[0].medianMs === "number" ? rows[0].medianMs : null;
      var slowest = rows.reduce(function (max, r) { return typeof r.medianMs === "number" && r.medianMs > max ? r.medianMs : max; }, 0);
      var figure = el("figure", "imo-bench-figure");
      if (key) figure.appendChild(el("figcaption", null, key));
      var wrap = el("div", "imo-bench-table-wrap");
      var table = el("table", "imo-bench-table");
      var head = el("thead");
      var hr = el("tr");
      ["Library", "Median time", "Compared with fastest", "Allocated", ""].forEach(function (label, i) {
        var th = el("th", i === 4 ? "imo-bench-table__bar" : null, label);
        th.scope = "col";
        if (i === 4) th.setAttribute("aria-hidden", "true");
        hr.appendChild(th);
      });
      head.appendChild(hr);
      table.appendChild(head);
      var body = el("tbody");
      rows.forEach(function (row) {
        var tr = el("tr", isOfficeImo(row.scenario) ? "is-officeimo" : null);
        var name = el("th", null, libraryName(row.scenario));
        name.scope = "row";
        var version = libraryVersion(metadata, row.scenario);
        if (version) name.appendChild(el("small", null, " " + version));
        tr.appendChild(name);
        var failed = row.status && row.status !== "Succeeded";
        tr.appendChild(el("td", "imo-bench-num", failed ? "Failed" : formatMs(row.medianMs)));
        var ratio = !failed && fastest && typeof row.medianMs === "number" ? row.medianMs / fastest : null;
        tr.appendChild(el("td", "imo-bench-num", ratio === null ? "—" : ratio < 1.005 ? "Fastest" : ratio.toFixed(2) + "× slower"));
        tr.appendChild(el("td", "imo-bench-num", formatBytes(allocated(row))));
        var barCell = el("td", "imo-bench-table__bar");
        barCell.setAttribute("aria-hidden", "true");
        if (!failed && typeof row.medianMs === "number" && slowest > 0) {
          var bar = el("span", "imo-bench-bar");
          bar.style.width = Math.max(1.5, row.medianMs / slowest * 100).toFixed(1) + "%";
          barCell.appendChild(bar);
        }
        tr.appendChild(barCell);
        body.appendChild(tr);
      });
      table.appendChild(body);
      wrap.appendChild(table);
      figure.appendChild(wrap);
      target.appendChild(figure);
    });

    var links = el("p", "imo-bench-links");
    var sha = metadata.gitSha || metadata["benchmark.workload.sourceCommit"];
    if (sha) {
      var commit = el("a", null, "Source " + String(sha).substring(0, 12));
      commit.href = "https://github.com/EvotecIT/OfficeIMO/tree/" + sha;
      links.appendChild(commit);
      links.appendChild(document.createTextNode(" · "));
    }
    var raw = el("a", null, "Result JSON");
    raw.href = run.path;
    links.appendChild(raw);
    if (versions.length) links.appendChild(document.createTextNode(" · Versions: " + versions.join(", ")));
    target.appendChild(links);
  }

  function enhance(section) {
    var buttons = Array.prototype.slice.call(section.querySelectorAll(".imo-bench-run"));
    var target = section.querySelector("[data-benchmark-chart]");
    if (!buttons.length || !target) return null;
    var active = 0;
    function select(button) {
      var request = ++active;
      buttons.forEach(function (b) { b.setAttribute("aria-pressed", b === button ? "true" : "false"); });
      var run = { path: button.getAttribute("data-result"), mode: button.getAttribute("data-mode"), platform: button.getAttribute("data-platform") };
      target.setAttribute("aria-busy", "true");
      load(run.path).then(function (result) {
        if (request !== active) return;
        render(target, result, run);
      }).catch(function (error) {
        if (request !== active) return;
        target.textContent = "";
        var p = el("p", "imo-bench-note", "This result couldn't be loaded (" + error.message + "). ");
        var a = el("a", null, "Open the result JSON");
        a.href = run.path;
        p.appendChild(a);
        target.appendChild(p);
      }).then(function () { if (request === active) target.removeAttribute("aria-busy"); });
    }
    buttons.forEach(function (b) { b.addEventListener("click", function () { select(b); }); });
    function score(b) {
      var s = 0;
      if (b.getAttribute("data-mode") === "full") s += 4;
      if (b.getAttribute("data-published") === "true") s += 2;
      if (b.getAttribute("data-platform") === "windows") s += 1;
      return s;
    }
    var preferred = buttons.slice().sort(function (a, b) { return score(b) - score(a); })[0];
    return function () { select(preferred); };
  }

  function init() {
    var sections = Array.prototype.slice.call(document.querySelectorAll("[data-benchmark-workload]"));
    // Result files can be several megabytes, so each chart loads when its section approaches the viewport.
    var starters = sections.map(enhance);
    if (!("IntersectionObserver" in window)) { starters.forEach(function (start) { if (start) start(); }); return; }
    var observer = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (!entry.isIntersecting) return;
        observer.unobserve(entry.target);
        var start = starters[sections.indexOf(entry.target)];
        if (start) start();
      });
    }, { rootMargin: "600px 0px" });
    sections.forEach(function (section) { observer.observe(section); });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
