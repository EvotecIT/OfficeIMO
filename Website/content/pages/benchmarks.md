---
title: "Benchmarks"
description: "Compare reproducible OfficeIMO performance evidence across document and data workloads, with measured environments, datasets, results, and repeatable commands."
layout: page
meta.raw_html: true
---

<div class="imo-benchmark-hub">
  <section class="imo-benchmark-hub__hero">
    <p class="imo-benchmark-eyebrow">Performance evidence</p>
    <h2>Use the benchmark that matches your workload.</h2>
    <p>Office documents, delimited data, extraction, and rendering have different cost profiles. Every result below links to a committed suite, a reproducible command, or a documented performance contract.</p>
    <div class="imo-benchmark-principles" aria-label="Benchmark principles">
      <span>Equivalent work</span>
      <span>Validated output</span>
      <span>Committed artifacts</span>
      <span>Platforms kept separate</span>
    </div>
  </section>

  <section class="imo-benchmark-evidence-model" aria-labelledby="benchmark-evidence-model-title">
    <div class="imo-benchmark-section-heading">
      <p class="imo-benchmark-eyebrow">Evidence model</p>
      <h2 id="benchmark-evidence-model-title">Broad coverage, with every claim kept in its lane.</h2>
      <p>The source tree, publishable comparison artifacts, and older engineering runs answer different questions. The website shows all three without presenting a runnable suite as a measured result or an older workstation run as a current cross-platform ranking.</p>
    </div>
    <div class="imo-benchmark-evidence-model__grid">
      <article>
        <span class="imo-benchmark-evidence-model__index">01</span>
        <strong>33 source suites</strong>
        <h3>What can be measured</h3>
        <p>Benchmark projects in the current repository define runnable workloads, validation, and reproduction commands. Their existence is coverage—not a performance claim.</p>
      </article>
      <article>
        <span class="imo-benchmark-evidence-model__index">02</span>
        <strong data-published-full-count>Cataloged evidence</strong>
        <h3>What can be compared</h3>
        <p>Full artifacts include the measured source revision, runtime, hardware, operating system, and validated result rows. Diagnostic runs stay visible but do not support public ranking claims.</p>
      </article>
      <article>
        <span class="imo-benchmark-evidence-model__index">03</span>
        <strong>Full historical matrix</strong>
        <h3>What remains useful</h3>
        <p>Older Excel evidence keeps hundreds of scenarios and every participating library available. Missing environment provenance is displayed beside it, so the detail survives without overstating certainty.</p>
      </article>
    </div>
  </section>

  <section class="imo-benchmark-coverage" aria-labelledby="benchmark-coverage-title">
    <div class="imo-benchmark-section-heading">
      <p class="imo-benchmark-eyebrow">Source coverage by family</p>
      <h2 id="benchmark-coverage-title">What the current source tree can measure</h2>
      <p>A comparison measures equivalent libraries. A regression suite protects OfficeIMO against its own baseline. A guardrail enforces a budget. The source inventory keeps those contracts discoverable even when a current public result has not been committed.</p>
    </div>
    <div class="imo-benchmark-coverage__grid">
      <article data-family="data-readers"><span>Published and diagnostic evidence</span><h3>CSV and Excel reads and writes</h3><p>Pinned 65K-record read fixtures and validated 25,000-row IDataReader write contracts are cataloged per workload, source revision, operating system, and run mode. The coverage grid shows full, diagnostic, and unpublished lanes.</p><a href="#library-comparison-evidence">Inspect published coverage</a></article>
      <article data-family="excel"><span>Current and historical evidence</span><h3>Excel create and write</h3><p>The current compact IDataReader write lane records OS and run mode. Additional create, write, package, and engineering scenarios remain available as explicitly historical snapshots.</p><a href="#library-comparison-evidence">Select current XLSX write evidence</a></article>
      <article data-family="csv"><span>Current and historical evidence</span><h3>CSV write workloads</h3><p>The current IDataReader lane records OS and run mode. Additional wide and database-shaped snapshots are retained as historical engineering evidence because their original environment dimensions were not recorded.</p><a href="#library-comparison-evidence">Select current CSV write evidence</a></article>
      <article data-family="reader"><span>Regression baseline</span><h3>Reader</h3><p>25 cases across 14 document formats, with detection, chunking, and transport lanes. Timings are a local regression baseline, not a cross-machine promise.</p><a href="https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/benchmarks/officeimo.reader.foundation-2026-07-10.md" target="_blank" rel="noopener">Open the Reader evidence</a></article>
      <article data-family="guardrails"><span>Performance guardrails</span><h3>PDF, RTF, and Email</h3><p>Budget and regression checks cover representative rendering, parsing, memory, and I/O behavior without presenting unrelated engines as equivalent competitors.</p><a href="/docs/capabilities/benchmarks/#performance-guardrails">Run the guardrail suites</a></article>
      <article data-family="formats"><span>Repeatable suites</span><h3>Markdown, HTML, and open formats</h3><p>Dedicated projects exercise Markdown, HTML, OneNote, OpenDocument, and drawing workloads with scenario-specific validation.</p><a href="/docs/capabilities/benchmarks/#additional-benchmark-projects">Find the projects and commands</a></article>
      <article data-family="powershell"><span>Published comparisons</span><h3>PSWriteOffice</h3><p>PowerForge-backed Excel and CSV comparisons cover PowerShell object, DataTable, compression, workbook, and database-shaped workflows with validation.</p><a href="/docs/workflows/powershell-benchmarks/">Reproduce the PSWriteOffice suites</a></article>
      <article data-family="boundaries"><span>Coverage boundary</span><h3>Word and PowerPoint</h3><p>Performance-focused tests protect known workflows. We do not publish a cross-library ranking until equivalent workloads and committed result artifacts exist.</p><a href="/docs/capabilities/benchmarks/#word-and-powerpoint">See what is verified today</a></article>
    </div>
  </section>

  <div id="library-comparison-evidence">
    {{< include path="../../themes/officeimo/partials/library-comparison-benchmarks.html" >}}
  </div>

  <section class="imo-benchmark-archive" id="historical-benchmark-evidence" aria-labelledby="historical-benchmark-title">
    <header class="imo-benchmark-archive__header">
      <div>
        <p class="imo-benchmark-eyebrow">Complete engineering matrix</p>
        <h2 id="historical-benchmark-title">Explore every measured Excel scenario.</h2>
      </div>
      <p>The full matrix keeps all 274 scenario rows, 1,106 measurements, eight libraries, row tiers, workloads, categories, package sizes, and relative timings available for filtering and sorting. Its original artifact did not record operating system or run mode, so it remains separate from the platform-specific selector above.</p>
    </header>
    {{< include path="../../themes/officeimo/partials/generated/benchmarks-excel.html" >}}
    <details class="imo-benchmark-snapshots">
      <summary>
        <span><strong>Focused Excel and CSV snapshots</strong><small>Open the compact scenario cards derived from older committed artifacts.</small></span>
        <span aria-hidden="true">Open snapshots</span>
      </summary>
      {{< include path="../../themes/officeimo/partials/generated/benchmarks-overview.html" >}}
    </details>
  </section>

  <section class="imo-benchmark-method" aria-labelledby="benchmark-method-title">
    <p class="imo-benchmark-eyebrow">Read results responsibly</p>
    <h2 id="benchmark-method-title">Reproduce before you decide.</h2>
    <p>Hardware, runtime, data shape, enabled features, and output validation all affect results. Use these snapshots to choose the right suite, then run its documented scenario on infrastructure close to your own.</p>
    <p><a class="imo-action-link" href="/docs/capabilities/benchmarks/">Open the benchmark reproduction guide <span aria-hidden="true">→</span></a></p>
  </section>
</div>
