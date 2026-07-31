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

  <section class="imo-benchmark-coverage" aria-labelledby="benchmark-coverage-title">
    <div class="imo-benchmark-section-heading">
      <p class="imo-benchmark-eyebrow">Coverage by family</p>
      <h2 id="benchmark-coverage-title">What is measured today</h2>
      <p>A comparison measures equivalent libraries. A regression suite protects OfficeIMO against its own baseline. We keep those claims separate.</p>
    </div>
    <div class="imo-benchmark-coverage__grid">
      <article data-family="data-readers"><span>Cross-platform comparison</span><h3>CSV and Excel reads and writes</h3><p>Pinned 65K-record read fixtures and the validated 25,000-row CSV IDataReader write contract are measured across capable libraries, separately on Windows, Linux, and macOS. Missing lanes remain visible.</p><a href="#library-comparison-evidence">Select a workload and platform</a></article>
      <article data-family="excel"><span>Focused workstation suite</span><h3>Excel create and write</h3><p>25,000-row create, write, and package scenarios, plus a detailed engineering matrix. These historical snapshots are kept separate from the cross-platform read ranking.</p><a href="#excel-evidence">See Excel evidence</a></article>
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

  <details class="imo-benchmark-explorer" id="historical-benchmark-evidence">
    <summary>
      <span><strong>Open legacy snapshots with incomplete environment provenance</strong><small>These older Excel and CSV artifacts did not record OS and/or run mode, so they are excluded from the current comparison selector.</small></span>
      <span aria-hidden="true">Open history</span>
    </summary>
    {{< include path="../../themes/officeimo/partials/generated/benchmarks-overview.html" >}}
    {{< include path="../../themes/officeimo/partials/generated/benchmarks-excel.html" >}}
  </details>

  <section class="imo-benchmark-method" aria-labelledby="benchmark-method-title">
    <p class="imo-benchmark-eyebrow">Read results responsibly</p>
    <h2 id="benchmark-method-title">Reproduce before you decide.</h2>
    <p>Hardware, runtime, data shape, enabled features, and output validation all affect results. Use these snapshots to choose the right suite, then run its documented scenario on infrastructure close to your own.</p>
    <p><a class="imo-action-link" href="/docs/capabilities/benchmarks/">Open the benchmark reproduction guide <span aria-hidden="true">→</span></a></p>
  </section>
</div>
