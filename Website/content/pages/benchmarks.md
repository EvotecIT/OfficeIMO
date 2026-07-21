---
title: "Benchmarks"
description: "Reproducible performance evidence for OfficeIMO document and data workloads."
layout: page
meta.raw_html: true
---

<div class="imo-benchmark-hub">
  <section class="imo-benchmark-hub__hero">
    <p class="imo-benchmark-eyebrow">Performance evidence</p>
    <h2>Use the benchmark that matches your workload.</h2>
    <p>Office documents, delimited data, extraction, and rendering have different cost profiles. We publish format-specific evidence where the repository has a repeatable, validated suite and say plainly when a public comparison is not ready.</p>
    <div class="imo-benchmark-principles" aria-label="Benchmark principles">
      <span>Equivalent work</span>
      <span>Validated output</span>
      <span>Committed artifacts</span>
      <span>Machine context retained</span>
    </div>
  </section>

  <section class="imo-benchmark-coverage" aria-labelledby="benchmark-coverage-title">
    <div class="imo-benchmark-section-heading">
      <p class="imo-benchmark-eyebrow">Coverage by family</p>
      <h2 id="benchmark-coverage-title">What is measured today</h2>
      <p>A comparison measures equivalent libraries. A regression suite protects OfficeIMO against its own baseline. We keep those claims separate.</p>
    </div>
    <div class="imo-benchmark-coverage__grid">
      <article data-family="excel"><span>Published comparison</span><h3>Excel</h3><p>25,000-row create, write, and read scenarios, plus a detailed engineering matrix across spreadsheet libraries.</p><a href="#excel-evidence">See Excel evidence</a></article>
      <article data-family="csv"><span>Published comparison</span><h3>CSV</h3><p>Wide reads and three write contracts, with field traversal and semantic output validation included.</p><a href="#csv-evidence">See CSV evidence</a></article>
      <article data-family="reader"><span>Regression baseline</span><h3>Reader</h3><p>25 cases across 14 document formats, with detection, chunking, and transport lanes. Timings are a local regression baseline, not a cross-machine promise.</p><a href="https://github.com/EvotecIT/OfficeIMO/blob/main/Docs/benchmarks/officeimo.reader.foundation-2026-07-10.md" target="_blank" rel="noopener">Inspect Reader baseline</a></article>
      <article data-family="guardrails"><span>Performance guardrails</span><h3>PDF, RTF, and Email</h3><p>Budget and regression checks cover representative rendering, parsing, memory, and I/O behavior without presenting unrelated engines as equivalent competitors.</p><a href="https://github.com/EvotecIT/OfficeIMO/blob/main/Docs/officeimo.email-performance.md" target="_blank" rel="noopener">Inspect an evidence note</a></article>
      <article data-family="formats"><span>Repeatable suites</span><h3>Markdown, HTML, and open formats</h3><p>Dedicated projects exercise Markdown, HTML, OneNote, OpenDocument, and drawing workloads. Public comparison snapshots will appear here only after their contracts and artifacts are stable.</p><a href="https://github.com/EvotecIT/OfficeIMO/tree/main" target="_blank" rel="noopener">Browse benchmark projects</a></article>
      <article data-family="planned"><span>Public baseline planned</span><h3>Word and PowerPoint</h3><p>The libraries have performance-focused tests, but no publication-grade comparison suite is committed yet. We do not substitute guessed timings or unrelated micro-tests.</p><a href="https://github.com/EvotecIT/OfficeIMO/issues" target="_blank" rel="noopener">Discuss a workload</a></article>
    </div>
  </section>

  {{< include path="../../themes/officeimo/partials/generated/benchmarks-overview.html" >}}

  <details class="imo-benchmark-explorer" id="excel-matrix">
    <summary>
      <span><strong>Explore the full Excel matrix</strong><small>Filter and sort every committed scenario when you need engineering-level detail.</small></span>
      <span aria-hidden="true">Open explorer</span>
    </summary>
    {{< include path="../../themes/officeimo/partials/generated/benchmarks-excel.html" >}}
  </details>

  <section class="imo-benchmark-method" aria-labelledby="benchmark-method-title">
    <p class="imo-benchmark-eyebrow">Read results responsibly</p>
    <h2 id="benchmark-method-title">Reproduce before you decide.</h2>
    <p>Hardware, runtime, data shape, enabled features, and output validation all affect results. Use these snapshots to choose the right suite, then run its documented scenario on infrastructure close to your own.</p>
  </section>
</div>
