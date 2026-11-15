---
title: "Benchmarks"
description: "Measured OfficeIMO performance across PDF, Excel, CSV and other formats, compared with other .NET libraries on Windows, Linux and macOS."
meta.seo_title: "OfficeIMO benchmarks for .NET document workflows"
layout: page
meta.raw_html: true
meta.content_engine: scriban
---

<div class="imo-benchmark-hub imo-bench-overview">
  <p class="imo-bench-lead">Every result compares libraries doing the same validated work, on one machine and one operating system, with the source commit and raw data attached. Pick an area to see each workload, including the ones where another library is faster.</p>

  {{ include "benchmark-overview" }}

  <section class="imo-benchmark-method" aria-labelledby="benchmark-method-title">
    <h2 id="benchmark-method-title">Reproduce before you decide.</h2>
    <p>Hardware, runtime, data shape, enabled features and output validation all affect results. Use these results to choose a suite, then run it on infrastructure close to your own.</p>
    <p><a class="imo-action-link" href="/docs/capabilities/benchmarks/">Benchmark reproduction guide <span aria-hidden="true">→</span></a> · <a class="imo-action-link" href="/benchmarks/archive/">Earlier Excel and CSV measurements <span aria-hidden="true">→</span></a></p>
  </section>
</div>
