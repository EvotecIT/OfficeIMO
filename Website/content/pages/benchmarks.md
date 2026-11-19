---
title: "Benchmarks"
description: "Benchmarks comparing OfficeIMO with other .NET libraries for PDF, Excel, CSV, Word and more document formats, with separate Windows, Linux and macOS runs."
meta.seo_title: "OfficeIMO benchmarks for .NET document workflows"
layout: page
meta.raw_html: true
meta.content_engine: scriban
---

<div class="imo-benchmark-hub imo-bench-overview">
  <p class="imo-bench-lead">A comparison runs OfficeIMO and other .NET libraries on the same validated work, on one machine and one operating system at a time; a few suites measure OfficeIMO alone and are listed without being published. A result appears here only when its run meets the publishing rules below; until then a comparison shows that it has not been measured or published yet. Open an area to see every comparison, including the ones where another library is faster.</p>

  {{ include "powerforge/benchmarks/index" data.benchmarks }}

  {{ include "benchmark-method" }}
</div>
