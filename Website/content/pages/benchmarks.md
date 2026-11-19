---
title: "Benchmarks"
description: "OfficeIMO performance compared with other .NET libraries across PDF, Excel, CSV, Word and other formats, measured separately on Windows, Linux and macOS."
meta.seo_title: "OfficeIMO benchmarks for .NET document workflows"
layout: page
meta.raw_html: true
meta.content_engine: scriban
---

<div class="imo-benchmark-hub imo-bench-overview">
  <p class="imo-bench-lead">Every comparison runs OfficeIMO and other .NET libraries on the same validated work, on one machine and one operating system at a time. A result appears here only when its run meets the publishing rules below; until then a comparison shows that it has not been measured or published yet. Open an area to see every comparison, including the ones where another library is faster.</p>

  {{ include "powerforge/benchmarks/index" data.benchmarks }}

  {{ include "benchmark-method" }}
</div>
