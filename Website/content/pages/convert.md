---
title: "Browser Document Workspace"
description: "Convert documents and run practical PDF workflows locally in your browser with OfficeIMO WebAssembly, without uploading document bytes."
layout: playground
meta.seo_title: "Browser document converter | OfficeIMO"
---

<section class="imo-converter-launch">
  <div class="imo-container imo-converter-launch__intro">
    <div><p class="imo-converter-launch__eyebrow">Browser tools</p><h1>Convert, organize, and inspect your files.</h1><p>Your files stay in this browser. Choose a conversion, PDF tool, or provenance inspection below.</p></div>
  </div>
  <div class="imo-container imo-converter-launch__frame-shell"><iframe class="imo-converter-launch__frame" data-workspace-src="/apps/officeimo-converter/?embedded=1" title="OfficeIMO browser document workspace" loading="lazy"></iframe></div>
  <section class="imo-container imo-converter-launch__details" aria-labelledby="browser-workspace-details">
    <div><p class="imo-converter-launch__eyebrow">What runs in this tab</p><h2 id="browser-workspace-details">Keep working with your files</h2><p>Convert between twelve document routes, use twelve PDF tools, or inspect and remove supported provenance. Reuse a result in another compatible tool and restore the original files when needed.</p></div>
    <div class="imo-converter-launch__detail-grid">
      <article><strong>Convert</strong><p>Office, HTML, Markdown, and PDF routes return real downloads with fidelity diagnostics.</p></article>
      <article><strong>Organize PDFs</strong><p>Merge, split, extract, delete, reorder, and rotate separate output copies.</p></article>
      <article><strong>Review and secure</strong><p>Inspect, compare, optimize, protect, unlock, and redact with operation evidence.</p></article>
      <article><strong>Inspect provenance</strong><p>Inspect supported images and Office files, choose provenance carriers to remove, and download a re-inspected copy.</p></article>
      <article><strong>Know the boundary</strong><p>OCR, lossy scan compression, and cryptographic signing stay outside the generic browser workflow.</p></article>
    </div>
    <p class="imo-converter-launch__boundary">Each file is limited to 25 MiB. Multi-file PDF workflows accept up to ten files and 75 MiB combined; PDF parsing is capped at 500 pages, split at 100 outputs, and visual comparison at 25 pages. <a href="/docs/converters/browser-playground/">Read the complete browser contract</a>.</p>
  </section>
  <noscript><div class="imo-container"><p><a href="/apps/officeimo-converter/">Open the document workspace</a></p></div></noscript>
</section>
