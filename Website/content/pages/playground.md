---
title: "Conversion Playground"
description: "Choose OfficeIMO conversion routes immediately, then load the WebAssembly engine only when a route is used."
layout: playground
meta.head_html: '<link rel="alternate" hreflang="en" href="https://officeimo.com/playground/" /><link rel="alternate" hreflang="x-default" href="https://officeimo.com/playground/" />'
---

<section class="ocx-launcher ocx-launcher--routes" data-ocx-launcher aria-label="OfficeIMO Browser Converter">
<div class="ocx-container">
<section class="ocx-route-hero">
<div>
<p class="ocx-eyebrow">OfficeIMO WebAssembly</p>
<h1 class="ocx-title">Choose a conversion route.</h1>
<p class="ocx-lede">Pick the format path first. The page shows the OfficeIMO conversion surface immediately, then loads the browser engine only when you run a live route.</p>
</div>
<div class="ocx-status-strip" aria-label="Playground status">
<div class="ocx-status-card"><span class="ocx-status-label">Engine load</span><span class="ocx-status-value">On convert</span></div>
<div class="ocx-status-card"><span class="ocx-status-label">Privacy</span><span class="ocx-status-value">Local only</span></div>
<div class="ocx-status-card"><span class="ocx-status-label">Live</span><span class="ocx-status-value">6 routes</span></div>
</div>
</section>
<section class="ocx-route-board" aria-label="Live conversion routes">
<article class="ocx-route-card ocx-route-card--primary">
<span class="ocx-chip ocx-chip--good">Live browser</span>
<h2>Office to PDF</h2>
<p>Convert DOCX, XLSX, or PPTX streams locally with the OfficeIMO PDF engines.</p>
<div class="ocx-route-formats"><span>DOCX</span><span>XLSX</span><span>PPTX</span><span>PDF</span></div>
<button class="ocx-button ocx-button--primary" type="button" data-ocx-start data-ocx-route="office-pdf">Use Office to PDF</button>
</article>
<article class="ocx-route-card">
<span class="ocx-chip ocx-chip--good">Live browser</span>
<h2>Markdown to HTML</h2>
<p>Render Markdown to an HTML preview and downloadable HTML output.</p>
<div class="ocx-route-formats"><span>MD</span><span>HTML</span></div>
<button class="ocx-button" type="button" data-ocx-start data-ocx-route="markdown-html">Use Markdown to HTML</button>
</article>
<article class="ocx-route-card">
<span class="ocx-chip ocx-chip--good">Live browser</span>
<h2>HTML to Markdown</h2>
<p>Convert HTML into portable Markdown for docs, content pipelines, and cleanup flows.</p>
<div class="ocx-route-formats"><span>HTML</span><span>MD</span></div>
<button class="ocx-button" type="button" data-ocx-start data-ocx-route="html-markdown">Use HTML to Markdown</button>
</article>
<article class="ocx-route-card">
<span class="ocx-chip ocx-chip--good">Live browser</span>
<h2>Markdown to DOCX</h2>
<p>Generate a Word document package from Markdown without a server round trip.</p>
<div class="ocx-route-formats"><span>MD</span><span>DOCX</span></div>
<button class="ocx-button" type="button" data-ocx-start data-ocx-route="markdown-docx">Use Markdown to DOCX</button>
</article>
</section>
<section class="ocx-route-backlog" aria-label="Planned conversion routes">
<div class="ocx-section-head">
<p class="ocx-eyebrow">More engine routes</p>
<h2>Next browser, MCP, skill, plugin, and server paths.</h2>
<p>These routes are visible now so the playground feels like a conversion map, not a hidden demo. They should become live as each browser-safe or tool-backed path is validated.</p>
</div>
<div class="ocx-showcase-grid">
<article class="ocx-capability-card"><span class="ocx-chip ocx-chip--warn">Next browser route</span><h3>DOCX to Markdown</h3><p>Use OfficeIMO.Word.Markdown with asset sidecars and Markdown download handling.</p><code>WordDocument.Load(stream).ToMarkdown(options)</code></article>
<article class="ocx-capability-card"><span class="ocx-chip ocx-chip--warn">Next browser route</span><h3>DOCX to HTML</h3><p>Route Word content through the Markdown bridge and render a browser preview.</p><code>WordDocument.Load(stream).ToHtmlViaMarkdown(options)</code></article>
<article class="ocx-capability-card"><span class="ocx-chip ocx-chip--warn">Next browser route</span><h3>HTML to DOCX</h3><p>Complete the web-content bridge after style and image handling are validated.</p><code>HtmlConversionDocument.Parse(html).ToWordDocument().ToBytes()</code></article>
<article class="ocx-capability-card"><span class="ocx-chip ocx-chip--warn">Next browser route</span><h3>Markdown to PDF</h3><p>Use the Markdown PDF converter directly and keep the raw source explicit.</p><code>markdown.ToPdfFromMarkdown(options)</code></article>
<article class="ocx-capability-card"><span class="ocx-chip ocx-chip--cold">Engine route</span><h3>Excel to CSV / JSON / HTML</h3><p>Preview sheets, extract tables, export records, and produce HTML table output.</p><code>ExcelDocument.Load(stream).Worksheets</code></article>
<article class="ocx-capability-card"><span class="ocx-chip ocx-chip--cold">Engine route</span><h3>CSV / JSON to Excel</h3><p>Build workbooks from structured data with schema validation and cleanup options.</p><code>CsvDocument / ExcelDocument</code></article>
<article class="ocx-capability-card"><span class="ocx-chip ocx-chip--cold">Engine route</span><h3>Reader extraction</h3><p>Extract Markdown, JSON, chunks, tables, and assets from document families.</p><code>reader.Read(...)</code></article>
<article class="ocx-capability-card"><span class="ocx-chip ocx-chip--cold">MCP / skill route</span><h3>Agent conversion tools</h3><p>Expose repeatable conversions and diagnostics through MCP tools, Codex skills, and plugins.</p><code>OfficeIMO conversion tools</code></article>
</div>
</section>
<div class="ocx-engine-dock" data-ocx-frame-host hidden>
<div class="ocx-engine-dock-head">
<div>
<p class="ocx-eyebrow">Loaded engine</p>
<h2 data-ocx-loaded-title>OfficeIMO converter</h2>
</div>
</div>
</div>
<noscript><p class="ocx-noscript"><a href="/apps/officeimo-converter/">Open the converter app</a></p></noscript>
</div>
</section>
