---
title: "Conversion Playground"
description: "Run selected OfficeIMO document conversions locally in the browser with WebAssembly."
layout: playground
---

<section class="ocx-launcher" data-ocx-launcher aria-label="OfficeIMO Browser Converter">
  <div class="ocx-container">
    <section class="ocx-launcher-hero">
      <div>
        <p class="ocx-eyebrow">OfficeIMO WebAssembly</p>
        <h1 class="ocx-title">Show the OfficeIMO conversion engine in the browser.</h1>
        <p class="ocx-lede">Run selected OfficeIMO engines locally: DOCX, XLSX, and PPTX to PDF; Markdown to HTML; HTML to Markdown; and Markdown to DOCX. The app also maps the wider engine surface: Word to Markdown, Word to HTML, Excel to CSV/JSON/HTML-style outputs, Reader extraction, PDF operations, Markup exporters, MCP tools, skills, and server workflows.</p>
        <div class="ocx-hero-actions">
          <button class="ocx-button ocx-button--primary" type="button" data-ocx-start>Start converter</button>
          <a class="ocx-button" href="/apps/officeimo-converter/" data-ocx-direct>Open full app</a>
        </div>
      </div>
      <div class="ocx-status-strip" aria-label="Playground status">
        <div class="ocx-status-card">
          <span class="ocx-status-label">Runtime</span>
          <span class="ocx-status-value">Loaded on demand</span>
        </div>
        <div class="ocx-status-card">
          <span class="ocx-status-label">Privacy</span>
          <span class="ocx-status-value">Local only</span>
        </div>
        <div class="ocx-status-card">
          <span class="ocx-status-label">Formats</span>
          <span class="ocx-status-value">PDF MD HTML DOCX</span>
        </div>
      </div>
    </section>
    <section class="ocx-launcher-grid" aria-label="Conversion capabilities">
      <div class="ocx-matrix-item">
        <strong>Office to PDF</strong>
        <span>DOCX, XLSX, and PPTX streams convert locally with OfficeIMO.Word.Pdf, OfficeIMO.Excel.Pdf, and OfficeIMO.PowerPoint.Pdf.</span>
      </div>
      <div class="ocx-matrix-item">
        <strong>Markdown and HTML</strong>
        <span>Markdown renders to HTML, HTML converts back to Markdown, and Markdown can generate a downloadable DOCX package.</span>
      </div>
      <div class="ocx-matrix-item">
        <strong>Engine roadmap</strong>
        <span>The in-app map separates live browser routes from next browser routes and CLI, PowerShell, MCP, plugin, skill, and server candidates.</span>
      </div>
    </section>
    <div class="ocx-embed-host" data-ocx-frame-host hidden></div>
    <noscript>
      <p class="ocx-noscript"><a href="/apps/officeimo-converter/">Open the converter app</a></p>
    </noscript>
  </div>
</section>
