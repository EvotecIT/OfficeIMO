---
title: "Conversion Playground"
description: "Convert DOCX, XLSX, and PPTX files to PDF locally in the browser with OfficeIMO WebAssembly."
layout: playground
---

<section class="ocx-launcher" data-ocx-launcher aria-label="OfficeIMO Browser Converter">
  <div class="ocx-container">
    <section class="ocx-launcher-hero">
      <div>
        <p class="ocx-eyebrow">OfficeIMO WebAssembly</p>
        <h1 class="ocx-title">Convert Office files to PDF in your browser.</h1>
        <p class="ocx-lede">Run the OfficeIMO converter locally for DOCX, XLSX, and PPTX files. The WebAssembly app loads only when you start it, so this page stays lightweight until you need the engine.</p>
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
          <span class="ocx-status-value">DOCX XLSX PPTX</span>
        </div>
      </div>
    </section>
    <section class="ocx-launcher-grid" aria-label="Conversion capabilities">
      <div class="ocx-matrix-item">
        <strong>Word to PDF</strong>
        <span>Simple DOCX files convert in the browser. Rich Word files can still expose Unicode font embedding gaps.</span>
      </div>
      <div class="ocx-matrix-item">
        <strong>Excel to PDF</strong>
        <span>Workbook streams flow through OfficeIMO.Excel.Pdf and download as local PDF output.</span>
      </div>
      <div class="ocx-matrix-item">
        <strong>PowerPoint to PDF</strong>
        <span>Presentation streams flow through OfficeIMO.PowerPoint.Pdf with no Office or server process.</span>
      </div>
    </section>
    <div class="ocx-embed-host" data-ocx-frame-host hidden></div>
    <noscript>
      <p class="ocx-noscript"><a href="/apps/officeimo-converter/">Open the converter app</a></p>
    </noscript>
  </div>
</section>
