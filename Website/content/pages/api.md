---
title: "API Reference"
description: "Browse the generated API reference for OfficeIMO .NET libraries and the PSWriteOffice PowerShell module."
layout: page
---

## Reference Areas

The OfficeIMO website ships two API surfaces:

- .NET library reference generated from the compiled assemblies.
- PowerShell cmdlet reference for **PSWriteOffice**.

<div class="imo-api-overview">
  <section class="imo-api-overview__panel">
    <div>
      <span class="imo-api-overview__eyebrow">Choose your starting point</span>
      <h2 class="imo-api-overview__title">Reference pages when you know the package, guides when you need the workflow</h2>
      <p class="imo-api-overview__copy">Start with the API when you already know the library you want to call. If you are still shaping the workflow, open the guide first and then jump into the matching reference pages.</p>
    </div>
    <div class="imo-api-overview__actions">
      <a href="/docs/" class="imo-btn imo-btn-primary">Open Documentation</a>
      <a href="/downloads/" class="imo-btn imo-btn-ghost">View Packages and Downloads</a>
    </div>
  </section>

  <section class="imo-api-overview__group" aria-labelledby="api-group-dotnet">
    <div class="imo-api-overview__group-header">
      <div>
        <span class="imo-api-overview__eyebrow">.NET libraries</span>
        <h2 id="api-group-dotnet">Open the package reference that matches your build target</h2>
      </div>
      <p>These pages are generated from the current OfficeIMO assemblies and XML documentation during website CI.</p>
    </div>
    <div class="imo-api-card-grid">
      <article class="imo-api-card" style="--api-accent: var(--imo-word);">
        <div class="imo-api-card__header">
          <h3>OfficeIMO.Word</h3>
          <span class="imo-badge">Core builder</span>
        </div>
        <p class="imo-api-card__desc">Word document creation, editing, formatting, bookmarks, tables, images, charts, and mail-merge style composition.</p>
        <p class="imo-api-card__best">Best for: document generation services, reports, and template-free Word authoring.</p>
        <div class="imo-api-card__links">
          <a href="/api/word/" class="imo-api-card__primary">Open Word API Reference</a>
          <a href="/docs/word/" class="imo-api-card__secondary">Word guides</a>
        </div>
      </article>
      <article class="imo-api-card" style="--api-accent: var(--imo-excel);">
        <div class="imo-api-card__header">
          <h3>OfficeIMO.Excel</h3>
          <span class="imo-badge">Reports</span>
        </div>
        <p class="imo-api-card__desc">Workbook generation, worksheets, tables, ranges, validation, charts, and extraction helpers for spreadsheet-heavy workflows.</p>
        <p class="imo-api-card__best">Best for: exports, tabular reporting, data validation, and chart-driven workbooks.</p>
        <div class="imo-api-card__links">
          <a href="/api/excel/" class="imo-api-card__primary">Open Excel API Reference</a>
          <a href="/docs/excel/" class="imo-api-card__secondary">Excel guides</a>
        </div>
      </article>
      <article class="imo-api-card" style="--api-accent: var(--imo-powerpoint);">
        <div class="imo-api-card__header">
          <h3>OfficeIMO.PowerPoint</h3>
          <span class="imo-badge">Slides</span>
        </div>
        <p class="imo-api-card__desc">Slides, layouts, themes, text frames, charts, tables, shapes, and presentation composition helpers.</p>
        <p class="imo-api-card__best">Best for: decks assembled from data, status updates, and repeatable presentation generation.</p>
        <div class="imo-api-card__links">
          <a href="/api/powerpoint/" class="imo-api-card__primary">Open PowerPoint API Reference</a>
          <a href="/docs/powerpoint/" class="imo-api-card__secondary">PowerPoint guides</a>
        </div>
      </article>
      <article class="imo-api-card" style="--api-accent: var(--imo-markdown);">
        <div class="imo-api-card__header">
          <h3>OfficeIMO.Markdown</h3>
          <span class="imo-badge">AST and rendering</span>
        </div>
        <p class="imo-api-card__desc">Markdown builder, parser, abstract syntax tree, HTML rendering, transforms, and host-friendly document tooling.</p>
        <p class="imo-api-card__best">Best for: markdown processing pipelines, renderer hosts, and structured content tooling.</p>
        <div class="imo-api-card__links">
          <a href="/api/markdown/" class="imo-api-card__primary">Open Markdown API Reference</a>
          <a href="/docs/markdown/" class="imo-api-card__secondary">Markdown guides</a>
        </div>
      </article>
      <article class="imo-api-card" style="--api-accent: var(--imo-csv);">
        <div class="imo-api-card__header">
          <h3>OfficeIMO.CSV</h3>
          <span class="imo-badge">Schemas</span>
        </div>
        <p class="imo-api-card__desc">CSV schema definition, typed mapping, validation, headers, formatting policies, and streaming workflows.</p>
        <p class="imo-api-card__best">Best for: stable import and export pipelines where CSV shape matters as much as the values.</p>
        <div class="imo-api-card__links">
          <a href="/api/csv/" class="imo-api-card__primary">Open CSV API Reference</a>
          <a href="/docs/csv/" class="imo-api-card__secondary">CSV guides</a>
        </div>
      </article>
      <article class="imo-api-card" style="--api-accent: var(--imo-visio);">
        <div class="imo-api-card__header">
          <h3>OfficeIMO.Visio</h3>
          <span class="imo-badge">Diagrams</span>
        </div>
        <p class="imo-api-card__desc">Diagram pages, shapes, connectors, masters, positioning helpers, and fluent Visio authoring primitives.</p>
        <p class="imo-api-card__best">Best for: topology maps, process diagrams, and generated documentation visuals.</p>
        <div class="imo-api-card__links">
          <a href="/api/visio/" class="imo-api-card__primary">Open Visio API Reference</a>
          <a href="/docs/visio/" class="imo-api-card__secondary">Visio guides</a>
        </div>
      </article>
      <article class="imo-api-card" style="--api-accent: #8b5cf6;">
        <div class="imo-api-card__header">
          <h3>OfficeIMO.Reader</h3>
          <span class="imo-badge">Ingestion</span>
        </div>
        <p class="imo-api-card__desc">Unified extraction, chunking, and ingestion-oriented document processing across Office and markdown-adjacent inputs.</p>
        <p class="imo-api-card__best">Best for: read-only pipelines, content extraction, chunked processing, and AI-oriented ingestion stages.</p>
        <div class="imo-api-card__links">
          <a href="/api/reader/" class="imo-api-card__primary">Open Reader API Reference</a>
          <a href="/docs/reader/" class="imo-api-card__secondary">Reader guides</a>
        </div>
      </article>
    </div>
  </section>

  <section class="imo-api-overview__group" aria-labelledby="api-group-powershell">
    <div class="imo-api-overview__group-header">
      <div>
        <span class="imo-api-overview__eyebrow">PowerShell automation</span>
        <h2 id="api-group-powershell">Use PSWriteOffice when the workflow belongs in scripts</h2>
      </div>
      <p>The PowerShell reference is generated from synced help XML and example scripts, with checked-in fallback inputs for clean local builds.</p>
    </div>
    <div class="imo-api-card-grid imo-api-card-grid--single">
      <article class="imo-api-card" style="--api-accent: #38bdf8;">
        <div class="imo-api-card__header">
          <h3>PSWriteOffice</h3>
          <span class="imo-badge">PowerShell module</span>
        </div>
        <p class="imo-api-card__desc">Cmdlets and DSL-style helpers for Word, Excel, PowerPoint, Markdown, and CSV workflows on top of OfficeIMO.</p>
        <p class="imo-api-card__best">Best for: automation jobs, scheduled reporting, GitHub Actions, and script-first document pipelines.</p>
        <div class="imo-api-card__links">
          <a href="/api/powershell/" class="imo-api-card__primary">Open Cmdlet Reference</a>
          <a href="/docs/pswriteoffice/" class="imo-api-card__secondary">PowerShell guides</a>
        </div>
      </article>
    </div>
  </section>
</div>

## Snapshot Table

| Area | Focus | Link |
|------|-------|------|
| OfficeIMO.Word | Word document creation, editing, formatting, bookmarks, tables, images, charts, and more. | [Open OfficeIMO.Word API](/api/word/) |
| OfficeIMO.Excel | Workbook generation, worksheets, tables, charts, validation, and extraction helpers. | [Open OfficeIMO.Excel API](/api/excel/) |
| OfficeIMO.PowerPoint | Slides, layouts, themes, transitions, tables, charts, and shape composition. | [Open OfficeIMO.PowerPoint API](/api/powerpoint/) |
| OfficeIMO.Markdown | Markdown builder, parser, AST, HTML rendering, and transforms. | [Open OfficeIMO.Markdown API](/api/markdown/) |
| OfficeIMO.CSV | CSV schema definition, typed mapping, validation, and streaming workflows. | [Open OfficeIMO.CSV API](/api/csv/) |
| OfficeIMO.Visio | Diagram pages, shapes, connectors, and fluent Visio generation helpers. | [Open OfficeIMO.Visio API](/api/visio/) |
| OfficeIMO.Reader | Unified extraction, chunking, and ingestion-oriented document processing. | [Open OfficeIMO.Reader API](/api/reader/) |

## PowerShell Snapshot Table

| Area | Focus | Link |
|------|-------|------|
| PSWriteOffice | PowerShell cmdlets for Word, Excel, PowerPoint, Markdown, and CSV workflows on top of OfficeIMO. | [Open PSWriteOffice Cmdlets](/api/powershell/) |

## How To Use The API Docs

1. Start on the library landing page that matches the package you use.
2. Filter the sidebar by type name, namespace, or kind.
3. Jump into a type page for signatures, summaries, parameters, and source links.
4. Cross-reference the conceptual guides under [Getting Started](/docs/getting-started/) and the package docs under [Documentation](/docs/).

## How These Pages Are Generated

- .NET library reference is generated from the current `OfficeIMO` build outputs during website CI.
- PSWriteOffice reference is generated from synced PowerShell help XML and example scripts, with a checked-in fallback snapshot for local and clean-checkout builds.

## Need A Practical Entry Point?

- [Getting Started](/docs/getting-started/) for installation and first steps.
- [Documentation](/docs/) for task-based guides.
- [Downloads](/downloads/) for NuGet and PowerShell Gallery entry points.
