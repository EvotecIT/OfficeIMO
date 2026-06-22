---
title: "Conversion Playground"
description: "Run selected OfficeIMO document conversions locally in the browser with the planned static WebAssembly app."
layout: page
---

OfficeIMO is preparing a static browser conversion playground for DOCX, XLSX, and PPTX files. The app is designed for GitHub Pages-style hosting: files stay in the browser, conversion calls OfficeIMO byte and stream APIs, and generated output is downloaded locally.

The current browser proof validates DOCX, XLSX, and PPTX to PDF conversion in Blazor WebAssembly for representative fixtures. Rich Word files still need the Unicode font embedding path before the playground should be treated as production-grade for every document.

The static app mount for the future published WebAssembly app is:

```text
/apps/officeimo-converter/
```

Implementation details are tracked in the [browser playground docs](/docs/converters/browser-playground/). The static app mount currently exposes a [machine-readable manifest](/apps/officeimo-converter/manifest.json) until the Blazor WebAssembly app is published there.
