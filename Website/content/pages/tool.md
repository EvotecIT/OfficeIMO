---
title: "OfficeIMO.Tool"
description: "Convert, extract, and inspect documents from your terminal with OfficeIMO.Tool. Install the .NET tool from NuGet and automate repeatable document jobs."
meta.seo_title: "OfficeIMO.Tool: document conversion and inspection CLI"
layout: application
meta.application: "tool"
---

## Install once. Use the same command in your scripts.

Install the tool from NuGet with the .NET SDK, then inspect the available commands:

```powershell
dotnet tool install --global OfficeIMO.Tool
officeimo --version
officeimo help
```

For a repository-pinned installation, use a .NET tool manifest and commit it with your project. Update a global installation with `dotnet tool update --global OfficeIMO.Tool`.

## Start with a document

Convert a Word document to PDF:

```powershell
officeimo convert report.docx report.pdf
```

Extract document content as Markdown:

```powershell
officeimo convert report.pdf report.md
```

Discover the format handlers available in your installed version:

```powershell
officeimo reader capabilities
```

Conversion refuses to replace an existing destination unless you explicitly pass `--force`. Output formats have different fidelity boundaries: Markdown is a semantic projection, while PDF preserves a fixed page layout.

## More than conversion

The command areas cover document reading, tabular extraction, HTML/PDF workflows, provenance inspection, PDF redaction, page-image export, document assembly, and print planning. The same executable offers bounded agent operations and an MCP server over standard input/output.

```powershell
officeimo mcp serve --stdio
```

Use the [command guide](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Tool#readme) for supported arguments and optional dependencies. OCR requires an explicitly configured provider; the default tool does not include an OCR runtime or model. Studio features are not automatically CLI commands.

## Portable downloads and package managers

Install from NuGet today. Portable binaries and WinGet are planned; available channels will be listed on the [downloads page](/downloads/).

Want to work visually instead? Explore [OfficeIMO Studio](/studio/). For an embedded API, choose a focused [NuGet library](/downloads/#downloads-packages).
