---
title: Documentation
description: Guided documentation and reference links for the OfficeIMO package family.
layout: docs-home
order: 1
slug: index
---

Use the guided sections above when you are choosing a workflow or package family. The reference notes below are here to help when you need package feeds, repo-backed source-of-truth pointers, or a reminder of which areas still live more in code than in narrative guides.

## Package Feeds

- Core builders:
  [OfficeIMO.Word on NuGet](https://www.nuget.org/packages/OfficeIMO.Word),
  [OfficeIMO.Excel on NuGet](https://www.nuget.org/packages/OfficeIMO.Excel),
  [OfficeIMO.PowerPoint on NuGet](https://www.nuget.org/packages/OfficeIMO.PowerPoint),
  [OfficeIMO.Markdown on NuGet](https://www.nuget.org/packages/OfficeIMO.Markdown),
  [OfficeIMO.CSV on NuGet](https://www.nuget.org/packages/OfficeIMO.CSV),
  [OfficeIMO.Visio on NuGet](https://www.nuget.org/packages/OfficeIMO.Visio)
- Workflow and conversion packages:
  [OfficeIMO.Reader on NuGet](https://www.nuget.org/packages/OfficeIMO.Reader),
  [OfficeIMO.Word.Html on NuGet](https://www.nuget.org/packages/OfficeIMO.Word.Html),
  [OfficeIMO.Word.Markdown on NuGet](https://www.nuget.org/packages/OfficeIMO.Word.Markdown)
- PowerShell automation:
  [PSWriteOffice on PowerShell Gallery](https://www.powershellgallery.com/packages/PSWriteOffice)

## When the Guides End

This site gives the best narrative coverage to the packages most teams start with:

- `OfficeIMO.Word`
- `OfficeIMO.Excel`
- `OfficeIMO.PowerPoint`
- `OfficeIMO.Markdown`
- `OfficeIMO.CSV`
- `OfficeIMO.Reader`
- `PSWriteOffice`

The repo also includes adjacent packages such as `OfficeIMO.Word.Pdf`, `OfficeIMO.Word.Html`, `OfficeIMO.Word.Markdown`, `OfficeIMO.Visio`, specialized reader extensions, and renderer projects. Some of those are linked from the package feeds below and the API reference, but not all of them have full narrative guides on the website yet. When a package is not fully covered here, the most accurate sources are usually:

- the package README in the repo,
- the generated API reference,
- the example projects under `OfficeIMO.Examples`,
- and the test suite for current behavior.

## Fast Reference

- [Installation](/docs/getting-started/installation) for package/module install commands.
- [Quick Start](/docs/getting-started/quickstart) for a first successful output file.
- [Platform Support](/docs/getting-started/platform-support) for frameworks, OS expectations, and AOT notes.
- [AOT and Trimming](/docs/advanced/aot-trimming) for deployment-sensitive workloads.
- [PSWriteOffice Docs](/docs/pswriteoffice/) for script-first automation.
- [API Reference](/api/) when you already know the package and need type-level details.

## License

OfficeIMO is licensed under the [MIT License](https://github.com/EvotecIT/OfficeIMO/blob/master/LICENSE). Copyright (c) Przemyslaw Klys @ Evotec. If you need to review upstream runtime dependencies such as Open XML SDK, ImageSharp, or QuestPDF, see the [Third-Party Dependencies](/third-party/) page.

## Source Code

The full source code is available on GitHub: [https://github.com/EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
