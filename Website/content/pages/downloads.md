---
title: "Downloads"
description: "Download OfficeIMO NuGet packages and PSWriteOffice PowerShell module."
layout: page
---

Use NuGet or the PowerShell Gallery for normal installs, and GitHub release archives when you want packaged artifacts for a specific release.

## Latest Stable Archives

{{< release-button placement="downloads.latest_word" >}}
{{< release-button placement="downloads.latest_excel" >}}
{{< release-button placement="downloads.latest_powerpoint" >}}
{{< release-button placement="downloads.latest_markdown" >}}

## Converter, Renderer, and PDF Packages

{{< release-button placement="downloads.latest_word_html" >}}
{{< release-button placement="downloads.latest_word_markdown" >}}
{{< release-button placement="downloads.latest_word_pdf" >}}
{{< release-button placement="downloads.latest_markdown_html" >}}
{{< release-button placement="downloads.latest_pdf" >}}
{{< release-button placement="downloads.latest_renderer" >}}
{{< release-button placement="downloads.latest_renderer_wpf" >}}

## NuGet Packages (.NET)

Install any OfficeIMO package via the .NET CLI or NuGet Package Manager:

| Package | Install Command | Package Page |
|---------|----------------|--------------|
| OfficeIMO.Word | `dotnet add package OfficeIMO.Word` | [OfficeIMO.Word on NuGet](https://www.nuget.org/packages/OfficeIMO.Word) |
| OfficeIMO.Excel | `dotnet add package OfficeIMO.Excel` | [OfficeIMO.Excel on NuGet](https://www.nuget.org/packages/OfficeIMO.Excel) |
| OfficeIMO.PowerPoint | `dotnet add package OfficeIMO.PowerPoint` | [OfficeIMO.PowerPoint on NuGet](https://www.nuget.org/packages/OfficeIMO.PowerPoint) |
| OfficeIMO.Markdown | `dotnet add package OfficeIMO.Markdown` | [OfficeIMO.Markdown on NuGet](https://www.nuget.org/packages/OfficeIMO.Markdown) |
| OfficeIMO.CSV | `dotnet add package OfficeIMO.CSV` | [OfficeIMO.CSV on NuGet](https://www.nuget.org/packages/OfficeIMO.CSV) |
| OfficeIMO.Visio | `dotnet add package OfficeIMO.Visio` | [OfficeIMO.Visio on NuGet](https://www.nuget.org/packages/OfficeIMO.Visio) |
| OfficeIMO.Reader | `dotnet add package OfficeIMO.Reader` | [OfficeIMO.Reader on NuGet](https://www.nuget.org/packages/OfficeIMO.Reader) |

### Converters

| Package | Install Command | Package Page |
|---------|----------------|--------------|
| OfficeIMO.Word.Html | `dotnet add package OfficeIMO.Word.Html` | [OfficeIMO.Word.Html on NuGet](https://www.nuget.org/packages/OfficeIMO.Word.Html) |
| OfficeIMO.Word.Markdown | `dotnet add package OfficeIMO.Word.Markdown` | [OfficeIMO.Word.Markdown on NuGet](https://www.nuget.org/packages/OfficeIMO.Word.Markdown) |
| OfficeIMO.Word.Pdf | `dotnet add package OfficeIMO.Word.Pdf` | [OfficeIMO.Word.Pdf on NuGet](https://www.nuget.org/packages/OfficeIMO.Word.Pdf) |
| OfficeIMO.Markdown.Html | `dotnet add package OfficeIMO.Markdown.Html` | [OfficeIMO.Markdown.Html on NuGet](https://www.nuget.org/packages/OfficeIMO.Markdown.Html) |
| OfficeIMO.Pdf | `dotnet add package OfficeIMO.Pdf` | [OfficeIMO.Pdf on NuGet](https://www.nuget.org/packages/OfficeIMO.Pdf) |
| OfficeIMO.MarkdownRenderer | `dotnet add package OfficeIMO.MarkdownRenderer` | [OfficeIMO.MarkdownRenderer on NuGet](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer) |
| OfficeIMO.MarkdownRenderer.Wpf | `dotnet add package OfficeIMO.MarkdownRenderer.Wpf` | [OfficeIMO.MarkdownRenderer.Wpf on NuGet](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer.Wpf) |

## PowerShell Module

Install PSWriteOffice from the PowerShell Gallery:

```powershell
Install-Module PSWriteOffice -Force
```

Or for the current user only:

```powershell
Install-Module PSWriteOffice -Scope CurrentUser
```

[View on PowerShell Gallery](https://www.powershellgallery.com/packages/PSWriteOffice)

## GitHub Releases

The matrix below is built from the latest stable GitHub release assets so you can grab exact package archives without browsing the repository manually.

{{< release-buttons placement="downloads.current_release" >}}
