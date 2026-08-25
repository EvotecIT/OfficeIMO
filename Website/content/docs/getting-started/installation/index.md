---
title: Installation
description: How to install the OfficeIMO command-line tool, .NET packages, or PowerShell module from their public package feeds.
order: 1
---

Released OfficeIMO .NET packages are distributed through [NuGet.org](https://www.nuget.org/profiles/EvotecIT). The PowerShell module is distributed through the [PowerShell Gallery](https://www.powershellgallery.com/packages/PSWriteOffice).

Install only the packages an application needs and keep coordinated OfficeIMO package references on the same version. The commands below select the current stable release from NuGet.

## OfficeIMO.Tool

Install the command-line tool globally when `officeimo` should be available from any directory:

```powershell
dotnet tool install --global OfficeIMO.Tool
officeimo --version
```

For a repository-local installation, create or reuse a .NET tool manifest:

```powershell
dotnet new tool-manifest
dotnet tool install OfficeIMO.Tool
dotnet tool run officeimo help
```

The SDK also resolves the local tool through this shorthand:

```powershell
dotnet officeimo help
```

The [OfficeIMO.Tool guide](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Tool) covers conversion, extraction, inspection, and MCP workflows.

## .NET Packages

### OfficeIMO.Word

The core Word document library. Create, read, and modify `.docx` files.

**.NET CLI**

```bash
dotnet add package OfficeIMO.Word
```

**Package Manager Console**

```powershell
Install-Package OfficeIMO.Word
```

### OfficeIMO.Excel

Create and manipulate Excel `.xlsx` workbooks.

**.NET CLI**

```bash
dotnet add package OfficeIMO.Excel
```

**Package Manager Console**

```powershell
Install-Package OfficeIMO.Excel
```

### OfficeIMO.Markdown

Fluent Markdown builder, typed reader/AST, and HTML renderer. Zero external dependencies.

**.NET CLI**

```bash
dotnet add package OfficeIMO.Markdown
```

**Package Manager Console**

```powershell
Install-Package OfficeIMO.Markdown
```

### OfficeIMO.CSV

Strongly-typed CSV document model with validation and streaming.

**.NET CLI**

```bash
dotnet add package OfficeIMO.CSV
```

**Package Manager Console**

```powershell
Install-Package OfficeIMO.CSV
```

### OfficeIMO.Word.Html

Bidirectional Word-to-HTML conversion powered by AngleSharp.

**.NET CLI**

```bash
dotnet add package OfficeIMO.Word.Html
```

### OfficeIMO.Word.Markdown

Bidirectional Word-to-Markdown conversion built on OfficeIMO.Markdown.

**.NET CLI**

```bash
dotnet add package OfficeIMO.Word.Markdown
```

### OfficeIMO.Word.Pdf

Word-to-PDF conversion built on the first-party OfficeIMO.Pdf engine.

**.NET CLI**

```bash
dotnet add package OfficeIMO.Word.Pdf
```

### OfficeIMO.Excel.Pdf

Excel workbook-to-PDF conversion built on the first-party OfficeIMO.Pdf engine.

**.NET CLI**

```bash
dotnet add package OfficeIMO.Excel.Pdf
```

### OfficeIMO.Pdf

Direct PDF generation, reading, editing, rendering, and signature workflows.

**.NET CLI**

```bash
dotnet add package OfficeIMO.Pdf
```

## PSWriteOffice (PowerShell Module)

PSWriteOffice wraps OfficeIMO for use from PowerShell. Install it from the PowerShell Gallery:

```powershell
Install-Module -Name PSWriteOffice -Scope CurrentUser
```

To install for all users (requires elevation):

```powershell
Install-Module -Name PSWriteOffice -Scope AllUsers
```

Update to the latest version:

```powershell
Update-Module -Name PSWriteOffice
```

## Verifying Installation

After installing a .NET package, verify it builds correctly:

```bash
dotnet build
```

For PSWriteOffice, verify the module loads:

```powershell
Import-Module PSWriteOffice
Get-Module PSWriteOffice
```

## Dependencies

OfficeIMO.Word and OfficeIMO.Excel depend on:

- **DocumentFormat.OpenXml** (`[3.5.1, 4.0.0)`) — The Microsoft Open XML SDK.
- **OfficeIMO.Core** — The zero-dependency shared lifecycle, package-safety, data-mapping, color, image, and drawing foundation used by the document packages.

OfficeIMO.Word and OfficeIMO.Excel also use a compatibility helper on older targets:

- **Microsoft.Bcl.AsyncInterfaces** (`10.0.9`) — Async interface compatibility for `netstandard2.0` and `net472`.

OfficeIMO.Excel additionally uses **System.Text.Json** (`[10.0.7,11.0.0)`) for JSON support on `netstandard2.0` and `net472`.

OfficeIMO.Word.Html uses the first-party OfficeIMO.Html package, which depends on:

- **AngleSharp** (`1.5.2`) — HTML parsing and DOM manipulation.
- **AngleSharp.Css** (`1.0.0-beta.216`) — CSS parsing for style mapping.

OfficeIMO.Pdf depends on the first-party OfficeIMO.Core and OfficeIMO.Security packages. OfficeIMO.Security brings **BouncyCastle.Cryptography** (`[2.6.2,3.0.0)`) for cryptographic and signature support.

OfficeIMO.CSV uses **System.Buffers** (`4.5.1`) on `netstandard2.0` and .NET Framework compatibility targets. OfficeIMO.Markdown has no third-party runtime dependency beyond the .NET runtime and first-party OfficeIMO.Core package.
