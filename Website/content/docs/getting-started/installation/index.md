---
title: Installation
description: How to install OfficeIMO packages via NuGet, Package Manager Console, or PowerShell Gallery.
order: 1
---

# Installation

All OfficeIMO .NET packages are published to [NuGet.org](https://www.nuget.org/profiles/EvotecIT). The PowerShell module is published to the [PowerShell Gallery](https://www.powershellgallery.com/packages/PSWriteOffice).

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

**PackageReference (csproj)**

```xml
<PackageReference Include="OfficeIMO.Word" Version="1.0.39" />
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

**PackageReference**

```xml
<PackageReference Include="OfficeIMO.Excel" Version="0.6.19" />
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

**PackageReference**

```xml
<PackageReference Include="OfficeIMO.Markdown" Version="0.6.6" />
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

**PackageReference**

```xml
<PackageReference Include="OfficeIMO.CSV" Version="0.1.19" />
```

### OfficeIMO.Word.Html

Bidirectional Word-to-HTML conversion powered by AngleSharp.

**.NET CLI**

```bash
dotnet add package OfficeIMO.Word.Html
```

**PackageReference**

```xml
<PackageReference Include="OfficeIMO.Word.Html" Version="1.0.13" />
```

### OfficeIMO.Word.Markdown

Bidirectional Word-to-Markdown conversion built on OfficeIMO.Markdown.

**.NET CLI**

```bash
dotnet add package OfficeIMO.Word.Markdown
```

**PackageReference**

```xml
<PackageReference Include="OfficeIMO.Word.Markdown" Version="1.0.13" />
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

- **DocumentFormat.OpenXml** (>= 3.3.0, < 4.0.0) -- The Microsoft Open XML SDK.
- **SixLabors.ImageSharp** (2.1.11) -- Cross-platform image processing for image insertion and measurement.

OfficeIMO.Word.Html additionally depends on:

- **AngleSharp** (1.3.0) -- HTML parsing and DOM manipulation.
- **AngleSharp.Css** (1.0.0-beta.157) -- CSS parsing for style mapping.

OfficeIMO.Markdown and OfficeIMO.CSV have **no external dependencies** beyond the .NET runtime.
