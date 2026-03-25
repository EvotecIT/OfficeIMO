---
title: "Converting Word to PDF Without Office on Linux"
description: "A practical guide to converting DOCX files to PDF on Linux using OfficeIMO.Word.Pdf, including a Docker-based workflow for CI/CD pipelines."
date: 2025-09-10
tags: [pdf, linux, cross-platform, converters]
categories: [Tutorial]
author: "Przemyslaw Klys"
---

Converting Word documents to PDF is one of the most requested features in any document automation library. Traditionally it required a Windows machine with Microsoft Office or a LibreOffice sidecar. **OfficeIMO.Word.Pdf** provides an in-process conversion path that works on Linux, macOS, and Windows.

## Installation

```bash
dotnet add package OfficeIMO.Word.Pdf
```

The package depends on OfficeIMO.Word together with QuestPDF and SkiaSharp. You still do not need Microsoft Office or LibreOffice, but on Linux containers you should provision fonts so text measurement and output quality stay predictable.

## Basic Conversion

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

// Create or open a Word document
using var doc = WordDocument.Create("Invoice.docx");
doc.AddParagraph("Invoice #10042").Bold = true;
doc.AddParagraph("Date: 2025-09-10");
doc.AddParagraph("Amount Due: $1,250.00");
doc.Save();

// Export the document directly to PDF
doc.SaveAsPdf("Invoice.pdf");
```

That is it. Two lines to go from DOCX to PDF once the document exists.

## Stream-Based Conversion

For web applications and cloud functions, you often work with streams rather than files:

```csharp
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var doc = WordDocument.Create("Invoice.docx");
doc.AddParagraph("Invoice #10042");
doc.Save();

using var docxStream = new MemoryStream();
doc.Save(docxStream);
docxStream.Position = 0;

using var loaded = WordDocument.Load(docxStream);
using var pdfStream = new MemoryStream();
loaded.SaveAsPdf(pdfStream);

// Upload pdfStream to blob storage or return as HTTP response
```

## Running in Docker

Here is a minimal Dockerfile for a PDF conversion microservice:

```dockerfile
FROM mcr.microsoft.com/dotnet/sdk:8.0-alpine AS build
WORKDIR /src
COPY . .
RUN dotnet publish -c Release -o /app

FROM mcr.microsoft.com/dotnet/aspnet:8.0-alpine
WORKDIR /app
COPY --from=build /app .

# Install font packages for decent PDF rendering
RUN apk add --no-cache icu-libs fontconfig ttf-liberation

ENTRYPOINT ["dotnet", "PdfService.dll"]
```

The key detail is the font packages. PDF rendering needs real font files to compute glyph widths and embed subsets. The `ttf-liberation` package provides metrically compatible substitutes for Times New Roman, Arial, and Courier New.

## ASP.NET Core Endpoint

A minimal API endpoint that accepts a DOCX upload and returns a PDF:

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

app.MapPost("/convert", async (IFormFile file) =>
{
    await using var input = new MemoryStream();
    await file.CopyToAsync(input);
    input.Position = 0;

    using var document = WordDocument.Load(input);
    var output = new MemoryStream();
    document.SaveAsPdf(output);
    output.Position = 0;

    return Results.File(output, "application/pdf", "converted.pdf");
});
```

Deploy this behind a load balancer and you have a scalable, stateless conversion service.

## Font Handling Tips

- **Embed fonts in the DOCX when possible.** That reduces surprises, but host fonts still matter for fallback and for documents that reference fonts that are not embedded.
- **Fallback fonts.** If a font is missing and not embedded, the converter substitutes the closest available font. Install `ttf-liberation` or `ttf-dejavu` to minimise mismatches.
- **CJK text.** For Chinese, Japanese, or Korean content, install `font-noto-cjk` on Alpine or `fonts-noto-cjk` on Debian.

## Limitations

The conversion engine handles paragraphs, tables, images, headers, footers, page breaks, and basic styles well. Complex features like SmartArt, embedded OLE objects, and advanced text effects may render with reduced fidelity. For those edge cases, consider Aspose or a LibreOffice sidecar container.

For the vast majority of business documents, invoices, letters, reports, and contracts, OfficeIMO.Word.Pdf produces clean, professional PDFs without a commercial dependency.
