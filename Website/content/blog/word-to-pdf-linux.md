---
title: "Converting Word to PDF Without Office on Linux"
description: "A practical guide to converting DOCX files to PDF on Linux using OfficeIMO.Word.Pdf, including a Docker-based workflow for CI/CD pipelines."
date: 2025-09-10
tags: [pdf, linux, cross-platform, converters]
categories: [Tutorial]
author: "Przemyslaw Klys"
---

Converting Word documents to PDF is one of the most requested features in any document automation library. Traditionally it required a Windows machine with Microsoft Office or LibreOffice installed. **OfficeIMO.Word.Pdf** changes that by providing a managed-code conversion path that works on Linux, macOS, and Windows.

## Installation

```bash
dotnet add package OfficeIMO.Word.Pdf
```

The package depends on OfficeIMO.Word and a lightweight layout engine. No native binaries to install, no `apt-get` packages to manage.

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

// Convert to PDF
PdfConverter.Convert("Invoice.docx", "Invoice.pdf");
```

That is it. Two lines to go from DOCX to PDF.

## Stream-Based Conversion

For web applications and cloud functions, you often work with streams rather than files:

```csharp
using var docxStream = new MemoryStream();
doc.Save(docxStream);
docxStream.Position = 0;

using var pdfStream = new MemoryStream();
PdfConverter.Convert(docxStream, pdfStream);

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
app.MapPost("/convert", async (IFormFile file) =>
{
    using var input = new MemoryStream();
    await file.CopyToAsync(input);
    input.Position = 0;

    var output = new MemoryStream();
    PdfConverter.Convert(input, output);
    output.Position = 0;

    return Results.File(output, "application/pdf", "converted.pdf");
});
```

Deploy this behind a load balancer and you have a scalable, stateless conversion service.

## Font Handling Tips

- **Embed fonts in the DOCX.** If the source document embeds its fonts, OfficeIMO.Word.Pdf uses them directly, producing pixel-identical output regardless of what is installed on the host.
- **Fallback fonts.** If a font is missing and not embedded, the converter substitutes the closest available font. Install `ttf-liberation` or `ttf-dejavu` to minimise mismatches.
- **CJK text.** For Chinese, Japanese, or Korean content, install `font-noto-cjk` on Alpine or `fonts-noto-cjk` on Debian.

## Limitations

The conversion engine handles paragraphs, tables, images, headers, footers, page breaks, and basic styles well. Complex features like SmartArt, embedded OLE objects, and advanced text effects may render with reduced fidelity. For those edge cases, consider Aspose or a LibreOffice sidecar container.

For the vast majority of business documents, invoices, letters, reports, and contracts, OfficeIMO.Word.Pdf produces clean, professional PDFs without any commercial dependency.
