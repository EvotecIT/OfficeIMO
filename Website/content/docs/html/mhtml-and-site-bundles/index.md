---
title: Read MHTML Web Archives and ZIP Site Bundles
description: Open MHT and MHTML web archives or zipped websites, read their HTML and embedded resources, create new archives, and convert them to PDF or images without network access.
order: 44
search_terms:
  - mhtml
  - mht file
  - web archive
  - mhtml to pdf
  - saved web page
  - zip website to pdf
packages:
  - OfficeIMO.Mhtml
  - OfficeIMO.Html
  - OfficeIMO.Html.Pdf
cmdlets:
  - Read-OfficeDocument
commands:
  - officeimo reader read
  - officeimo html convert
  - officeimo html render
related:
  - /docs/html/html-to-pdf/
  - /docs/html/render-and-convert/
limits: Archives are rendered offline from their saved state; scripts do not run, and missing resources are fetched only when you supply a fetcher through MhtmlRemoteResourcePolicy.
verified: 2026-11-14
---

A saved web page usually arrives in one of two containers. An MHT/MHTML file is a MIME message that holds the page and its images and stylesheets; `OfficeIMO.Mhtml` reads and writes it as an `MhtmlDocument`. A ZIP of a website holds the same files as ordinary entries; `HtmlSiteBundle` in `OfficeIMO.Html` loads it without extracting anything to disk. Both expose the page as an `HtmlConversionDocument`, so every HTML conversion works on them.

```sh
dotnet add package OfficeIMO.Mhtml
dotnet add package OfficeIMO.Html.Pdf
```

## Create and save an MHTML archive

Build an archive from HTML plus the resources it references by `cid:` or `Content-Location`.

### C#

```csharp
using System.Text;
using OfficeIMO.Mhtml;

var stylesheet = new MhtmlResource(
    Encoding.UTF8.GetBytes("h1 { color: navy }"),
    contentType: "text/css",
    contentLocation: "https://example.com/site.css");

var archive = new MhtmlDocument(
    "<html><head><link rel=\"stylesheet\" href=\"site.css\"></head><body><h1>Saved page</h1></body></html>",
    new[] { stylesheet },
    contentLocation: "https://example.com/index.html",
    subject: "Saved page");

archive.Save("saved-page.mhtml");
```

## Read an MHTML file

`Load` parses the MIME structure, picks the root HTML part and resolves its base URI. `MimeDiagnostics` lists recovered or ambiguous parts, such as duplicate `Content-ID` values.

### C#

```csharp
using OfficeIMO.Mhtml;

MhtmlDocument archive = MhtmlDocument.Load("saved-page.mhtml");

Console.WriteLine($"Subject: {archive.Subject}");
Console.WriteLine($"Base URI: {archive.BaseUri}");
Console.WriteLine(archive.HtmlDocument.NormalizedHtml);

foreach (MhtmlResource resource in archive.Resources) {
    Console.WriteLine($"{resource.ContentType} {resource.ContentLocation ?? resource.ContentId}");
}

foreach (var diagnostic in archive.MimeDiagnostics) {
    Console.WriteLine(diagnostic);
}
```

### PowerShell

```powershell
Read-OfficeDocument -Path .\saved-page.mhtml -As Text
```

### CLI

```sh
officeimo reader read saved-page.mhtml --format markdown --output saved-page.md
```

## Convert an MHTML file to PDF

`ConfigureRenderOptions` points the renderer at the archive's embedded resources. The same options object is then passed to the PDF conversion.

### C#

```csharp
using OfficeIMO.Html.Pdf;
using OfficeIMO.Mhtml;

MhtmlDocument archive = MhtmlDocument.Load("saved-page.mhtml");

var options = new HtmlPdfExportOptions();
archive.ConfigureRenderOptions(options);

archive.HtmlDocument.SaveAsPdf("saved-page.pdf", options);
```

### CLI

```sh
officeimo html convert saved-page.mhtml --output saved-page.pdf
```

## Render an MHTML file to images

### C#

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Mhtml;

MhtmlDocument archive = MhtmlDocument.Load("saved-page.mhtml");

var options = new HtmlRenderOptions {
    ViewportWidth = 1280,
    Mode = HtmlRenderMode.Continuous,
    Format = OfficeImageExportFormat.Png
};
archive.ConfigureRenderOptions(options);

archive.HtmlDocument.SaveAsImage("saved-page.png", options);
```

### CLI

```sh
officeimo html render saved-page.mhtml --profile screen-full-page --encoder png --output saved-page-images.zip
```

## Convert a zipped website to PDF

`HtmlSiteBundle` selects `index.html` (or the single HTML entry) unless you name one in `EntryPath`. It checks entry counts, decoded sizes, compression ratios and checksums, and rejects path traversal. Archive images, CSS and fonts are used only when `ResourcePolicy.AllowEmbeddedPackageResources` is on.

### C#

```csharp
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

HtmlSiteBundle bundle = await HtmlSiteBundle.LoadAsync("saved-site.zip",
    new HtmlSiteBundleOptions { EntryPath = "articles/report.html" });

Console.WriteLine($"{bundle.EntryPath}: {bundle.EntryNames.Count} entries");

var options = new HtmlPdfExportOptions();
options.ResourcePolicy.AllowEmbeddedPackageResources = true;

HtmlPdfRenderRequestResult result = await bundle.RenderToPdfResultAsync(
    HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));
File.WriteAllBytes("saved-site.pdf", result.ToBytes());
```

### CLI

```sh
officeimo html convert saved-site.zip --input-format site-bundle --entry-path articles/report.html --output saved-site.pdf
```
