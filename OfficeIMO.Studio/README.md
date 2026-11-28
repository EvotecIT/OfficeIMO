# OfficeIMO Studio

A desktop app for reading, marking up, filling, organizing, converting, protecting and signing documents, built on OfficeIMO's own engines. It runs on Windows and Linux, works on your files locally, and doesn't need Microsoft Office, a PDF library or an external converter.

## Download

Windows and Linux packages for x64 and Arm64 are on [officeimo.com/downloads](https://officeimo.com/downloads/). The packaged editions include the .NET runtime. Native apps for Mac, iPad and iPhone are in development; see [Apple/OfficeIMOStudio](../Apple/OfficeIMOStudio/README.md).

## What you can do

| Area | Features |
| --- | --- |
| Read and find | Open PDFs, Search, Select and copy text, Reading layouts, Bookmarks and outline, Document properties |
| Mark up and review | Highlight, underline and strike through, Freehand drawing, Notes, Shapes, text boxes and stamps, Comment threads, Compare documents |
| Organize pages | Rotate, reorder, duplicate and delete pages, Insert pages, Extract pages, Split a document, Combine documents, Crop and resize pages, Page numbers, headers and watermarks |
| Edit content | Edit existing text, Add and replace images, Links and attachments |
| Forms | Fill forms, Create form fields, Flatten forms and annotations |
| Scan and recognize text | Make scans searchable |
| Convert | Convert to PDF, Open Office documents, Convert from PDF |
| Protect and sign | Password protection, Redaction, Handwritten signature, Certificate signature, Validate signatures and standards |
| Print, share and automate | Print, Automation, Document assistant |
| Workbenches for larger jobs | Document health, Conversion workbench, Assemble a PDF, Batch and page-image export, Print, export and assemble in one place, Review recognized text, OCR setup check, Prepare scans, Saved jobs and run history, Publish books, Electronic invoices, File origin and Content Credentials |
| Workspace | Keep unsaved work, English, Polish, German, French, Spanish, Italian and Brazilian Portuguese, Light, dark and high contrast, Screen readers |

The [Studio page](https://officeimo.com/studio/) shows screenshots and which features are available on each platform, from [`Website/data/studio_features.json`](../Website/data/studio_features.json). The [Studio guide](../Docs/officeimo.studio-guide.md) describes each workspace in detail.

## Requirements

Conversions run in process through OfficeIMO's engines, so no office suite or converter needs to be installed. Text recognition is optional and uses an installed Tesseract; Studio can download checksum-pinned language data but never installs the OCR engine itself. Printing uses the platform printer service, and preparing and saving print sheets works without a printer.

## Languages and local data

Studio is available in English, Polish, German, French, Spanish, Italian and Brazilian Portuguese, with system, light, dark and high-contrast appearance and screen reader support. Preferences, recovery data and diagnostics stay in the operating system's local application-data folder; set `OFFICEIMO_STUDIO_DATA_ROOT` to use another folder. See [Studio privacy](https://officeimo.com/studio-privacy/).

## Build and run from source

```powershell
dotnet run --project OfficeIMO.Studio/OfficeIMO.Studio.csproj
dotnet test OfficeIMO.Studio.Tests/OfficeIMO.Studio.Tests.csproj -c Release
./Build/Studio/Build-Studio.ps1 -Plan
```

[`Build/Studio/README.md`](../Build/Studio/README.md) covers release targets, signing and packaging. Interface text is translated into every shipped language in the same change; see the Studio translation scripts in `Build/Studio`.

## Learn more

- Planned work is in the **Desktop Studio** section of the [roadmap](../Docs/ROADMAP.md#desktop-studio).
- Studio has no document engine of its own: PDF, conversion, workflow, security and OCR behaviour lives in the OfficeIMO packages it calls. The [Studio guide](../Docs/officeimo.studio-guide.md#capability-ownership) lists which package owns what.
