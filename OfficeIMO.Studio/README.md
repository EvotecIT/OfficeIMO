# OfficeIMO Studio

OfficeIMO Studio is the cross-platform desktop surface for OfficeIMO's document engines. The application presents the existing PDF, conversion, workflow, security, and OCR capabilities; reusable document behavior remains in the owning OfficeIMO packages.

## Current workspace

- Open several PDFs in document tabs and use single-page, continuous, two-page, or grid reading modes.
- Search, follow bookmarks and links, navigate by keyboard, use page night mode, and compare two PDFs in synchronized panes.
- Select existing text, images, and annotations with visible bounds and handles. Replace, move, resize, recolor, flatten, or remove supported objects through the canonical PDF editors.
- Add text, images, links, notes, markup, shapes, ink, stamps, signature appearances, watermarks, and page numbers.
- Organize pages by drag and drop; rotate, crop, duplicate, reorder, import, extract, split, and insert blank pages with undo, redo, save, and recovery support.
- Fill, flatten, and author supported AcroForm controls.
- Inspect protection and signatures; protect or decrypt a copy; apply certificate signatures, validate signatures, add Bates numbering, sanitize, repair, optimize, and perform verified redaction.
- Convert supported document formats, export PDF pages to supported image formats, preview print sheets, and assemble PDFs from ordered PDFs, images, Office files, folders, and ZIP archives.
- Create searchable PDFs with the public `OfficeIMO.Reader.Ocr` facade. The workbench supports page ranges, 28 typed language choices, confidence and rendering controls, checksum-verified language-data provisioning, cancellation, and safe output conflict policy.

## Capability ownership

Studio does not contain a second document engine:

- `OfficeIMO.Pdf` owns PDF reading, rendering, interaction geometry, editing, forms, security, signatures, redaction, and output mutation.
- `OfficeIMO.Workflows` owns reusable conversion, inspection, repair, optimization, comparison, sanitization, image-export, print-planning, and mixed-source assembly workflows.
- `OfficeIMO.Reader.Ocr` owns the easy searchable-PDF flow over the engine-neutral Reader OCR contracts and the optional Tesseract CLI provider.
- Avalonia owns only the cross-platform windowing and control layer. Studio renders OfficeIMO's retained page scene; it does not use PDFium or convert pages to images merely to display them.

Tesseract is an explicit OCR runtime prerequisite. OfficeIMO discovers an installed executable and can provision checksum-pinned language data, but the application does not silently install or bundle a native OCR executable.

## Run and verify

```powershell
dotnet run --project OfficeIMO.Studio/OfficeIMO.Studio.csproj
dotnet test OfficeIMO.Studio.Tests/OfficeIMO.Studio.Tests.csproj -c Release
```

The document workspace currently opens PDF files directly. Other supported formats enter through conversion or mixed-source assembly. Print planning and preview are implemented; native operating-system printer submission remains open work.

## Next product outcomes

The repository's single open-work plan is [`Docs/ROADMAP.md`](../Docs/ROADMAP.md). Its **Desktop Studio** section tracks distribution, richer OCR review, native print and scan intake, reusable workflow recipes, and cross-platform accessibility and usability evidence.
