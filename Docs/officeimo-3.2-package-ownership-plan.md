# OfficeIMO 3.2 package ownership plan

## Goal

OfficeIMO 3.2 will separate base format engines, optional cross-format bridges, and Reader-specific behavior. Installing a base format package must not pull unrelated formats merely because an optional conversion exists.

The cleanup has two related outcomes:

- format conversions use a neutral document model owned by `OfficeIMO.Core`, not `OfficeIMO.Reader.*`;
- `OfficeIMO.Html` remains the normal lean HTML package, while MHTML, RTF, email-image, and MHTML-to-PDF behavior moves to explicit bridge packages.

No published package is being removed. `OfficeIMO.Html.Core`, `OfficeIMO.Reader.Mhtml`, and a separate document-model package will not be introduced.

## Target package graph

### Neutral document model and conversions

```text
OfficeIMO.Core
  OfficeDocumentModel and neutral document primitives

OfficeIMO.Visio
  VisioDocument -> OfficeDocumentModel

OfficeIMO.Pdf
  OfficeDocumentModel -> PdfDocument

OfficeIMO.Visio.Pdf
  VisioDocument -> OfficeDocumentModel -> PdfDocument
```

`OfficeIMO.Visio.Pdf` will depend on `OfficeIMO.Visio` and `OfficeIMO.Pdf`. It will not depend on `OfficeIMO.Reader.Visio` or `OfficeIMO.Reader.Pdf`.

The neutral Core model will include document source information, pages, blocks, tables, visuals, assets, links, forms, locations, metadata, and diagnostics. Reader chunks, routing, registration, processing, search, JSON transport, and reader options remain in `OfficeIMO.Reader.Core`.

`OfficeIMO.Core` will remain free of NuGet package dependencies. JSON-specific attributes and serialization behavior stay outside the neutral model.

### HTML, MHTML, RTF, email, and EPUB

```text
OfficeIMO.Reader.Epub
  -> OfficeIMO.Reader.Html
      -> OfficeIMO.Html
          -> OfficeIMO.Core + AngleSharp + AngleSharp.Css
```

EPUB chapters are XHTML documents. `OfficeIMO.Reader.Epub` will continue to reuse the HTML reader, but the HTML reader will no longer include MHTML or Email behavior.

```text
OfficeIMO.Html.Rtf
  -> OfficeIMO.Html + OfficeIMO.Rtf

OfficeIMO.Mhtml
  -> OfficeIMO.Html + OfficeIMO.Email

OfficeIMO.Email.Image
  -> OfficeIMO.Email + OfficeIMO.Html + OfficeIMO.Html.Rtf

OfficeIMO.Mhtml.Pdf
  -> OfficeIMO.Mhtml + OfficeIMO.Html.Pdf + OfficeIMO.Pdf
```

These bridge packages are optional edges between otherwise independent formats. They prevent either base package from acquiring the other format's dependencies.

## New packages

| Package | Responsibility | Direct dependencies |
|---|---|---|
| `OfficeIMO.Html.Rtf` | HTML to RTF and RTF to HTML conversion | Html, Rtf, Core |
| `OfficeIMO.Mhtml` | MHTML load/save, HTML document projection, and embedded MIME resources | Html, Email, Core |
| `OfficeIMO.Email.Image` | Render email bodies to supported image formats, including HTML and RTF body fallback | Email, Html, Html.Rtf, Core |
| `OfficeIMO.Mhtml.Pdf` | Direct MHTML to PDF conversion | Mhtml, Html.Pdf, Pdf, Core |

Each new package will use the repository's standard target matrix: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.

## Existing package changes

| Existing package | Change | Resulting important dependencies |
|---|---|---|
| `OfficeIMO.Core` | Gains neutral document contracts; gains no package dependencies | None |
| `OfficeIMO.Reader.Core` | Retains Reader result types, execution, and transport behavior; it is no longer the conversion boundary | Core; System.Text.Json on legacy targets |
| `OfficeIMO.Visio` | Gains Visio to neutral-document projection | Core and existing Visio dependencies |
| `OfficeIMO.Pdf` | Gains neutral-document to PDF projection and destination-owned projection options | Core and existing PDF dependencies |
| `OfficeIMO.Visio.Pdf` | Becomes a thin format bridge | Core, Visio, Pdf |
| `OfficeIMO.Reader.Visio` | Reuses the Visio-owned neutral projection and adds Reader behavior | Reader.Core, Visio |
| `OfficeIMO.Reader.Pdf` | Retains PDF reading and a thin Reader-result compatibility bridge; the projection engine and options move to PDF | Reader.Core, Pdf |
| `OfficeIMO.Html` | Keeps the HTML engine; loses Email, Rtf, MHTML, and email-image code | Core, AngleSharp, AngleSharp.Css |
| `OfficeIMO.Html.Pdf` | Keeps plain HTML/PDF conversion; MHTML/PDF APIs move out | Html, Pdf, Core |
| `OfficeIMO.Reader.Html` | Handles only HTML, HTM, and XHTML | Reader.Core, Html, Markdown.Html |
| `OfficeIMO.Reader.Email` | Owns email and MHTML reader registration; MHTML uses the Mhtml and HTML reader owners | Reader.Core, Email, Mhtml, Reader.Html |
| `OfficeIMO.Reader.Epub` | Keeps EPUB orchestration over the lean HTML reader | Reader.Core, Reader.Html, Epub |
| `OfficeIMO.Reader.All` | Continues to compose all Reader capabilities | Existing Reader packages; MHTML arrives through Reader.Email |
| `OfficeIMO.Email` | Remains the base email/MIME engine | Existing dependencies only; no Html, Rtf, or Pdf |
| `OfficeIMO.Rtf` | Remains the base RTF engine | Existing dependencies only; no Html |

## Packages that lose transitive Email and RTF dependencies

Once `OfficeIMO.Html` is lean, these packages can reference it directly without inheriting Email or RTF:

- `OfficeIMO.Adf`
- `OfficeIMO.Epub.Image`
- `OfficeIMO.Excel.Html`
- `OfficeIMO.Markdown.Html`
- `OfficeIMO.OneNote.Html`
- `OfficeIMO.PowerPoint.Html`
- `OfficeIMO.Word.Html`
- `OfficeIMO.Html.Pdf`
- `OfficeIMO.Reader.Html`
- `OfficeIMO.Reader.Epub`

Their downstream consumers also lose those transitive dependencies unless they independently request an Email, MHTML, or RTF bridge.

## Public API moves

### Neutral document model

- Introduce `OfficeDocumentModel` and unambiguously named `OfficeDocumentModel*` child types in `OfficeIMO.Core`.
- Replace Reader-prefixed model types used by conversions, such as `ReaderTable`, `ReaderVisual`, `ReaderLocation`, and `ReaderInputKind`, with neutral equivalents.
- Keep `ReaderChunk`, `ReaderOptions`, Reader processors, and Reader serialization in `OfficeIMO.Reader.Core`.
- Reader APIs keep their Reader-specific result and chunks. Their compatibility adapter maps that result into `OfficeDocumentModel`; format converters accept the neutral model.

### Visio and PDF

- Move `VisioDocument -> OfficeDocumentModel` from `OfficeIMO.Reader.Visio` to `OfficeIMO.Visio`.
- Move `OfficeDocumentModel -> PdfDocument` from `OfficeIMO.Reader.Pdf` to `OfficeIMO.Pdf`.
- Replace `ReaderPdfProjectionOptions` with destination-owned `PdfProjectionOptions`.
- Remove Reader option types from `VisioPdfSaveOptions`; source options belong to Visio and output options belong to PDF.

### HTML integrations

- Move HTML/RTF conversion APIs from `OfficeIMO.Html` to `OfficeIMO.Html.Rtf`.
- Move `MhtmlDocument` and `MhtmlResource` from `OfficeIMO.Html` to `OfficeIMO.Mhtml`.
- Move email image-export APIs from `OfficeIMO.Html` to `OfficeIMO.Email.Image`.
- Move MHTML/PDF extension methods from `OfficeIMO.Html.Pdf` to `OfficeIMO.Mhtml.Pdf`.
- `OfficeDocumentReaderBuilder.WithHtml()` will register `.html`, `.htm`, and `.xhtml` only.
- MHTML registration will move to `OfficeIMO.Reader.Email`, with an explicit `AddMhtmlHandler()` entry point. `OfficeIMO.Reader.All` will register it automatically.

## Compatibility and migration

This is an intentional 3.2 breaking cleanup. A compatibility facade must not make `OfficeIMO.Html` depend on every bridge again.

The migration guide will include these package-reference changes:

| 3.1 usage | 3.2 package |
|---|---|
| Plain HTML parsing, layout, normalization, or image rendering | `OfficeIMO.Html` |
| HTML and RTF conversion | Add `OfficeIMO.Html.Rtf` |
| MHTML load/save | Add `OfficeIMO.Mhtml` |
| Email body image rendering | Add `OfficeIMO.Email.Image` |
| Plain HTML/PDF conversion | `OfficeIMO.Html.Pdf` |
| MHTML/PDF conversion | Add `OfficeIMO.Mhtml.Pdf` |
| HTML Reader registration | `OfficeIMO.Reader.Html` and `WithHtml()` |
| MHTML Reader registration | `OfficeIMO.Reader.Email` and `AddMhtmlHandler()` |
| Visio/PDF conversion | `OfficeIMO.Visio.Pdf`; Reader packages are no longer required |

Applications upgrading the moved public types must rebuild against the coordinated 3.2 package set. Mixing pre-3.2 satellite binaries with the new ownership graph will not be supported.

## Implementation and release sequence

The work should be reviewed as two ownership changes and released together as OfficeIMO 3.2.0.

### Change 1: neutral conversion model

- [x] Add the neutral document model to Core without external dependencies.
- [x] Keep Reader-only chunks, routing, processing, and serialization outside the neutral model.
- [x] Move the Visio source projection into `OfficeIMO.Visio`.
- [x] Move the PDF destination projection into `OfficeIMO.Pdf`.
- [x] Remove Reader dependencies and Reader options from `OfficeIMO.Visio.Pdf`.
- [x] Update affected tests, docs, AOT roots, build configuration, and package smoke coverage.

### Change 2: HTML integration packages

- [x] Keep the existing `OfficeIMO.Html` assembly and package as the lean HTML engine.
- [x] Remove the unpublished `OfficeIMO.Html.Core` experiment.
- [x] Add Html.Rtf, Mhtml, Email.Image, and Mhtml.Pdf projects.
- [x] Move APIs and tests to their owning packages.
- [x] Make Reader.Html HTML-only and move MHTML registration to Reader.Email.
- [x] Verify Reader.Epub uses the lean Reader.Html path.
- [x] Update the solution, package configuration, workflows, AOT smoke projects, website catalog, and generated API documentation.

### Release proof

- [x] Build and test all affected projects for every supported target framework.
- [x] Pack the coordinated 3.2.0 package set and inspect every generated nuspec dependency group.
- [x] Prove a clean HTML consumer does not receive Email, Rtf, Mhtml, or Pdf.
- [x] Prove a clean EPUB Reader consumer does not receive Email, Rtf, or Mhtml.
- [x] Prove Visio-to-PDF runs without Reader assemblies present.
- [x] Prove each optional bridge works when its package is installed: HTML/RTF, MHTML, email/image, and MHTML/PDF.
- [x] Run clean consumer builds on `net472`, `net8.0`, and `net10.0`, plus the relevant NativeAOT smoke checks.
- [ ] Publish all coordinated packages before updating downstream consumers.
- [x] Update website package/API source and migration documentation from the validated target package graph.
