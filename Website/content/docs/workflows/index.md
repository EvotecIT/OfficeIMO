---
title: "Workflow Finder"
description: "Choose an OfficeIMO or PSWriteOffice workflow by the input you have and the artifact you need."
order: 5
meta.seo_title: "OfficeIMO examples by document workflow"
---

Start with the job, not the package name. Each route below links to the smallest useful package set, a runnable pattern, and the evidence available for the generated result.

## Build and deliver

| Starting point | Result | Recommended route |
|---|---|---|
| Application objects or a `DataTable` | XLSX report | `OfficeIMO.Excel`, or `Export-OfficeExcel` from PowerShell |
| Template plus business data | DOCX report or contract | `OfficeIMO.Word`, or the PSWriteOffice Word commands |
| Structured report model | PDF | `OfficeIMO.Pdf` for native authoring; add a focused converter for Word, HTML, RTF, or OpenDocument input |
| Brand brief and metrics | PPTX presentation | `OfficeIMO.PowerPoint` designer and template APIs |
| Topology or inventory | VSDX plus image preview | `OfficeIMO.Visio` diagram builders |
| SQL query | Excel or CSV, optionally written back to SQL | [DbaClientX and PSWriteOffice reporting](/docs/workflows/database-reporting/) |

PowerShell users can start with the [PSWriteOffice recipe gallery](/docs/workflows/powershell-recipes/) for Excel, Word/PDF, native PDF, Markdown, CSV, Reader, and PowerPoint patterns.

## Convert and publish

| Input | Useful outputs | Guide |
|---|---|---|
| HTML or MHTML | Images, PDF, Word, Markdown, RTF | [Render and convert HTML](/docs/html/render-and-convert/) |
| OneNote `.one`, `.onetoc2`, or `.onepkg` | Images, HTML, Markdown, PDF, Reader results | [Export and convert OneNote](/docs/onenote/export-and-convert/) |
| RTF | RTF edits, Word, Markdown, HTML, PDF, Reader results | [RTF workflows](/docs/rtf/) |
| OpenDocument ODT, ODS, or ODP | Office formats, PDF, inspection results | [OpenDocument workflows](/docs/open-document/) |
| PowerPoint | Slide images, PDF, content extraction | [PowerPoint image export](/docs/powerpoint/image-export/) |

## Extract, index, and validate

Use the [OfficeIMO.Reader documentation](/docs/reader/) when the destination is a normalized extraction model rather than another document. Install selective Reader adapters for a service with a known format set, or `OfficeIMO.Reader.All` for a host that deliberately accepts the complete adapter family.

For generated output, keep validation close to the workflow:

- Validate DOCX, XLSX, PPTX, VSDX, and package relationships before delivery.
- Inspect conversion diagnostics instead of assuming every source feature maps losslessly.
- Apply bounded reader profiles to uploaded or otherwise untrusted documents.
- Keep a representative output artifact or visual baseline when layout matters.

See the [generated output gallery](/showcase/) for real previews tied to examples and tests.
