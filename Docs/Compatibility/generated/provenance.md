# OfficeIMO.Provenance capability contract

Schema version: 1

Only formats with a named OfficeIMO owner appear in this contract. Memory-only and browser support require separate qualification.

| Capability | Extension | Structural format | Owner | Inspect | Assess | Remove | Memory-only | Browser | Boundary |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Word Open XML | `.docm` | ZipPackage | `OfficeIMO.Word` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| Word Open XML | `.docx` | ZipPackage | `OfficeIMO.Word` | Yes | Yes | Yes | Yes | Yes | Package signatures block mutation unless removal is explicitly authorized. |
| Word Open XML | `.dotm` | ZipPackage | `OfficeIMO.Word` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| Word Open XML | `.dotx` | ZipPackage | `OfficeIMO.Word` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| Excel workbook package | `.xlam` | ZipPackage | `OfficeIMO.Excel` | Yes | Yes | Yes | No | No | SpreadsheetML and XLSB package identity are validated before mutation. |
| Excel workbook package | `.xlsb` | ZipPackage | `OfficeIMO.Excel` | Yes | Yes | Yes | No | No | SpreadsheetML and XLSB package identity are validated before mutation. |
| Excel workbook package | `.xlsm` | ZipPackage | `OfficeIMO.Excel` | Yes | Yes | Yes | No | No | SpreadsheetML and XLSB package identity are validated before mutation. |
| Excel workbook package | `.xlsx` | ZipPackage | `OfficeIMO.Excel` | Yes | Yes | Yes | Yes | Yes | SpreadsheetML and XLSB package identity are validated before mutation. |
| Excel workbook package | `.xltm` | ZipPackage | `OfficeIMO.Excel` | Yes | Yes | Yes | No | No | SpreadsheetML and XLSB package identity are validated before mutation. |
| Excel workbook package | `.xltx` | ZipPackage | `OfficeIMO.Excel` | Yes | Yes | Yes | No | No | SpreadsheetML and XLSB package identity are validated before mutation. |
| PowerPoint Open XML | `.potm` | ZipPackage | `OfficeIMO.PowerPoint` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| PowerPoint Open XML | `.potx` | ZipPackage | `OfficeIMO.PowerPoint` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| PowerPoint Open XML | `.ppam` | ZipPackage | `OfficeIMO.PowerPoint` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| PowerPoint Open XML | `.ppsm` | ZipPackage | `OfficeIMO.PowerPoint` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| PowerPoint Open XML | `.ppsx` | ZipPackage | `OfficeIMO.PowerPoint` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| PowerPoint Open XML | `.pptm` | ZipPackage | `OfficeIMO.PowerPoint` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| PowerPoint Open XML | `.pptx` | ZipPackage | `OfficeIMO.PowerPoint` | Yes | Yes | Yes | Yes | Yes | Package signatures block mutation unless removal is explicitly authorized. |
| Visio Open XML | `.vsdm` | ZipPackage | `OfficeIMO.Visio` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| Visio Open XML | `.vsdx` | ZipPackage | `OfficeIMO.Visio` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| Visio Open XML | `.vssm` | ZipPackage | `OfficeIMO.Visio` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| Visio Open XML | `.vssx` | ZipPackage | `OfficeIMO.Visio` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| Visio Open XML | `.vstm` | ZipPackage | `OfficeIMO.Visio` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| Visio Open XML | `.vstx` | ZipPackage | `OfficeIMO.Visio` | Yes | Yes | Yes | No | No | Package signatures block mutation unless removal is explicitly authorized. |
| OpenDocument | `.odg` | ZipPackage | `OfficeIMO.OpenDocument` | Yes | Yes | Yes | No | No | Encrypted OpenDocument packages cannot be rewritten. |
| OpenDocument | `.odp` | ZipPackage | `OfficeIMO.OpenDocument` | Yes | Yes | Yes | No | No | Encrypted OpenDocument packages cannot be rewritten. |
| OpenDocument | `.ods` | ZipPackage | `OfficeIMO.OpenDocument` | Yes | Yes | Yes | No | No | Encrypted OpenDocument packages cannot be rewritten. |
| OpenDocument | `.odt` | ZipPackage | `OfficeIMO.OpenDocument` | Yes | Yes | Yes | No | No | Encrypted OpenDocument packages cannot be rewritten. |
| OpenDocument | `.otg` | ZipPackage | `OfficeIMO.OpenDocument` | Yes | Yes | Yes | No | No | Encrypted OpenDocument packages cannot be rewritten. |
| OpenDocument | `.otp` | ZipPackage | `OfficeIMO.OpenDocument` | Yes | Yes | Yes | No | No | Encrypted OpenDocument packages cannot be rewritten. |
| OpenDocument | `.ots` | ZipPackage | `OfficeIMO.OpenDocument` | Yes | Yes | Yes | No | No | Encrypted OpenDocument packages cannot be rewritten. |
| OpenDocument | `.ott` | ZipPackage | `OfficeIMO.OpenDocument` | Yes | Yes | Yes | No | No | Encrypted OpenDocument packages cannot be rewritten. |
| EPUB | `.epub` | ZipPackage | `OfficeIMO.Epub` | Yes | Yes | Yes | No | No | Package structure is validated before inspection or mutation. |
| PDF | `.pdf` | Pdf | `OfficeIMO.Pdf` | Yes | Yes | Yes | Yes | Yes | Removal is limited to provenance associations supported by the PDF owner. |
| HTML | `.htm` | Html | `OfficeIMO.Html` | Yes | Yes | Yes | No | No | External resources are not fetched during inspection. |
| HTML | `.html` | Html | `OfficeIMO.Html` | Yes | Yes | Yes | No | No | External resources are not fetched during inspection. |
| Markdown | `.markdown` | StructuredText | `OfficeIMO.Markdown` | Yes | Yes | Yes | No | No | Original BOM-aware UTF encoding is preserved by file mutation. |
| Markdown | `.md` | StructuredText | `OfficeIMO.Markdown` | Yes | Yes | Yes | No | No | Original BOM-aware UTF encoding is preserved by file mutation. |
| Image provenance | `.gif` | Gif | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | File contents must match the registered image format; unrelated media containers are not inferred from signatures. |
| Image provenance | `.jpeg` | Jpeg | `OfficeIMO.Core` | Yes | Yes | Yes | Yes | Yes | File contents must match the registered image format; unrelated media containers are not inferred from signatures. |
| Image provenance | `.jpg` | Jpeg | `OfficeIMO.Core` | Yes | Yes | Yes | Yes | Yes | File contents must match the registered image format; unrelated media containers are not inferred from signatures. |
| Image provenance | `.png` | Png | `OfficeIMO.Core` | Yes | Yes | Yes | Yes | Yes | File contents must match the registered image format; unrelated media containers are not inferred from signatures. |
| Image provenance | `.svg` | Svg | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | File contents must match the registered image format; unrelated media containers are not inferred from signatures. |
| Image provenance | `.tif` | Tiff | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | File contents must match the registered image format; unrelated media containers are not inferred from signatures. |
| Image provenance | `.tiff` | Tiff | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | File contents must match the registered image format; unrelated media containers are not inferred from signatures. |
| Image provenance | `.webp` | Webp | `OfficeIMO.Core` | Yes | Yes | Yes | Yes | Yes | File contents must match the registered image format; unrelated media containers are not inferred from signatures. |
| Structured text | `.adoc` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.asciidoc` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.bat` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.c` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.cjs` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.cmd` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.cpp` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.cs` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.css` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.go` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.h` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.hpp` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.ini` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.java` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.js` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.json` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.lua` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.mjs` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.ps1` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.py` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.rb` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.rs` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.sh` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.sql` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.tex` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.toml` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.ts` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.txt` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.vb` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.xml` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.yaml` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
| Structured text | `.yml` | StructuredText, UnstructuredText | `OfficeIMO.Core` | Yes | Yes | Yes | No | No | Only standards-defined structured or wrapped text carriers are changed. |
