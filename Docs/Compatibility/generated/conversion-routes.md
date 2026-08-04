# OfficeIMO conversion routes

Schema version: 1

| Route | Source | Target | Package | Fidelity | Browser | API | Result contract |
| --- | --- | --- | --- | --- | --- | --- | --- |
| docx-pdf | DOCX | PDF | OfficeIMO.Word.Pdf | FixedLayout | Yes | `WordDocument.Load(stream).ToPdfDocumentResult(options)` | PdfDocumentConversionResult |
| xlsx-pdf | XLSX | PDF | OfficeIMO.Excel.Pdf | FixedLayout | Yes | `ExcelDocument.Load(stream).ToPdfDocumentResult(options)` | PdfDocumentConversionResult |
| pptx-pdf | PPTX | PDF | OfficeIMO.PowerPoint.Pdf | FixedLayout | Yes | `PowerPointPresentation.Load(stream).ToPdfDocumentResult(options)` | PdfDocumentConversionResult |
| html-pdf | HTML | PDF | OfficeIMO.Html.Pdf | FixedLayout | Yes | `HtmlConversionDocument.Parse(html).ToPdfDocumentResult(options)` | PdfDocumentConversionResult |
| markdown-html | Markdown | HTML | OfficeIMO.MarkdownRenderer | Semantic | Yes | `MarkdownRenderer.RenderBodyHtml(markdown, options)` | string |
| html-markdown | HTML | Markdown | OfficeIMO.Markdown.Html | Semantic | Yes | `HtmlConversionDocument.Parse(html).ToMarkdownDocumentResult(options)` | HtmlToMarkdownResult |
| markdown-docx | Markdown | DOCX | OfficeIMO.Word.Markdown | Editable | Yes | `MarkdownReader.Parse(markdown).ToWordDocumentResult(options)` | MarkdownToWordResult |
| docx-html | DOCX | HTML | OfficeIMO.Word.Html | Semantic | No | `WordDocument.Load(stream).ToHtmlResult(options)` | HtmlTextConversionResult |
| docx-markdown | DOCX | Markdown | OfficeIMO.Word.Markdown | Semantic | No | `WordDocument.Load(stream).ToMarkdownDocumentResult(options)` | WordToMarkdownResult |
| docx-odt | DOCX | ODT | OfficeIMO.Word.OpenDocument | Editable | No | `WordDocument.Load(stream).ToOpenDocumentResult(options)` | OdfConversionResult<OdtDocument> |
| odt-docx | ODT | DOCX | OfficeIMO.Word.OpenDocument | Editable | No | `OdtDocument.Load(stream).ToWordDocumentResult(options)` | OdfConversionResult<WordDocument> |
| docx-rtf | DOCX | RTF | OfficeIMO.Word.Rtf | Editable | No | `WordDocument.Load(stream).ToRtfDocumentResult()` | RtfConversionResult<RtfDocument> |
| rtf-docx | RTF | DOCX | OfficeIMO.Word.Rtf | Editable | No | `RtfDocument.Load(stream, readOptions).ToWordDocumentResult(sourcePath)` | RtfConversionResult<WordDocument> |
| xlsx-html | XLSX | HTML | OfficeIMO.Excel.Html | Semantic | No | `ExcelDocument.Load(stream).ToHtmlResult(options)` | HtmlTextConversionResult |
| xlsx-ods | XLSX | ODS | OfficeIMO.Excel.OpenDocument | Editable | No | `ExcelDocument.Load(stream).ToOpenDocumentResult(options)` | OdfConversionResult<OdsDocument> |
| ods-xlsx | ODS | XLSX | OfficeIMO.Excel.OpenDocument | Editable | No | `OdsDocument.Load(stream).ToExcelDocumentResult(options)` | OdfConversionResult<ExcelDocument> |
| pptx-html | PPTX | HTML | OfficeIMO.PowerPoint.Html | Semantic | No | `PowerPointPresentation.Load(stream).ToHtmlResult(options)` | PowerPointToHtmlResult |
| pptx-odp | PPTX | ODP | OfficeIMO.PowerPoint.OpenDocument | Editable | No | `PowerPointPresentation.Load(stream).ToOpenDocumentResult(options)` | OdfConversionResult<OdpPresentation> |
| odp-pptx | ODP | PPTX | OfficeIMO.PowerPoint.OpenDocument | Editable | No | `OdpPresentation.Load(stream).ToPowerPointPresentationResult(options)` | OdfConversionResult<PowerPointPresentation> |
| markdown-pdf | Markdown | PDF | OfficeIMO.Markdown.Pdf | FixedLayout | No | `MarkdownReader.Parse(markdown).ToPdfDocumentResult(options)` | PdfDocumentConversionResult |
| rtf-markdown | RTF | Markdown | OfficeIMO.Rtf.Markdown | Semantic | No | `RtfDocument.Load(stream, readOptions).Document.ToMarkdownResult(options)` | RtfConversionResult<string> |
| rtf-pdf | RTF | PDF | OfficeIMO.Rtf.Pdf | FixedLayout | No | `RtfDocument.Load(stream, readOptions).Document.ToPdfDocumentResult(options)` | PdfDocumentConversionResult |
| markdown-rtf | Markdown | RTF | OfficeIMO.Rtf.Markdown | Editable | No | `MarkdownReader.Parse(markdown).ToRtfDocumentResult(options)` | RtfConversionResult<RtfDocument> |
| html-docx | HTML | DOCX | OfficeIMO.Word.Html | Editable | No | `HtmlConversionDocument.Parse(html).ToWordDocumentResult(options)` | HtmlToWordResult |
| html-xlsx | HTML | XLSX | OfficeIMO.Excel.Html | Editable | No | `HtmlConversionDocument.Parse(html).ToExcelDocumentResult(options)` | HtmlToExcelResult |
| html-pptx | HTML | PPTX | OfficeIMO.PowerPoint.Html | Editable | No | `HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(options)` | HtmlToPowerPointResult |
| html-rtf | HTML | RTF | OfficeIMO.Html | Editable | No | `HtmlConversionDocument.Parse(html).ToRtfDocumentResult(options)` | HtmlToRtfResult |
| rtf-html | RTF | HTML | OfficeIMO.Html | Semantic | No | `RtfDocument.Load(stream, readOptions).Document.ToHtmlResult(options)` | RtfToHtmlResult |
| asciidoc-markdown | AsciiDoc | Markdown | OfficeIMO.AsciiDoc.Markdown | Semantic | No | `AsciiDocDocument.Parse(source).Document.ToMarkdownDocumentResult(options)` | AsciiDocToMarkdownResult |
| markdown-asciidoc | Markdown | AsciiDoc | OfficeIMO.AsciiDoc.Markdown | Semantic | No | `MarkdownReader.Parse(markdown).ToAsciiDocDocumentResult(options)` | MarkdownToAsciiDocResult |
| asciidoc-pdf | AsciiDoc | PDF | OfficeIMO.AsciiDoc.Pdf | FixedLayout | No | `AsciiDocDocument.Parse(source).Document.ToPdfDocumentResult(options)` | PdfDocumentConversionResult |
| latex-markdown | LaTeX | Markdown | OfficeIMO.Latex.Markdown | Semantic | No | `LatexDocument.Parse(source).Document.ToMarkdownDocumentResult(options)` | LatexToMarkdownResult |
| markdown-latex | Markdown | LaTeX | OfficeIMO.Latex.Markdown | Semantic | No | `MarkdownReader.Parse(markdown).ToLatexDocumentResult(options)` | MarkdownToLatexResult |
| latex-pdf | LaTeX | PDF | OfficeIMO.Latex.Pdf | FixedLayout | No | `LatexDocument.Parse(source).Document.ToPdfDocumentResult(options)` | PdfDocumentConversionResult |
| onenote-html | OneNote | HTML | OfficeIMO.OneNote.Html | Semantic | No | `section.ToHtmlDocumentResult(projectionOptions, htmlOptions)` | HtmlTextConversionResult |
| html-onenote | HTML | OneNote | OfficeIMO.OneNote.Html | Editable | No | `HtmlConversionDocument.Parse(html).ToOneNoteSectionResult(options)` | HtmlToOneNoteSectionResult |
| onenote-markdown | OneNote | Markdown | OfficeIMO.OneNote.Markdown | Semantic | No | `section.ToMarkdownDocumentResult(options)` | OneNoteMarkdownConversionResult |
| onenote-pdf | OneNote | PDF | OfficeIMO.OneNote.Pdf | FixedLayout | No | `section.ToPdfDocumentResult(options)` | PdfDocumentConversionResult |
| odt-pdf | ODT | PDF | OfficeIMO.OpenDocument.Odt.Pdf | FixedLayout | No | `OdtDocument.Load(stream).ToPdfDocumentResult(conversionOptions, pdfOptions)` | PdfDocumentConversionResult |
| ods-pdf | ODS | PDF | OfficeIMO.OpenDocument.Ods.Pdf | FixedLayout | No | `OdsDocument.Load(stream).ToPdfDocumentResult(conversionOptions, pdfOptions)` | PdfDocumentConversionResult |
| odp-pdf | ODP | PDF | OfficeIMO.OpenDocument.Odp.Pdf | FixedLayout | No | `OdpPresentation.Load(stream).ToPdfDocumentResult(conversionOptions, pdfOptions)` | PdfDocumentConversionResult |
| pdf-docx | PDF | DOCX | OfficeIMO.Word.Pdf | Editable | No | `PdfDocument.Open(stream).ToWordDocumentResult(options)` | PdfWordConversionResult |
| pdf-xlsx | PDF | XLSX | OfficeIMO.Excel.Pdf | Editable | No | `PdfDocument.Open(stream).ImportTablesToExcelDocumentResult(options)` | PdfExcelTableImportResult |
| pdf-pptx | PDF | PPTX | OfficeIMO.PowerPoint.Pdf | Editable | No | `PdfDocument.Open(stream).ToPowerPointPresentationResult(options)` | PdfPowerPointConversionResult |
| pdf-html | PDF | HTML | OfficeIMO.Html.Pdf | Semantic | No | `PdfDocument.Open(stream).ToHtmlResult(options)` | PdfHtmlConversionResult |
| pdf-rtf | PDF | RTF | OfficeIMO.Rtf.Pdf | Editable | No | `PdfDocument.Open(stream).ToRtfDocumentResult(options)` | PdfRtfConversionResult |
| pdf-odt | PDF | ODT | OfficeIMO.OpenDocument.Odt.Pdf | Editable | No | `PdfDocument.Open(stream).ToOdtDocumentResult(pdfOptions, openDocumentOptions)` | PdfOdtConversionResult |
| pdf-ods | PDF | ODS | OfficeIMO.OpenDocument.Ods.Pdf | Editable | No | `PdfDocument.Open(stream).ToOdsDocumentResult(pdfOptions, openDocumentOptions)` | PdfOdsConversionResult |
| pdf-odp | PDF | ODP | OfficeIMO.OpenDocument.Odp.Pdf | Editable | No | `PdfDocument.Open(stream).ToOdpPresentationResult(pdfOptions, openDocumentOptions)` | PdfOdpConversionResult |
| mhtml-pdf | MHTML | PDF | OfficeIMO.Html.Pdf | FixedLayout | No | `MhtmlDocument.Load(stream, options).ToPdfDocumentResult(pdfOptions)` | PdfDocumentConversionResult |
| visio-pdf | Visio | PDF | OfficeIMO.Visio.Pdf | FixedLayout | No | `VisioDocument.Load(stream).ToPdfDocumentResult(options)` | PdfDocumentConversionResult |
