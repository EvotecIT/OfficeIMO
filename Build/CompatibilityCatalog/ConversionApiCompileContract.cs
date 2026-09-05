using OfficeIMO.Adf;
using OfficeIMO.AsciiDoc;
using OfficeIMO.Confluence;
using OfficeIMO.CSV;
using OfficeIMO.Drawing;
using OfficeIMO.Email;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.Excel.Csv;
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Pdf.Browser;
using OfficeIMO.Epub;
using OfficeIMO.Latex;
using OfficeIMO.Markdown;
using OfficeIMO.Markup;
using OfficeIMO.Mhtml;
using OfficeIMO.OneNote;
using OfficeIMO.OpenDocument;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.PowerPoint.GoogleSlides;
using OfficeIMO.Rtf;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.GoogleDocs;
using HtmlTinkerX;

internal static class ConversionApiCompileContract {
    // This method is intentionally never executed. Its body makes every representative
    // catalog entry a compile-time dependency, so generated API guidance cannot drift
    // into a plausible-looking but invalid signature.
    internal static void Verify(
        Stream stream,
        string source,
        AdfDocument adf,
        AsciiDocDocument asciiDoc,
        ConfluencePage confluencePage,
        CsvDocument csv,
        EmailDocument email,
        global::OfficeIMO.Epub.EpubDocument epub,
        ExcelDocument excel,
        HtmlConversionDocument html,
        HtmlBrowserPdfResult browserPdf,
        LatexDocument latex,
        MarkdownDoc markdown,
        OfficeMarkupDocument officeMarkup,
        MhtmlDocument mhtml,
        OneNoteSection oneNote,
        OdtDocument odt,
        OdsDocument ods,
        OdpPresentation odp,
        PdfDocument pdf,
        PowerPointPresentation powerPoint,
        RtfDocument rtf,
        RtfReadResult rtfRead,
        VisioDocument visio,
        WordDocument word,
        GoogleWorkspaceSession session) {
        OfficeHtmlDocumentOptions output = new OfficeHtmlDocumentOptions {
            EmitDocumentShell = true,
            IncludeDefaultStyles = true,
            Language = "en",
            NewLine = "\n"
        };
        WordToHtmlOptions wordOptions = WordToHtmlOptions.CreateDocumentRoundTripProfile();
        wordOptions.DocumentOutput = output.Clone();
        ExcelHtmlSaveOptions excelOptions = ExcelHtmlSaveOptions.CreateVisualReviewProfile();
        excelOptions.DocumentOutput = output.Clone();
        PowerPointHtmlSaveOptions powerPointOptions = PowerPointHtmlSaveOptions.CreateVisualReviewProfile();
        powerPointOptions.DocumentOutput = output.Clone();
        RtfToHtmlOptions rtfOptions = RtfToHtmlOptions.CreatePrintReviewProfile();
        rtfOptions.DocumentOutput = output.Clone();
        PdfToHtmlOptions pdfOptions = PdfToHtmlOptions.CreatePositionedReviewProfile();
        pdfOptions.DocumentOutput = output.Clone();
        HtmlConversionProfile wordSharedProfile = wordOptions.SharedProfile;
        HtmlConversionProfile excelSharedProfile = excelOptions.SharedProfile;
        HtmlConversionProfile powerPointSharedProfile = powerPointOptions.SharedProfile;
        HtmlConversionProfile rtfSharedProfile = rtfOptions.SharedProfile;
        HtmlTargetCapabilityContract target = HtmlTargetCapabilityContracts.Get(HtmlConversionTarget.Pdf);
        HtmlToTargetCapabilityContract htmlToTarget = target.HtmlToTarget;
        TargetToHtmlCapabilityContract? targetToHtml = target.TargetToHtml;
        _ = htmlToTarget.DiagnosticsContract;
        _ = targetToHtml?.DiagnosticsContract;
        _ = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(source);
        _ = OfficeIMO.AsciiDoc.Markdown.AsciiDocMarkdownConverterExtensions.ToMarkdownDocumentResult(asciiDoc);
        _ = OfficeIMO.AsciiDoc.Markdown.AsciiDocMarkdownConverterExtensions.ToAsciiDocDocumentResult(markdown);
        _ = OfficeIMO.AsciiDoc.Pdf.AsciiDocPdfConverterExtensions.ToPdfDocumentResult(asciiDoc);
        _ = OfficeIMO.Latex.Markdown.LatexMarkdownConverterExtensions.ToMarkdownDocumentResult(latex);
        _ = OfficeIMO.Latex.Markdown.LatexMarkdownConverterExtensions.ToLatexDocumentResult(markdown);
        _ = OfficeIMO.Latex.Pdf.LatexPdfConverterExtensions.ToPdfDocumentResult(latex);
        _ = OfficeIMO.Markdown.Html.HtmlMarkdownConverterExtensions.ToMarkdownDocumentResult(html);
        _ = OfficeIMO.Markdown.Pdf.MarkdownPdfConverterExtensions.ToPdfDocumentResult(markdown);
        _ = OfficeIMO.Rtf.Markdown.RtfMarkdownConverterExtensions.ToMarkdownResult(rtf);
        _ = OfficeIMO.Rtf.Markdown.RtfMarkdownConverterExtensions.ToRtfDocumentResult(markdown);
        _ = OfficeIMO.Word.Markdown.WordMarkdownConverterExtensions.ToMarkdownDocumentResult(word);
        _ = OfficeIMO.Word.Markdown.WordMarkdownConverterExtensions.ToWordDocumentResult(markdown);
        _ = OfficeIMO.Word.OpenDocument.WordOpenDocumentConversionExtensions.ToOpenDocumentResult(word);
        _ = OfficeIMO.Word.OpenDocument.WordOpenDocumentConversionExtensions.ToWordDocumentResult(odt);
        _ = OfficeIMO.Html.HtmlRtfConverterExtensions.ToRtfDocumentResult(html);
        _ = OfficeIMO.Html.HtmlRtfConverterExtensions.ToHtmlResult(rtf);
        _ = OfficeIMO.Word.Rtf.WordRtfConverterExtensions.ToRtfDocumentResult(word);
        _ = OfficeIMO.Word.Rtf.WordRtfConverterExtensions.ToWordDocumentResult(rtfRead);
        _ = OfficeIMO.Excel.Html.ExcelHtmlConverterExtensions.ToHtmlResult(excel);
        _ = OfficeIMO.Excel.Html.HtmlExcelConverterExtensions.ToExcelDocumentResult(html);
        _ = OfficeIMO.Excel.OpenDocument.ExcelOpenDocumentConversionExtensions.ToOpenDocumentResult(excel);
        _ = OfficeIMO.Excel.OpenDocument.ExcelOpenDocumentConversionExtensions.ToExcelDocumentResult(ods);
        _ = OfficeIMO.Word.Html.WordHtmlConverterExtensions.ToHtmlResult(word);
        _ = OfficeIMO.Word.Html.WordHtmlConverterExtensions.ToWordDocumentResult(html);
        _ = OfficeIMO.PowerPoint.Html.PowerPointHtmlConverterExtensions.ToHtmlResult(powerPoint);
        _ = OfficeIMO.PowerPoint.Html.HtmlPowerPointConverterExtensions.ToPowerPointPresentationResult(html);
        _ = OfficeIMO.PowerPoint.OpenDocument.PowerPointOpenDocumentConversionExtensions.ToOpenDocumentResult(powerPoint);
        _ = OfficeIMO.PowerPoint.OpenDocument.PowerPointOpenDocumentConversionExtensions.ToPowerPointPresentationResult(odp);
        _ = OfficeIMO.OneNote.Markdown.OneNoteMarkdownConverterExtensions.ToMarkdownDocumentResult(oneNote);
        _ = OfficeIMO.OneNote.Html.OneNoteHtmlConverterExtensions.ToHtmlDocumentResult(oneNote);
        _ = OfficeIMO.OneNote.Html.HtmlOneNoteConverterExtensions.ToOneNoteSectionResult(html);
        _ = OfficeIMO.OneNote.Pdf.OneNoteSectionPdfConverterExtensions.ToPdfDocumentResult(oneNote);
        _ = OfficeIMO.Word.Pdf.WordPdfConverterExtensions.ToPdfDocumentResult(word);
        _ = OfficeIMO.Word.Pdf.PdfWordConverterExtensions.ToWordDocumentResult(pdf);
        _ = OfficeIMO.Excel.Pdf.ExcelPdfConverterExtensions.ToPdfDocumentResult(excel);
        _ = OfficeIMO.Excel.Pdf.PdfExcelTableConverterExtensions.ImportTablesToExcelDocumentResult(pdf);
        _ = OfficeIMO.PowerPoint.Pdf.PowerPointPdfConverterExtensions.ToPdfDocumentResult(powerPoint);
        _ = OfficeIMO.PowerPoint.Pdf.PowerPointPdfConverterExtensions.ToPowerPointPresentationResult(pdf);
        _ = OfficeIMO.Html.Pdf.HtmlPdfConverterExtensions.ToPdfDocumentResult(html);
        _ = browserPdf.ToPdfDocumentResult();
        _ = OfficeIMO.Mhtml.MhtmlPdfConverterExtensions.ToPdfDocumentResult(mhtml);
        PdfHtmlConversionResult pdfHtmlResult = OfficeIMO.Html.Pdf.PdfHtmlConverterExtensions.ToHtmlResult(pdf);
        PdfConversionReport pdfHtmlReport = pdfHtmlResult.Report;
        PdfConversionReport pdfHtmlSaveReport = OfficeIMO.Html.Pdf.PdfHtmlConverterExtensions.SaveAsHtml(pdf, stream, pdfOptions);
        _ = pdfHtmlReport.Warnings;
        _ = pdfHtmlSaveReport.Warnings;
        var galleryResult = new HtmlCapabilityGalleryResult(
            new HtmlCapabilityGalleryScenario("compat", "Compatibility", "HTML", "Established gallery API"));
        galleryResult.AddArtifact(new HtmlCapabilityGalleryArtifact(
            "source", "html", "source.html", "text/html", 1, new string('0', 64)));
        galleryResult.Diagnostics.Add(new HtmlDiagnostic("Compatibility", "EstablishedApi", "Established gallery diagnostics API"));
        _ = OfficeIMO.Rtf.Pdf.RtfPdfConverterExtensions.ToPdfDocumentResult(rtf);
        _ = OfficeIMO.Rtf.Pdf.RtfPdfConverterExtensions.ToRtfDocumentResult(pdf);
        _ = OfficeIMO.OpenDocument.Odt.Pdf.OdtPdfConversionExtensions.ToPdfDocumentResult(odt);
        _ = OfficeIMO.OpenDocument.Odt.Pdf.OdtPdfConversionExtensions.ToOdtDocumentResult(pdf);
        _ = OfficeIMO.OpenDocument.Ods.Pdf.OdsPdfConversionExtensions.ToPdfDocumentResult(ods);
        _ = OfficeIMO.OpenDocument.Ods.Pdf.OdsPdfConversionExtensions.ToOdsDocumentResult(pdf);
        _ = OfficeIMO.OpenDocument.Odp.Pdf.OdpPdfConversionExtensions.ToPdfDocumentResult(odp);
        _ = OfficeIMO.OpenDocument.Odp.Pdf.OdpPdfConversionExtensions.ToOdpPresentationResult(pdf);
        _ = OfficeIMO.Visio.Pdf.VisioPdfConverterExtensions.ToPdfDocumentResult(visio);
        _ = WordGoogleDocsExtensions.ExportToGoogleDocsAsync(word, session);
        _ = WordGoogleDocsExtensions.ImportGoogleDocAsync(session, "document-id");
        _ = ExcelGoogleSheetsExtensions.ExportToGoogleSheetsAsync(excel, session);
        _ = ExcelGoogleSheetsExtensions.ImportGoogleSheetAsync(session, "spreadsheet-id");
        _ = PowerPointGoogleSlidesExtensions.ExportToGoogleSlidesAsync(powerPoint, session);
        _ = PowerPointGoogleSlidesExtensions.ImportGoogleSlidesAsync(session, "presentation-id");
        _ = AdfConverter.ToMarkdown(adf);
        _ = AdfConverter.FromMarkdown(source);
        _ = AdfConverter.ToHtml(adf);
        _ = AdfConverter.FromHtml(source);
        _ = ConfluenceContentConverter.FromMarkdown(source);
        _ = ConfluenceContentConverter.FromHtml(source);
        _ = ConfluenceContentConverter.ToMarkdown(confluencePage);
        _ = ConfluenceContentConverter.ToHtml(confluencePage);
        _ = ExcelDocumentCsvExtensions.ToExcelDocument(csv);
        _ = ExcelSheetCsvExtensions.ToCsv(excel.Sheets[0]);
        _ = OfficeIMO.Markup.Word.OfficeMarkupWordConverterExtensions.ToWordDocumentResult(officeMarkup);
        _ = OfficeIMO.Markup.Excel.OfficeMarkupExcelConverterExtensions.ToExcelDocumentResult(officeMarkup);
        _ = OfficeIMO.Markup.PowerPoint.OfficeMarkupPowerPointConverterExtensions.ToPowerPointPresentationResult(officeMarkup);

        // Every cataloged image source compiles against the same five-format owner
        // contract. Runtime tests enumerate PNG, SVG, JPEG, TIFF, and WebP.
        const OfficeImageExportFormat imageFormat = OfficeImageExportFormat.Png;
        _ = word.ExportImages(imageFormat);
        _ = excel.ExportImages(imageFormat);
        _ = powerPoint.ExportImages(imageFormat);
        _ = OfficeIMO.Html.HtmlImageExportExtensions.ExportImages(html, imageFormat);
        _ = OfficeIMO.OneNote.OneNoteImageExportExtensions.ExportImages(oneNote, imageFormat);
        _ = OfficeIMO.Visio.VisioImageExportExtensions.ExportImages(visio, imageFormat);
        _ = OfficeIMO.Email.EmailImageExportExtensions.ExportImages(email, imageFormat);
        _ = OfficeIMO.Epub.Image.EpubImageExportExtensions.ExportImages(epub, imageFormat);
        _ = OfficeIMO.Word.OpenDocument.WordOpenDocumentImageExportExtensions.ExportImages(odt, imageFormat);
        _ = OfficeIMO.Excel.OpenDocument.ExcelOpenDocumentImageExportExtensions.ExportImages(ods, imageFormat);
        _ = OfficeIMO.PowerPoint.OpenDocument.PowerPointOpenDocumentImageExportExtensions.ExportImages(odp, imageFormat);
        _ = pdf.Render.ExportImages(imageFormat);
    }
}
