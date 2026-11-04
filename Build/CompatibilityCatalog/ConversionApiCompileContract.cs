using OfficeIMO.Pdf.Word;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf.PowerPoint;
using OfficeIMO;
using OfficeIMO.Adf;
using DBAClientX.Dbf;
using OfficeIMO.AsciiDoc;
using OfficeIMO.Confluence;
using OfficeIMO.CSV;
using OfficeIMO.Chm;
using OfficeIMO.Drawing;
using OfficeIMO.DjVu;
using OfficeIMO.DjVu.Pdf;
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
using OfficeIMO.Publisher;
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
        ChmDocument chm,
        EmailDocument email,
        global::OfficeIMO.Epub.EpubDocument epub,
        ExcelDocument excel,
        HtmlConversionDocument html,
        HtmlBrowserPdfResult browserPdf,
        LatexDocument latex,
        MarkdownDocument markdown,
        OfficeMarkupDocument officeMarkup,
        MhtmlDocument mhtml,
        OneNoteSection oneNote,
        OdtDocument odt,
        OdsDocument ods,
        OdpPresentation odp,
        OdgDocument odg,
        PdfDocument pdf,
        PowerPointDocument powerPoint,
        PublisherDocument publisher,
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
        _ = DjVuDocument.Load(stream).SaveAsPdf(stream, new DjVuPdfExportOptions());
        WordHtmlExportOptions wordOptions = WordHtmlExportOptions.CreateDocumentRoundTripProfile();
        wordOptions.DocumentOutput = output.Clone();
        ExcelHtmlExportOptions excelOptions = ExcelHtmlExportOptions.CreateVisualReviewProfile();
        excelOptions.DocumentOutput = output.Clone();
        PowerPointHtmlExportOptions powerPointOptions = PowerPointHtmlExportOptions.CreateVisualReviewProfile();
        powerPointOptions.DocumentOutput = output.Clone();
        RtfHtmlExportOptions rtfOptions = RtfHtmlExportOptions.CreatePrintReviewProfile();
        rtfOptions.DocumentOutput = output.Clone();
        PdfHtmlExportOptions pdfOptions = PdfHtmlExportOptions.CreatePositionedReviewProfile();
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
        _ = OfficeIMO.AsciiDoc.Pdf.AsciiDocPdfConverterExtensions.ToPdf(asciiDoc);
        _ = OfficeIMO.Latex.Markdown.LatexMarkdownConverterExtensions.ToMarkdownDocumentResult(latex);
        _ = OfficeIMO.Latex.Markdown.LatexMarkdownConverterExtensions.ToLatexDocumentResult(markdown);
        _ = OfficeIMO.Latex.Pdf.LatexPdfConverterExtensions.ToPdf(latex);
        _ = OfficeIMO.Markdown.Html.HtmlMarkdownConverterExtensions.ToMarkdown(html);
        _ = OfficeIMO.Markdown.Pdf.MarkdownPdfConverterExtensions.ToPdf(markdown);
        _ = OfficeIMO.Rtf.Markdown.RtfMarkdownConverterExtensions.ToMarkdown(rtf);
        _ = OfficeIMO.Rtf.Markdown.RtfMarkdownConverterExtensions.ToRtf(markdown);
        _ = OfficeIMO.Word.Markdown.WordMarkdownConverterExtensions.ToMarkdown(word);
        _ = OfficeIMO.Word.Markdown.WordMarkdownConverterExtensions.ToWord(markdown);
        _ = OfficeIMO.Word.OpenDocument.WordOpenDocumentConversionExtensions.ToOpenDocument(word);
        _ = OfficeIMO.Word.OpenDocument.WordOpenDocumentConversionExtensions.ToWord(odt);
        _ = OfficeIMO.Html.HtmlRtfConverterExtensions.ToRtf(html);
        _ = OfficeIMO.Html.HtmlRtfConverterExtensions.ToHtml(rtf);
        _ = OfficeIMO.Word.Rtf.WordRtfConverterExtensions.ToRtf(word);
        _ = OfficeIMO.Word.Rtf.WordRtfConverterExtensions.ToWord(rtfRead);
        _ = OfficeIMO.Excel.Html.ExcelHtmlConverterExtensions.ToHtml(excel);
        _ = OfficeIMO.Excel.Html.HtmlExcelConverterExtensions.ToExcel(html);
        _ = OfficeIMO.Excel.OpenDocument.ExcelOpenDocumentConversionExtensions.ToOpenDocument(excel);
        _ = OfficeIMO.Excel.OpenDocument.ExcelOpenDocumentConversionExtensions.ToExcel(ods);
        _ = OfficeIMO.Word.Html.WordHtmlConverterExtensions.ToHtml(word);
        _ = OfficeIMO.Word.Html.WordHtmlConverterExtensions.ToWord(html);
        _ = OfficeIMO.PowerPoint.Html.PowerPointHtmlConverterExtensions.ToHtml(powerPoint, null, out OfficeConversionReport powerPointHtmlReport);
        _ = OfficeIMO.PowerPoint.Html.HtmlPowerPointConverterExtensions.ToPowerPoint(html);
        _ = OfficeIMO.PowerPoint.OpenDocument.PowerPointOpenDocumentConversionExtensions.ToOpenDocument(powerPoint);
        _ = OfficeIMO.PowerPoint.OpenDocument.PowerPointOpenDocumentConversionExtensions.ToPowerPoint(odp);
        _ = OfficeIMO.OneNote.Markdown.OneNoteMarkdownConverterExtensions.ToMarkdown(oneNote);
        _ = OfficeIMO.OneNote.Html.OneNoteHtmlConverterExtensions.ToHtml(oneNote);
        _ = OfficeIMO.OneNote.Html.HtmlOneNoteConverterExtensions.ToOneNoteSection(html);
        _ = OfficeIMO.OneNote.Pdf.OneNoteSectionPdfConverterExtensions.ToPdf(oneNote);
        _ = OfficeIMO.Word.Pdf.WordPdfConverterExtensions.ToPdf(word);
        _ = OfficeIMO.Pdf.Word.PdfWordConverterExtensions.ToWord(pdf);
        _ = OfficeIMO.Excel.Pdf.ExcelPdfConverterExtensions.ToPdf(excel);
        _ = OfficeIMO.Pdf.Excel.PdfExcelTableConverterExtensions.ToExcel(pdf);
        _ = OfficeIMO.PowerPoint.Pdf.PowerPointPdfConverterExtensions.ToPdf(powerPoint);
        _ = OfficeIMO.Pdf.PowerPoint.PdfPowerPointConverterExtensions.ToPowerPoint(pdf);
        _ = OfficeIMO.Html.Pdf.HtmlPdfConverterExtensions.ToPdf(html);
        _ = browserPdf.ToPdfDocumentResult();
        _ = OfficeIMO.Mhtml.MhtmlPdfConverterExtensions.ToPdf(mhtml);
        _ = ChmDocument.Load(stream).ToPdf();
        _ = chm.ToMarkdown();
        _ = chm.ToEpub();
        string pdfHtml = OfficeIMO.Html.Pdf.PdfHtmlConverterExtensions.ToHtml(pdf, null, out OfficeConversionReport pdfHtmlReport);
        OfficeConversionReport pdfHtmlOutput = OfficeIMO.Html.Pdf.PdfHtmlConverterExtensions.SaveAsHtml(pdf, stream, pdfOptions);
        _ = pdfHtmlReport.Diagnostics;
        _ = pdfHtmlOutput.Diagnostics;
        var galleryResult = new HtmlCapabilityGalleryResult(
            new HtmlCapabilityGalleryScenario("compat", "Compatibility", "HTML", "Established gallery API"));
        galleryResult.AddArtifact(new HtmlCapabilityGalleryArtifact(
            "source", "html", "source.html", "text/html", 1, new string('0', 64)));
        galleryResult.Diagnostics.Add(new HtmlDiagnostic("Compatibility", "EstablishedApi", "Established gallery diagnostics API"));
        _ = OfficeIMO.Rtf.Pdf.RtfPdfConverterExtensions.ToPdf(rtf);
        _ = OfficeIMO.Rtf.Pdf.RtfPdfConverterExtensions.ToRtfDocumentResult(pdf);
        _ = OfficeIMO.OpenDocument.Odg.Pdf.OdgPdfConversionExtensions.ToPdf(odg);
        _ = OfficeIMO.OpenDocument.Odt.Pdf.OdtPdfConversionExtensions.ToPdf(odt);
        _ = OfficeIMO.OpenDocument.Odt.Pdf.OdtPdfConversionExtensions.ToOdt(pdf);
        _ = OfficeIMO.OpenDocument.Ods.Pdf.OdsPdfConversionExtensions.ToPdf(ods);
        _ = OfficeIMO.OpenDocument.Ods.Pdf.OdsPdfConversionExtensions.ToOds(pdf);
        _ = OfficeIMO.OpenDocument.Odp.Pdf.OdpPdfConversionExtensions.ToPdf(odp);
        _ = OfficeIMO.OpenDocument.Odp.Pdf.OdpPdfConversionExtensions.ToOdp(pdf);
        OfficeConversionResult<VisioDocument, OfficeLegacyImportReport> binaryVisio = VisioDocument.LoadLegacyBinary(stream);
        OfficeLegacyImportReport binaryImportReport = binaryVisio.Report;
        binaryVisio.Value.Save(stream);
        OfficeConversionResult<VisioDocument, OfficeLegacyImportReport> binaryStencil = VisioDocument.LoadLegacyBinary(stream, VisioPackageType.Stencil);
        OfficeLegacyImportReport stencilImportReport = binaryStencil.Report;
        binaryStencil.Value.Save(stream);
        OfficeConversionResult<VisioDocument, OfficeLegacyImportReport> binaryTemplate = VisioDocument.LoadLegacyBinary(stream, VisioPackageType.Template);
        OfficeLegacyImportReport templateImportReport = binaryTemplate.Report;
        binaryTemplate.Value.Save(stream);
        OfficeConversionResult<VisioDocument, OfficeLegacyImportReport> binaryPageSource = VisioDocument.LoadLegacyBinary(source);
        OfficeLegacyImportReport pageSourceImportReport = binaryPageSource.Report;
        string binarySvg = binaryPageSource.Value.Pages[0].ToSvg();
        PdfDocument binaryPdf = OfficeIMO.Visio.Pdf.VisioPdfConverterExtensions.ToPdf(binaryPageSource.Value,
            new OfficeIMO.Visio.Pdf.VisioPdfExportOptions { Mode = OfficeIMO.Visio.Pdf.VisioPdfProjectionMode.DiagramPages }, out OfficeConversionReport binaryPdfReport);
        _ = OfficeIMO.Visio.Pdf.VisioPdfConverterExtensions.ToPdf(visio);
        _ = OfficeIMO.Publisher.Pdf.PublisherPdfConversionExtensions.ToPdf(publisher);
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
        using (DbfDataReader dbf = DbfDataReader.Open(source))
        using (StringWriter csvOutput = new StringWriter()) {
            CsvDocument.WriteDataReader(csvOutput, dbf);
        }
        using (DbfDataReader dbf = DbfDataReader.Open(source)) {
            var destination1 = stream;

            var writeOptions1 = new ExcelWriterOptions {
                UseSharedStrings = false,
                CreateTable = false
            };

            using (var workbookWriter1 = ExcelDocument.CreateWriter(destination1, writeOptions1, leaveOpen: true)) {
                _ = workbookWriter1.WriteTable(writeOptions1.SheetName, dbf);
            }

            if (destination1.CanSeek)
                destination1.Position = 0;
        }
        _ = OfficeIMO.Markup.Word.OfficeMarkupWordConverterExtensions.ToWordDocumentResult(officeMarkup);
        _ = OfficeIMO.Markup.Excel.OfficeMarkupExcelConverterExtensions.ToExcelDocumentResult(officeMarkup);
        _ = OfficeIMO.Markup.PowerPoint.OfficeMarkupPowerPointConverterExtensions.ToPowerPointDocumentResult(officeMarkup);

        // Every cataloged image source compiles against the same five-format owner
        // contract. Runtime tests enumerate PNG, SVG, JPEG, TIFF, and WebP.
        const OfficeImageExportFormat imageFormat = OfficeImageExportFormat.Png;
        _ = word.ExportImages(imageFormat);
        _ = excel.ExportImages(imageFormat);
        _ = powerPoint.ExportImages(imageFormat);
        _ = OfficeIMO.Html.Pdf.HtmlImageExportExtensions.ExportImages(html, imageFormat);
        _ = OfficeIMO.OneNote.OneNoteImageExportExtensions.ExportImages(oneNote, imageFormat);
        _ = OfficeIMO.Visio.VisioImageExportExtensions.ExportImages(visio, imageFormat);
        _ = OfficeIMO.Email.EmailImageExportExtensions.ExportImages(email, imageFormat);
        _ = OfficeIMO.Epub.Image.EpubImageExportExtensions.ExportImages(epub, imageFormat);
        _ = OfficeIMO.Word.OpenDocument.WordOpenDocumentImageExportExtensions.ExportImages(odt, imageFormat);
        _ = OfficeIMO.Excel.OpenDocument.ExcelOpenDocumentImageExportExtensions.ExportImages(ods, imageFormat);
        _ = OfficeIMO.PowerPoint.OpenDocument.PowerPointOpenDocumentImageExportExtensions.ExportImages(odp, imageFormat);
        _ = pdf.ExportImages(new PdfImageExportOptions { Format = imageFormat }, out _);
    }
}
