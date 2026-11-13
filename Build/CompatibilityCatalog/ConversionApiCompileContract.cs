using OfficeIMO.Rtf.Pdf;
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
        _ = OfficeIMO.AsciiDoc.Markdown.AsciiDocMarkdownConverterExtensions.ToMarkdown(asciiDoc);
        _ = OfficeIMO.AsciiDoc.Markdown.AsciiDocMarkdownConverterExtensions.ToAsciiDoc(markdown);
        _ = OfficeIMO.AsciiDoc.Pdf.AsciiDocPdfConverterExtensions.ToPdf(asciiDoc);
        _ = OfficeIMO.Latex.Markdown.LatexMarkdownConverterExtensions.ToMarkdown(latex);
        _ = OfficeIMO.Latex.Markdown.LatexMarkdownConverterExtensions.ToLatex(markdown);
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
        using RtfDocument pdfRtf = OfficeIMO.Rtf.Pdf.RtfPdfConverterExtensions.ToRtf(pdf, null, out OfficeConversionReport pdfRtfReport);
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
        using ExcelDocument csvExcel = ExcelDocumentCsvExtensions.ToExcel(csv, null, out OfficeConversionReport csvExcelReport);
        using CsvDocument excelCsv = ExcelSheetCsvExtensions.ToCsv(excel.Sheets[0]);
        var dataOptions = new OfficeIMO.Data.Conversions.DataConversionOptions { JsonLines = true };
        _ = OfficeIMO.Data.Conversions.YamlDataConversions.SaveAsJson(OfficeIMO.Yaml.YamlDocument.Load(stream), stream, dataOptions);
        _ = OfficeIMO.Data.Conversions.XmlDataConversions.SaveAsCsv(System.Xml.Linq.XDocument.Load(stream), stream, dataOptions);
        _ = OfficeIMO.Data.Conversions.CsvDataConversions.SaveAsXml(csv, stream, dataOptions);
        _ = OfficeIMO.Data.Conversions.JsonDataConversions.SaveAsYaml(System.Text.Json.Nodes.JsonNode.Parse(stream), stream, dataOptions);
        _ = OfficeIMO.Data.Conversions.JsonDataConversions.SaveAsCsv(System.Text.Json.Nodes.JsonNode.Parse(stream), stream, dataOptions);
        _ = OfficeIMO.Data.Conversions.ExcelDataConversions.SaveAsExcel(System.Text.Json.Nodes.JsonNode.Parse(stream), stream, dataOptions);
        _ = OfficeIMO.Data.Conversions.ExcelDataConversions.SaveAsJson(excel, stream, dataOptions);
        using var dataRows = OfficeIMO.Data.Conversions.StructuredDataReader.Open(stream, OfficeDocumentFormat.Json);
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
        using WordDocument markupWord = OfficeIMO.Markup.Word.OfficeMarkupWordConverterExtensions.ToWord(officeMarkup, null, out OfficeConversionReport markupWordReport);
        using ExcelDocument markupExcel = OfficeIMO.Markup.Excel.OfficeMarkupExcelConverterExtensions.ToExcel(officeMarkup, null, out OfficeConversionReport markupExcelReport);
        using PowerPointDocument markupPowerPoint = OfficeIMO.Markup.PowerPoint.OfficeMarkupPowerPointConverterExtensions.ToPowerPoint(officeMarkup, null, out OfficeConversionReport markupPowerPointReport);

        // Every cataloged image source compiles against the same five-format owner
        // contract. Runtime tests enumerate PNG, SVG, JPEG, TIFF, and WebP.
        const OfficeImageExportFormat imageFormat = OfficeImageExportFormat.Png;
        _ = word.ToImages(imageFormat);
        _ = excel.ToImages(imageFormat);
        _ = powerPoint.ToImages(imageFormat);
        _ = OfficeIMO.Html.Pdf.HtmlImageExportExtensions.ToImages(html, imageFormat);
        _ = OfficeIMO.OneNote.OneNoteImageExportExtensions.ToImages(oneNote, imageFormat);
        _ = OfficeIMO.Visio.VisioImageExportExtensions.ToImages(visio, imageFormat);
        _ = OfficeIMO.Email.EmailImageExportExtensions.ToImages(email, imageFormat);
        _ = OfficeIMO.Epub.Image.EpubImageExportExtensions.ToImages(epub, imageFormat);
        _ = OfficeIMO.Word.OpenDocument.WordOpenDocumentImageExportExtensions.ToImages(odt, imageFormat);
        _ = OfficeIMO.Excel.OpenDocument.ExcelOpenDocumentImageExportExtensions.ToImages(ods, imageFormat);
        _ = OfficeIMO.PowerPoint.OpenDocument.PowerPointOpenDocumentImageExportExtensions.ToImages(odp, imageFormat);
        _ = pdf.ToImages(new PdfImageExportOptions { Format = imageFormat }, out _);
        _ = rtf.ToImages(new PdfImageExportOptions { Format = imageFormat });
        _ = mhtml.ToImages(new PdfImageExportOptions { Format = imageFormat });
        _ = odg.ToImages(new OfficeImageExportOptions { Format = imageFormat });
        _ = publisher.ToImages(new OfficeImageExportOptions { Format = imageFormat });
        _ = DjVuDocument.Load(stream).ToImages(new OfficeImageExportOptions { Format = imageFormat });
        _ = OfficeIMO.Xps.XpsDocumentImageConversionExtensions.ToImages(OfficeIMO.Xps.XpsDocument.Load(stream), new OfficeImageExportOptions { Format = imageFormat });
    }
}
