using OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf.PowerPoint;
using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Packaging;
using System.Net;
using System.Text.RegularExpressions;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf.Word;
using PdfCore = OfficeIMO.Pdf;
using A = DocumentFormat.OpenXml.Drawing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Measures equivalent parse, editable projection, and serialization work across
/// deterministic PDFs emitted by each isolated benchmark producer.
/// </summary>
[MemoryDiagnoser]
public class PdfReverseConversionBenchmarks {
    private byte[] _source = Array.Empty<byte>();
    private PdfBenchmarkScenario _scenario = null!;

    [Params(PdfBenchmarkProducer.OfficeIMO, PdfBenchmarkProducer.QuestPDF, PdfBenchmarkProducer.PeachPDF, PdfBenchmarkProducer.MigraDoc, PdfBenchmarkProducer.IText)]
    public PdfBenchmarkProducer Producer { get; set; }

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = PdfBenchmarkScenario.Get(Scale);
        _source = PdfDocumentGenerators.Generate(Producer, _scenario);
        PdfBenchmarkValidation.ValidateGenerated(_source, _scenario, Producer + " reverse-conversion source");
        ValidateOutputs();
    }

    [Benchmark]
    public int PdfToDocx() {
        using var input = new MemoryStream(_source, writable: false);
        using PdfCore.PdfDocument source = PdfCore.PdfDocument.Load(input);
        using OfficeIMO.Word.WordDocument document = source.ToWord();
        return document.ToBytes().Length;
    }

    [Benchmark]
    public int PdfToHtml() {
        using var input = new MemoryStream(_source, writable: false);
        using PdfCore.PdfDocument source = PdfCore.PdfDocument.Load(input);
        return source.ToHtml(new PdfToHtmlOptions { Profile = PdfHtmlProfile.Semantic }).Length;
    }

    [Benchmark]
    public int PdfToXlsx() {
        using var input = new MemoryStream(_source, writable: false);
        using PdfCore.PdfDocument source = PdfCore.PdfDocument.Load(input);
        using var stream = new MemoryStream();
        source.SaveAsExcel(stream);
        return checked((int)stream.Length);
    }

    [Benchmark]
    public int PdfToPptx() {
        using var pptInput = new MemoryStream(_source, writable: false);
        using PdfCore.PdfDocument pptSource = PdfCore.PdfDocument.Load(pptInput);
        using (var presentation = pptSource.ToPowerPoint(PowerPointPdfImportOptions.CreateEditableContent())) {
            using var stream = new MemoryStream();
            presentation.Save(stream);
            return checked((int)stream.Length);
        }
    }

    [Benchmark]
    public int PdfToOdt() {
        using var input = new MemoryStream(_source, writable: false);
        using PdfCore.PdfDocument source = PdfCore.PdfDocument.Load(input);
        var document = source.ToOdtDocumentResult().Value;
        return document.ToBytes().Length;
    }

    [Benchmark]
    public int PdfToOds() {
        using var input = new MemoryStream(_source, writable: false);
        using PdfCore.PdfDocument source = PdfCore.PdfDocument.Load(input);
        var document = source.ToOdsDocumentResult().Value;
        return document.ToBytes().Length;
    }

    [Benchmark]
    public int PdfToOdp() {
        using var input = new MemoryStream(_source, writable: false);
        using PdfCore.PdfDocument source = PdfCore.PdfDocument.Load(input);
        var document = source.ToOdpPresentationResult(PowerPointPdfImportOptions.CreateEditableContent()).Value;
        return document.ToBytes().Length;
    }

    [Benchmark]
    public int PdfToPng() {
        using var input = new MemoryStream(_source, writable: false);
        using PdfCore.PdfDocument source = PdfCore.PdfDocument.Load(input);
        IReadOnlyList<PdfCore.PdfPageRenderResult> pages = source.Render.Pages(
            options: new PdfCore.PdfPageRenderOptions { Format = global::OfficeIMO.Drawing.OfficeImageExportFormat.Png, Dpi = 72D });
        return checked(pages.Sum(static page => page.Bytes?.Length ?? 0));
    }

    private void ValidateOutputs() {
        using var input = new MemoryStream(_source, writable: false);
        using PdfCore.PdfDocument source = PdfCore.PdfDocument.Load(input);
        using (OfficeIMO.Word.WordDocument document = source.ToWord()) {
            using var packageInput = new MemoryStream(document.ToBytes(), writable: false);
            using WordprocessingDocument package = WordprocessingDocument.Open(packageInput, false);
            W.Body body = package.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("DOCX reverse conversion did not produce a document body.");
            ValidateContent(string.Join(" ", body.Descendants<W.Text>().Select(static text => text.Text)), "DOCX");
        }
        string html = source.ToHtml(new PdfToHtmlOptions { Profile = PdfHtmlProfile.Semantic });
        int htmlPages = Regex.Matches(html, "class=\"pdf-page\"[^>]*data-page-number=", RegexOptions.CultureInvariant).Count;
        if (htmlPages != _scenario.PageCount) throw new InvalidOperationException($"HTML reverse conversion retained {htmlPages} of {_scenario.PageCount} page scopes.");
        ValidateContent(WebUtility.HtmlDecode(Regex.Replace(html, "<[^>]+>", " ")), "HTML");
        using (var stream = new MemoryStream()) {
            OfficeConversionReport report = source.SaveAsExcel(stream);
            if (report.ItemsWritten == 0) throw new InvalidOperationException("XLSX reverse conversion did not recover benchmark tables.");
            using var packageInput = new MemoryStream(stream.ToArray(), writable: false);
            using SpreadsheetDocument package = SpreadsheetDocument.Open(packageInput, false);
            if (package.WorkbookPart?.Workbook is null) throw new InvalidOperationException("XLSX reverse conversion did not produce a workbook.");
            if (!package.WorkbookPart.WorksheetParts.SelectMany(static worksheet => worksheet.TableDefinitionParts).Any()) throw new InvalidOperationException("XLSX reverse conversion did not produce editable table definitions.");
            PdfBenchmarkValidation.ValidateTableScenarioContent(GetSpreadsheetText(package), _scenario, Producer + " " + Scale + " XLSX");
        }
        using var pptInput = new MemoryStream(_source, writable: false);
        using PdfCore.PdfDocument pptSource = PdfCore.PdfDocument.Load(pptInput);
        using (var presentation = pptSource.ToPowerPoint(PowerPointPdfImportOptions.CreateEditableContent())) {
            using var stream = new MemoryStream();
            presentation.Save(stream);
            using var packageInput = new MemoryStream(stream.ToArray(), writable: false);
            using PresentationDocument package = PresentationDocument.Open(packageInput, false);
            int slideCount = package.PresentationPart?.SlideParts.Count() ?? 0;
            if (slideCount != _scenario.PageCount) throw new InvalidOperationException($"PPTX reverse conversion produced {slideCount} of {_scenario.PageCount} slides.");
            ValidateContent(string.Join(" ", package.PresentationPart!.SlideParts.SelectMany(static slide => slide.Slide?.Descendants<A.Text>() ?? Enumerable.Empty<A.Text>()).Select(static text => text.Text)), "PPTX");
        }
        var odtOutput = source.ToOdtDocumentResult().Value;
        byte[] odt = odtOutput.ToBytes();
        using var odtInput = new MemoryStream(odt, writable: false);
        OdtDocument odtDocument = OdtDocument.Load(odtInput);
        ValidateContent(ReadOpenDocumentText(odt), "ODT");
        var odsOutput = source.ToOdsDocumentResult().Value;
        byte[] ods = odsOutput.ToBytes();
        using var odsInput = new MemoryStream(ods, writable: false);
        OdsDocument odsDocument = OdsDocument.Load(odsInput);
        if (odsDocument.Sheets.Count == 0) throw new InvalidOperationException("ODS reverse conversion did not produce a sheet.");
        PdfBenchmarkValidation.ValidateTableScenarioContent(ReadOpenDocumentText(ods), _scenario, Producer + " " + Scale + " ODS");
        var odpOutput = source.ToOdpPresentationResult(PowerPointPdfImportOptions.CreateEditableContent()).Value;
        byte[] odp = odpOutput.ToBytes();
        using var odpInput = new MemoryStream(odp, writable: false);
        OdpPresentation odpDocument = OdpPresentation.Load(odpInput);
        if (odpDocument.Slides.Count != _scenario.PageCount) throw new InvalidOperationException($"ODP reverse conversion produced {odpDocument.Slides.Count} of {_scenario.PageCount} slides.");
        ValidateContent(ReadOpenDocumentText(odp), "ODP");
        IReadOnlyList<PdfCore.PdfPageRenderResult> pngPages = source.Render.Pages(
            options: new PdfCore.PdfPageRenderOptions { Format = global::OfficeIMO.Drawing.OfficeImageExportFormat.Png, Dpi = 72D });
        if (pngPages.Count != _scenario.PageCount || pngPages.Any(static page => !page.Succeeded || page.Bytes is null || page.Bytes.Length <= 24)) {
            throw new InvalidOperationException("PNG reverse conversion did not produce one valid visual artifact per source page.");
        }
    }

    private void ValidateContent(string text, string target) =>
        PdfBenchmarkValidation.ValidateScenarioContent(text, _scenario, Producer + " " + Scale + " " + target);

    private static string GetSpreadsheetText(SpreadsheetDocument package) {
        string[] sharedStrings = package.WorkbookPart?.SharedStringTablePart?.SharedStringTable?
            .Elements<S.SharedStringItem>()
            .Select(static item => item.InnerText)
            .ToArray() ?? Array.Empty<string>();
        var values = new List<string>();
        foreach (S.Cell cell in package.WorkbookPart!.WorksheetParts.SelectMany(static worksheet => worksheet.Worksheet?.Descendants<S.Cell>() ?? Enumerable.Empty<S.Cell>())) {
            if (cell.DataType?.Value == S.CellValues.SharedString &&
                int.TryParse(cell.CellValue?.Text, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int sharedIndex) &&
                sharedIndex >= 0 && sharedIndex < sharedStrings.Length) {
                values.Add(sharedStrings[sharedIndex]);
            } else if (cell.InlineString is not null) {
                values.Add(cell.InlineString.InnerText);
            } else if (cell.CellValue is not null) {
                values.Add(cell.CellValue.Text);
            }
        }
        return string.Join(" ", values);
    }

    private static string ReadOpenDocumentText(byte[] artifact) {
        using var input = new MemoryStream(artifact, writable: false);
        using var archive = new System.IO.Compression.ZipArchive(input, System.IO.Compression.ZipArchiveMode.Read);
        System.IO.Compression.ZipArchiveEntry content = archive.GetEntry("content.xml")
            ?? throw new InvalidDataException("OpenDocument package did not contain content.xml.");
        using var reader = new StreamReader(content.Open());
        return Regex.Replace(reader.ReadToEnd(), "<[^>]+>", " ");
    }
}
