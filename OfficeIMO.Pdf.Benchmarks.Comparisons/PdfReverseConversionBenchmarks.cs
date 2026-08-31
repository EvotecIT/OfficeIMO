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
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocument.Load(_source).Read();
        using OfficeIMO.Word.WordDocument document = logical.ToWordDocument();
        return document.ToBytes().Length;
    }

    [Benchmark]
    public int PdfToHtml() {
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocument.Load(_source).Read();
        return logical.ToHtml(new PdfHtmlSaveOptions { Profile = PdfHtmlProfile.Semantic }).Length;
    }

    [Benchmark]
    public int PdfToXlsx() {
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocument.Load(_source).Read();
        using var stream = new MemoryStream();
        logical.SaveTablesAsExcel(stream);
        return checked((int)stream.Length);
    }

    [Benchmark]
    public int PdfToPptx() {
        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(_source)
            .ToPowerPointPresentationResult(PdfPowerPointImportOptions.CreateEditableContent());
        using (result.Value) {
            using var stream = new MemoryStream();
            result.Value.Save(stream);
            return checked((int)stream.Length);
        }
    }

    [Benchmark]
    public int PdfToOdt() {
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocument.Load(_source).Read();
        return logical.ToOdtDocumentResult().Value.ToBytes().Length;
    }

    [Benchmark]
    public int PdfToOds() {
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocument.Load(_source).Read();
        return logical.ToOdsDocumentResult().Value.ToBytes().Length;
    }

    [Benchmark]
    public int PdfToOdp() {
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocument.Load(_source).Read();
        return logical.ToOdpPresentationResult(PdfPowerPointImportOptions.CreateEditableContent()).Value.ToBytes().Length;
    }

    [Benchmark]
    public int PdfToPng() {
        IReadOnlyList<PdfCore.PdfPageRenderResult> pages = PdfCore.PdfDocument.Load(_source).Render.Pages(
            options: new PdfCore.PdfPageRenderOptions { Format = PdfCore.PdfPageRenderFormat.Png, Dpi = 72D });
        return checked(pages.Sum(static page => page.Bytes?.Length ?? 0));
    }

    private void ValidateOutputs() {
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocument.Load(_source).Read();
        using (OfficeIMO.Word.WordDocument document = logical.ToWordDocument()) {
            using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(document.ToBytes()), false);
            W.Body body = package.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("DOCX reverse conversion did not produce a document body.");
            ValidateContent(string.Join(" ", body.Descendants<W.Text>().Select(static text => text.Text)), "DOCX");
        }
        string html = logical.ToHtml(new PdfHtmlSaveOptions { Profile = PdfHtmlProfile.Semantic });
        int htmlPages = Regex.Matches(html, "class=\"pdf-page\"[^>]*data-page-number=", RegexOptions.CultureInvariant).Count;
        if (htmlPages != _scenario.PageCount) throw new InvalidOperationException($"HTML reverse conversion retained {htmlPages} of {_scenario.PageCount} page scopes.");
        ValidateContent(WebUtility.HtmlDecode(Regex.Replace(html, "<[^>]+>", " ")), "HTML");
        using (var stream = new MemoryStream()) {
            PdfExcelTableImportReport report = logical.SaveTablesAsExcel(stream);
            if (report.Entries.Count == 0) throw new InvalidOperationException("XLSX reverse conversion did not recover benchmark tables.");
            using SpreadsheetDocument package = SpreadsheetDocument.Open(new MemoryStream(stream.ToArray()), false);
            if (package.WorkbookPart?.Workbook is null) throw new InvalidOperationException("XLSX reverse conversion did not produce a workbook.");
            if (!package.WorkbookPart.WorksheetParts.SelectMany(static worksheet => worksheet.TableDefinitionParts).Any()) throw new InvalidOperationException("XLSX reverse conversion did not produce editable table definitions.");
            PdfBenchmarkValidation.ValidateTableScenarioContent(GetSpreadsheetText(package), _scenario, Producer + " " + Scale + " XLSX");
        }
        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(_source)
            .ToPowerPointPresentationResult(PdfPowerPointImportOptions.CreateEditableContent());
        using (result.Value) {
            using var stream = new MemoryStream();
            result.Value.Save(stream);
            using PresentationDocument package = PresentationDocument.Open(new MemoryStream(stream.ToArray()), false);
            int slideCount = package.PresentationPart?.SlideParts.Count() ?? 0;
            if (slideCount != _scenario.PageCount) throw new InvalidOperationException($"PPTX reverse conversion produced {slideCount} of {_scenario.PageCount} slides.");
            ValidateContent(string.Join(" ", package.PresentationPart!.SlideParts.SelectMany(static slide => slide.Slide?.Descendants<A.Text>() ?? Enumerable.Empty<A.Text>()).Select(static text => text.Text)), "PPTX");
        }
        byte[] odt = logical.ToOdtDocumentResult().Value.ToBytes();
        _ = OdtDocument.Load(new MemoryStream(odt));
        ValidateContent(ReadOpenDocumentText(odt), "ODT");
        byte[] ods = logical.ToOdsDocumentResult().Value.ToBytes();
        OdsDocument odsDocument = OdsDocument.Load(new MemoryStream(ods));
        if (odsDocument.Sheets.Count == 0) throw new InvalidOperationException("ODS reverse conversion did not produce a sheet.");
        PdfBenchmarkValidation.ValidateTableScenarioContent(ReadOpenDocumentText(ods), _scenario, Producer + " " + Scale + " ODS");
        byte[] odp = logical.ToOdpPresentationResult(PdfPowerPointImportOptions.CreateEditableContent()).Value.ToBytes();
        OdpPresentation odpDocument = OdpPresentation.Load(new MemoryStream(odp));
        if (odpDocument.Slides.Count != _scenario.PageCount) throw new InvalidOperationException($"ODP reverse conversion produced {odpDocument.Slides.Count} of {_scenario.PageCount} slides.");
        ValidateContent(ReadOpenDocumentText(odp), "ODP");
        IReadOnlyList<PdfCore.PdfPageRenderResult> pngPages = PdfCore.PdfDocument.Load(_source).Render.Pages(
            options: new PdfCore.PdfPageRenderOptions { Format = PdfCore.PdfPageRenderFormat.Png, Dpi = 72D });
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
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(artifact), System.IO.Compression.ZipArchiveMode.Read);
        System.IO.Compression.ZipArchiveEntry content = archive.GetEntry("content.xml")
            ?? throw new InvalidDataException("OpenDocument package did not contain content.xml.");
        using var reader = new StreamReader(content.Open());
        return Regex.Replace(reader.ReadToEnd(), "<[^>]+>", " ");
    }
}
