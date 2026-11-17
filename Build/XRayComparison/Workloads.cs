using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Word;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using XRay.Content.Core;
using XRay.Office;
using XRay.Office.Core.Documents;
using XRay.Office.Rendering.Pdf;
using CoreDocumentAccessMode = OfficeIMO.DocumentAccessMode;
using PowerPointDocument = OfficeIMO.PowerPoint.PowerPointDocument;
using XRayLayoutOptions = XRay.Office.Core.Documents.LayoutOptions;

namespace OfficeIMO.XRayComparison;

/// <summary>Extracted text or Markdown with the tables each library reports.</summary>
public sealed record ExtractionOutput(string Content, IReadOnlyList<IReadOnlyList<string>> TableCells, int TableCount, int PageCount);

/// <summary>
/// The measured operations. Each pair does the same job through each library's public API, starting from
/// the same in-memory bytes (or the same file path where the API is path-first) and consuming the full result.
/// </summary>
internal static class Workloads {
    private static readonly OfficeDocumentReader Reader = new OfficeDocumentReaderBuilder()
        .AddWordHandler()
        .AddExcelHandler()
        .AddPowerPointHandler()
        .AddPdfHandler()
        .Build();

    private static readonly Extractor XRayExtractor = new();

    // OfficeIMO caps tables at 200 rows and chunks at 8,000 characters by default, and the XLSX Markdown
    // projection drops whatever does not fit a chunk. X-Ray returns the whole document, so both caps are lifted
    // here (set XRAY_BENCH_READER_DEFAULTS=1 to measure the defaults). Hashing is not part of extraction.
    private static ReaderOptions ReaderOptions(ReaderProjection projection) => new() {
        Projection = projection,
        ComputeHashes = false,
        MaxTableRows = UseReaderDefaults ? 200 : int.MaxValue,
        MaxChars = UseReaderDefaults ? 8_000 : 50_000_000
    };

    private static readonly bool UseReaderDefaults = Environment.GetEnvironmentVariable("XRAY_BENCH_READER_DEFAULTS") == "1";

    // ---- 1/2. Text and Markdown extraction -------------------------------------------------------------

    public static ExtractionOutput OfficeImoText(byte[] bytes, BenchCase item) {
        OfficeDocumentReadResult result = Reader.ReadDocument(bytes, item.FileName, ReaderOptions(ReaderProjection.Text));
        return new ExtractionOutput(result.GetText(), [], 0, result.Pages.Count);
    }

    public static ExtractionOutput OfficeImoMarkdown(byte[] bytes, BenchCase item) {
        OfficeDocumentReadResult result = Reader.ReadDocument(bytes, item.FileName, ReaderOptions(ReaderProjection.Full));
        string markdown = result.Markdown ?? string.Join("\n\n", result.Chunks.Select(chunk => chunk.Markdown ?? chunk.Text));
        var cells = result.Tables.Select(table => (IReadOnlyList<string>)table.Columns.Concat(table.Rows.SelectMany(row => row)).ToArray()).ToArray();
        return new ExtractionOutput(markdown, cells, result.Tables.Count, result.Pages.Count);
    }

    public static ExtractionOutput XRayText(byte[] bytes, BenchCase item) => XRayExtract(bytes, item, OutputFormat.Plain);

    public static ExtractionOutput XRayMarkdown(byte[] bytes, BenchCase item) => XRayExtract(bytes, item, OutputFormat.Markdown);

    private static ExtractionOutput XRayExtract(byte[] bytes, BenchCase item, OutputFormat format) {
        ExtractionResult result = XRayExtractor.Extract(
            ExtractInput.FromBytes(bytes, Corpus.MimeType(item), item.FileName),
            new ExtractionConfig { OutputFormat = format });
        if (result.Results.Count == 0) {
            throw new InvalidOperationException("X-Ray returned no result: " + string.Join("; ", result.Errors.Select(error => error.ToString())));
        }
        var document = result.Results[0];
        var cells = (document.Tables ?? []).Select(table => (IReadOnlyList<string>)(table.Cells ?? []).SelectMany(row => row).ToArray()).ToArray();
        return new ExtractionOutput(document.Content ?? string.Empty, cells, document.Tables?.Count ?? 0, document.Pages?.Count ?? 0);
    }

    // ---- 3. Office to PDF ------------------------------------------------------------------------------

    public static byte[] OfficeImoPdf(BenchCase item) => OfficeImoPdfWithReport(item).Pdf;

    /// <summary>The same conversion, also returning OfficeIMO's conversion report for validation.</summary>
    public static (byte[] Pdf, OfficeConversionReport Report) OfficeImoPdfWithReport(BenchCase item) {
        using var output = new MemoryStream();
        OfficeConversionReport report;
        switch (item.Format) {
            case "docx":
                using (WordDocument document = WordDocument.Load(item.Path, new WordLoadOptions { AccessMode = CoreDocumentAccessMode.ReadOnly })) {
                    report = document.SaveAsPdf(output);
                }
                break;
            case "xlsx":
                using (ExcelDocument document = ExcelDocument.Load(item.Path, new ExcelLoadOptions { AccessMode = CoreDocumentAccessMode.ReadOnly })) {
                    report = document.SaveAsPdf(output);
                }
                break;
            case "pptx":
                using (PowerPointDocument document = PowerPointDocument.Load(item.Path, new PowerPointLoadOptions { AccessMode = CoreDocumentAccessMode.ReadOnly })) {
                    report = document.SaveAsPdf(output);
                }
                break;
            default:
                throw new NotSupportedException(item.Format);
        }
        return (output.ToArray(), report);
    }

    public static byte[] XRayPdf(BenchCase item) {
        using var output = new MemoryStream();
        using (IDocument document = OfficeDocument.Open(item.Path)) {
            IPageSequence pages = ((IPaginatedDocument)document).Layout(new XRayLayoutOptions());
            new PdfRenderer(PdfRenderOptions.Default).Render(pages, output);
        }
        return output.ToArray();
    }

}
