using System.Globalization;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

public enum PdfWordStructureKind {
    List,
    Table,
    NestedTable,
    MultiParagraphCell
}

/// <summary>
/// Measures complete DOCX loading, structural projection, PDF layout, and
/// serialization across adjacent and larger list/table workloads. Validation
/// reopens the PDF and requires every generated paragraph marker, while list
/// cases must also preserve distinct marker lines and multi-level indentation.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfWordStructureScalingBenchmarks {
    private byte[] _sourceBytes = null!;
    private byte[]? _result;

    [Params(
        PdfWordStructureKind.List,
        PdfWordStructureKind.Table,
        PdfWordStructureKind.NestedTable,
        PdfWordStructureKind.MultiParagraphCell)]
    public PdfWordStructureKind Structure { get; set; }

    [Params(5, 6, 30, 120)]
    public int RecordCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        BenchmarkAffinityGuard.Validate();
        _sourceBytes = CreateSource(Structure, RecordCount);
    }

    [Benchmark]
    public byte[] ConvertWordToPdf() {
        using var stream = new MemoryStream(_sourceBytes, writable: false);
        using WordDocument document = WordDocument.Load(stream);
        return _result = document.ToPdfBytes();
    }

    [GlobalCleanup]
    public void Validate() {
        if (_result == null) {
            throw new InvalidDataException($"Word/{Structure}/{RecordCount} did not return a PDF result.");
        }

        PdfReadObservation observation = PdfBenchmarkValidation.ReadWithPdfPig(_result);
        foreach (string expectedText in GetExpectedText(Structure, RecordCount)) {
            RequireText(observation, expectedText);
        }

        if (Structure == PdfWordStructureKind.List) {
            ValidateListStructure(_result, RecordCount);
        }

        Console.WriteLine(
            $"WORD_STRUCTURE_PDF_EVIDENCE structure={Structure} records={RecordCount} " +
            $"sourceBytes={_sourceBytes.Length} pdfBytes={_result.Length} " +
            $"pages={observation.PageCount} textLength={observation.TextLength}");
    }

    private static byte[] CreateSource(PdfWordStructureKind structure, int recordCount) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("WORD STRUCTURE PDF BENCHMARK").SetStyle(WordParagraphStyles.Heading1);

        switch (structure) {
            case PdfWordStructureKind.List:
                WordList list = document.AddList(WordListStyle.Bulleted);
                for (int index = 1; index <= recordCount; index++) {
                    list.AddItem(
                        CreateMarker(index) + " structured list value for pagination and wrapping",
                        level: (index - 1) % 3);
                }
                break;
            case PdfWordStructureKind.Table:
                CreateTable(document, recordCount, includeNestedTables: false);
                break;
            case PdfWordStructureKind.NestedTable:
                CreateTable(document, recordCount, includeNestedTables: true);
                break;
            case PdfWordStructureKind.MultiParagraphCell:
                CreateMultiParagraphCell(document, recordCount);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(structure));
        }

        return document.ToBytes();
    }

    private static void CreateTable(WordDocument document, int recordCount, bool includeNestedTables) {
        WordTable table = document.AddTable(recordCount, 2, WordTableStyle.TableGrid);
        for (int index = 1; index <= recordCount; index++) {
            WordTableCell markerCell = table.Rows[index - 1].Cells[0];
            markerCell.Paragraphs[0].Text = CreateMarker(index);
            markerCell.AddParagraph(CreateSecondParagraphMarker(index));

            WordTableCell valueCell = table.Rows[index - 1].Cells[1];
            valueCell.Paragraphs[0].Text = CreateValueMarker(index);
            valueCell.AddParagraph(CreateWrappedMarker(index) + " layout evidence for pagination");
            if (includeNestedTables) {
                WordTable nested = valueCell.AddTable(1, 1, WordTableStyle.TableGrid);
                nested.Rows[0].Cells[0].Paragraphs[0].Text = CreateNestedMarker(index);
                nested.Rows[0].Cells[0].AddParagraph(CreateNestedDetailMarker(index));
            }
        }
    }

    private static void CreateMultiParagraphCell(WordDocument document, int paragraphCount) {
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
        WordTableCell cell = table.Rows[0].Cells[0];
        for (int index = 1; index <= paragraphCount; index++) {
            string text = CreateParagraphMarker(index) + " contextual layout evidence";
            if (index == 1) {
                cell.Paragraphs[0].Text = text;
            } else {
                cell.AddParagraph(text);
            }
        }
    }

    private static IEnumerable<string> GetExpectedText(PdfWordStructureKind structure, int recordCount) {
        for (int index = 1; index <= recordCount; index++) {
            if (structure == PdfWordStructureKind.MultiParagraphCell) {
                yield return CreateParagraphMarker(index);
                continue;
            }

            yield return CreateMarker(index);
            if (structure is PdfWordStructureKind.Table or PdfWordStructureKind.NestedTable) {
                yield return CreateSecondParagraphMarker(index);
                yield return CreateValueMarker(index);
                yield return CreateWrappedMarker(index);
            }

            if (structure == PdfWordStructureKind.NestedTable) {
                yield return CreateNestedMarker(index);
                yield return CreateNestedDetailMarker(index);
            }
        }
    }

    private static void RequireText(PdfReadObservation observation, string expectedText) {
        string normalized = PdfBenchmarkValidation.Normalize(expectedText);
        if (!observation.NormalizedText.Contains(normalized, StringComparison.Ordinal)) {
            throw new InvalidDataException($"Word output did not preserve marker {normalized}.");
        }
    }

    private static void ValidateListStructure(byte[] pdf, int recordCount) {
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        StructuredLine[] recordLines = document
            .ExtractStructuredPages(new PdfTextLayoutOptions { ForceSingleColumn = true })
            .SelectMany(page => page.LinesDetailed)
            .Where(line => line.Text.Contains("WORD-RECORD-", StringComparison.Ordinal))
            .ToArray();

        if (recordLines.Length < recordCount) {
            throw new InvalidDataException(
                $"Word/List/{recordCount} preserved only {recordLines.Length} distinct rendered marker lines.");
        }

        for (int index = 1; index <= recordCount; index++) {
            string marker = CreateMarker(index);
            if (!recordLines.Any(line => line.Text.Contains(marker, StringComparison.Ordinal))) {
                throw new InvalidDataException(
                    $"Word/List/{recordCount} did not preserve a rendered marker line for {marker}.");
            }
        }

        int expectedLevels = Math.Min(3, recordCount);
        int actualLevels = recordLines
            .Select(line => Math.Round(line.XStart, 1))
            .Distinct()
            .Count();
        if (actualLevels < expectedLevels) {
            throw new InvalidDataException(
                $"Word/List/{recordCount} preserved only {actualLevels} of {expectedLevels} expected indentation levels.");
        }
    }

    private static string CreateMarker(int index) =>
        "WORD-RECORD-" + index.ToString("0000", CultureInfo.InvariantCulture);

    private static string CreateNestedMarker(int index) =>
        "WORD-NESTED-" + index.ToString("0000", CultureInfo.InvariantCulture);

    private static string CreateSecondParagraphMarker(int index) =>
        "WORD-SECOND-" + index.ToString("0000", CultureInfo.InvariantCulture);

    private static string CreateValueMarker(int index) =>
        "WORD-VALUE-" + index.ToString("0000", CultureInfo.InvariantCulture);

    private static string CreateWrappedMarker(int index) =>
        "WORD-WRAPPED-" + index.ToString("0000", CultureInfo.InvariantCulture);

    private static string CreateNestedDetailMarker(int index) =>
        "WORD-NESTED-DETAIL-" + index.ToString("0000", CultureInfo.InvariantCulture);

    private static string CreateParagraphMarker(int index) =>
        "WORD-PARAGRAPH-" + index.ToString("0000", CultureInfo.InvariantCulture);
}
