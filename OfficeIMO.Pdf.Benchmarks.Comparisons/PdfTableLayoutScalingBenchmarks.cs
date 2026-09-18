using System.Globalization;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

public enum PdfTableLayoutCellMode {
    Plain,
    ExplicitRich
}

/// <summary>
/// Measures worksheet-like table layout, pagination, and serialization across
/// adjacent and larger row counts. Every body cell carries an explicit font
/// size so the shared shrink-to-fit path used by spreadsheet conversion is
/// exercised rather than only the plain-text fast path.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfTableLayoutScalingBenchmarks {
    private IReadOnlyList<PdfTableCell[]> _rows = null!;
    private PdfTableStyle _style = null!;
    private byte[]? _result;

    [Params(5, 6, 30, 120)]
    public int RecordCount { get; set; }

    [Params(PdfTableLayoutCellMode.Plain, PdfTableLayoutCellMode.ExplicitRich)]
    public PdfTableLayoutCellMode CellMode { get; set; }

    [GlobalSetup]
    public void Setup() {
        BenchmarkAffinityGuard.Validate();
        _rows = CreateRows(RecordCount, CellMode);
        _style = TableStyles.TableGrid();
        _style.HeaderRowCount = 1;
        _style.RepeatHeaderRowCount = 1;
        _style.ColumnWidthPoints = new List<double?> { 54D, 58D, 72D, 80D, 72D, 78D, 54D };
        _style.PreserveWidth = true;
        _style.ShrinkTextToFit = true;
        _style.MinimumShrinkFontSize = 6D;
        _style.FontSize = 9D;
        _style.LineHeight = 1.15D;
        _style.CellPaddingX = 2D;
        _style.CellPaddingY = 1D;
    }

    [Benchmark]
    public byte[] ComposeTablePdf() {
        PdfDocument document = PdfDocument.Create(pdf => pdf.Content(content => content.Table(_rows, style: _style)), new PdfOptions {
            PageWidth = 612D,
            PageHeight = 792D,
            MarginLeft = 54D,
            MarginRight = 54D,
            MarginTop = 54D,
            MarginBottom = 54D,
            DefaultFontSize = 9D,
            FileVersion = PdfFileVersion.Pdf17,
            ObjectSerializationMode = PdfObjectSerializationMode.ForwardOnly
        });
        return _result = document.ToBytes();
    }

    [GlobalCleanup]
    public void Validate() {
        if (_result == null) {
            throw new InvalidDataException($"Table/{RecordCount} did not return a PDF result.");
        }

        PdfReadObservation observation = PdfBenchmarkValidation.ReadWithPdfPig(_result);
        for (int index = 1; index <= RecordCount; index++) {
            string marker = PdfBenchmarkValidation.Normalize(CreateMarker(index));
            if (!observation.NormalizedText.Contains(marker, StringComparison.Ordinal)) {
                throw new InvalidDataException(
                    $"Table/{CellMode}/{RecordCount} did not preserve row marker {marker}.");
            }
        }

        Console.WriteLine(
            $"TABLE_LAYOUT_PDF_EVIDENCE mode={CellMode} records={RecordCount} pdfBytes={_result.Length} " +
            $"pages={observation.PageCount} textLength={observation.TextLength}");
    }

    private static IReadOnlyList<PdfTableCell[]> CreateRows(int recordCount, PdfTableLayoutCellMode cellMode) {
        var rows = new List<PdfTableCell[]>(recordCount + 1) {
            new[] {
                Cell("Record", cellMode, bold: true),
                Cell("Status", cellMode, bold: true),
                Cell("Owner", cellMode, bold: true),
                Cell("Service", cellMode, bold: true),
                Cell("Region", cellMode, bold: true),
                Cell("Evidence", cellMode, bold: true),
                Cell("Score", cellMode, bold: true)
            }
        };

        for (int index = 1; index <= recordCount; index++) {
            rows.Add(new[] {
                Cell(CreateMarker(index), cellMode),
                Cell(index % 3 == 0 ? "NeedsReview" : "Approved", cellMode),
                Cell(index % 2 == 0 ? "OperationsTeam" : "ComplianceTeam", cellMode),
                Cell("IdentityProvisioningService", cellMode),
                Cell(index % 2 == 0 ? "NorthRegion" : "SouthRegion", cellMode),
                Cell("/evidence/controls/2026/section-" + index.ToString("00", CultureInfo.InvariantCulture), cellMode),
                Cell((90 + index % 10).ToString(CultureInfo.InvariantCulture), cellMode)
            });
        }

        return rows;
    }

    private static PdfTableCell Cell(string text, PdfTableLayoutCellMode cellMode, bool bold = false) =>
        cellMode == PdfTableLayoutCellMode.Plain
            ? PdfTableCell.TextCell(text).WithNoWrap()
            : PdfTableCell.RichTextCell(new[] {
                new PdfTextRun(text, bold: bold, fontSize: 11D)
            }).WithNoWrap();

    private static string CreateMarker(int index) =>
        "ROW-" + index.ToString("0000", CultureInfo.InvariantCulture);
}
