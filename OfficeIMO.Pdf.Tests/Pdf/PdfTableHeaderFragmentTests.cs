using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableHeaderFragmentTests {
    [Theory]
    [InlineData("flow")]
    [InlineData("row")]
    [InlineData("columns")]
    public void HeaderReservationUsesTheFirstLegalBodyFragment(string mode) {
        using var pdf = PdfPigDocument.Open(Render(mode, 110, 16));

        if (mode == "columns") {
            var headers = pdf.GetPage(1).GetWords().Where(word => word.Text == "HeaderMarker").ToArray();
            Assert.DoesNotContain(headers, word => Math.Abs(word.Letters[0].StartBaseLine.X - 40D) < 0.1D);
            Assert.Contains(headers, word => Math.Abs(word.Letters[0].StartBaseLine.X - 210D) < 0.1D);
        } else {
            Assert.DoesNotContain("HeaderMarker", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("HeaderMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
            Assert.Contains("BodyLine01", pdf.GetPage(2).Text, StringComparison.Ordinal);
        }
        Assert.Contains("BodyLine16", string.Concat(pdf.GetPages().Select(page => page.Text)), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("row")]
    public void HeaderReservationDoesNotForceTheWholeBodyRowToMove(string mode) {
        using var pdf = PdfPigDocument.Open(Render(mode, 90, 4));

        Assert.Contains("HeaderMarker", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("BodyLine01", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("BodyLine02", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("BodyLine04", pdf.GetPage(2).Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("row")]
    public void IncomingKeepReservesTheHeaderAndFirstBodyFragment(string mode) {
        using var pdf = PdfPigDocument.Open(Render(mode, 85, 2, incomingKeep: true));

        Assert.DoesNotContain("PrecedingMarker", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("PrecedingMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("HeaderMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("BodyLine01", pdf.GetPage(2).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void HeaderReservationIncludesContentBeforeTheContainingRow() {
        using var pdf = PdfPigDocument.Open(Render("row", 110, 1, outerFiller: true));

        Assert.DoesNotContain("HeaderMarker", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("HeaderMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("BodyLine01", pdf.GetPage(2).Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("flow", false)]
    [InlineData("row", false)]
    [InlineData("flow", true)]
    [InlineData("row", true)]
    public void HeaderReservationIncludesTheCompleteVisibleCellSpan(string mode, bool incomingKeep) {
        using var pdf = PdfPigDocument.Open(RenderViewport(mode, incomingKeep));

        Assert.DoesNotContain("HeaderMarker", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("HeaderMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("RowOneMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("RowTwoMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        if (incomingKeep) {
            Assert.DoesNotContain("PrecedingMarker", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("PrecedingMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HeaderReservationIncludesSpacingAfterTheFirstCompleteBodyRow(bool incomingKeep) {
        using var pdf = PdfPigDocument.Open(Render("row", incomingKeep ? 74 : 94, 1,
            incomingKeep: incomingKeep, cellSpacing: 5, trailingRow: true));

        Assert.DoesNotContain("HeaderMarker", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("HeaderMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("BodyLine01", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("LastRowMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        if (incomingKeep) {
            Assert.DoesNotContain("PrecedingMarker", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("PrecedingMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void FlowHeaderReservationAllowsTheBodyRowWhenOnlyItsTrailingSpacingDoesNotFit() {
        using var pdf = PdfPigDocument.Open(Render("flow", 94, 1, cellSpacing: 5, trailingRow: true));

        Assert.Contains("HeaderMarker", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("BodyLine01", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("LastRowMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
    }

    private static byte[] Render(string mode, double fillerHeight, int bodyLines, bool incomingKeep = false,
        bool outerFiller = false, double cellSpacing = 0, bool trailingRow = false) {
        PdfTableStyle style = HeaderStyle();
        style.CellSpacing = P(cellSpacing);
        var rows = new List<PdfTableCell[]> {
            new[] { Cell("HeaderMarker") },
            new[] { Cell(string.Join("\n", Enumerable.Range(1, bodyLines).Select(line => $"BodyLine{line:D2}"))) }
        };
        if (trailingRow) rows.Add(new[] { Cell("LastRowMarker") });
        Action<PdfContentBuilder> content = builder => {
            if (!outerFiller) builder.AddParagraph(p => p.AddText("FillerMarker"), style: Line(fillerHeight));
            if (incomingKeep) builder.AddParagraph(p => p.AddText("PrecedingMarker"), style: Line(20, keepWithNext: true));
            builder.AddTable(rows, style: style);
        };
        return PdfDocument.Create(document => document.Content(root => {
            if (outerFiller) root.AddParagraph(p => p.AddText("FillerMarker"), style: Line(fillerHeight));
            if (mode == "row") root.Row(row => row.PercentColumn(100, content));
            else if (mode == "columns") root.Columns(content, new PdfMultiColumnOptions { ColumnCount = 2, Gap = P(20), BalanceLastPage = false });
            else content(root);
        }), Options()).ToBytes();
    }

    private static byte[] RenderViewport(string mode, bool incomingKeep) {
        PdfTableStyle style = HeaderStyle();
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { P(160), P(160) };
        style.FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { P(20), P(40), P(40) };
        var rows = new[] {
            new[] { Cell("HeaderMarker"), Cell("OtherHeader") },
            new[] { new PdfTableCell("SpanMarker", rowSpan: 2)
                .WithViewport(new PdfTableCellViewport(P(160), P(160), P(160), P(80), offsetY: P(80))), Cell("RowOneMarker") },
            new[] { Cell("RowTwoMarker") }
        };
        Action<PdfContentBuilder> content = builder => {
            builder.AddParagraph(p => p.AddText("FillerMarker"), style: Line(incomingKeep ? 50 : 70));
            if (incomingKeep) builder.AddParagraph(p => p.AddText("PrecedingMarker"), style: Line(20, keepWithNext: true));
            builder.AddTable(rows, style: style);
        };
        return PdfDocument.Create(document => document.Content(root => {
            if (mode == "row") root.Row(row => row.PercentColumn(100, content));
            else content(root);
        }), Options()).ToBytes(new PdfSaveOptions { AllowLoss = true });
    }

    private static PdfTableStyle HeaderStyle() => new() {
        HeaderRowCount = 1, RepeatHeaderRowCount = 1, KeepHeaderWithFirstBodyFragment = true,
        MinimumBodyRowsOnFirstPage = 0, MinimumBodyRowsOnLastPage = 0, CellPaddingX = P(0), CellPaddingY = P(0),
        FontSize = P(12), LineHeight = 20D / 12D, SpacingBefore = P(0), SpacingAfter = P(0), HeaderBold = false,
        BorderWidth = P(0), RowSeparatorWidth = P(0)
    };

    private static PdfTableCell Cell(string text) {
        var runs = new[] { PdfTextRun.Normal(text) };
        return new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs, fontSize: P(12),
            lineSpacing: PdfLineSpacing.Exact(P(20)), widowControl: false) });
    }

    private static global::OfficeIMO.OfficeLength P(double points) => global::OfficeIMO.OfficeLength.Points(points);

    private static PdfParagraphStyle Line(double height, bool keepWithNext = false) => new() {
        FontSize = P(12), LineSpacing = PdfLineSpacing.Exact(P(height)), SpacingBefore = P(0),
        SpacingAfter = P(0), WidowControl = false, KeepTogether = true, KeepWithNext = keepWithNext
    };

    private static PdfCreateOptions Options() => new() {
        PageSetup = {
            Size = new PdfPageSize(P(400), P(220)),
            Margins = { Left = P(40), Right = P(40), Top = P(40), Bottom = P(40) }
        },
        DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = P(12), MaxGeneratedPages = 8
    };
}
