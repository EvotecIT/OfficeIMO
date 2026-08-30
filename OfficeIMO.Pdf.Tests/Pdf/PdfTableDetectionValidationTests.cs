using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableDetectionValidationTests {
    [Fact]
    public void LogicalTables_DoNotTreatSparseTwoColumnParagraphsAsATable() {
        byte[] pdf = PdfDocument.Create(document => document.Content(content => content.Canvas(canvas => canvas
            .Text("Left opening paragraph", 36D, 90D, 220D, 24D, fontSize: 11D)
            .Text("Left indented paragraph", 66D, 140D, 190D, 24D, fontSize: 11D)
            .Text("Left closing paragraph", 36D, 190D, 220D, 24D, fontSize: 11D)
            .Text("Right opening paragraph", 326D, 90D, 220D, 24D, fontSize: 11D)
            .Text("Right middle paragraph", 326D, 140D, 220D, 24D, fontSize: 11D)
            .Text("Right closing paragraph", 326D, 190D, 220D, 24D, fontSize: 11D)
            .Table(new[] {
                new[] { "TABLE-MARKER", "Owner", "Amount", "Status" },
                new[] { "ACC-01", "Owner 01", "1037.25", "Approved" },
                new[] { "ACC-02", "Owner 02", "1074.50", "Review" }
            }, 36D, 300D, 510D, 150D, style: new PdfTableStyle { HeaderRowCount = 1 })))).ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf);

        PdfLogicalTable table = Assert.Single(logical.Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("TABLE-MARKER", data.Columns.Concat(data.Rows.SelectMany(static row => row)));
    }

    [Fact]
    public void LogicalTables_RetainRegularFontProseHeavyTablesWithStableColumns() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Assigned owner", "Current workflow" },
                new[] { "North region coordinator", "Review pending requests" },
                new[] { "South region coordinator", "Approve completed requests" }
            }, style: new PdfTableStyle {
                HeaderBold = false,
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 220D, 220D }
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfLogicalDocument.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("North region coordinator", data.Rows.SelectMany(static row => row));
        Assert.Contains("Approve completed requests", data.Rows.SelectMany(static row => row));
    }

    [Fact]
    public void LogicalTables_RetainRegularFontQualitativeTablesWithDigitBearingHeaders() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Phase 2026", "Phase 2027" },
                new[] { "Planning", "Ready" },
                new[] { "Review", "Complete" }
            }, style: new PdfTableStyle {
                HeaderBold = false,
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 220D, 220D }
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfLogicalDocument.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("Phase 2026", data.Columns.Concat(data.Rows.SelectMany(static row => row)));
        Assert.Contains("Complete", data.Rows.SelectMany(static row => row));
    }

    [Fact]
    public void LogicalTables_RetainRegularFontQualitativeTablesWithRepeatedHeaders() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Actual", "Actual", "Forecast" },
                new[] { "North", "Ready", "Planned" },
                new[] { "South", "Complete", "Pending" }
            }, style: new PdfTableStyle {
                HeaderBold = false,
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 140D, 140D, 140D }
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfLogicalDocument.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("Forecast", data.Columns.Concat(data.Rows.SelectMany(static row => row)));
        Assert.Contains("Pending", data.Rows.SelectMany(static row => row));
    }

    [Fact]
    public void LogicalTables_RetainSparseSpanningRowsWhenTheTableHasStrongEvidence() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Account", "Owner", "Amount", "Status" },
                new[] { "SECTION-A", "", "", "" },
                new[] { "ACC-01", "Owner 01", "1037.25", "Approved" }
            }, style: new PdfTableStyle { HeaderRowCount = 1 })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfLogicalDocument.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("SECTION-A", data.Rows.SelectMany(static row => row));
        Assert.Contains("1037.25", data.Rows.SelectMany(static row => row));
    }

    [Theory]
    [InlineData("Intervening narrative text must remain a paragraph.")]
    [InlineData("Intervening narrative remains independent")]
    public void LogicalTables_DoNotConsumeNarrativeBetweenAlignedTables(string narrative) {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "First owner", "First status" },
                new[] { "North region", "Review pending" },
                new[] { "South region", "Approved" }
            }, style: new PdfTableStyle {
                HeaderRowCount = 1,
                HeaderBold = false,
                ColumnWidthPoints = new List<double?> { 220D, 220D }
            })
            .Paragraph(paragraph => paragraph.Text(narrative))
            .Table(new[] {
                new[] { "Second owner", "Second status" },
                new[] { "East region", "In progress" },
                new[] { "West region", "Complete" }
            }, style: new PdfTableStyle {
                HeaderRowCount = 1,
                HeaderBold = false,
                ColumnWidthPoints = new List<double?> { 220D, 220D }
            })
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf);

        Assert.Equal(2, logical.Tables.Count);
        Assert.Contains(logical.Paragraphs,
            paragraph => paragraph.Text.Contains(narrative, StringComparison.Ordinal));
    }

    [Fact]
    public void LogicalTables_RetainTwoRowTablesWithStableColumns() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Metric", "Value" },
                new[] { "Quality", "Premium" }
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfLogicalDocument.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("Quality", data.Rows.SelectMany(static row => row));
        Assert.Contains("Premium", data.Rows.SelectMany(static row => row));
    }

    [Fact]
    public void LogicalTables_RetainTwoRowTablesWithDigitBearingHeadersAndValues() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Phase 2026", "Phase 2027" },
                new[] { "10", "20" }
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfLogicalDocument.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("Phase 2026", data.Columns.Concat(data.Rows.SelectMany(static row => row)));
        Assert.Contains("20", data.Rows.SelectMany(static row => row));
    }

    [Theory]
    [InlineData(PdfColumnAlign.Center)]
    [InlineData(PdfColumnAlign.Right)]
    public void LogicalTables_RetainAlignedRegularFontProseTables(PdfColumnAlign alignment) {
        var cellAlignments = new Dictionary<(int Row, int Column), PdfColumnAlign>();
        for (int row = 0; row < 3; row++) {
            cellAlignments[(row, 0)] = alignment;
            cellAlignments[(row, 1)] = alignment;
        }
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Assigned owner", "Current workflow" },
                new[] { "North region coordinator", "Review pending requests" },
                new[] { "South team", "Approve all completed requests" }
            }, style: new PdfTableStyle {
                HeaderBold = false,
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 220D, 220D },
                CellAlignments = cellAlignments
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfLogicalDocument.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("North region coordinator", data.Rows.SelectMany(static row => row));
        Assert.Contains("Approve all completed requests", data.Rows.SelectMany(static row => row));
    }

    [Fact]
    public void TableDetector_DoesNotUseBoldAloneAsTwoRowTableEvidence() {
        List<List<TextLayoutEngine.TextLine>> bands = new() {
            new() { CreateLine(520D, ("Summary", 50D, 55D, "Helvetica-Bold"), ("Notes", 220D, 40D, "Helvetica-Bold")) },
            new() { CreateLine(500D, ("Management", 75D, 70D, "Helvetica"), ("review", 245D, 38D, "Helvetica")) }
        };

        Assert.Empty(TableDetector.DetectTablesFromBands(bands));
    }

    [Fact]
    public void TableDetector_DoesNotMergeAcrossOrdinaryTightRhythmProse() {
        List<List<TextLayoutEngine.TextLine>> bands = new() {
            new() { CreateLine(520D, ("Account", 50D, 55D, "Helvetica"), ("Amount", 220D, 48D, "Helvetica")) },
            new() { CreateLine(500D, ("A-1", 50D, 24D, "Helvetica"), ("100", 220D, 24D, "Helvetica")) },
            new() { CreateLine(480D, ("Management review remains pending.", 50D, 250D, "Helvetica")) },
            new() { CreateLine(460D, ("A-2", 50D, 24D, "Helvetica"), ("200", 220D, 24D, "Helvetica")) }
        };

        List<StructuredTable> tables = TableDetector.DetectTablesFromBands(bands);

        Assert.DoesNotContain(tables.SelectMany(static table => table.Rows).SelectMany(static row => row),
            cell => cell.Contains("Management", StringComparison.Ordinal));
    }

    [Fact]
    public void TableDetector_DoesNotTreatParallelPageColumnsAsCompactTable() {
        List<List<TextLayoutEngine.TextLine>> bands = new() {
            new() { CreateLine(520D, ("Left one", 50D, 110D, "Helvetica"), ("Right one", 320D, 110D, "Helvetica")) },
            new() { CreateLine(500D, ("Left two", 80D, 110D, "Helvetica"), ("Right two", 320D, 110D, "Helvetica")) },
            new() { CreateLine(480D, ("Left three", 50D, 110D, "Helvetica"), ("Right three", 320D, 110D, "Helvetica")) }
        };

        Assert.Empty(TableDetector.DetectTablesFromBands(bands));
    }

    [Fact]
    public void TableDetector_RetainsPunctuatedSpanningRowsWithGeometricEvidence() {
        List<List<TextLayoutEngine.TextLine>> bands = new() {
            new() { CreateLine(520D, ("Account", 50D, 55D, "Helvetica"), ("Amount", 220D, 48D, "Helvetica")) },
            new() { CreateLine(500D, ("A-1", 50D, 24D, "Helvetica"), ("100", 220D, 24D, "Helvetica")) },
            new() { CreateLine(480D, ("Amounts exclude tax.", 50D, 250D, "Helvetica")) },
            new() { CreateLine(460D, ("A-2", 50D, 24D, "Helvetica"), ("200", 220D, 24D, "Helvetica")) }
        };

        StructuredTable table = Assert.Single(TableDetector.DetectTablesFromBands(bands));

        Assert.Contains(table.Rows.SelectMany(static row => row),
            cell => cell.Contains("Amounts exclude tax.", StringComparison.Ordinal));
    }

    [Fact]
    public void TableOwnership_RequiresHorizontalAsWellAsVerticalOverlap() {
        var table = new StructuredTable { YTop = 500D, YBottom = 450D };
        table.Columns.Add(new StructuredTableColumn { From = 50D, To = 150D });
        table.Columns.Add(new StructuredTableColumn { From = 150D, To = 250D });
        TextLayoutEngine.TextLine adjacent = CreateLine(475D, ("Adjacent prose", 400D, 90D, "Helvetica"));
        TextLayoutEngine.TextLine overlapping = CreateLine(475D, ("Table prose", 100D, 90D, "Helvetica"));

        Assert.False(ContentStructureExtractor.IsInsideTable(adjacent, new[] { table }));
        Assert.True(ContentStructureExtractor.IsInsideTable(overlapping, new[] { table }));
    }

    private static TextLayoutEngine.TextLine CreateLine(
        double y,
        params (string Text, double X, double Advance, string Font)[] values) {
        var spans = values
            .Select(value => new PdfTextSpan(value.Text, value.Font, 11D, value.X, y, value.Advance))
            .ToList();
        return new TextLayoutEngine.TextLine(
            y,
            spans.Min(static span => span.X),
            spans.Max(static span => span.X + span.Advance),
            string.Join(" ", values.Select(static value => value.Text)),
            spans);
    }
}
