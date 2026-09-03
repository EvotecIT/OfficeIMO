using OfficeIMO.Pdf;
using System.Threading;
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

        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(pdf);

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

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf).Tables);
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

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("Phase 2026", data.Columns.Concat(data.Rows.SelectMany(static row => row)));
        Assert.Contains("Complete", data.Rows.SelectMany(static row => row));
    }

    [Fact]
    public void LogicalTables_RetainRegularFontQualitativeTablesWithNumericHeaders() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "2025", "2026" },
                new[] { "Planned", "Ready" },
                new[] { "Review", "Complete" }
            }, style: new PdfTableStyle {
                HeaderBold = false,
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 220D, 220D }
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("2025", data.Columns.Concat(data.Rows.SelectMany(static row => row)));
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

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("Forecast", data.Columns.Concat(data.Rows.SelectMany(static row => row)));
        Assert.Contains("Pending", data.Rows.SelectMany(static row => row));
    }

    [Fact]
    public void PositionedRecovery_RetainsTextOnlyCategoricalTablesFromStructuralEvidence() {
        TextLayoutEngine.TextLine[] lines = {
            CreateLine(520D, ("Feature", 50D, 50D, "Helvetica"), ("Enabled", 220D, 52D, "Helvetica")),
            CreateLine(500D, ("Search", 50D, 42D, "Helvetica"), ("Yes", 220D, 24D, "Helvetica")),
            CreateLine(480D, ("Export", 50D, 42D, "Helvetica"), ("No", 220D, 18D, "Helvetica")),
            CreateLine(460D, ("Archive", 50D, 46D, "Helvetica"), ("Yes", 220D, 24D, "Helvetica"))
        };

        StructuredTable table = Assert.Single(TableDetector.DetectPositionedCellTables(lines));

        Assert.Equal("positioned-cells-bounded", table.Kind);
        Assert.Equal(4, table.Rows.Count);
        Assert.Contains("Archive", table.Rows.SelectMany(static row => row));
    }

    [Fact]
    public void PositionedRecovery_RecognizesSupplementaryPlaneDecimalDigits() {
        TextLayoutEngine.TextLine[] lines = {
            CreateLine(520D, ("Code", 50D, 50D, "Helvetica"), ("Value", 220D, 52D, "Helvetica")),
            CreateLine(500D, ("Alpha", 50D, 42D, "Helvetica"), ("𝟙𝟚", 220D, 24D, "Helvetica")),
            CreateLine(480D, ("Beta", 50D, 42D, "Helvetica"), ("𝟛𝟜", 220D, 24D, "Helvetica"))
        };

        StructuredTable table = Assert.Single(TableDetector.DetectPositionedCellTables(lines));

        Assert.Contains("𝟙𝟚", table.Rows.SelectMany(static row => row));
        Assert.Contains("𝟛𝟜", table.Rows.SelectMany(static row => row));
    }

    [Fact]
    public void LeaderTables_RecognizeSupplementaryPlaneDecimalDigits() {
        var lines = new List<TextLayoutEngine.TextLine> {
            CreateLine(520D, ("Section", 50D, 50D, "Helvetica"), (".....", 140D, 50D, "Helvetica"), ("𝟙𝟚", 240D, 24D, "Helvetica")),
            CreateLine(500D, ("Appendix", 50D, 56D, "Helvetica"), (".....", 140D, 50D, "Helvetica"), ("𝟛𝟜", 240D, 24D, "Helvetica"))
        };

        StructuredTable table = Assert.IsType<StructuredTable>(TableDetector.DetectLeaderTable(lines));

        Assert.Equal(new[] { "Section", "𝟙𝟚" }, table.Rows[0]);
        Assert.Equal(new[] { "Appendix", "𝟛𝟜" }, table.Rows[1]);
    }

    [Fact]
    public void LeaderTables_PreservePunctuationInLabelsAndValues() {
        var lines = new List<TextLayoutEngine.TextLine> {
            CreateLine(520D, ("Release 1.2.3.", 50D, 76D, "Helvetica"), (".....", 150D, 50D, "Helvetica"), ("1.2.3.4", 240D, 42D, "Helvetica")),
            CreateLine(500D, ("Wait... what?", 50D, 76D, "Helvetica"), (".....", 150D, 50D, "Helvetica"), ("0.0.0.1", 240D, 42D, "Helvetica"))
        };

        StructuredTable table = Assert.IsType<StructuredTable>(TableDetector.DetectLeaderTable(lines));

        Assert.Equal(new[] { "Release 1.2.3.", "1.2.3.4" }, table.Rows[0]);
        Assert.Equal(new[] { "Wait... what?", "0.0.0.1" }, table.Rows[1]);
    }

    [Fact]
    public void DetectedTableNormalization_DoesNotRewriteCellPunctuation() {
        var table = new StructuredTable();
        table.Rows.Add(new[] { " Release  1.2.3.4 ", "Wait... what?", "3 . 14" });

        ContentStructureExtractor.NormalizeDetectedTable(table);

        Assert.Equal(new[] { "Release 1.2.3.4", "Wait... what?", "3 . 14" }, table.Rows[0]);
    }

    [Fact]
    public void TableDetector_ReconstructsSparseResponseFormsFromPositionedGaps() {
        List<List<TextLayoutEngine.TextLine>> bands = new() {
            new() { CreateLine(520D,
                ("D", 50D, 8D, "Helvetica"), ("eli", 58D, 16D, "Helvetica"),
                ("v", 74D, 8D, "Helvetica"), ("ery", 82D, 20D, "Helvetica"),
                ("W", 106D, 10D, "Helvetica"), ("o", 116D, 8D, "Helvetica"),
                ("rksh", 124D, 24D, "Helvetica"), ("ee", 148D, 12D, "Helvetica"),
                ("t", 160D, 6D, "Helvetica"), ("Response", 300D, 50D, "Helvetica")) },
            new() { CreateLine(500D, ("Auditor notes", 50D, 116D, "Helvetica"), (" ", 300D, 0D, "Helvetica")) },
            new() { CreateLine(480D, ("Client response", 50D, 116D, "Helvetica"), (" ", 300D, 0D, "Helvetica")) }
        };

        StructuredTable table = Assert.Single(TableDetector.DetectTablesFromBands(bands));
        Assert.Equal(2, table.Columns.Count);
        Assert.Contains(table.Rows.SelectMany(static row => row),
            cell => cell == "Delivery Worksheet");
        Assert.Contains(table.Rows.SelectMany(static row => row),
            cell => ContentStructureExtractor.NormalizeShattered(cell) == "Client response");
    }

    [Fact]
    public void TableDetector_RetainsSparseSignOffFormsWithSeveralBlankColumns() {
        List<List<TextLayoutEngine.TextLine>> bands = new() {
            new() { CreateLine(520D,
                ("Role", 50D, 40D, "Helvetica"), ("Name", 150D, 40D, "Helvetica"),
                ("Decision", 250D, 50D, "Helvetica"), ("Date", 350D, 40D, "Helvetica"),
                ("Notes", 450D, 40D, "Helvetica")) },
            new() { CreateLine(500D,
                ("Lead auditor", 50D, 40D, "Helvetica"), (" ", 150D, 40D, "Helvetica"),
                (" ", 250D, 40D, "Helvetica"), (" ", 350D, 40D, "Helvetica"),
                (" ", 450D, 40D, "Helvetica")) },
            new() { CreateLine(480D,
                ("Client owner", 50D, 40D, "Helvetica"), (" ", 150D, 40D, "Helvetica"),
                (" ", 250D, 40D, "Helvetica"), (" ", 350D, 40D, "Helvetica"),
                (" ", 450D, 40D, "Helvetica")) }
        };

        StructuredTable table = Assert.Single(TableDetector.DetectTablesFromBands(bands));
        Assert.Equal(5, table.Columns.Count);
        Assert.Contains("Lead auditor", table.Rows.SelectMany(static row => row));
        Assert.Contains("Client owner", table.Rows.SelectMany(static row => row));
    }

    [Fact]
    public void TableDetector_RetainsSparseFormsWhenLabelsAreNotInTheFirstColumn() {
        List<List<TextLayoutEngine.TextLine>> bands = new() {
            new() { CreateLine(520D,
                ("Notes", 50D, 40D, "Helvetica"), ("Date", 150D, 40D, "Helvetica"),
                ("Decision", 250D, 50D, "Helvetica"), ("Name", 350D, 40D, "Helvetica"),
                ("Role", 450D, 40D, "Helvetica")) },
            new() { CreateLine(500D,
                (" ", 50D, 40D, "Helvetica"), (" ", 150D, 40D, "Helvetica"),
                (" ", 250D, 40D, "Helvetica"), (" ", 350D, 40D, "Helvetica"),
                ("Lead auditor", 450D, 40D, "Helvetica")) },
            new() { CreateLine(480D,
                (" ", 50D, 40D, "Helvetica"), (" ", 150D, 40D, "Helvetica"),
                (" ", 250D, 40D, "Helvetica"), (" ", 350D, 40D, "Helvetica"),
                ("Client owner", 450D, 40D, "Helvetica")) }
        };

        StructuredTable table = Assert.Single(TableDetector.DetectTablesFromBands(bands));
        Assert.Equal(5, table.Columns.Count);
        Assert.Contains("Lead auditor", table.Rows.SelectMany(static row => row));
        Assert.Contains("Client owner", table.Rows.SelectMany(static row => row));
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

        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(pdf);

        Assert.Equal(2, logical.Tables.Count);
        Assert.Contains(logical.Paragraphs,
            paragraph => paragraph.Text.Contains(narrative, StringComparison.Ordinal));
        Assert.All(logical.Pages[0].Analysis.TableCandidates, candidate =>
            Assert.DoesNotContain(candidate.SourceLines, line => line.Text.Contains(narrative, StringComparison.Ordinal)));
    }

    [Fact]
    public void LogicalTables_RetainCompactSparseRowsUsingStructuralEvidence() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Code", "Owner", "Amount", "State" },
                new[] { "AREA-77", "", "", "" },
                new[] { "ACC-01", "Owner 01", "4100.25", "Accepted" }
            }, style: new PdfTableStyle { HeaderRowCount = 1 })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("AREA-77", data.Rows.SelectMany(static row => row));
        Assert.Contains("4100.25", data.Rows.SelectMany(static row => row));
    }

    [Fact]
    public void LogicalTables_RetainTwoRowTablesWithStableColumns() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Metric", "Value" },
                new[] { "Quality", "Premium" }
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf).Tables);
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

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf).Tables);
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

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf).Tables);
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
    public void TableDetector_DoesNotMergeAcrossNonoverlappingMarginNotes() {
        List<List<TextLayoutEngine.TextLine>> bands = new() {
            new() { CreateLine(520D, ("Account", 50D, 55D, "Helvetica"), ("Amount", 220D, 48D, "Helvetica")) },
            new() { CreateLine(500D, ("A-1", 50D, 24D, "Helvetica"), ("100", 220D, 24D, "Helvetica")) },
            new() { CreateLine(480D, ("Margin note", 400D, 80D, "Helvetica")) },
            new() { CreateLine(460D, ("Account", 50D, 55D, "Helvetica"), ("Amount", 220D, 48D, "Helvetica")) },
            new() { CreateLine(440D, ("B-1", 50D, 24D, "Helvetica"), ("200", 220D, 24D, "Helvetica")) }
        };

        List<StructuredTable> tables = TableDetector.DetectTablesFromBands(bands);

        Assert.Equal(2, tables.Count);
        Assert.DoesNotContain(
            tables.SelectMany(static table => table.Rows).SelectMany(static row => row),
            cell => cell.Contains("Margin note", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("Section summary", "Helvetica-Bold")]
    [InlineData("SECTION SUMMARY", "Helvetica")]
    public void TableDetector_DoesNotMergeAcrossSingleColumnEmphasizedLabels(string label, string font) {
        List<List<TextLayoutEngine.TextLine>> bands = new() {
            new() { CreateLine(520D, ("Account", 50D, 55D, "Helvetica"), ("Amount", 220D, 48D, "Helvetica")) },
            new() { CreateLine(500D, ("A-1", 50D, 55D, "Helvetica"), ("100", 220D, 48D, "Helvetica")) },
            new() { CreateLine(480D, (label, 50D, 100D, font)) },
            new() { CreateLine(460D, ("Account", 50D, 55D, "Helvetica"), ("Amount", 220D, 48D, "Helvetica")) },
            new() { CreateLine(440D, ("B-1", 50D, 55D, "Helvetica"), ("200", 220D, 48D, "Helvetica")) }
        };

        List<StructuredTable> tables = TableDetector.DetectTablesFromBands(bands);

        Assert.Equal(2, tables.Count);
        Assert.DoesNotContain(
            tables.SelectMany(static table => table.Rows).SelectMany(static row => row),
            cell => cell.Contains("SECTION SUMMARY", StringComparison.OrdinalIgnoreCase));
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
            new() { CreateLine(480D, ("Amounts exclude tax.", 50D, 250D, "Helvetica-Bold")) },
            new() { CreateLine(460D, ("A-2", 50D, 24D, "Helvetica"), ("200", 220D, 24D, "Helvetica")) }
        };

        StructuredTable table = Assert.Single(TableDetector.DetectTablesFromBands(bands));

        Assert.Contains(table.Rows.SelectMany(static row => row),
            cell => cell.Contains("Amounts exclude tax.", StringComparison.Ordinal));
    }

    [Fact]
    public void TableDetector_DoesNotSwallowWideNarrativeProseBetweenAlignedBands() {
        List<List<TextLayoutEngine.TextLine>> bands = new() {
            new() { CreateLine(520D, ("Account", 50D, 55D, "Helvetica"), ("Amount", 220D, 48D, "Helvetica")) },
            new() { CreateLine(500D, ("A-1", 50D, 24D, "Helvetica"), ("100", 220D, 24D, "Helvetica")) },
            new() { CreateLine(480D, ("This narrative sentence explains the surrounding report without qualifying the table.", 50D, 360D, "Helvetica")) },
            new() { CreateLine(460D, ("A-2", 50D, 24D, "Helvetica"), ("200", 220D, 24D, "Helvetica")) }
        };

        IReadOnlyList<StructuredTable> tables = TableDetector.DetectTablesFromBands(bands);

        Assert.DoesNotContain(
            tables.SelectMany(static table => table.Rows).SelectMany(static row => row),
            cell => cell.Contains("narrative sentence", StringComparison.OrdinalIgnoreCase));
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

    [Fact]
    public void TableOwnership_ChargesTableAndColumnTraversal() {
        StructuredTable[] tables = Enumerable.Range(0, 40)
            .Select(index => {
                var table = new StructuredTable { YTop = 500D, YBottom = 450D };
                for (int column = 0; column < 8; column++) {
                    table.Columns.Add(new StructuredTableColumn {
                        From = 50D + column * 25D,
                        To = 75D + column * 25D
                    });
                }
                return table;
            })
            .ToArray();
        TextLayoutEngine.TextLine line = CreateLine(475D, ("Outside every table", 500D, 100D, "Helvetica"));
        long consumed = 0L;

        Assert.Throws<InvalidOperationException>(() => ContentStructureExtractor.IsInsideTable(
            line,
            tables,
            units => {
                consumed += units;
                if (consumed > 50L) throw new InvalidOperationException("test work limit");
            },
            cancellationCheck: null));

        Assert.Equal(51L, consumed);
    }

    [Fact]
    public void TableOwnership_ObservesCancellationDuringTableTraversal() {
        var table = new StructuredTable { YTop = 500D, YBottom = 450D };
        for (int column = 0; column < 100; column++) {
            table.Columns.Add(new StructuredTableColumn { From = column, To = column + 1D });
        }
        TextLayoutEngine.TextLine line = CreateLine(475D, ("Outside table", 500D, 100D, "Helvetica"));
        using var cancellation = new CancellationTokenSource();
        int polls = 0;

        Assert.Throws<OperationCanceledException>(() => ContentStructureExtractor.IsInsideTable(
            line,
            new[] { table },
            consumeWork: null,
            cancellationCheck: () => {
                if (++polls == 10) cancellation.Cancel();
                cancellation.Token.ThrowIfCancellationRequested();
            }));

        Assert.Equal(10, polls);
    }

    [Fact]
    public void SplitBySplits_PreservesExplicitWhitespaceWhenAdvanceConsumesTheGap() {
        var first = new PdfTextSpan(
            "North",
            "Helvetica",
            11D,
            50D,
            500D,
            30D,
            null,
            true,
            0D,
            null,
            null,
            logicalTrailingSpace: true);
        var second = new PdfTextSpan("Region", "Helvetica", 11D, 80D, 500D, 36D);
        var value = new PdfTextSpan("42", "Helvetica", 11D, 250D, 500D, 12D);
        var line = new TextLayoutEngine.TextLine(
            500D,
            50D,
            262D,
            "North Region 42",
            new List<PdfTextSpan> { first, second, value });

        string[] cells = TableDetector.SplitBySplits(line, new List<double> { 200D });

        Assert.Equal(new[] { "North Region", "42" }, cells);
    }

    private static TextLayoutEngine.TextLine CreateLine(
        double y,
        params (string Text, double X, double Advance, string Font)[] values) {
        var spans = values
            .Select(value => new PdfTextSpan(value.Text, value.Font, 11D, value.X, y, value.Advance, baseFont: value.Font))
            .ToList();
        return new TextLayoutEngine.TextLine(
            y,
            spans.Min(static span => span.X),
            spans.Max(static span => span.X + span.Advance),
            string.Join(" ", values.Select(static value => value.Text)),
            spans);
    }
}
