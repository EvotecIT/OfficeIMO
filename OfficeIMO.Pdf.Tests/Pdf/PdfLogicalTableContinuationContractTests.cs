using System.Globalization;
using System.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfLogicalTableContinuationContractTests {
    [Fact]
    public void TableContinuations_ExposeTypedEvidenceConfidenceAndPageScope() {
        PdfDocumentReadResult document = PdfDocumentReadResult.Load(BuildMultiPageTablePdf());

        PdfLogicalTableContinuationGroup group = Assert.Single(document.GetTableContinuationGroups());

        Assert.True(group.SpansPages);
        Assert.True(group.Segments.Count > 1);
        Assert.Equal(1, group.FirstPageNumber);
        Assert.Equal(group.Segments.Count, group.LastPageNumber);
        Assert.Equal(30, group.TotalRowCount);
        Assert.InRange(group.Confidence, 0.75D, 1D);
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.AdjacentPages));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.BoundaryTables));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.PageEdges));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.MatchingColumnCount));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.MatchingDetectionKind));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.CompatibleGeometry));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.CompatibleHeaders));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.RepeatedHeaders));
    }

    [Fact]
    public void TableContinuations_CanDisableCrossPageInference() {
        PdfDocumentReadResult document = PdfDocumentReadResult.Load(BuildMultiPageTablePdf());

        IReadOnlyList<PdfLogicalTableContinuationGroup> groups = document.GetTableContinuationGroups(
            new PdfLogicalTableContinuationOptions { MergePageContinuations = false });

        Assert.True(groups.Count > 1);
        Assert.All(groups, group => {
            Assert.False(group.SpansPages);
            Assert.Equal(1D, group.Confidence);
            Assert.Equal(PdfLogicalTableContinuationEvidence.None, group.Evidence);
        });
    }

    [Fact]
    public void TableContinuations_PublicReaderSupportsSelectorsAndPreflight() {
        PdfDocument source = PdfDocument.Load(BuildMultiPageTablePdf());

        PdfLogicalTableContinuationGroup group = Assert.Single(source.Reader.TableContinuations(PdfPageSelector.Parse("all")));
        PdfOperationResult<IReadOnlyList<PdfLogicalTableContinuationGroup>> attempt = source.Reader.TableContinuationsResult();

        Assert.True(group.SpansPages);
        Assert.True(attempt.Succeeded);
        Assert.Equal(PdfPreflightCapability.ReadLogicalObjects, attempt.Capability);
        Assert.True(Assert.Single(attempt.RequireValue()).SpansPages);
    }

    [Fact]
    public void TableContinuations_RejectInvalidConfidence() {
        PdfDocumentReadResult document = PdfDocumentReadResult.Load(BuildMultiPageTablePdf());

        Assert.Throws<ArgumentOutOfRangeException>(() => document.GetTableContinuationGroups(
            new PdfLogicalTableContinuationOptions { MinimumConfidence = double.NaN }));
    }

    [Fact]
    public void TableContinuations_UseBoundedFuzzyHeaderSignatures() {
        Assert.True(PdfLogicalTableContinuations.HeadersEqual(
            new[] { "Transaction description", "Amount" },
            new[] { "Transaction descripton", "Amount" }));
        Assert.False(PdfLogicalTableContinuations.HeadersEqual(
            new[] { "Transaction description", "Amount" },
            new[] { "Customer identifier", "Status" }));
        Assert.False(PdfLogicalTableContinuations.HeadersEqual(
            new[] { "ID", "Amount" },
            new[] { "IP", "Amount" }));
        Assert.False(PdfLogicalTableContinuations.HeadersEqual(
            new[] { "Region", "Revenue 2023" },
            new[] { "Region", "Revenue 2024" }));
    }

    [Theory]
    [InlineData("Amount page 1", "Amount page 2")]
    [InlineData("Kwota strona 1 z 2", "Kwota strona 2 z 2")]
    [InlineData("Belopp sida 1/2", "Belopp sida 2/2")]
    [InlineData("Сумма страница 1", "Сумма страница 2")]
    [InlineData("金额 第1页", "金额 第2页")]
    public void TableContinuations_DoNotEraseNumericSuffixesBasedOnVocabulary(string previous, string current) {
        Assert.False(PdfLogicalTableContinuations.HeadersEqual(
            new[] { "Transaction description", previous },
            new[] { "Transaction description", current }));
    }

    [Fact]
    public void TableContinuations_KeepComparingDigitsOutsidePaginationSuffixes() {
        Assert.False(PdfLogicalTableContinuations.HeadersEqual(
            new[] { "Region", "Revenue 2023 page 1/2" },
            new[] { "Region", "Revenue 2024 page 2/2" }));
    }

    [Theory]
    [InlineData("Amount 1", "Amount ١")]
    [InlineData("Amount 3", "Amount ３")]
    public void TableContinuations_NormalizeEquivalentUnicodeDecimalDigits(string previous, string current) {
        Assert.True(PdfLogicalTableContinuations.HeadersEqual(
            new[] { "Description", previous },
            new[] { "Description", current }));
    }

    [Fact]
    public void TableContinuations_DoNotMergeTablesWhoseHeaderNumbersDiffer() {
        PdfDocumentReadResult document = PdfDocumentReadResult.Load(BuildSlashPaginationTablePdf());

        IReadOnlyList<PdfLogicalTableContinuationGroup> groups = document.GetTableContinuationGroups();
        Assert.True(groups.Count == 2, string.Join(" | ", PdfLogicalTableAnalysis.ExtractTables(document, 0).Select(extraction =>
            "page=" + extraction.PageNumber.ToString(CultureInfo.InvariantCulture) +
            ",kind=" + extraction.DetectionKind +
            ",top=" + extraction.Table.YTop.ToString(CultureInfo.InvariantCulture) +
            ",bottom=" + extraction.Table.YBottom.ToString(CultureInfo.InvariantCulture) +
            ",header=" + extraction.Data.Structure.HasHeaderRow +
            ",columns=" + string.Join("/", extraction.Data.Columns) +
            ",geometry=" + string.Join("/", extraction.Table.Columns.Select(column =>
                column.From.ToString(CultureInfo.InvariantCulture) + "-" + column.To.ToString(CultureInfo.InvariantCulture))))));
        Assert.All(groups, static group => Assert.False(group.SpansPages));
    }

    [Fact]
    public void TableContinuations_CompareColumnsAfterCropOriginNormalization() {
        byte[] first = BuildSinglePageTablePdf(marginLeft: 30D);
        byte[] second = BuildSinglePageTablePdf(marginLeft: 50D);
        byte[] merged = PdfDocument.Merge(PdfDocument.Load(first), PdfDocument.Load(second))
            .Pages.SetCropBox(20D, 0D, 320D, 220D, 2)
            .ToBytes();
        PdfDocumentReadResult document = PdfDocumentReadResult.Load(merged);

        Assert.True(PdfLogicalTableContinuations.HasCompatibleColumns(
            Assert.Single(document.Pages[0].Tables),
            document.Pages[0],
            Assert.Single(document.Pages[1].Tables),
            document.Pages[1],
            tolerance: 4D));
    }

    [Fact]
    public void TableContinuations_RejectRawColumnsShiftedByCropOrigin() {
        byte[] page = BuildSinglePageTablePdf(marginLeft: 30D);
        byte[] merged = PdfDocument.Merge(PdfDocument.Load(page), PdfDocument.Load(page))
            .Pages.SetCropBox(20D, 0D, 320D, 220D, 2)
            .ToBytes();
        PdfDocumentReadResult document = PdfDocumentReadResult.Load(merged);

        Assert.False(PdfLogicalTableContinuations.HasCompatibleColumns(
            Assert.Single(document.Pages[0].Tables),
            document.Pages[0],
            Assert.Single(document.Pages[1].Tables),
            document.Pages[1],
            tolerance: 4D));
    }

    [Fact]
    public void TableContinuations_CompareColumnsAfterUserUnitNormalization() {
        byte[] scaled = WithUserUnit(
            BuildSinglePageTablePdf(30D, 60D, 50D, pageWidth: 600D),
            userUnit: 2D);
        byte[] unscaled = BuildSinglePageTablePdf(64D, 144D, 140D, pageWidth: 600D);
        PdfLogicalPage scaledPage = Assert.Single(PdfDocumentReadResult.Load(scaled).Pages);
        PdfLogicalPage unscaledPage = Assert.Single(PdfDocumentReadResult.Load(unscaled).Pages);

        PdfLogicalTable scaledTable = Assert.Single(scaledPage.Tables);
        PdfLogicalTable unscaledTable = Assert.Single(unscaledPage.Tables);
        Assert.True(PdfLogicalTableContinuations.HasCompatibleColumns(
            scaledTable,
            scaledPage,
            unscaledTable,
            unscaledPage,
            tolerance: 4D),
            "scaled=" + string.Join("/", scaledTable.Columns.Select(static column => column.From + "-" + column.To)) +
            "; unscaled=" + string.Join("/", unscaledTable.Columns.Select(static column => column.From + "-" + column.To)));
    }

    [Fact]
    public void TableContinuations_DoNotTransformOcrVisualColumnsAgain() {
        PdfLogicalPage scaledPage = Assert.Single(PdfDocumentReadResult.Load(WithUserUnit(
            BuildSinglePageTablePdf(30D),
            userUnit: 2D)).Pages);
        PdfLogicalPage unscaledPage = Assert.Single(PdfDocumentReadResult.Load(BuildSinglePageTablePdf(30D)).Pages);
        IReadOnlyList<IReadOnlyList<string>> rows = [new[] { "Description", "Amount" }];
        PdfLogicalTable scaledTable = CreateOcrTable(rows);
        PdfLogicalTable unscaledTable = CreateOcrTable(rows);

        Assert.True(PdfLogicalTableContinuations.HasCompatibleColumns(
            scaledTable,
            scaledPage,
            unscaledTable,
            unscaledPage,
            tolerance: 4D));
    }

    [Fact]
    public void TableContinuations_UseOcrVisualBoundsForPageEdgeInference() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second page"))
            .ToBytes();
        PdfDocumentReadResult source = PdfDocumentReadResult.Load(pdf);
        PdfLogicalPage firstSource = source.Pages[0];
        double firstHeight = firstSource.GetVisualPageSize().Height;
        IReadOnlyList<IReadOnlyList<string>> firstRows = [
            new[] { "Code", "Amount" },
            new[] { "A-1", "10" },
            new[] { "A-2", "20" }
        ];
        IReadOnlyList<IReadOnlyList<string>> secondRows = [
            new[] { "Code", "Amount" },
            new[] { "A-3", "30" },
            new[] { "A-4", "40" }
        ];
        var pages = new[] {
            new PdfOcrPageMergeResult(
                1,
                CreateOcrTableWords(firstRows, firstHeight - 55D),
                0,
                0,
                Array.Empty<string>(),
                string.Empty),
            new PdfOcrPageMergeResult(
                2,
                CreateOcrTableWords(secondRows, 5D),
                0,
                0,
                Array.Empty<string>(),
                string.Empty)
        };
        PdfDocumentReadResult enriched = BuildOcrDocument(pdf, pages);

        PdfLogicalTableContinuationGroup group = Assert.Single(enriched.GetTableContinuationGroups());

        Assert.True(group.SpansPages);
        Assert.Equal(new[] { 1, 2 }, group.Segments.Select(static segment => segment.PageNumber));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.PageEdges));
    }

    [Fact]
    public void TableContinuations_CanCrossNativeAndOcrDetectionBoundaries() {
        byte[] pdf = BuildNativeTableThenBlankPagePdf();
        PdfDocumentReadResult native = PdfDocumentReadResult.Load(pdf);
        PdfLogicalTable firstTable = Assert.Single(native.Pages[0].Tables);
        Assert.Empty(native.Pages[1].Tables);
        IReadOnlyList<IReadOnlyList<string>> rows = [
            new[] { "Code", "Amount" },
            new[] { "A-3", "30" },
            new[] { "A-4", "40" }
        ];
        var words = new List<PdfRecognizedWord>();
        int sequence = 0;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            for (int columnIndex = 0; columnIndex < firstTable.Columns.Count; columnIndex++) {
                PdfLogicalTableColumn column = firstTable.Columns[columnIndex];
                words.Add(new PdfRecognizedWord(
                    rows[rowIndex][columnIndex],
                    column.From,
                    5D + (rowIndex * 20D),
                    Math.Min(20D, column.To - column.From),
                    10D,
                    0.9D,
                    sequence++,
                    "table",
                    "row-" + rowIndex.ToString(CultureInfo.InvariantCulture),
                    "line-" + rowIndex.ToString(CultureInfo.InvariantCulture)));
            }
        }
        PdfDocumentReadResult enriched = BuildOcrDocument(pdf, new[] {
            new PdfOcrPageMergeResult(
                2,
                words.AsReadOnly(),
                0,
                0,
                Array.Empty<string>(),
                string.Empty)
        });

        Assert.Single(enriched.Pages[0].Tables);
        Assert.Single(enriched.Pages[1].Tables);
        PdfLogicalTable enrichedFirst = enriched.Pages[0].Tables[0];
        PdfLogicalTable enrichedSecond = enriched.Pages[1].Tables[0];
        Assert.True(
            PdfLogicalTableContinuations.HasCompatibleColumns(
                enrichedFirst,
                enriched.Pages[0],
                enrichedSecond,
                enriched.Pages[1],
                tolerance: 4D),
            "native=" + string.Join("/", enrichedFirst.Columns.Select(static column => column.From + "-" + column.To)) +
            "; ocr=" + string.Join("/", enrichedSecond.Columns.Select(static column => column.From + "-" + column.To)) +
            "; nativeY=" + enrichedFirst.YTop + "-" + enrichedFirst.YBottom +
            "; ocrY=" + enrichedSecond.YTop + "-" + enrichedSecond.YBottom);

        PdfLogicalTableContinuationGroup group = Assert.Single(
            enriched.GetTableContinuationGroups(), static candidate => candidate.SpansPages);

        Assert.Equal(2, group.Segments.Count);
        Assert.NotEqual(group.Segments[0].DetectionKind, group.Segments[1].DetectionKind);
        Assert.False(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.MatchingDetectionKind));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.CompatibleGeometry));
    }

    private static IReadOnlyList<PdfRecognizedWord> CreateOcrTableWords(
        IReadOnlyList<IReadOnlyList<string>> rows,
        double top) {
        var words = new List<PdfRecognizedWord>(rows.Count * 2);
        int sequence = 0;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            double y = top + (rowIndex * 20D);
            for (int columnIndex = 0; columnIndex < 2; columnIndex++) {
                string text = rows[rowIndex][columnIndex];
                words.Add(new PdfRecognizedWord(
                    text,
                    columnIndex == 0 ? 30D : 150D,
                    y,
                    50D,
                    10D,
                    0.9D,
                    sequence++,
                    "table",
                    "row-" + rowIndex.ToString(CultureInfo.InvariantCulture),
                    "line-" + rowIndex.ToString(CultureInfo.InvariantCulture)));
            }
        }
        return words.AsReadOnly();
    }

    private static PdfLogicalTable CreateOcrTable(IReadOnlyList<IReadOnlyList<string>> rows) =>
        PdfLogicalTable.From(
            1,
            PdfUnderstandingTableCandidate.FromOcr(
                "ocr-aligned-geometry",
                20D,
                80D,
                new PdfLogicalVisualBounds(20D, 20D, 140D, 80D),
                [(20D, 80D), (80D, 140D)],
                rows,
                0.8D,
                new[] { new PdfInferenceEvidence("test.ocr-table", "Test OCR table geometry.", 1D) }));

    private static PdfDocumentReadResult BuildOcrDocument(
        byte[] pdf,
        IReadOnlyList<PdfOcrPageMergeResult> pages) {
        PdfReadDocument source = PdfReadDocument.Open(pdf, null, CancellationToken.None);
        var layoutOptions = new PdfTextLayoutOptions();
        var pipelineOptions = new PdfUnderstandingPipelineOptions();
        PdfDocumentReadResult native = PdfDocumentReadEngine.Read(
            source,
            new PdfReadOptions {
                Profile = PdfReadProfile.Structured,
                LayoutOptions = layoutOptions,
                Pipeline = pipelineOptions
            },
            out IReadOnlyList<PdfUnderstandingPageResult> nativePageAnalyses);
        return PdfOcrLogicalDocumentBuilder.Build(
            source,
            native,
            nativePageAnalyses,
            pages,
            layoutOptions,
            pipelineOptions,
            CancellationToken.None);
    }

    private static byte[] BuildMultiPageTablePdf() {
        var rows = new List<string[]> {
            new[] { "Group", "State" },
            new[] { "Metric", "Owner" }
        };
        for (int index = 1; index <= 30; index++) {
            rows.Add(new[] {
                "Check " + index.ToString(CultureInfo.InvariantCulture),
                "Team " + index.ToString(CultureInfo.InvariantCulture)
            });
        }

        return PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(rows, style: new PdfTableStyle {
                HeaderRowCount = 2,
                RepeatHeaderRowCount = 2,
                ColumnWidthPoints = new List<double?> { 120, 120 },
                CellPaddingX = 5,
                CellPaddingY = 3
            })
            .ToBytes();
    }

    private static byte[] BuildSlashPaginationTablePdf() {
        PdfDocument document = PdfDocument.Create(new PdfOptions {
            PageWidth = 500,
            PageHeight = 320,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFontSize = 9
        });
        for (int index = 0; index < 10; index++) {
            document.Paragraph(paragraph => paragraph.Text("Lead-in line " + index.ToString(CultureInfo.InvariantCulture)));
        }
        return document
            .Table(new[] {
                new[] { "Description", "Amount page 1/2", "State" },
                new[] { "Segment A", "10", "Open" },
                new[] { "Segment B", "11", "Open" },
                new[] { "Segment C", "12", "Open" }
            }, style: new PdfTableStyle {
                HeaderRowCount = 1,
                HeaderBold = true,
                ColumnWidthPoints = new List<double?> { 200, 150, 80 },
                CellPaddingX = 4,
                CellPaddingY = 2
            })
            .PageBreak()
            .Table(new[] {
                new[] { "Description", "Amount page 2/2", "State" },
                new[] { "Segment A", "20", "Open" },
                new[] { "Segment B", "21", "Open" },
                new[] { "Segment C", "22", "Open" }
            }, style: new PdfTableStyle {
                HeaderRowCount = 1,
                HeaderBold = true,
                ColumnWidthPoints = new List<double?> { 200, 150, 80 },
                CellPaddingX = 4,
                CellPaddingY = 2
            })
            .ToBytes();
    }

    private static byte[] BuildNativeTableThenBlankPagePdf() {
        PdfDocument document = PdfDocument.Create(new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 20,
            MarginBottom = 20,
            DefaultFontSize = 9
        });
        for (int index = 0; index < 7; index++) {
            document.Paragraph(paragraph => paragraph.Text("Lead line " + index.ToString(CultureInfo.InvariantCulture)));
        }
        return document
            .Table(new[] {
                new[] { "Code", "Amount" },
                new[] { "A-1", "10" },
                new[] { "A-2", "20" }
            }, style: new PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 100D, 100D },
                CellPaddingX = 4D,
                CellPaddingY = 2D
            })
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Scanned page placeholder."))
            .ToBytes();
    }

    private static byte[] BuildSinglePageTablePdf(
        double marginLeft,
        double firstColumnWidth = 120D,
        double secondColumnWidth = 100D,
        double pageWidth = 320D) =>
        PdfDocument.Create(new PdfOptions {
                PageWidth = pageWidth,
                PageHeight = 220,
                MarginLeft = marginLeft,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "Description", "Amount" },
                new[] { "First item", "10" },
                new[] { "Second item", "20" }
            }, style: new PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { firstColumnWidth, secondColumnWidth },
                CellPaddingX = 4,
                CellPaddingY = 2
            })
            .ToBytes();

    private static byte[] WithUserUnit(byte[] source, double userUnit) =>
        PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            PdfIndirectObject page = Assert.Single(objects.Values, static item =>
                item.Value is PdfDictionary dictionary &&
                string.Equals(dictionary.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal));
            Assert.IsType<PdfDictionary>(page.Value).Items["UserUnit"] = new PdfNumber(userUnit);
            return security.InfoObjectNumber;
        });
}
