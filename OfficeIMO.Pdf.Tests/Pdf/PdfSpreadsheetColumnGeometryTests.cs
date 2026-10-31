using OfficeIMO.Pdf.Excel;
using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.Pdf;
using Xunit;
using Xunit.Abstractions;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSpreadsheetColumnGeometryTests {
    private readonly ITestOutputHelper _output;

    public PdfSpreadsheetColumnGeometryTests(ITestOutputHelper output) => _output = output;

    [Theory]
    [InlineData("ruled-ledger-0", "xlsx")]
    [InlineData("ruled-ledger-90", "xlsx")]
    [InlineData("ruled-ledger-0", "ods")]
    [InlineData("ruled-ledger-90", "ods")]
    public void SavedSpreadsheetsPreserveDetectedProgressionAndEditableNumbers(string id, string route) {
        Fixture fixture = ReadFixture(id);
        PdfDocumentReadResult source = Read(fixture);
        PdfLogicalTable table = Assert.Single(source.Tables);
        AssertSourceRows(fixture, table);
        double[] detected = DetectedShares(table);

        // The independent authored grid remains a separate oracle. Retain any
        // recovery miss; mapping detected geometry does not qualify grid recovery.
        _output.WriteLine(JsonSerializer.Serialize(new {
            fixture = id,
            authoredGrid = fixture.AuthoredShares,
            detected,
            maximumSourceGridShareError = MaximumError(fixture.AuthoredShares, detected),
            sourceGridRecoveryQualified = false
        }));

        Artifact artifact = RoundTrip(source, route, new ExcelPdfImportOptions {
            UseFirstRowAsHeader = true,
            AutoFitColumns = false,
            PreserveColumnProportions = true
        });

        Assert.True(artifact.Entry.ColumnProportionsApplied);
        AssertShares(detected, artifact.Widths, fixture.Tolerance(route));
        AssertCells(fixture, artifact);
    }

    [Theory]
    [InlineData("latin-90", "xlsx")]
    [InlineData("latin-270", "xlsx")]
    [InlineData("latin-90", "ods")]
    [InlineData("latin-270", "ods")]
    public void ContentQuarterTurnUsesVerticalColumnStripsInsteadOfCoincidentXProjections(string id, string route) {
        Fixture fixture = ReadFixture(id, "MultilingualLayout");
        PdfDocumentReadResult source = Read(fixture);
        PdfLogicalTable table = Assert.Single(source.Tables);
        AssertSourceRows(fixture, table);
        Assert.All(table.Columns, column => {
            Assert.NotNull(column.VisualBounds);
            Assert.Equal(table.Columns[0].From, column.From, 6);
            Assert.Equal(table.Columns[0].To, column.To, 6);
        });

        Artifact artifact = RoundTrip(source, route, new ExcelPdfImportOptions {
            UseFirstRowAsHeader = true,
            AutoFitColumns = false,
            PreserveColumnProportions = true
        });

        Assert.True(artifact.Entry.ColumnProportionsApplied);
        AssertShares(DetectedShares(table, vertical: true), artifact.Widths, fixture.Tolerance(route));
        AssertCells(fixture, artifact);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("ods")]
    public void AutoFitSuppliesTotalBudgetBeforeDetectedProportionsDetermineWidths(string route) {
        Fixture fixture = ReadFixture("ruled-ledger-0");
        PdfDocumentReadResult source = Read(fixture);
        Artifact fitted = RoundTrip(source, route, new ExcelPdfImportOptions { UseFirstRowAsHeader = true });
        Artifact mapped = RoundTrip(source, route, new ExcelPdfImportOptions {
            UseFirstRowAsHeader = true,
            PreserveColumnProportions = true
        });

        Assert.False(fitted.Entry.ColumnProportionsApplied);
        Assert.True(mapped.Entry.ColumnProportionsApplied);
        // ODF length serialization rounds each physical width to 0.001 point.
        double budgetTolerance = route == "ods" ? mapped.Widths.Length * 0.001D : 0.000001D;
        Assert.InRange(Math.Abs(fitted.Widths.Sum() - mapped.Widths.Sum()), 0D, budgetTolerance);
        AssertShares(DetectedShares(Assert.Single(source.Tables)), mapped.Widths, fixture.Tolerance(route));
    }

    [Theory]
    [InlineData("xlsx", false)]
    [InlineData("xlsx", true)]
    [InlineData("ods", false)]
    [InlineData("ods", true)]
    public void DefaultOptionRetainsExistingAutoFitAndDefaultWidthPolicies(string route, bool autoFit) {
        Fixture fixture = ReadFixture("ruled-ledger-0");
        PdfDocumentReadResult source = Read(fixture);
        Artifact defaults = RoundTrip(source, route, new ExcelPdfImportOptions {
            UseFirstRowAsHeader = true, AutoFitColumns = autoFit
        });
        Artifact disabled = RoundTrip(source, route, new ExcelPdfImportOptions {
            UseFirstRowAsHeader = true, AutoFitColumns = autoFit, PreserveColumnProportions = false
        });

        Assert.False(defaults.Entry.ColumnProportionsApplied);
        Assert.False(disabled.Entry.ColumnProportionsApplied);
        Assert.Equal(defaults.Widths, disabled.Widths);
        if (autoFit) Assert.True(defaults.Widths.Max() > defaults.Widths.Min());
        else Assert.All(defaults.Widths, width => Assert.Equal(defaults.Widths[0], width));
        AssertCells(fixture, defaults);
    }

    [Theory]
    [InlineData("xlsx", false)]
    [InlineData("xlsx", true)]
    [InlineData("ods", false)]
    [InlineData("ods", true)]
    public void IncompleteCustomDetectorGeometryRetainsSizingAndNumericContent(string route, bool autoFit) {
        Fixture fixture = ReadFixture("ruled-ledger-0");
        var pipeline = PdfUnderstandingPipelineOptions.Structured();
        pipeline.TableDetection = new IncompleteGeometryTableStage();
        PdfDocumentReadResult source = PdfDocument.Load(new System.IO.MemoryStream(fixture.Pdf)).Read(new PdfLoadOptions {
            Profile = PdfReadProfile.Structured, Pipeline = pipeline
        });
        Assert.Equal(0D, Assert.Single(source.Tables).Columns[1].To - source.Tables[0].Columns[1].From);
        Artifact baseline = RoundTrip(source, route, new ExcelPdfImportOptions {
            UseFirstRowAsHeader = true, AutoFitColumns = autoFit
        });
        Artifact fallback = RoundTrip(source, route, new ExcelPdfImportOptions {
            UseFirstRowAsHeader = true, AutoFitColumns = autoFit, PreserveColumnProportions = true
        });

        Assert.False(fallback.Entry.ColumnProportionsApplied);
        Assert.Equal(baseline.Widths, fallback.Widths);
        AssertCells(fixture, fallback);
    }

    private static PdfDocumentReadResult Read(Fixture fixture) =>
        PdfDocument.Load(new System.IO.MemoryStream(fixture.Pdf)).Read(new PdfLoadOptions { Profile = PdfReadProfile.Structured });

    private static double[] DetectedShares(PdfLogicalTable table, bool vertical = false) {
        // Page /Rotate leaves the ledger's native content coordinates unchanged.
        // The Pango quarter-turn inputs instead rotate content, so their restored
        // visual strips progress vertically; this expectation comes from the fixture.
        double[] lengths = table.Columns.Select(column => column.VisualBounds is null
            ? Math.Abs(column.To - column.From)
            : vertical ? column.VisualBounds.Height : column.VisualBounds.Width).ToArray();
        Assert.All(lengths, length => Assert.True(IsPositiveFinite(length)));
        Assert.True(lengths.Max() > lengths.Min());
        return Normalize(lengths);
    }

    private static void AssertSourceRows(Fixture fixture, PdfLogicalTable table) {
        Assert.Equal(fixture.Rows.Length, table.Rows.Count);
        for (int row = 0; row < fixture.Rows.Length; row++) Assert.Equal(fixture.Rows[row], table.Rows[row]);
    }

    private static void AssertCells(Fixture fixture, Artifact artifact) {
        Assert.Equal(fixture.Rows.Length, artifact.Text.Length);
        Assert.Equal(fixture.Rows.Length - 1, artifact.Entry.RowCount);
        for (int row = 0; row < fixture.Rows.Length; row++) {
            Assert.Equal(fixture.Rows[row].Length, artifact.Text[row].Length);
            for (int column = 0; column < fixture.Rows[row].Length; column++) {
                if (row > 0 && column > 0) {
                    Assert.NotNull(artifact.Numbers[row][column]);
                    Assert.Equal(decimal.Parse(fixture.Rows[row][column], CultureInfo.InvariantCulture),
                        artifact.Numbers[row][column]!.Value);
                } else Assert.Equal(fixture.Rows[row][column], artifact.Text[row][column]);
            }
        }
    }

    private static void AssertShares(double[] expected, double[] widths, double tolerance) {
        Assert.Equal(expected.Length, widths.Length);
        Assert.All(widths, width => Assert.True(IsPositiveFinite(width)));
        double[] actual = Normalize(widths);
        for (int index = 0; index < expected.Length; index++) {
            Assert.InRange(actual[index], expected[index] - tolerance, expected[index] + tolerance);
        }
    }

    private static double[] Normalize(double[] values) {
        double total = values.Sum();
        return values.Select(value => value / total).ToArray();
    }
    private static bool IsPositiveFinite(double value) => value > 0D && !double.IsNaN(value) && !double.IsInfinity(value);
    private static double MaximumError(double[] expected, double[] actual) =>
        expected.Zip(actual, (left, right) => Math.Abs(left - right)).Max();

    private static Artifact RoundTrip(PdfDocumentReadResult source, string route, ExcelPdfImportOptions options) =>
        route == "xlsx" ? RoundTripXlsx(source, options) : RoundTripOds(source, options);

    private static Artifact RoundTripXlsx(PdfDocumentReadResult source, ExcelPdfImportOptions options) {
        using var bytes = new MemoryStream();
        var conversion = source.ConvertToExcelWithEvidence(options);
        using var convertedWorkbook = conversion.Value;
        convertedWorkbook.Save(bytes);
        PdfExcelTableImportReport report = conversion.Report;
        PdfExcelTableImportEntry entry = Assert.Single(report.Entries);
        bytes.Position = 0;
        using SpreadsheetDocument reopened = SpreadsheetDocument.Open(bytes, false);
        WorkbookPart workbook = reopened.WorkbookPart!;
        S.Worksheet worksheet = Assert.Single(workbook.WorksheetParts).Worksheet!;
        S.Column[] definitions = worksheet.GetFirstChild<S.Columns>()?.Elements<S.Column>().ToArray() ?? Array.Empty<S.Column>();
        double defaultWidth = worksheet.GetFirstChild<S.SheetFormatProperties>()?.DefaultColumnWidth?.Value ?? 8.43D;
        double[] widths = Enumerable.Range(1, entry.ColumnCount).Select(index => definitions
            .FirstOrDefault(column => column.Min!.Value <= index && column.Max!.Value >= index)?.Width?.Value ?? defaultWidth).ToArray();
        S.SharedStringItem[] strings = workbook.SharedStringTablePart?.SharedStringTable?.Elements<S.SharedStringItem>().ToArray()
            ?? Array.Empty<S.SharedStringItem>();
        S.Row[] rows = worksheet.Descendants<S.Row>().ToArray();
        string[][] text = rows.Select(row => row.Elements<S.Cell>().Select(cell => cell.DataType?.Value == S.CellValues.SharedString
            ? strings[int.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture)].InnerText
            : cell.InlineString?.InnerText ?? cell.CellValue?.Text ?? string.Empty).ToArray()).ToArray();
        decimal?[][] numbers = rows.Select(row => row.Elements<S.Cell>().Select(cell =>
            cell.DataType == null || cell.DataType.Value == S.CellValues.Number
                ? (decimal?)decimal.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture) : null).ToArray()).ToArray();
        return new Artifact(widths, entry, text, numbers);
    }

    private static Artifact RoundTripOds(PdfDocumentReadResult source, ExcelPdfImportOptions options) {
        PdfOdsConversionResult result = source.ToOdsDocumentResult(options);
        PdfExcelTableImportEntry entry = Assert.Single(result.Report.PdfReport.Entries);
        using var bytes = new MemoryStream(result.Value.ToBytes());
        OdsDocument reopened = OdsDocument.Load(bytes);
        OdsSheet sheet = Assert.Single(reopened.Sheets);
        double[] widths = Enumerable.Range(0, entry.ColumnCount).Select(index => {
            OdfLength? length = sheet.Column(index).Width;
            if (length.HasValue) {
                Assert.True(length.Value.TryToPoints(out double points));
                return points;
            }
            // The unchanged no-custom-width route uses the spreadsheet owner's
            // existing default-width approximation when displayed as points.
            return (8.43D * 7D + 5D) * 72D / 96D;
        }).ToArray();
        OdsUsedRange range = sheet.UsedRange!.Value;
        Assert.Equal(0L, range.FirstRow);
        Assert.Equal(0L, range.FirstColumn);
        Assert.Equal((long)entry.ColumnCount, range.ColumnCount);
        var text = new string[checked((int)range.RowCount)][];
        var numbers = new decimal?[text.Length][];
        for (int row = 0; row < text.Length; row++) {
            text[row] = new string[entry.ColumnCount];
            numbers[row] = new decimal?[entry.ColumnCount];
            for (int column = 0; column < entry.ColumnCount; column++) {
                OdsCell cell = sheet.Cell(row, column);
                text[row][column] = cell.Text;
                if (cell.Value.Kind == OdsCellValueKind.Number) numbers[row][column] = cell.Value.AsDecimal();
            }
        }
        return new Artifact(widths, entry, text, numbers);
    }

    private static Fixture ReadFixture(string id, string family = "TableGeometry") {
        string root = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", family);
        using JsonDocument document = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(root, "manifest.json")));
        JsonElement item = document.RootElement.GetProperty("cases").EnumerateArray()
            .Single(value => value.GetProperty("id").GetString() == id);
        byte[] pdf = File.ReadAllBytes(Path.Combine(root, id + "-native.pdf"));
        using SHA256 hash = SHA256.Create();
        Assert.Equal(item.GetProperty("files").GetProperty("native.pdf").GetString(),
            string.Concat(hash.ComputeHash(pdf).Select(value => value.ToString("x2", CultureInfo.InvariantCulture))));
        bool hasAcceptance = item.TryGetProperty("acceptance", out JsonElement acceptance);
        return new Fixture(pdf,
            item.GetProperty("table").EnumerateArray().Select(row => row.EnumerateArray().Select(cell => cell.GetString()!).ToArray()).ToArray(),
            item.TryGetProperty("geometry", out JsonElement geometry)
                ? geometry.GetProperty("normalizedColumnWidths").EnumerateArray().Select(value => value.GetDouble()).ToArray()
                : Array.Empty<double>(),
            hasAcceptance ? acceptance.GetProperty("maximumXlsxNormalizedShareError").GetDouble() : 0.000001D,
            hasAcceptance ? acceptance.GetProperty("maximumOdsNormalizedShareError").GetDouble() : 0.02D);
    }

    private sealed class Fixture {
        internal Fixture(byte[] pdf, string[][] rows, double[] authoredShares,
            double xlsxTolerance, double odsTolerance) {
            Pdf = pdf; Rows = rows; AuthoredShares = authoredShares;
            XlsxTolerance = xlsxTolerance; OdsTolerance = odsTolerance;
        }
        internal byte[] Pdf { get; }
        internal string[][] Rows { get; }
        internal double[] AuthoredShares { get; }
        private double XlsxTolerance { get; }
        private double OdsTolerance { get; }
        internal double Tolerance(string route) => route == "xlsx" ? XlsxTolerance : OdsTolerance;
    }
    private sealed class Artifact {
        internal Artifact(double[] widths, PdfExcelTableImportEntry entry, string[][] text, decimal?[][] numbers) {
            Widths = widths; Entry = entry; Text = text; Numbers = numbers;
        }
        internal double[] Widths { get; }
        internal PdfExcelTableImportEntry Entry { get; }
        internal string[][] Text { get; }
        internal decimal?[][] Numbers { get; }
    }

    // A caller-supplied detector is a supported production boundary. It can
    // retain correct cells while supplying incomplete geometry; no test hook is needed.
    private sealed class IncompleteGeometryTableStage : IPdfTableDetectionStage {
        public IReadOnlyList<PdfUnderstandingTableCandidate> DetectTables(
            PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingLine> lines) {
            PdfUnderstandingTableCandidate candidate = Assert.Single(
                PdfAdvancedUnderstandingStages.TableDetection.DetectTables(context, lines));
            PdfUnderstandingTableColumn[] columns = candidate.Columns.Select((column, index) =>
                new PdfUnderstandingTableColumn(column.From, index == 1 ? column.From : column.To)).ToArray();
            return new[] { new PdfUnderstandingTableCandidate(candidate.DetectionKind, candidate.YTop, candidate.YBottom,
                columns, candidate.Rows, candidate.SourceLines, candidate.Confidence, candidate.Evidence) };
        }
    }
}
