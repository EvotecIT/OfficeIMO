using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfClosedRulingGridTests {
    private const string RulingEvidence = "table-closed-ruling-grid";

    [Theory]
    [InlineData("ruled-ledger-0")]
    [InlineData("ruled-ledger-90")]
    public void StructuredReadRecoversIndependentAuthoredStrokeCentreGrid(string id) {
        Fixture fixture = LoadFixture(id, "manifest.json");
        PdfLogicalTable table = Assert.Single(Read(fixture.Pdf).Tables);

        AssertRows(fixture, table);
        Assert.Contains(table.Evidence, evidence => evidence.Code == RulingEvidence);
        double[] boundaries = Numbers(fixture.Label.GetProperty("geometry"), "columnBoundariesPoints");
        double[] rows = Numbers(fixture.Label.GetProperty("geometry"), "rowBoundariesPoints");
        for (int index = 0; index < table.Columns.Count; index++) {
            // The independent oracle is the authored PDF grid, not a second OfficeIMO
            // extraction or the detected proportions propagated by an importer.
            Assert.Equal(boundaries[index], table.Columns[index].From, 6);
            Assert.Equal(boundaries[index + 1], table.Columns[index].To, 6);
        }
        Assert.Equal(rows[0], table.YTop, 6);
        Assert.Equal(rows[rows.Length - 1], table.YBottom, 6);
    }

    public static IEnumerable<object[]> Controls() {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(AssetRoot, "ruling-controls-manifest.json")));
        foreach (JsonElement item in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            yield return new object[] { item.GetProperty("id").GetString()!, item.GetProperty("expectedAdmission").GetBoolean() };
        }
    }

    [Theory]
    [MemberData(nameof(Controls))]
    public void RulingAdmissionRequiresCompleteUnambiguousVisibleGridWithUnchangedCells(string id, bool admitted) {
        Fixture fixture = LoadFixture(id, "ruling-controls-manifest.json");
        PdfLogicalTable table = Assert.Single(Read(fixture.Pdf).Tables);
        AssertRows(fixture, table);
        Assert.Equal(admitted, table.Evidence.Any(evidence => evidence.Code == RulingEvidence));

        if (!admitted) {
            // Adding insufficient ruling evidence must preserve the supported
            // text-only result, including its geometry, order and duplicate values.
            Fixture control = LoadFixture("ruled-control-no-rules", "ruling-controls-manifest.json");
            PdfLogicalTable fallback = Assert.Single(Read(control.Pdf).Tables);
            Assert.Equal(fallback.Columns.Select(column => (column.From, column.To)),
                table.Columns.Select(column => (column.From, column.To)));
            Assert.Equal(fallback.YTop, table.YTop);
            Assert.Equal(fallback.YBottom, table.YBottom);
            return;
        }

        JsonElement geometry = fixture.Label.GetProperty("geometry");
        JsonElement[] expectedStrips = geometry.GetProperty("columnStripsVisualPoints").EnumerateArray().ToArray();
        PdfReadPage raw = PdfReadDocument.Open(fixture.Pdf).Pages[0];
        for (int index = 0; index < table.Columns.Count; index++) {
            PdfLogicalTableColumn column = table.Columns[index];
            PdfVisualBounds actual = column.VisualBounds is PdfLogicalVisualBounds visual
                ? new PdfVisualBounds(visual.Left, visual.Top, visual.Right, visual.Bottom)
                : raw.TransformBoundsToVisual(column.From, table.YBottom, column.To, table.YTop);
            double[] expected = expectedStrips[index].EnumerateArray().Select(value => value.GetDouble()).ToArray();
            Assert.Equal(expected[0], actual.Left, 6);
            Assert.Equal(expected[1], actual.Top, 6);
            Assert.Equal(expected[2], actual.Right, 6);
            Assert.Equal(expected[3], actual.Bottom, 6);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CollectorHonorsOperationWorkLimitAndCancellation(bool cancel) {
        Fixture fixture = LoadFixture("ruled-ledger-0", "manifest.json");
        PdfReadPage page = PdfReadDocument.Open(fixture.Pdf).Pages[0];
        using var cancellation = new CancellationTokenSource();
        var context = new PdfUnderstandingPageContext(page, 1, new PdfTextLayoutOptions(),
            100_000, 10_000, cancel ? 10_000 : 1, cancellation.Token);
        if (cancel) cancellation.Cancel();

        if (cancel) Assert.Throws<OperationCanceledException>(() => page.GetTableRulings(context));
        else {
            PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => page.GetTableRulings(context));
            Assert.Equal(PdfReadLimitKind.UnderstandingWork, exception.Kind);
        }
    }

    [Fact]
    public void RefinerWorkExhaustionDoesNotPartiallyChangeRecoveredGeometry() {
        Fixture fixture = LoadFixture("ruled-ledger-0", "manifest.json");
        PdfReadPage page = PdfReadDocument.Open(fixture.Pdf).Pages[0];
        IReadOnlyList<PdfTextSpan> runs = page.GetTextSpans();
        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(runs);
        PdfStructuredTable table = Assert.Single(TableDetector.DetectTablesFromBands(TextLayoutEngine.BandLines(lines), 792D));
        var context = new PdfUnderstandingPageContext(page, 1, new PdfTextLayoutOptions(),
            100_000, 10_000, 1) { DecodedRuns = runs };
        double[] x = Numbers(fixture.Label.GetProperty("geometry"), "columnBoundariesPoints");
        double[] y = Numbers(fixture.Label.GetProperty("geometry"), "rowBoundariesPoints");
        PdfTableRuling[] rulings = y.Select(value => new PdfTableRuling(x[0], value, x[x.Length - 1], value))
            .Concat(x.Select(value => new PdfTableRuling(value, y[0], value, y[y.Length - 1]))).ToArray();
        (double From, double To)[] original = table.Columns.Select(column => (column.From, column.To)).ToArray();
        double top = table.YTop, bottom = table.YBottom;

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            TableDetector.TryRefineClosedRulingGrid(table, rulings, context));

        Assert.Equal(PdfReadLimitKind.UnderstandingWork, exception.Kind);
        Assert.Equal(original, table.Columns.Select(column => (column.From, column.To)));
        Assert.Equal(top, table.YTop);
        Assert.Equal(bottom, table.YBottom);
    }

    private static string AssetRoot => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "TableGeometry");
    private static Fixture LoadFixture(string id, string manifestName) {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(AssetRoot, manifestName)));
        JsonElement label = manifest.RootElement.GetProperty("cases").EnumerateArray()
            .Single(item => item.GetProperty("id").GetString() == id).Clone();
        byte[] pdf = File.ReadAllBytes(Path.Combine(AssetRoot, id + "-native.pdf"));
        using SHA256 hash = SHA256.Create();
        Assert.Equal(label.GetProperty("files").GetProperty("native.pdf").GetString(),
            string.Concat(hash.ComputeHash(pdf).Select(value => value.ToString("x2", CultureInfo.InvariantCulture))));
        return new Fixture(pdf, label);
    }
    private static PdfDocumentReadResult Read(byte[] pdf) =>
        PdfDocument.Load(pdf).Read(new PdfLoadOptions { Profile = PdfReadProfile.Structured });
    private static double[] Numbers(JsonElement item, string name) =>
        item.GetProperty(name).EnumerateArray().Select(value => value.GetDouble()).ToArray();
    private static void AssertRows(Fixture fixture, PdfLogicalTable table) {
        string[][] expected = fixture.Label.GetProperty("table").EnumerateArray()
            .Select(row => row.EnumerateArray().Select(cell => cell.GetString()!).ToArray()).ToArray();
        Assert.Equal(expected.Length, table.Rows.Count);
        Assert.Equal(expected[0].Length, table.Columns.Count);
        for (int row = 0; row < expected.Length; row++) Assert.Equal(expected[row], table.Rows[row]);
        Assert.Equal(16, table.Cells.Count);
        Assert.Equal(2, table.Rows.SelectMany(row => row).Count(value => value == "89.00"));
        Assert.Equal(-13.50M, decimal.Parse(table.Rows[2][3], CultureInfo.InvariantCulture));
    }
    private sealed class Fixture {
        internal Fixture(byte[] pdf, JsonElement label) { Pdf = pdf; Label = label; }
        internal byte[] Pdf { get; }
        internal JsonElement Label { get; }
    }
}
