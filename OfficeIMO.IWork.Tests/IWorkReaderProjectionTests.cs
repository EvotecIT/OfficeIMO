using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using System.Threading;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData("◦", IWorkListMarkerKind.Text, "    - Nested")]
    [InlineData("1)", IWorkListMarkerKind.Number, "    1. Nested")]
    [InlineData("1)", IWorkListMarkerKind.Text, "    - Nested")]
    public void Reader_lists_use_markdown_markers_at_the_source_depth(
        string sourceLabel, IWorkListMarkerKind kind, string expected) {
        var style = new IWorkTextStyle(null, null, null, null, null,
            null, null, null, null);
        var paragraphStyle = new IWorkParagraphStyle(null, null, null, null,
            null, null, null, null, null, null, style);
        var paragraph = new IWorkTextParagraph(
            new[] { new IWorkTextRun("Nested", style, null) }, paragraphStyle,
            listIdentifier: null, listLevel: 2, listLabel: sourceLabel,
            breakKind: IWorkParagraphBreakKind.None, listMarkerKind: kind);

        Assert.Equal(expected, IWorkReadProjection.RichTextMarkdown(paragraph));
        Assert.Equal(sourceLabel, paragraph.ListLabel);
    }

    [Fact]
    public void Reader_table_preserves_description_and_null_formula_cache() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.numbers"), IWorkDocumentKind.Numbers);
        var cells = new[] {
            new IWorkTableCell(1, 1, IWorkCellKind.Formula, null,
                formula: "=A2", valueKind: IWorkCellKind.Text),
            new IWorkTableCell(1, 2, IWorkCellKind.Formula, null,
                formula: "=B2", error: "#DIV/0!", valueKind: IWorkCellKind.Error)
        };
        var table = new IWorkTable("Results", 1, 2, cells,
            accessibilityDescription: "Quarterly results");
        var sheet = new IWorkNumbersSheet("Sheet 1", new[] { table }, Array.Empty<string>());
        var numbers = new IWorkNumbersProjection(source, new[] { sheet },
            Array.Empty<IWorkDiagnostic>(), supportsEditableReconstruction: true);
        var result = new OfficeDocumentReadResult { Kind = ReaderInputKind.IWork };
        var projection = new IWorkReadProjection(result, "results.numbers",
            new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);

        projection.AddNumbers(numbers);
        projection.Complete(source);

        ReaderTable output = Assert.Single(result.Tables);
        Assert.Equal("Quarterly results", output.Summary);
        Assert.Equal(new[] { string.Empty, "#DIV/0!" }, Assert.Single(output.Rows));
        Assert.DoesNotContain("=A2", result.Markdown);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_FORMULA_CACHE");
    }

    [Fact]
    public void Reader_table_reports_unrepresented_header_column_and_footer_row_roles() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.numbers"), IWorkDocumentKind.Numbers);
        var table = new IWorkTable("Results", 3, 2, new[] {
            new IWorkTableCell(1, 1, IWorkCellKind.Text, "Label"),
            new IWorkTableCell(3, 2, IWorkCellKind.Number, 42d)
        }, headerRowCount: 1, headerColumnCount: 1, footerRowCount: 1);
        var sheet = new IWorkNumbersSheet("Sheet 1", new[] { table }, Array.Empty<string>());
        var numbers = new IWorkNumbersProjection(source, new[] { sheet },
            Array.Empty<IWorkDiagnostic>(), supportsEditableReconstruction: true);
        var result = new OfficeDocumentReadResult { Kind = ReaderInputKind.IWork };
        var projection = new IWorkReadProjection(result, "results.numbers",
            new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);

        projection.AddNumbers(numbers);
        projection.Complete(source);

        OfficeDocumentDiagnostic diagnostic = Assert.Single(result.Diagnostics,
            item => item.Code == "IWORK_READER_TABLE_ROLES_UNSUPPORTED");
        Assert.Equal("1", diagnostic.Attributes["headerColumnCount"]);
        Assert.Equal("1", diagnostic.Attributes["footerRowCount"]);
        Assert.Equal("42", Assert.Single(result.Tables).Rows[1][1]);
    }

    [Fact]
    public void Numbers_reader_keeps_table_before_later_text_shape() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("First", 1, 1, 42d)
        }, textBox: "After table");
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers);
        IWorkNumbersProjection numbers = source.ReadNumbers();
        Assert.Collection(Assert.Single(numbers.Sheets).Drawables,
            drawable => Assert.Equal(IWorkNumbersDrawableKind.Table, drawable.Kind),
            drawable => Assert.Equal(IWorkNumbersDrawableKind.TextBox, drawable.Kind));
        var result = new OfficeDocumentReadResult { Kind = ReaderInputKind.IWork };
        var projection = new IWorkReadProjection(result, "ordered.numbers",
            new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);

        projection.AddNumbers(numbers);
        projection.Complete(source);

        Assert.Equal(new[] { "table", "text-box" }, result.Blocks.Select(block => block.Kind));
        Assert.True(result.Markdown!.IndexOf("42", StringComparison.Ordinal)
            < result.Markdown.IndexOf("After table", StringComparison.Ordinal));
    }

    [Fact]
    public void Reader_table_exposes_profiles_and_recovered_geometry() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.numbers"), IWorkDocumentKind.Numbers);
        var table = new IWorkTable("Profiled", 2, 2, new[] {
            new IWorkTableCell(1, 1, IWorkCellKind.Number, 42d),
            new IWorkTableCell(1, 2, IWorkCellKind.Text, "A"),
            new IWorkTableCell(2, 1, IWorkCellKind.Number, 7d),
            new IWorkTableCell(2, 2, IWorkCellKind.Text, "B")
        }, geometry: new IWorkGeometry(12, 24, 300, 120, 0));
        var sheet = new IWorkNumbersSheet("Sheet 1", new[] { table }, Array.Empty<string>());
        var numbers = new IWorkNumbersProjection(source, new[] { sheet },
            Array.Empty<IWorkDiagnostic>(), supportsEditableReconstruction: true);
        var result = new OfficeDocumentReadResult { Kind = ReaderInputKind.IWork };
        var projection = new IWorkReadProjection(result, "profiled.numbers",
            new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);

        projection.AddNumbers(numbers);
        projection.Complete(source);

        ReaderTable output = Assert.Single(result.Tables);
        Assert.Equal(ReaderTableColumnKind.Numeric, output.ColumnProfiles[0].Kind);
        Assert.Equal(ReaderTableColumnKind.Text, output.ColumnProfiles[1].Kind);
        ReaderTableDiagnostics geometry = Assert.IsType<ReaderTableDiagnostics>(output.Diagnostics);
        Assert.True(geometry.HasGeometry);
        Assert.Equal((12d, 312d, 24d, 144d),
            (geometry.XStart, geometry.XEnd, geometry.YTop, geometry.YBottom));
        Assert.Equal((300d, 120d), (geometry.Width, geometry.Height));
        Assert.Equal(12d, Assert.Single(result.Blocks).Region!.X);
    }

    [Fact]
    public void Reader_table_reports_saturated_counts_for_large_sparse_geometry() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.numbers"), IWorkDocumentKind.Numbers);
        var table = new IWorkTable("Sparse", 1_048_576, 16_384, new[] {
            new IWorkTableCell(1, 1, IWorkCellKind.Number, 1d)
        }, geometry: new IWorkGeometry(1, 2, 3, 4, 0));
        var sheet = new IWorkNumbersSheet("Sheet 1", new[] { table }, Array.Empty<string>());
        var numbers = new IWorkNumbersProjection(source, new[] { sheet },
            Array.Empty<IWorkDiagnostic>(), supportsEditableReconstruction: true);
        var result = new OfficeDocumentReadResult { Kind = ReaderInputKind.IWork };
        var projection = new IWorkReadProjection(result, "sparse.numbers",
            new ReaderOptions { MaxTableRows = 1 }, new ReaderIWorkOptions {
                MaximumTableColumns = 1,
                MaximumProjectedTableCells = 1
            }, CancellationToken.None);

        projection.AddNumbers(numbers);
        projection.Complete(source);

        ReaderTableDiagnostics diagnostics = Assert.IsType<ReaderTableDiagnostics>(
            Assert.Single(result.Tables).Diagnostics);
        Assert.Equal(int.MaxValue, diagnostics.ExpectedCellCount);
        Assert.Equal(int.MaxValue, diagnostics.MissingCellCount);
        Assert.Equal(1, diagnostics.FilledCellCount);
        Assert.Equal(1d / 17_179_869_184d, diagnostics.CellCompleteness, 12);
        OfficeDocumentDiagnostic notice = Assert.Single(result.Diagnostics,
            item => item.Code == "IWORK_READER_TABLE_CELL_COUNTS_SATURATED");
        Assert.Equal("17179869184", notice.Attributes["sourceLogicalCellCount"]);
    }
}
