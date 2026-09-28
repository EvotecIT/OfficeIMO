using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using System.Threading;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData("◦", "    - Nested")]
    [InlineData("3)", "    3. Nested")]
    public void Reader_lists_use_markdown_markers_at_the_source_depth(
        string sourceLabel, string expected) {
        var style = new IWorkTextStyle(null, null, null, null, null,
            null, null, null, null);
        var paragraphStyle = new IWorkParagraphStyle(null, null, null, null,
            null, null, null, null, null, null, style);
        var paragraph = new IWorkTextParagraph(
            new[] { new IWorkTextRun("Nested", style, null) }, paragraphStyle,
            listIdentifier: null, listLevel: 2, listLabel: sourceLabel,
            breakKind: IWorkParagraphBreakKind.None);

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
}
