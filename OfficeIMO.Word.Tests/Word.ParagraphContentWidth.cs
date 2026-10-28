using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordParagraphContentWidthTests {
    [Fact]
    public void ContentWidthUsesOwningSectionAndParagraphIndents() {
        using var document = WordDocument.Create();
        SetPage(document.Sections[0], 12000);
        WordParagraph first = document.AddParagraph("First section");
        first.IndentationBefore = global::OfficeIMO.OfficeLength.Points(20D);
        first.IndentationAfter = global::OfficeIMO.OfficeLength.Points(10D);
        first.IndentationFirstLine = global::OfficeIMO.OfficeLength.Points(15D);
        WordSection second = document.AddSection();
        SetPage(second, 18000);
        WordParagraph last = document.AddParagraph("Second section");

        Assert.Equal(455D, first.GetContentWidthPoints());
        Assert.Equal(800D, last.GetContentWidthPoints());
        second.Margins.Gutter = global::OfficeIMO.OfficeLength.Points(50D);
        Assert.Equal(750D, last.GetContentWidthPoints());
        document.Settings.GutterAtTop = true;
        Assert.Equal(800D, last.GetContentWidthPoints());
    }

    [Fact]
    public void ContentWidthReservesHangingIndentWithoutAddingFirstLineIndent() {
        using var document = WordDocument.Create();
        SetPage(document.Sections[0], 12000);
        WordParagraph paragraph = document.AddParagraph("Hanging paragraph");
        paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points(30D);
        paragraph.IndentationAfter = global::OfficeIMO.OfficeLength.Points(10D);
        paragraph.IndentationHanging = global::OfficeIMO.OfficeLength.Points(15D);
        paragraph.IndentationFirstLine = global::OfficeIMO.OfficeLength.Points(25D);

        Assert.Equal(460D, paragraph.GetContentWidthPoints());
    }

    [Fact]
    public void ContentWidthFitsEqualAndUnequalFlowingColumns() {
        using var document = WordDocument.Create();
        WordSection section = document.Sections[0];
        SetPage(section, 18000);
        section.ColumnCount = 3;
        WordParagraph paragraph = document.AddParagraph("Column content");
        section.ColumnsSpace = null;
        Assert.Equal(728D / 3D, paragraph.GetContentWidthPoints(), 6);
        section.ColumnsSpace = OfficeLength.Zero;
        Assert.Equal(800D / 3D, paragraph.GetContentWidthPoints(), 6);
        section.ColumnsSpace = OfficeLength.Points(-1);
        Assert.Equal(800D / 3D, paragraph.GetContentWidthPoints(), 6);
        section.ColumnsSpace = OfficeLength.Points(15);
        Assert.Equal(770D / 3D, paragraph.GetContentWidthPoints(), 6);
        section.ColumnDefinitions = new[] {
            new WordSectionColumn(OfficeLength.Points(180), OfficeLength.Points(15)),
            new WordSectionColumn(OfficeLength.Points(300), OfficeLength.Points(30))
        };
        Assert.Equal(180D, paragraph.GetContentWidthPoints());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ContentWidthUsesContainingCellAndMergedCellGrid(bool gridSpan) {
        using var document = WordDocument.Create();
        SetPage(document.Sections[0], 18000);
        WordTable table = document.AddTable(1, 2);
        table.GridColumnWidth = new List<int> { 2400, 4800 };
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.MarginLeft = global::OfficeIMO.OfficeLength.Points(10D);
        cell.MarginRight = global::OfficeIMO.OfficeLength.Points(10D);
        WordParagraph paragraph = cell.AddParagraph("Cell content");
        double initialWidth = paragraph.GetContentWidthPoints();
        Assert.InRange(initialWidth, 90D, 110D);
        paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points(10D);
        Assert.Equal(initialWidth - 10D, paragraph.GetContentWidthPoints());

        if (gridSpan) {
            cell._tableCell.TableCellProperties!.GridSpan = new DocumentFormat.OpenXml.Wordprocessing.GridSpan { Val = 2 };
            table.Rows[0].Cells[1]._tableCell.Remove();
        } else {
            table.MergeCells(0, 0, 1, 2, copyParagraphs: true);
        }
        Assert.Equal(2, cell.ColumnSpan);
        Assert.InRange(paragraph.GetContentWidthPoints(), 320D, 350D);
    }

    [Theory]
    [InlineData(false, 396D)]
    [InlineData(true, 169.5D)]
    public void NativeChartContentWidthPersistsItsSectionOrCellGeometry(bool inCell, double expectedWidthPoints) {
        using var document = WordDocument.Create();
        SetPage(document.Sections[0], 12000);
        document.AddParagraph("First section");
        WordSection second = document.AddSection();
        SetPage(second, 17840);
        WordParagraph paragraph;
        if (inCell) {
            WordTable table = document.AddTable(1, 1);
            table.GridColumnWidth = new List<int> { 7200 };
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.MarginLeft = global::OfficeIMO.OfficeLength.Points(10D);
            cell.MarginRight = global::OfficeIMO.OfficeLength.Points(10D);
            paragraph = cell.AddParagraph("Cell chart");
        } else {
            paragraph = document.AddParagraph("Second section chart");
        }
        WordChart chart = paragraph.AddChart(OfficeChartType.Line,            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeSeries("Values", new[] { 2D, 4D })
            }));
        chart.SetWidthToPageContent(0.5,height: OfficeIMO.OfficeLength.Pixels(160));
        using var stream = new System.IO.MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        using WordDocument reloaded = WordDocument.Load(stream);
        WordChart persisted = Assert.Single(reloaded.Charts);
        Assert.True(persisted.TryGetLayoutSnapshot(out WordDrawingLayoutInfo layout));
        Assert.Equal(expectedWidthPoints, layout.WidthPoints, 6);
        Assert.Equal(120D, layout.HeightPoints, 6);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ContentWidthRejectsAmbiguousInheritedHeaderOrFooter(bool footer) {
        using var document = WordDocument.Create();
        SetPage(document.Sections[0], 12000);
        document.AddParagraph("Body");
        document.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? document.Footer.Default! : document.Header.Default!;
        WordParagraph paragraph = story.AddParagraph("Shared story");
        WordSection second = document.AddSection();
        SetPage(second, 12000);
        Assert.Equal(500D, paragraph.GetContentWidthPoints());

        second.PageSettings.Width = global::OfficeIMO.OfficeLength.Points(900D);
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => paragraph.GetContentWidthPoints());
        Assert.Contains("explicit drawing dimensions", error.Message);
    }

    [Fact]
    public void ContentWidthRejectsParagraphWithNoUsableSpace() {
        using var document = WordDocument.Create();
        SetPage(document.Sections[0], 12000);
        WordParagraph paragraph = document.AddParagraph("No usable space");
        paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points(500D);
        Assert.Throws<InvalidOperationException>(() => paragraph.GetContentWidthPoints());
    }

    private static void SetPage(WordSection section, uint widthTwips) {
        section.PageSettings.Width = global::OfficeIMO.OfficeLength.Points((widthTwips) / 20D);
        section.Margins.Left = global::OfficeIMO.OfficeLength.Points(50D);
        section.Margins.Right = global::OfficeIMO.OfficeLength.Points(50D);
        section.Margins.Gutter = global::OfficeIMO.OfficeLength.Zero;
    }
}
