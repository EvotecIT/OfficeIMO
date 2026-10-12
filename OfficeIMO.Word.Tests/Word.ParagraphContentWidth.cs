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
        first.IndentationBeforePoints = 20;
        first.IndentationAfterPoints = 10;
        first.IndentationFirstLinePoints = 15;
        WordSection second = document.AddSection();
        SetPage(second, 18000);
        WordParagraph last = document.AddParagraph("Second section");

        Assert.Equal(455D, first.GetContentWidthPoints());
        Assert.Equal(800D, last.GetContentWidthPoints());
        second.Margins.Gutter = 1000;
        Assert.Equal(750D, last.GetContentWidthPoints());
        document.Settings.GutterAtTop = true;
        Assert.Equal(800D, last.GetContentWidthPoints());
    }

    [Fact]
    public void ContentWidthReservesHangingIndentWithoutAddingFirstLineIndent() {
        using var document = WordDocument.Create();
        SetPage(document.Sections[0], 12000);
        WordParagraph paragraph = document.AddParagraph("Hanging paragraph");
        paragraph.IndentationBeforePoints = 30;
        paragraph.IndentationAfterPoints = 10;
        paragraph.IndentationHangingPoints = 15;
        paragraph.IndentationFirstLinePoints = 25;

        Assert.Equal(460D, paragraph.GetContentWidthPoints());
    }

    [Fact]
    public void ContentWidthFitsEqualAndUnequalFlowingColumns() {
        using var document = WordDocument.Create();
        WordSection section = document.Sections[0];
        SetPage(section, 18000);
        section.ColumnCount = 3;
        section.ColumnsSpace = 300;
        WordParagraph paragraph = document.AddParagraph("Column content");

        Assert.Equal(770D / 3D, paragraph.GetContentWidthPoints(), 6);
        section.ColumnDefinitions = new[] {
            new WordSectionColumn(3600, 300), new WordSectionColumn(6000, 600)
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
        cell.MarginLeftWidth = 200;
        cell.MarginRightWidth = 200;
        WordParagraph paragraph = cell.AddParagraph("Cell content");
        double initialWidth = paragraph.GetContentWidthPoints();
        Assert.InRange(initialWidth, 90D, 110D);
        paragraph.IndentationBeforePoints = 10;
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
            cell.MarginLeftWidth = 200;
            cell.MarginRightWidth = 200;
            paragraph = cell.AddParagraph("Cell chart");
        } else {
            paragraph = document.AddParagraph("Second section chart");
        }
        WordChart chart = paragraph.AddChart(OfficeChartKind.Line,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Values", new[] { 2D, 4D })
            }));
        chart.SetWidthToPageContent(0.5, heightPx: 160);
        using var stream = new System.IO.MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        using WordDocument reloaded = WordDocument.Load(stream);
        WordChart persisted = Assert.Single(reloaded.Charts);
        Assert.True(persisted.TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot layout));
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

        second.PageSettings.Width = 18000;
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => paragraph.GetContentWidthPoints());
        Assert.Contains("explicit drawing dimensions", error.Message);
    }

    [Fact]
    public void ContentWidthRejectsParagraphWithNoUsableSpace() {
        using var document = WordDocument.Create();
        SetPage(document.Sections[0], 12000);
        WordParagraph paragraph = document.AddParagraph("No usable space");
        paragraph.IndentationBeforePoints = 500;
        Assert.Throws<InvalidOperationException>(() => paragraph.GetContentWidthPoints());
    }

    private static void SetPage(WordSection section, uint widthTwips) {
        section.PageSettings.Width = widthTwips;
        section.Margins.Left = 1000;
        section.Margins.Right = 1000;
        section.Margins.Gutter = 0;
    }
}
