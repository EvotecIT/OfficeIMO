using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(1, 60, 132D)]
    [InlineData(2, 60, 258D)]
    [InlineData(3, 60, 384D)]
    [InlineData(1, 120, 144D)]
    [InlineData(2, 120, 276D)]
    [InlineData(3, 120, 408D)]
    public void SaveAsPdf_SpacedTableRetainsCellGapsAndIndependentPerimeter(int columns, short spacing, double width) {
        using WordDocument source = WordDocument.Create();
        CreateBorderFrameControl(source, 2, columns, spacing);
        using WordDocument imported = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdfBytes(BorderFramePdfOptions()));
        Assert.Equal(xml, imported._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        var page = pdf.GetPage(1);
        var boxes = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(box => box.HasValue).Select(box => box!.Value).ToArray();
        Assert.Equal(width, boxes.Max(box => box.Right) - boxes.Min(box => box.Left), 2);
        var words = page.GetWords().ToArray();
        Assert.Equal(2 * columns, words.Length);
        if (columns > 1) {
            double distance = words.Single(word => word.Text == "Frame1").BoundingBox.Left -
                words.Single(word => word.Text == "Frame0").BoundingBox.Left;
            Assert.Equal(120D + spacing / 10D, distance, 3);
        }
        Assert.True(boxes.Min(box => box.Bottom) < words.Min(word => word.BoundingBox.Bottom) - spacing / 10D);
    }

    [Theory]
    [InlineData(false, 480, 24D, 30.5D)]
    [InlineData(false, 960, 48D, 54.5D)]
    [InlineData(true, 480, 37D, 37D)]
    [InlineData(true, 960, 61D, 61D)]
    public void SaveAsPdf_SpacedTableHonorsExactAndMinimumRowHeights(bool minimum, int height, double firstAdvance, double nextAdvance) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 3, 2, 120);
        foreach (WordTableRow row in table.Rows) {
            if (minimum) row.MinimumHeight = height;
            else row.Height = height;
        }
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(BorderFramePdfOptions()));
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.Equal(6, words.Length); // The first exact-height cell retains its partial line.
        double[] baseline = Enumerable.Range(0, 3).Select(row =>
            words.Single(word => word.Text == $"Frame{row * 2}").Letters[0].StartBaseLine.Y).ToArray();
        Assert.Equal(firstAdvance, baseline[0] - baseline[1], 3);
        Assert.Equal(nextAdvance, baseline[1] - baseline[2], 3);
    }

    [Fact]
    public void SaveAsPdf_SpacedTablePreservesDistinctOuterBordersAndCellBorders() {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 2, 2, 120);
        table._tableProperties!.TableBorders = new W.TableBorders(
            new W.TopBorder { Val = W.BorderValues.Single, Color = "FF0000", Size = 16U },
            new W.LeftBorder { Val = W.BorderValues.Single, Color = "0000FF", Size = 24U },
            new W.BottomBorder { Val = W.BorderValues.Single, Color = "800080", Size = 16U },
            new W.RightBorder { Val = W.BorderValues.Single, Color = "00FF00", Size = 16U });
        foreach (WordTableRow row in table.Rows)
            foreach (WordTableCell cell in row.Cells) {
                cell.Borders.TopStyle = cell.Borders.BottomStyle = cell.Borders.LeftStyle = cell.Borders.RightStyle = WordBorderStyle.Single;
                cell.Borders.TopSize = cell.Borders.BottomSize = cell.Borders.LeftSize = cell.Borders.RightSize = 4U;
                cell.Borders.TopColorHex = cell.Borders.BottomColorHex = cell.Borders.LeftColorHex = cell.Borders.RightColorHex = "000000";
            }
        string raw = PdfOperatorSearchText.From(document.ToPdfBytes(BorderFramePdfOptions()));
        Assert.Contains("1 0 0 RG", raw);
        Assert.Contains("0 0 1 RG", raw);
        Assert.Contains("0 1 0 RG", raw);
        Assert.Contains("0.502 0 0.502 RG", raw);
        Assert.Contains("0 0 0 RG", raw);
    }

    [Theory]
    [InlineData(WordTableAlignment.Center, 185.9D)]
    [InlineData(WordTableAlignment.Right, 281.65D)]
    public void SaveAsPdf_SpacedTableAlignsItsWholeFrame(WordTableAlignment alignment, double firstTextX) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 2, 2, 120);
        table.Alignment = alignment;
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(BorderFramePdfOptions()));
        var first = pdf.GetPage(1).GetWords().Single(word => word.Text == "Frame0");
        Assert.Equal(firstTextX, first.Letters[0].StartBaseLine.X, 2);
    }

    [Fact]
    public void SaveAsPdf_SpacedTableDoubleCellBordersReserveTheirWholePaintWidth() {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 2, 2, 120);
        using PdfPigDocument before = PdfPigDocument.Open(document.ToPdfBytes(BorderFramePdfOptions()));
        double textX = before.GetPage(1).GetWords().Single(word => word.Text == "Frame0").Letters[0].StartBaseLine.X;
        foreach (WordTableRow row in table.Rows)
            foreach (WordTableCell cell in row.Cells) {
                cell.Borders.TopStyle = cell.Borders.BottomStyle = cell.Borders.LeftStyle = cell.Borders.RightStyle = WordBorderStyle.Double;
                cell.Borders.TopSize = cell.Borders.BottomSize = cell.Borders.LeftSize = cell.Borders.RightSize = 12U;
                cell.Borders.TopColorHex = cell.Borders.BottomColorHex = cell.Borders.LeftColorHex = cell.Borders.RightColorHex = "000000";
            }
        using PdfPigDocument after = PdfPigDocument.Open(document.ToPdfBytes(BorderFramePdfOptions()));
        double borderedX = after.GetPage(1).GetWords().Single(word => word.Text == "Frame0").Letters[0].StartBaseLine.X;
        Assert.Equal(6D, borderedX - textX, 3);
        Assert.Equal(4, after.GetPage(1).GetWords().Count());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void SaveAsPdf_SpacedTableFlowsBetweenSectionColumns(bool exact, bool repeatHeader) {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].ColumnCount = 2;
        document.Sections[0].ColumnsSpace = 720;
        WordTable table = CreateBorderFrameControl(document, 40, 2, 120);
        for (int row = 0; row < 40; row++) {
            if (exact) table.Rows[row].Height = 480;
            for (int column = 0; column < 2; column++)
                table.Rows[row].Cells[column].Paragraphs[0].Text = $"M{row * 2 + column}";
        }
        table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = repeatHeader;
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(BorderFramePdfOptions()));
        Assert.Equal(1, pdf.NumberOfPages);
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.Equal(80, words.Select(word => word.Text).Distinct().Count());
        Assert.Equal(repeatHeader ? 2 : 1, words.Count(word => word.Text == "M0"));
        Assert.Contains(words, word => word.Text == "M79" && word.BoundingBox.Left > 306D);
        Assert.All(words, word => Assert.True(word.BoundingBox.Bottom >= 72D && word.BoundingBox.Right <= 540D));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_SpacedTableRepeatsHeadersAtTheSamePositionAndKeepsItsPerimeterInsideThePage(bool exact) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 40, 2, 120);
        table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
        if (exact) foreach (WordTableRow row in table.Rows) row.Height = 480;
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(BorderFramePdfOptions()));
        Assert.Equal(2, pdf.NumberOfPages);
        double? header = null;
        var markers = new HashSet<string>();
        foreach (var page in pdf.GetPages()) {
            var words = page.GetWords().ToArray();
            var currentHeader = words.Single(word => word.Text == "Frame0");
            if (header.HasValue) Assert.Equal(header.Value, currentHeader.Letters[0].StartBaseLine.Y, 3);
            else header = currentHeader.Letters[0].StartBaseLine.Y;
            foreach (var word in words) markers.Add(word.Text);
            double borderBottom = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
                .Where(box => box.HasValue).Min(box => box!.Value.Bottom);
            Assert.True(borderBottom >= 72D, $"Table perimeter exceeds the bottom margin: {borderBottom}");
        }
        Assert.Equal(80, markers.Count);
    }

    private static WordToPdfOptions BorderFramePdfOptions() => new() {
        IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
    };

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_SpacedTableRetainsGapShadingAndExplicitStyleClearing(bool inherited, bool clear) {
        using WordDocument source = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(source, 2, 2, 120);
        var shade = new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "80C0FF" };
        if (inherited) {
            source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
                new W.Style(new W.BasedOn { Val = "TableGrid" }, new W.StyleTableProperties(shade)) {
                    StyleId = "FrameGapFill", Type = W.StyleValues.Table
                });
            table._tableProperties!.TableStyle = new W.TableStyle { Val = "FrameGapFill" };
        } else table._tableProperties!.Append(shade);
        if (clear) table._tableProperties!.Append(new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "auto" });
        foreach (WordTableRow row in table.Rows)
            foreach (WordTableCell cell in row.Cells) cell.ShadingFillColorHex = "FFFFFF";
        using WordDocument imported = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        byte[] bytes = imported.ToPdfBytes(BorderFramePdfOptions());
        Assert.Equal(xml, imported._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        string raw = PdfOperatorSearchText.From(bytes);
        if (clear) Assert.DoesNotContain("0.502 0.753 1 rg", raw);
        else Assert.Contains("0.502 0.753 1 rg", raw);
        Assert.Contains("1 1 1 rg", raw);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(4, pdf.GetPage(1).GetWords().Count());
    }

    private static WordTable CreateBorderFrameControl(WordDocument document, int rows, int columns, short spacing) {
        WordTable table = CreateAutomaticWidthControl(document, Enumerable.Repeat(2400, columns).ToArray());
        while (table.Rows.Count < rows) table.AddRow();
        table.StyleDetails!.CellSpacing = spacing;
        for (int row = 0; row < rows; row++)
            for (int column = 0; column < columns; column++) {
                WordTableCell cell = table.Rows[row].Cells[column];
                cell.Width = 2400; cell.WidthType = WordTableWidthUnit.Dxa;
                cell.MarginLeftWidth = cell.MarginRightWidth = 108;
                cell.MarginTopWidth = cell.MarginBottomWidth = 0;
                WordParagraph paragraph = cell.Paragraphs[0];
                paragraph.Text = $"Frame{row * columns + column}";
                paragraph.FontFamily = "Arial"; paragraph.FontSize = 12;
                paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
                paragraph.LineSpacingRule = WordLineSpacingRule.Auto; paragraph.LineSpacing = 240;
            }
        table.GridColumnWidth = Enumerable.Repeat(2400, columns).ToList();
        return table;
    }
}
