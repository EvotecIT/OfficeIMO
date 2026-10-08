using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordTextDirection.TopToBottomRightToLeft)]
    [InlineData(WordTextDirection.BottomToTopLeftToRight)]
    public void SaveAsPdf_AutomaticTurnedGridWithoutPreferredWidthsUsesVisibleRunLineBox(WordTextDirection direction) {
        double first = AutomaticTurnedGridWidth(direction, 6);
        double second = AutomaticTurnedGridWidth(direction, 24);
        Assert.InRange(first, 18D, 30D);
        Assert.Equal(first, second, 3);
    }

    private static double AutomaticTurnedGridWidth(WordTextDirection direction, int markSize) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 1, 1, 0);
        table.LayoutMode = WordTableLayoutMode.AutoFit;
        table._tableProperties!.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Auto, Width = "0" };
        table.GridColumnWidth = new List<int> { 0 };
        table.Rows[0].Height = 2400;
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.WidthType = WordTableWidthUnit.Auto; cell.Width = 0;
        cell.TextDirection = direction;
        WordParagraph paragraph = cell.Paragraphs[0];
        paragraph.Text = "ABCDEFGHIJKLMNOPQRSTUVWXYZABCDEFGHIJKLMNOPQRSTUVWXYZABCDEFGHIJKLMNOPQRSTUVWXYZ";
        paragraph._paragraph.ParagraphProperties!.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties(
            new W.FontSize { Val = (markSize * 2).ToString(System.Globalization.CultureInfo.InvariantCulture) });
        string xml = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(BorderFramePdfOptions()));
        Assert.Equal(xml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        var page = Assert.Single(pdf.GetPages());
        var frame = Assert.Single(page.Paths.Where(path => path.IsStroked)
            .Select(path => path.GetBoundingRectangle()).Where(bounds => bounds.HasValue)
            .Select(bounds => bounds!.Value));
        var firstLetter = page.Letters.First(letter => letter.Value == "A");
        Assert.InRange(firstLetter.StartBaseLine.X, frame.Left, frame.Right);
        return frame.Width;
    }

    [Theory]
    [InlineData(WordTextDirection.LeftToRightTopToBottom, 0)]
    [InlineData(WordTextDirection.TopToBottomRightToLeft, -1)]
    [InlineData(WordTextDirection.BottomToTopLeftToRight, 1)]
    public void SaveAsPdf_CellDirectionPreservesTextAndSource(WordTextDirection direction, int verticalSign) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 1, 1, 0);
        table.Rows[0].Height = 2400;
        table.Rows[0].Cells[0].TextDirection = direction;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "ABCD";
        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(direction, imported.Tables[0].Rows[0].Cells[0].TextDirection);
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdfBytes(BorderFramePdfOptions()));
        Assert.Equal(xml, imported._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        var page = Assert.Single(pdf.GetPages());
        Assert.Equal("ABCD", string.Concat(page.Letters.Select(letter => letter.Value)));
        foreach (var letter in page.Letters) {
            double dx = letter.EndBaseLine.X - letter.StartBaseLine.X;
            double dy = letter.EndBaseLine.Y - letter.StartBaseLine.Y;
            if (verticalSign == 0) {
                Assert.True(dx > 0D);
                Assert.Equal(0D, dy, 3);
            } else {
                Assert.Equal(0D, dx, 3);
                Assert.True(dy * verticalSign > 0D);
            }
        }
    }

    [Theory]
    [InlineData(WordParagraphAlignment.Left)]
    [InlineData(WordParagraphAlignment.Center)]
    [InlineData(WordParagraphAlignment.Right)]
    public void SaveAsPdf_UprightCellPicturesRetainSizeAndPhysicalParagraphAlignment(WordParagraphAlignment alignment) {
        var clockwise = CellDirectionImageBounds(WordTextDirection.TopToBottomRightToLeft, alignment);
        var counterclockwise = CellDirectionImageBounds(WordTextDirection.BottomToTopLeftToRight, alignment);
        Assert.Equal(40D, clockwise.Width, 3);
        Assert.Equal(20D, clockwise.Height, 3);
        Assert.Equal(clockwise.Width, counterclockwise.Width, 3);
        Assert.Equal(clockwise.Height, counterclockwise.Height, 3);
        Assert.Equal(clockwise.Top, counterclockwise.Top, 3);
    }

    [Theory]
    [InlineData(WordTextDirection.LeftToRightTopToBottom, false)]
    [InlineData(WordTextDirection.TopToBottomRightToLeft, true)]
    [InlineData(WordTextDirection.BottomToTopLeftToRight, true)]
    public void SaveAsPdf_UprightTurnedCellPicturePaintsAboveOverlappingFollowingText(WordTextDirection direction, bool pictureAbove) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 1, 1, 0);
        table.Rows[0].Height = 2400;
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.TextDirection = direction;
        WordParagraph paragraph = cell.Paragraphs[0];
        paragraph.Text = "Before";
        using var image = new MemoryStream(Convert.FromBase64String(CellDirectionImagePng));
        paragraph.AddImage(image, "direction.png", 40D * 4D / 3D, 20D * 4D / 3D,
            description: "Red left half and blue right half");
        paragraph.AddText("After");
        byte[] bytes = document.ToPdfBytes(BorderFramePdfOptions());
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        Assert.Single(page.GetImages());
        Assert.Equal("BeforeAfter", string.Concat(page.Letters.Select(letter => letter.Value)));
        string operators = PdfOperatorSearchText.From(bytes);
        int picture = operators.LastIndexOf(" Do", StringComparison.Ordinal);
        int text = Math.Max(operators.LastIndexOf(" Tj", StringComparison.Ordinal), operators.LastIndexOf(" TJ", StringComparison.Ordinal));
        Assert.True(picture >= 0 && text >= 0);
        Assert.Equal(pictureAbove, picture > text);
    }

    [Theory]
    [InlineData(6)]
    [InlineData(12)]
    [InlineData(24)]
    public void SaveAsPdf_AutomaticTurnedCellRowUsesParagraphMarkInsteadOfVisibleRunSize(int markSize) {
        Assert.Equal(TurnedCellFollowingBaseline(markSize, 18), TurnedCellFollowingBaseline(markSize, 36), 3);
    }

    [Fact]
    public void SaveAsPdf_AutomaticTurnedCellMarkSizeChangesItsPhysicalFlowHeight() {
        Assert.True(TurnedCellFollowingBaseline(6, 18) > TurnedCellFollowingBaseline(24, 18));
    }

    private static double TurnedCellFollowingBaseline(int markSize, int runSize) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 1, 1, 0);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        table.Rows[0].Cells[0].TextDirection = WordTextDirection.TopToBottomRightToLeft;
        paragraph.Text = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        paragraph.FontSize = runSize;
        paragraph._paragraph.ParagraphProperties!.ParagraphMarkRunProperties =
            new W.ParagraphMarkRunProperties(new W.FontSize { Val = (markSize * 2).ToString(System.Globalization.CultureInfo.InvariantCulture) });
        WordParagraph after = document.AddParagraph("After");
        after.FontSize = 12;
        after.LineSpacingBeforePoints = after.LineSpacingAfterPoints = 0;
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(BorderFramePdfOptions()));
        return Assert.Single(Assert.Single(pdf.GetPages()).GetWords(), word => word.Text == "After")
            .Letters[0].StartBaseLine.Y;
    }

    private static (double Width, double Height, double Top) CellDirectionImageBounds(WordTextDirection direction, WordParagraphAlignment alignment) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 1, 1, 0);
        table.Rows[0].Height = 2400;
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.TextDirection = direction;
        WordParagraph paragraph = cell.Paragraphs[0];
        paragraph.Text = string.Empty;
        paragraph.ParagraphAlignment = alignment;
        using var imageStream = new MemoryStream(Convert.FromBase64String(CellDirectionImagePng));
        paragraph.AddImage(imageStream, "direction.png", 40D * 4D / 3D, 20D * 4D / 3D,
            description: "Red left half and blue right half");
        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdfBytes(BorderFramePdfOptions()));
        Assert.Equal(xml, imported._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        var image = Assert.Single(Assert.Single(pdf.GetPages()).GetImages());
        return (image.BoundingBox.Width, image.BoundingBox.Height, image.BoundingBox.Top);
    }

    [Theory]
    [InlineData(WordTextDirection.TopToBottomRightToLeft)]
    [InlineData(WordTextDirection.BottomToTopLeftToRight)]
    public void SaveAsPdf_AutomaticTurnedCellOmitsColumnsStartingOutsideItsContentFrame(WordTextDirection direction) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 1, 1, 0);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.TextDirection = direction;
        WordParagraph paragraph = cell.Paragraphs[0];
        paragraph.Text = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
        paragraph.LineSpacing = 300;
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(BorderFramePdfOptions()));
        // The 109.2-point content frame retains the partially visible eighth 15-point column.
        var page = Assert.Single(pdf.GetPages());
        Assert.Equal(8, page.Letters.Select(letter => Math.Round(letter.StartBaseLine.X, 3)).Distinct().Count());
        Assert.Equal("ABCDEFGHI", string.Concat(page.Letters.Select(letter => letter.Value)));
    }

    private const string CellDirectionImagePng = "iVBORw0KGgoAAAANSUhEUgAAAFAAAAAoCAIAAADmAupWAAAAWklEQVR4nOXOQQEAMAiAQKR/Z9diPrgCMMuN4aYsMRIjMRIjMRIjMRIjMRIjMRIjMRIjMRIjMRIjMRIjMRIjMRIjMRIjMRIjMRIjMRIjMRIjMRIjMRLj9cBvD7rzAk/vtHpzAAAAAElFTkSuQmCC";
}
