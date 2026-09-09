using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class PdfAnchoredImagePaginationTests {
    [Theory]
    [InlineData(false, "paragraph")]
    [InlineData(true, "paragraph")]
    [InlineData(false, "heading")]
    [InlineData(true, "heading")]
    [InlineData(false, "panel")]
    [InlineData(true, "panel")]
    [InlineData(false, "list")]
    [InlineData(true, "list")]
    [InlineData(false, "split")]
    [InlineData(true, "split")]
    [InlineData(false, "keepnext")]
    [InlineData(true, "keepnext")]
    [InlineData(false, "multiple")]
    [InlineData(true, "multiple")]
    public void ForegroundImageFollowsAnchorParagraphToItsFirstPage(bool columns, string kind) {
        using WordDocument word = WordDocument.Create();
        var section = word.Sections[0];
        section.PageSettings.Width = 6000;
        section.PageSettings.Height = 6000;
        if (columns) { section.ColumnCount = 2; section.ColumnsSpace = 120; }
        for (int index = 0; index < 16; index++) {
            var filler = word.AddParagraph("Filler " + index);
            filler.FontSize = 12;
            filler.LineSpacingAfterPoints = 0;
        }
        var anchor = kind == "list" ? word.AddList(WordListStyle.Bulleted).AddItem(string.Empty) : word.AddParagraph();
        anchor.FontSize = 12;
        anchor._paragraph!.ParagraphProperties ??= new ParagraphProperties();
        if (kind != "split") anchor._paragraph.ParagraphProperties.KeepLines = new KeepLines();
        if (kind == "keepnext") anchor._paragraph.ParagraphProperties.KeepNext = new KeepNext();
        if (kind == "heading") anchor.SetStyle(WordParagraphStyles.Heading1);
        if (kind == "panel") anchor.ShadingFillColorHex = "E6F2FF";
        byte[] bytes = OfficeIMO.Drawing.OfficeRasterImageEncoder.Encode(
            new OfficeIMO.Drawing.OfficeRasterImage(20, 20, OfficeIMO.Drawing.OfficeColor.Black), OfficeIMO.Drawing.OfficeImageExportFormat.Png);
        using var stream = new MemoryStream(bytes);
        var image = anchor.InsertImage(stream, "anchored.png", 24, 24, WordImageTextWrapping.InFrontOfText);
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        image.HorizontalPositionOffset = 10 * 12700;
        image.VerticalPositionOffset = 10 * 12700;
        if (kind == "multiple") {
            using var secondStream = new MemoryStream(bytes);
            var second = anchor.AddText(string.Empty).InsertImage(secondStream, "second.png", 24, 24, WordImageTextWrapping.InFrontOfText);
            second.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
            second.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
            second.HorizontalPositionOffset = 40 * 12700;
            second.VerticalPositionOffset = 10 * 12700;
        }
        anchor.AddText("ANCHOR TEXT " + string.Join(" ", Enumerable.Repeat("wrapped text", kind == "split" ? 80 : 4))).FontSize = 12;
        word.AddParagraph("Following paragraph");
        var result = word.ToPdfDocumentResult(new WordToPdfOptions { Margins = PdfCore.PageMargins.Uniform(30) });
        var pdf = PdfCore.PdfDocument.Load(result.Value.ToBytes());
        var anchorPage = Assert.Single(pdf.Reader.Pages(), page => pdf.Reader.Text(PdfCore.PdfPageSelection.From(page.PageNumber)).Contains("ANCHOR"));
        var placements = pdf.Images.Placements();
        Assert.Equal(kind == "multiple" ? 2 : 1, placements.Count);
        Assert.All(placements, item => Assert.Equal(anchorPage.PageNumber, item.PageNumber));
        var placement = placements[0];
        if (!columns && kind == "paragraph") Assert.True(anchorPage.PageNumber > 1, "The fixture must force the anchor onto a later page.");
        Assert.Equal(anchorPage.PageNumber, placement.PageNumber);
        Assert.Equal(10, placement.X, 3);
        Assert.Equal(272, placement.Y, 3);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeAnchoredImageFlowed");
        if (kind == "split") Assert.True(pdf.Reader.Pages().Count > anchorPage.PageNumber);
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_OUTPUT") is { Length: > 0 } output) {
            Directory.CreateDirectory(output);
            pdf.Save(Path.Combine(output, $"anchor-{kind}-{columns}.pdf"));
            foreach (var page in pdf.Render.Pages(PdfCore.PdfPageSelection.From(anchorPage.PageNumber), new PdfCore.PdfPageRenderOptions { Dpi = 96 }))
                File.WriteAllBytes(Path.Combine(output, $"anchor-{kind}-{columns}-{page.PageNumber}.png"), page.Bytes!);
        }
    }
}
