using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class PdfForegroundImageOrderTests {
    [Theory]
    [InlineData(false, false, 10U, 1U)]
    [InlineData(false, true, 10U, 1U)]
    [InlineData(true, false, 10U, 1U)]
    [InlineData(true, true, 10U, 1U)]
    [InlineData(false, false, 1U, 10U)]
    [InlineData(true, true, 10U, 10U)]
    public void ForegroundStackingUsesZOrderAcrossParagraphsAndColumns(bool sameParagraph, bool columns, uint firstOrder, uint secondOrder) {
        using WordDocument word = WordDocument.Create();
        word.Sections[0].PageSettings.Width = 6000;
        word.Sections[0].PageSettings.Height = 6000;
        if (columns) { word.Sections[0].ColumnCount = 2; word.Sections[0].ColumnsSpace = 120; }
        var first = word.AddParagraph("First anchor");
        AddImage(first, OfficeColor.Red, firstOrder, 10);
        var second = sameParagraph ? first.AddText("Second anchor") : word.AddParagraph("Second anchor");
        AddImage(second, OfficeColor.Blue, secondOrder, 28);
        var result = word.ToPdfDocumentResult(new WordToPdfOptions { Margins = PdfCore.PageMargins.Uniform(30) });
        var pdf = PdfCore.PdfDocument.Load(result.Value.ToBytes());
        Assert.Equal(2, pdf.Images.Placements().Count);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeAnchoredImageFlowed");
        byte[] rendered = pdf.Render.Pages(PdfCore.PdfPageSelection.From(1), new PdfCore.PdfPageRenderOptions { Dpi = 72 })[0].Bytes!;
        Assert.True(OfficePngReader.TryDecode(rendered, out var bitmap));
        byte[] pixels = bitmap!.GetPixels();
        int offset = (40 * bitmap.Width + 40) * 4;
        bool redAbove = firstOrder > secondOrder;
        Assert.Equal(redAbove ? (byte)255 : (byte)0, pixels[offset]);
        Assert.Equal((byte)0, pixels[offset + 1]);
        Assert.Equal(redAbove ? (byte)0 : (byte)255, pixels[offset + 2]);
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_OUTPUT") is { Length: > 0 } output) {
            Directory.CreateDirectory(output);
            File.WriteAllBytes(Path.Combine(output, $"zorder-{sameParagraph}-{columns}-{firstOrder}-{secondOrder}.png"), rendered);
        }
    }

    private static void AddImage(WordParagraph paragraph, OfficeColor color, uint order, int position) {
        using var stream = new MemoryStream(OfficeRasterImageEncoder.Encode(new OfficeRasterImage(20, 20, color), OfficeImageExportFormat.Png));
        var image = paragraph.InsertImage(stream, "layer.png", 72, 72, WordImageTextWrapping.InFrontOfText);
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        image.HorizontalPositionOffset = position * 12700;
        image.VerticalPositionOffset = position * 12700;
        image.ZOrder = order;
    }
}
