using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

// Noto Color Emoji and Apple Color Emoji glyphs are PNG images; each Word story places them on its page.
public sealed class WordPdfColorEmojiTests {
    private const string Text = "\U0001F600\U0001F600";

    [Theory]
    [InlineData("body", false)]
    [InlineData("body", true)]
    [InlineData("cell", false)]
    [InlineData("cell", true)]
    [InlineData("header", false)]
    [InlineData("header", true)]
    [InlineData("footer", false)]
    [InlineData("footer", true)]
    public void PngColorGlyphIsPaintedInEveryStory(string route, bool sbix) {
        using var document = WordDocument.Create();
        WordParagraph paragraph = route switch {
            "cell" => document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0],
            "header" => document.PageSetup.Header.AddParagraph(),
            "footer" => document.PageSetup.Footer.AddParagraph(),
            _ => document.AddParagraph()
        };
        paragraph._paragraph.RemoveAllChildren<Run>();
        paragraph._paragraph.Append(new Run(new RunProperties(
            new RunFonts { Ascii = ManagedTextShapingTestAssets.FamilyName, HighAnsi = ManagedTextShapingTestAssets.FamilyName },
            new FontSize { Val = "80" }), new Text(Text)));
        // Courier body text has no red ink, so any red comes from the emoji.
        if (route is "header" or "footer") document.AddParagraph("Body").FontFamily = "Courier";
        var options = new PdfCreateOptions { IncludeStandardFontToUnicodeMaps = true };
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateBitmapColorFont(0x1F600, sbix)));

        byte[] pdf = document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, PdfOptions = options, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes();

        // Headers and footers are pagination artifacts, which text extraction leaves out; the page text keeps them.
        using (PdfPigDocument read = PdfPigDocument.Open(pdf)) Assert.Contains(Text, read.GetPage(1).Text, StringComparison.Ordinal);
        Assert.True(CountRed(pdf) > 0, "The PNG glyph is painted.");
    }

    private static int CountRed(byte[] pdf) {
        OfficeRasterImage image = OfficeRasterImageDecoder.Decode(PdfPageImageRenderer.RenderPageAsPng(pdf, 1, 1D));
        int red = 0;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor color = image.GetPixel(x, y);
                if (color.R > 200 && color.G < 60 && color.B < 60) red++;
            }
        }
        return red;
    }
}
