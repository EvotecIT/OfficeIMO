using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTableAutoFitRunFontTests {
    private const int Computer = 0x1F4BB;
    private static readonly string ComputerText = char.ConvertFromUtf32(Computer);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AutoFitMeasuresGlyphsInTheFontThatPaintsThem(bool useFallbackSet) {
        // The row font is an embedded program without the emoji, so measuring
        // the cell in the row font throws. Auto-fit has to use the run's named
        // font or the fallback the renderer plans for that run.
        byte[] rowFont = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, ' ', 'A', 'B');
        byte[] emojiFont = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 1000, ' ', Computer);
        var options = new PdfCreateOptions()
            .EmbedStandardFont(PdfStandardFont.Helvetica, rowFont, "RowFont")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, rowFont, "RowFont-Bold");
        PdfTextRun run;
        if (useFallbackSet) {
            options.RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate("Emoji", emojiFont)
            }));
            run = PdfTextRun.Normal("AB" + ComputerText);
        } else {
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Emoji", emojiFont, emojiFont, emojiFont, emojiFont));
            run = PdfTextRun.Normal(ComputerText, fontFamily: "Emoji");
        }

        PdfTableCell[][] rows = { new[] { PdfTableCell.RichTextCell(new[] { run }) } };
        var documents = new[] {
            PdfDocument.Create(options).Table(rows, style: new PdfTableStyle { HeaderRowCount = 0 }),
            PdfDocument.Create(options).Table(rows, style: new PdfTableStyle { HeaderRowCount = 1 }),
            PdfDocument.Create(document => document.Page(page => page.Content(content =>
                content.Item(item => item.AddTable(rows, global::OfficeIMO.OfficeAlignment.Left)))), options)
        };
        foreach (PdfDocument document in documents) {
            byte[] pdf = document.ToBytes();

            using var independent = UglyToad.PdfPig.PdfDocument.Open(pdf);
            var letters = independent.GetPage(1).Letters;
            var emoji = Assert.Single(letters, letter => letter.Value == ComputerText);
            Assert.Contains("+Emoji", emoji.FontName ?? string.Empty, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void AutoFitMeasuresAsciiGlyphMissingFromEmbeddedRowFontInItsFallback() {
        int[] latinWithoutZ = Enumerable.Range(32, 95).Where(scalar => scalar != 'Z').ToArray();
        byte[] rowFont = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, latinWithoutZ);
        var options = new PdfCreateOptions()
            .EmbedStandardFont(PdfStandardFont.Helvetica, rowFont, "RowFont")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, rowFont, "RowFont-Bold")
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate("Zed", ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, ' ', 'Z'))
            }));

        byte[] pdf = PdfDocument.Create(document => document.Page(page => page.Content(content =>
            content.Item(item => item.AddTable(new[] { new[] { PdfTableCell.RichTextCell(new[] { PdfTextRun.Normal("AZ") }) } },
                global::OfficeIMO.OfficeAlignment.Left)))), options).ToBytes();

        using var independent = UglyToad.PdfPig.PdfDocument.Open(pdf);
        var zed = Assert.Single(independent.GetPage(1).Letters, letter => letter.Value == "Z");
        Assert.Contains("+Zed", zed.FontName ?? string.Empty, StringComparison.Ordinal);
    }
}
