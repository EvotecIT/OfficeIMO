using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRegionTextSourceReuseTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void RegionEditPreservesIndependentEmbeddedFontUnicodeAndNeighbor(bool move, bool encrypted) {
        string fixture = encrypted ? "source-font-aes-256.pdf" : "source-font-reportlab.pdf";
        var readOptions = new PdfLoadOptions { Password = encrypted ? "owner" : null };
        PdfDocument source = PdfDocument.Load(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "PdfEditing", fixture)), readOptions)
            .Text.Add(new PdfPageRegion(1, 50D, 590D, 200D, 30D), "kept neighbor", new PdfTextEditOptions {
                Font = PdfStandardFont.CourierBold,
                FontSize = 11D,
                Color = PdfColor.FromRgb(0, 0, 255)
            }).Document;
        byte[] original = source.ToBytes();
        byte[] originalSnapshot = original.ToArray();
        PdfFontInfo originalFont = Assert.Single(source.Resources.Fonts(new PdfFontInspectionOptions {
            IncludeEmbeddedProgramBytes = true
        }).Fonts, font => font.IsEmbedded);
        var independentOptions = new UglyToad.PdfPig.ParsingOptions { Password = encrypted ? "owner" : string.Empty };
        using var before = PdfPigDocument.Open(original, independentOptions);
        var originalWords = before.GetPage(1).GetWords().ToArray();
        var originalUnicode = Assert.Single(originalWords, word => word.Text == "Żółć");
        var originalNeighbor = Assert.Single(originalWords, word => word.Text == "neighbor");
        Assert.Equal("alpha beta gamma Żółć", string.Join(" ", originalWords
            .Where(word => word.Letters[0].StartBaseLine.Y > 650D)
            .OrderBy(word => word.BoundingBox.Left)
            .Select(word => word.Text)));

        var region = new PdfPageRegion(1, 49D, 685D, 250D, 30D);
        PdfTextEditResult result = move
            ? source.Text.Move(region, 20D, -30D)
            : source.Text.Replace(region, "gamma Żółć");
        byte[] output = result.Document.ToBytes();

        Assert.True(result.AffectedCount > 0);
        Assert.Empty(result.Warnings);
        Assert.Equal(originalSnapshot, source.ToBytes());
        PdfFontInfo editedFont = Assert.Single(result.Document.Resources.Fonts(new PdfFontInspectionOptions {
            IncludeEmbeddedProgramBytes = true
        }).Fonts, font => font.IsEmbedded);
        Assert.Equal(originalFont.BaseFontName, editedFont.BaseFontName);
        Assert.Equal(originalFont.EmbeddedProgramBytes, editedFont.EmbeddedProgramBytes);
        Assert.Equal(originalFont.ToUnicodeMappingCount, editedFont.ToUnicodeMappingCount);
        Assert.Equal(encrypted, PdfReadDocument.Open(output, readOptions).Security.HasEncryption);
        if (encrypted) Assert.Throws<PdfPasswordRequiredException>(() => PdfReadDocument.Open(output));

        using var after = PdfPigDocument.Open(output, independentOptions);
        var words = after.GetPage(1).GetWords().ToArray();
        var editedWords = words.Where(word => word.Letters[0].StartBaseLine.Y > 650D)
            .OrderBy(word => word.BoundingBox.Left).ToArray();
        Assert.Equal(move ? "alpha beta gamma Żółć" : "gamma Żółć", string.Join(" ", editedWords.Select(word => word.Text)));
        var editedLetters = editedWords.SelectMany(word => word.Letters).ToArray();
        Assert.Equal(move ? 70D : 50D, editedLetters.Min(letter => letter.StartBaseLine.X), 3);
        Assert.All(editedLetters, letter => {
            Assert.Equal(move ? 670D : 700D, letter.StartBaseLine.Y, 3);
            Assert.Equal(14D, letter.FontSize, 3);
            Assert.Equal(originalUnicode.Letters[0].FontName, letter.FontName);
            Assert.Equal(originalUnicode.Letters[0].Color.ToRGBValues(), letter.Color.ToRGBValues());
        });
        var neighbor = Assert.Single(words, word => word.Text == "neighbor");
        Assert.Equal(originalNeighbor.Letters.Select(letter => (
            letter.Value, letter.StartBaseLine.X, letter.StartBaseLine.Y, letter.FontSize, letter.FontName, letter.Color.ToRGBValues())),
            neighbor.Letters.Select(letter => (
                letter.Value, letter.StartBaseLine.X, letter.StartBaseLine.Y, letter.FontSize, letter.FontName, letter.Color.ToRGBValues())));
    }
}
