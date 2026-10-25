using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("header-anchor-explicit-defaults.docx", 21.45D, 40.20D, 66.36D)]
    [InlineData("header-anchor-explicit-defaults-preceding-line.docx", 33.45D, 52.20D, 67.32D)]
    public void SaveAsPdf_RunningHeaderPictureRetainsNativeParagraphAnchorAndRightTab(
        string fixture, double imageTop, double titleBaseline, double bodyBaseline) {
        // Fully synthetic sources explicitly declare paragraph defaults and
        // were schema-validated, exported by native Word, and inspected
        // independently. A preceding header paragraph moves
        // the paragraph-relative picture and title together, while body entry
        // follows the header's text rather than the picture's 49pt paint frame.
        string path = Path.Combine(AppContext.BaseDirectory, "WordLayoutFixtures", fixture);
        byte[] source = File.ReadAllBytes(path);
        using WordDocument document = WordDocument.Load(path);
        Assert.Empty(document.ValidateDocument());
        string sourceXml = AnchoredRunningImageSourceXml(document);

        var result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        var picture = Assert.Single(page.GetImages());
        Assert.InRange(Math.Abs(picture.BoundingBox.Left - 28.35D), 0D, .15D);
        Assert.InRange(Math.Abs(page.Height - picture.BoundingBox.Top - imageTop), 0D, .15D);
        Assert.Equal(128D, picture.BoundingBox.Width, 3);
        Assert.Equal(49D, picture.BoundingBox.Height, 3);

        var words = page.GetWords().ToArray();
        var title = Assert.Single(words, word => word.Text == "PROJECT");
        var titleEnd = Assert.Single(words, word => word.Text == "INITIATIVE");
        var body = Assert.Single(words, word => word.Text == "BODY");
        W.TabStop rightTab = Assert.Single(document._wordprocessingDocument.MainDocumentPart!
            .HeaderParts.Single().Header.Descendants<W.TabStop>(), tab => tab.Val?.Value == W.TabStopValues.Right);
        double rightTabAnchor = document.Sections[0].Margins.Left / 20D + rightTab.Position!.Value / 20D;
        // A right tab aligns the final advance, which is independent of the
        // selected face's ink overhang and the width of its preceding words.
        Assert.InRange(Math.Abs(titleEnd.Letters.Last().EndBaseLine.X - rightTabAnchor), 0D, .3D);
        Assert.InRange(Math.Abs(body.Letters[0].StartBaseLine.X - 28.35D), 0D, .15D);
        Assert.Contains("Generic form content.", page.Text);

        // Removing only the non-flow picture leaves the same ordinary text
        // paragraphs, mark/run formatting and tabs. These comparisons always
        // exercise line placement and body reservation, including font fallback.
        using WordDocument ordinaryHeader = WordDocument.Load(new MemoryStream(source));
        foreach (W.Drawing drawing in ordinaryHeader._wordprocessingDocument.MainDocumentPart!
                     .HeaderParts.SelectMany(part => part.Header.Descendants<W.Drawing>()).ToArray()) {
            drawing.Remove();
        }
        Assert.Empty(ordinaryHeader.ValidateDocument());
        using var ordinaryPdf = PdfPigDocument.Open(ordinaryHeader.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var ordinaryPage = Assert.Single(ordinaryPdf.GetPages());
        Assert.Empty(ordinaryPage.GetImages());
        var ordinaryWords = ordinaryPage.GetWords().ToArray();
        foreach (string token in new[] { "PROJECT", "INITIATIVE", "BODY" }) {
            AssertAnchoredRunningImageReferenceWord(
                Assert.Single(words, word => word.Text == token),
                Assert.Single(ordinaryWords, word => word.Text == token));
        }
        bool qualifiedCalibri = title.Letters.Concat(titleEnd.Letters).All(letter =>
                AnchoredRunningImageFontName(letter.FontName) == "Calibri-Bold") &&
            body.Letters.All(letter => IsAnchoredRunningImageRegularCalibri(letter.FontName));
        if (fixture.EndsWith("preceding-line.docx", StringComparison.Ordinal)) {
            var preceding = Assert.Single(words, word => word.Text == "HEADER");
            AssertAnchoredRunningImageReferenceWord(preceding,
                Assert.Single(ordinaryWords, word => word.Text == "HEADER"));
            qualifiedCalibri &= preceding.Letters.All(letter => IsAnchoredRunningImageRegularCalibri(letter.FontName));
            if (qualifiedCalibri) {
                Assert.InRange(Math.Abs(page.Height - preceding.Letters[0].StartBaseLine.Y - 30.84D), 0D, .3D);
            }
        }
        // Absolute native coordinates are qualified with these actual faces;
        // another installed face retains the source-relative contracts above.
        if (qualifiedCalibri) {
            Assert.InRange(Math.Abs(title.Letters[0].StartBaseLine.X - 416.35D), 0D, .3D);
            Assert.InRange(Math.Abs(page.Height - title.Letters[0].StartBaseLine.Y - titleBaseline), 0D, .3D);
            Assert.InRange(Math.Abs(page.Height - body.Letters[0].StartBaseLine.Y - bodyBaseline), 0D, .3D);
        }
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "NativeAnchoredImageFlowed");
        Assert.Equal(sourceXml, AnchoredRunningImageSourceXml(document));
        Assert.Equal(source, File.ReadAllBytes(path));
    }

    [Fact]
    public void SaveAsPdf_UnqualifiedRunningPictureKeepsBoundedImageAndTextProjection() {
        using WordDocument source = WordDocument.Load(Path.Combine(AppContext.BaseDirectory,
            "WordLayoutFixtures", "header-anchor-explicit-defaults.docx"));
        source.Sections[0].ColumnCount = 2;
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        string sourceXml = AnchoredRunningImageSourceXml(document);
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var page = Assert.Single(pdf.GetPages());
        var picture = Assert.Single(page.GetImages());
        // A not-yet-selected text column is outside the qualified anchor
        // contract. Preserve its readable projection rather than claim the
        // fixed left-margin placement before pagination selects a column.
        Assert.InRange(picture.BoundingBox.Left, 160D, page.Width - picture.BoundingBox.Width);
        foreach (string token in new[] { "PROJECT", "INITIATIVE", "BODY", "ANCHOR" })
            Assert.Single(page.GetWords(), word => word.Text == token);
        Assert.Equal(sourceXml, AnchoredRunningImageSourceXml(document));
    }

    private static string AnchoredRunningImageSourceXml(WordDocument document) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        return main.Document.OuterXml + main.DocumentSettingsPart!.Settings!.OuterXml +
            string.Concat(main.HeaderParts.Select(part => part.Header.OuterXml));
    }

    private static void AssertAnchoredRunningImageReferenceWord(
        UglyToad.PdfPig.Content.Word actual, UglyToad.PdfPig.Content.Word ordinary) {
        Assert.Equal(ordinary.Letters.Count, actual.Letters.Count);
        for (int index = 0; index < ordinary.Letters.Count; index++) {
            var expectedLetter = ordinary.Letters[index];
            var actualLetter = actual.Letters[index];
            Assert.Equal(expectedLetter.Value, actualLetter.Value);
            Assert.Equal(AnchoredRunningImageFontName(expectedLetter.FontName),
                AnchoredRunningImageFontName(actualLetter.FontName));
            Assert.Equal(expectedLetter.FontSize, actualLetter.FontSize, 3);
            Assert.Equal(expectedLetter.StartBaseLine.X, actualLetter.StartBaseLine.X, 3);
            Assert.Equal(expectedLetter.StartBaseLine.Y, actualLetter.StartBaseLine.Y, 3);
            Assert.Equal(expectedLetter.EndBaseLine.X, actualLetter.EndBaseLine.X, 3);
            Assert.Equal(expectedLetter.EndBaseLine.Y, actualLetter.EndBaseLine.Y, 3);
        }
    }

    private static bool IsAnchoredRunningImageRegularCalibri(string fontName) =>
        AnchoredRunningImageFontName(fontName) is "Calibri" or "Calibri-Regular";

    private static string AnchoredRunningImageFontName(string fontName) =>
        fontName.Length > 7 && fontName[6] == '+' &&
        fontName.Take(6).All(character => character >= 'A' && character <= 'Z')
            ? fontName.Substring(7) : fontName;
}
