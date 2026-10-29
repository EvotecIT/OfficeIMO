using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(100, 87.8D)]
    [InlineData(125, 109.7D)]
    public void SaveAsPdf_FontUnitPrecisionRetainsKerningAcrossPaintRuns(int scale, double bodyWidth) {
        int[] expected = new[] { 20, 20 };
        var actual = new List<int>();
        foreach (bool split in new[] { false, true }) {
            using WordDocument document = WordDocument.Create();
            document.Sections[0].PageSettings.Width = global::OfficeIMO.OfficeLength.Points(((uint)Math.Round((bodyWidth + 40D) * 20D)) / 20D);
            document.Sections[0].Margins.Left = global::OfficeIMO.OfficeLength.Points(20D); document.Sections[0].Margins.Right = global::OfficeIMO.OfficeLength.Points(20D);
            string text = string.Concat(Enumerable.Repeat("AV", 10));
            WordParagraph paragraph = document.AddParagraph();
            int index = 0;
            foreach (string fragment in split ? text.Select(c => c.ToString()) : new[] { text }) {
                WordParagraph run = paragraph.AddText(fragment);
                run.FontFamily = ManagedTextShapingTestAssets.FamilyName; run.FontSize = global::OfficeIMO.OfficeLength.Points(20);
                run.CharacterScale = scale; run.KerningMinimumFontSize = global::OfficeIMO.OfficeLength.Points(1);
                run.ColorHex = index++ % 2 == 0 ? "CC0000" : "0000CC";
            }
            byte[] font = ManagedTextShapingTestAssets.CreateFontWithFontUnitKerning(2048, 'A', 'V', -101);
            var options = new PdfCreateOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                ManagedTextShapingTestAssets.FamilyName, font));
            using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
                IncludePageNumbers = false, PdfOptions = options,
                ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
            }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
            var letters = Assert.Single(pdf.GetPages()).Letters.Where(letter => letter.Value is "A" or "V").ToArray();
            Assert.Equal(text, string.Concat(letters.Select(letter => letter.Value)));
            actual.Add(letters.Count(letter => Math.Abs(letter.StartBaseLine.Y - letters[0].StartBaseLine.Y) < .1D));
        }
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void SaveAsPdf_FontUnitPrecisionStartsSeparateOversizedWordsOnFreshLines() {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].PageSettings.Width = global::OfficeIMO.OfficeLength.Points(65D);
        document.Sections[0].Margins.Left = global::OfficeIMO.OfficeLength.Points(20D); document.Sections[0].Margins.Right = global::OfficeIMO.OfficeLength.Points(20D);
        WordParagraph paragraph = document.AddParagraph("AAAAAAAA BBBBBBBB");
        paragraph.FontFamily = ManagedTextShapingTestAssets.FamilyName; paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(12);
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(2048, 1229, ' ', 'A', 'B');
        var options = new PdfCreateOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
            ManagedTextShapingTestAssets.FamilyName, font));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, PdfOptions = options,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var lines = Assert.Single(pdf.GetPages()).Letters.Where(letter => letter.Value != " ")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 2))
            .Select(group => string.Concat(group.Select(letter => letter.Value))).ToArray();
        Assert.Equal(new[] { "AAA", "AAA", "AA", "BBB", "BBB", "BB" }, lines);
    }

    [Theory]
    [InlineData(1229, 10, 100, 0, 72D, 9, false)]
    [InlineData(1229, 10, 100, 0, 72D, 9, true)]
    [InlineData(1229, 20, 100, 0, 72D, 9, false)]
    [InlineData(1229, 20, 100, 0, 72D, 9, true)]
    [InlineData(1228, 20, 100, 0, 143.95D, 20, false)]
    [InlineData(1228, 20, 100, 0, 143.95D, 20, true)]
    [InlineData(1933, 15, 80, 0, 135.9D, 15, false)]
    [InlineData(1933, 15, 80, 0, 135.9D, 15, true)]
    [InlineData(1822, 15, 125, 4, 203.15D, 15, false)]
    [InlineData(1822, 15, 125, 4, 203.15D, 15, true)]
    public void SaveAsPdf_FontUnitPrecisionControlsWrappingAcrossRuns(
        ushort advance, int length, int scale, int spacing, double bodyWidth, int firstLineLetters, bool split) {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].PageSettings.Width = global::OfficeIMO.OfficeLength.Points(((uint)Math.Round((bodyWidth + 40D) * 20D)) / 20D);
        document.Sections[0].Margins.Left = global::OfficeIMO.OfficeLength.Points(20D); document.Sections[0].Margins.Right = global::OfficeIMO.OfficeLength.Points(20D);
        string text = new string('A', length) + " B";
        WordParagraph paragraph = document.AddParagraph(split ? text.Substring(0, length / 2) : text);
        Configure(paragraph);
        if (split) Configure(paragraph.AddText(text.Substring(length / 2)));
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(2048, advance, ' ', 'A', 'B');
        var options = new PdfCreateOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
            ManagedTextShapingTestAssets.FamilyName, font));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, PdfOptions = options,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = Assert.Single(pdf.GetPages()).Letters.Where(letter => letter.Value != " ").ToArray();
        Assert.Equal(new string('A', length) + "B", string.Concat(letters.Select(letter => letter.Value)));
        double firstBaseline = letters[0].StartBaseLine.Y;
        Assert.Equal(firstLineLetters, letters.Count(letter => letter.Value == "A"
            && Math.Abs(letter.StartBaseLine.Y - firstBaseline) < .1D));
        int[] lineCounts = letters.Where(letter => letter.Value == "A")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 2)).Select(group => group.Count()).ToArray();
        Assert.Equal(Enumerable.Range(0, (length + firstLineLetters - 1) / firstLineLetters)
            .Select(index => Math.Min(firstLineLetters, length - index * firstLineLetters)), lineCounts);

        void Configure(WordParagraph run) {
            run.FontFamily = ManagedTextShapingTestAssets.FamilyName; run.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            run.CharacterScale = scale; run.Spacing = spacing;
        }
    }
}
