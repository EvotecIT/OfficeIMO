using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordPdfPageFootnoteTests {
    [Fact]
    public void SaveAsPdf_FootnoteRuleSurvivesEquallyStyledFollowingTextInJoinedAndSplitRuns() {
        const string id = "geometry-indents-alignment-mixed";
        string folder = Path.Combine(AppContext.BaseDirectory, "Documents", "WordPageFootnoteSeparators");
        string inputPath = Path.Combine(folder, id + ".docx");
        string referencePath = Path.Combine(folder, id + ".microsoft-word.pdf");
        byte[] input = File.ReadAllBytes(inputPath);
        byte[] reference = File.ReadAllBytes(referencePath);
        Assert.Equal(FootnoteRuleFixtureHashes[id].Input, PageFootnoteSha256(input));
        Assert.Equal(FootnoteRuleFixtureHashes[id].Reference, PageFootnoteSha256(reference));

        PageFootnoteArtifact joined = ConvertFootnoteRuleRunMergingFixture(input, reference, splitRuns: false);
        PageFootnoteArtifact split = ConvertFootnoteRuleRunMergingFixture(input, reference, splitRuns: true);
        using var joinedPdf = PdfPigDocument.Open(joined.Actual);
        using var splitPdf = PdfPigDocument.Open(split.Actual);
        using var native = PdfPigDocument.Open(reference);
        Assert.Equal(1, joinedPdf.NumberOfPages);
        Assert.Equal(joinedPdf.NumberOfPages, splitPdf.NumberOfPages);
        Assert.Equal(PageFootnoteCompact(joinedPdf.GetPage(1).Text), PageFootnoteCompact(splitPdf.GetPage(1).Text));
        Assert.Equal(joinedPdf.GetPage(1).Letters.Select(ImplicitFootnoteGlyphGeometry).ToArray(),
            splitPdf.GetPage(1).Letters.Select(ImplicitFootnoteGlyphGeometry).ToArray());

        var rows = new[] {
            new FootnoteRuleBand(50.064, 363.460, 144.050, .600),
            new FootnoteRuleBand(146.090, 377.260, 132.050, .600),
            new FootnoteRuleBand(146.090, 391.060, 132.050, .600),
            new FootnoteRuleBand(146.090, 404.856, 132.050, .624),
            new FootnoteRuleBand(180.790, 418.680, 97.344, .600)
        };
        foreach (PageFootnoteArtifact artifact in new[] { joined, split }) {
            using var actual = PdfPigDocument.Open(artifact.Actual);
            var page = actual.GetPage(1);
            var lead = PageFootnoteToken(page, "LEAD");
            double noteBaseline = PageFootnoteToken(page, "SHORT NOTE BODY")[0].StartBaseLine.Y;
            var tail = page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value) &&
                letter.StartBaseLine.Y < lead[0].StartBaseLine.Y - .01D && letter.StartBaseLine.Y > noteBaseline &&
                Math.Abs(letter.PointSize - lead[0].PointSize) < .001D).ToArray();
            Assert.Equal("TAIL", string.Concat(tail.Select(letter => letter.Value)));
            AssertPageFootnoteNativeProfile(artifact, () => {
                AssertFootnoteRuleBands(page, native.GetPage(1), rows);
                AssertFootnoteRuleTextPosition(page, native.GetPage(1), "LEAD");
                AssertPageFootnoteNativeText(page, native.GetPage(1), "SHORT NOTE BODY",
                    PageFootnoteNativeAnchor.PhysicalBottom);
            });
        }
        Assert.Equal(input, File.ReadAllBytes(inputPath));
        Assert.Equal(reference, File.ReadAllBytes(referencePath));
    }

    private PageFootnoteArtifact ConvertFootnoteRuleRunMergingFixture(byte[] input, byte[] reference, bool splitRuns) {
        using WordDocument document = WordDocument.Load(new MemoryStream(input));
        var main = document._wordprocessingDocument.MainDocumentPart!;
        W.Footnote separator = Assert.Single(main.FootnotesPart!.Footnotes!.Elements<W.Footnote>(),
            note => note.Type?.Value == W.FootnoteEndnoteValues.Separator);
        W.Paragraph mixed = Assert.Single(separator.Elements<W.Paragraph>(),
            paragraph => paragraph.InnerText.Contains("LEAD", StringComparison.Ordinal));
        W.Run marker = Assert.Single(mixed.Elements<W.Run>(), run => run.Elements<W.SeparatorMark>().Any());
        var properties = (W.RunProperties)marker.RunProperties!.CloneNode(true);
        mixed.RemoveAllChildren<W.Run>();
        // The inherited and explicit styles are visually equivalent in the
        // qualified source. Make all three explicit and identical so the text
        // merger cannot rely on a color/style difference to preserve the rule.
        if (splitRuns) {
            mixed.Append(new W.Run(properties.CloneNode(true), new W.Text("LEAD ") { Space = SpaceProcessingModeValues.Preserve }),
                new W.Run(properties.CloneNode(true), new W.SeparatorMark()),
                new W.Run(properties.CloneNode(true), new W.Text(" TAIL") { Space = SpaceProcessingModeValues.Preserve }));
        } else {
            mixed.Append(new W.Run(properties.CloneNode(true),
                new W.Text("LEAD ") { Space = SpaceProcessingModeValues.Preserve },
                new W.SeparatorMark(), new W.Text(" TAIL") { Space = SpaceProcessingModeValues.Preserve }));
        }
        Assert.Empty(document.ValidateDocument());
        string[] before = { main.Document.OuterXml, main.FootnotesPart.Footnotes.OuterXml,
            main.StyleDefinitionsPart!.Styles!.OuterXml, main.DocumentSettingsPart!.Settings!.OuterXml,
            string.Join("\n", main.FooterParts.Select(part => part.Footer!.OuterXml)) };
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        byte[] actual = result.ToBytes();
        Assert.Equal(before, new[] { main.Document.OuterXml, main.FootnotesPart.Footnotes.OuterXml,
            main.StyleDefinitionsPart.Styles.OuterXml, main.DocumentSettingsPart.Settings.OuterXml,
            string.Join("\n", main.FooterParts.Select(part => part.Footer!.OuterXml)) });
        Assert.Empty(document.ValidateDocument());
        var qualification = QualifyPageFootnoteNativeArial("equal-style rule runs", actual, result.Warnings);
        AssertPageFootnoteAcceptedLosses("equal-style rule runs", result, qualification.Qualified, qualification.Substitutions);
        using var pdf = PdfPigDocument.Open(actual);
        AssertPageFootnoteTokenOnce(pdf, "LEAD");
        AssertPageFootnoteTokenOnce(pdf, "SHORT NOTE BODY");
        if (!result.Warnings.Any(warning => warning.Code == SeparatorWarningCode)) {
            // This owned geometry source has only five rule visuals. Count
            // their actual ink even on portable fonts before native pinning.
            var bands = pdf.GetPage(1).Paths.Where(path => path.IsFilled &&
                    path.FillColor.ToRGBValues() == (0D, 0D, 0D))
                .Select(path => path.GetBoundingRectangle())
                .Where(bounds => bounds is { Width: > 20D, Height: > 0D and < 2D }).ToArray();
            Assert.Equal(5, bands.Length);
        }
        return new PageFootnoteArtifact(actual, reference, ReadPageFootnoteFrame(input),
            qualification.Qualified, qualification.Qualification);
    }
}
