using System.Globalization;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordPdfPageFootnoteTests {
    private static readonly IReadOnlyDictionary<string, (string Input, string Reference)> ImplicitFootnoteFixtureHashes =
        new Dictionary<string, (string Input, string Reference)>(StringComparer.Ordinal) {
            ["implicit-base"] = ("d6dab1ace0debe528f11f9c77c298aa42eb42d60b1c0cb793a5aec09eb45fe29", "5e743ba6509aedabe0e522d2edb7e387366d1a3d5974e4dc67a87d03723851d0"),
            ["implicit-footnote-style20-line32"] = ("21a9909cae57254f91e878074fe5842d3bb53f2408f45a5096d0d64e3377c8d1", "4825248165086d8b6f53641672053350ca6779aa9d8e57ff2a78ba608749fd48"),
            ["implicit-normal-font18"] = ("2b1cfeed99b0febc2c6f934648f381df72f6b969affaa5fc58a7f3596a4a4cdc", "38799ec6192dd2f41baab6a83f94c0712e3c4cfc2aa7f0da22e4b1b57d786054"),
            ["implicit-explicit-equivalent"] = ("d913290a16c8b4a2cb51ec87933131af4a21ddffda0cb762b94d96c8386fea8a", "5ea9886b695898648d6c5a4cc3e6a0ecf2baedc4b8213844861a72b8b09f9671")
        };

    private static readonly string[] ImplicitFootnoteParagraphs = Enumerable.Range(1, 55)
        .Select(number => $"NOTE FIXED {number:000}").ToArray();
    private static readonly string[] ImplicitFootnoteBody = {
        "BODY FIXED FIRST", "BODY FIXED REFERENCE", "FOLLOWING BODY MARKER"
    };

    [Theory]
    [InlineData("implicit-base")]
    [InlineData("implicit-footnote-style20-line32")]
    [InlineData("implicit-normal-font18")]
    [InlineData("implicit-explicit-equivalent")]
    public void SaveAsPdf_ImplicitFootnoteSeparatorsPreserveQualifiedPaginationAndSource(string id) {
        PageFootnoteArtifact bytes = ConvertImplicitFootnoteFixture(id);
        using var actual = PdfPigDocument.Open(bytes.Actual);
        using var reference = PdfPigDocument.Open(bytes.Reference);

        AssertImplicitFootnoteFlow(actual, bytes.Frame);
        AssertImplicitFootnoteContent(reference);
        AssertImplicitFootnoteNativePartition(reference);
        foreach (string token in ImplicitFootnoteParagraphs.Concat(ImplicitFootnoteBody)) {
            Page page = actual.GetPage(PageFootnoteTokenPage(actual, token));
            Assert.All(PageFootnoteToken(page, token), letter => Assert.Equal(12D, letter.PointSize, 3));
        }
        AssertPageFootnoteNativeProfile(bytes, () => {
            AssertImplicitFootnoteNativePartition(actual);
            AssertImplicitFootnoteNativeRuleBands(actual, reference, bytes.Frame);
        });
    }

    [Fact]
    public void SaveAsPdf_ImplicitAndExplicitSeparatorsKeepFixedStoryGlyphGeometryAcrossStyles() {
        PageFootnoteArtifact baseline = ConvertImplicitFootnoteFixture("implicit-base");
        using var actualBaseline = PdfPigDocument.Open(baseline.Actual);
        using var nativeBaseline = PdfPigDocument.Open(baseline.Reference);
        foreach (string id in ImplicitFootnoteFixtureHashes.Keys.Where(id => id != "implicit-base")) {
            PageFootnoteArtifact candidate = ConvertImplicitFootnoteFixture(id);
            using var actual = PdfPigDocument.Open(candidate.Actual);
            using var native = PdfPigDocument.Open(candidate.Reference);
            // Native controls establish that these style changes leave fixed
            // body/note geometry alone. Compare each producer with itself;
            // this cohort adds no native-versus-converter coordinate budget.
            AssertImplicitFootnoteGlyphGeometry(nativeBaseline, native);
            AssertImplicitFootnoteGlyphGeometry(actualBaseline, actual);
        }
    }

    private PageFootnoteArtifact ConvertImplicitFootnoteFixture(string id) {
        string inputPath = PageFootnoteFixturePath(id + ".docx");
        byte[] input = File.ReadAllBytes(inputPath);
        byte[] reference = File.ReadAllBytes(PageFootnoteFixturePath(id + ".microsoft-word.pdf"));
        Assert.Equal(ImplicitFootnoteFixtureHashes[id].Input, PageFootnoteSha256(input));
        Assert.Equal(ImplicitFootnoteFixtureHashes[id].Reference, PageFootnoteSha256(reference));

        using WordDocument document = WordDocument.Load(new MemoryStream(input));
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        // These simple fixtures have no table borders that could be mistaken
        // for the native default separator's thin black filled band.
        Assert.Empty(main.Document.Descendants<W.Table>());
        Assert.Empty(main.FootnotesPart!.Footnotes!.Descendants<W.Table>());
        string bodyBefore = main.Document.OuterXml;
        string notesBefore = main.FootnotesPart!.Footnotes!.OuterXml;
        string stylesBefore = main.StyleDefinitionsPart!.Styles!.OuterXml;
        string settingsBefore = main.DocumentSettingsPart!.Settings!.OuterXml;
        W.SectionProperties section = Assert.Single(main.Document.Body!.Elements<W.SectionProperties>());
        W.PageMargin margin = section.GetFirstChild<W.PageMargin>()!;
        W.Footnote note = Assert.Single(main.FootnotesPart.Footnotes.Elements<W.Footnote>(), item => item.Id?.Value > 0);
        W.SpacingBetweenLines spacing = note.Elements<W.Paragraph>().First().ParagraphProperties!.SpacingBetweenLines!;
        Assert.Equal(W.LineSpacingRuleValues.Exact, spacing.LineRule!.Value);
        bool hasStoredSeparatorMark = main.FootnotesPart.Footnotes.Descendants().Any(element =>
            element.LocalName is "separator" or "continuationSeparator");
        var frame = new PageFootnoteFrame(margin.Top!.Value / 20D, margin.Bottom!.Value / 20D,
            double.Parse(spacing.Line!.Value!, CultureInfo.InvariantCulture) / 20D, hasStoredSeparatorMark);

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        byte[] actual = result.ToBytes();
        Assert.Equal("%PDF", Encoding.ASCII.GetString(actual, 0, 4));
        Assert.Equal(input, File.ReadAllBytes(inputPath));
        Assert.Equal(bodyBefore, main.Document.OuterXml);
        Assert.Equal(notesBefore, main.FootnotesPart.Footnotes.OuterXml);
        Assert.Equal(stylesBefore, main.StyleDefinitionsPart.Styles.OuterXml);
        Assert.Equal(settingsBefore, main.DocumentSettingsPart.Settings.OuterXml);
        Assert.Empty(document.ValidateDocument());

        var qualification = QualifyPageFootnoteNativeArial(id, actual, result.Warnings);
        AssertPageFootnoteAcceptedLosses(id, result, qualification.Qualified, qualification.Substitutions);
        return new PageFootnoteArtifact(actual, reference, frame, qualification.Qualified, qualification.Qualification);
    }

    private static void AssertImplicitFootnoteNativeRuleBands(PdfPigDocument actual, PdfPigDocument reference,
        PageFootnoteFrame frame) {
        foreach (int pageNumber in ImplicitFootnoteParagraphs.Select(token => PageFootnoteTokenPage(reference, token))
            .Distinct().OrderBy(number => number)) {
            string firstNote = ImplicitFootnoteParagraphs.First(token => PageFootnoteTokenPage(reference, token) == pageNumber);
            var expected = ImplicitFootnoteRuleBand(reference.GetPage(pageNumber), firstNote, frame);
            var observed = ImplicitFootnoteRuleBand(actual.GetPage(pageNumber), firstNote, frame);
            // A 600-dpi native device pixel is .12pt. Keep one-pixel print
            // precision plus .03pt for exported coordinates, independently of
            // the existing text geometry and pagination budgets.
            Assert.InRange(Math.Abs(observed.Left - expected.Left), 0D, .15D);
            Assert.InRange(Math.Abs(observed.Top - expected.Top), 0D, .15D);
            Assert.InRange(Math.Abs(observed.Width - expected.Width), 0D, .15D);
            Assert.InRange(Math.Abs(observed.Height - expected.Height), 0D, .15D);
        }
    }

    private static UglyToad.PdfPig.Core.PdfRectangle ImplicitFootnoteRuleBand(Page page, string firstNote,
        PageFootnoteFrame frame) {
        double noteInkTop = PageFootnoteToken(page, firstNote).Max(letter => letter.BoundingBox.Top);
        // Select each producer's band independently, above its first note and
        // within the authored story frame. Do not choose the nearest path or
        // filter by the eventual native-coordinate comparison budget.
        return Assert.Single(page.Paths.Where(path => path.IsFilled &&
                path.FillColor.ToRGBValues() == (0D, 0D, 0D))
            .Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds.HasValue && Math.Abs(bounds.Value.Width - 144D) <= .15D &&
                bounds.Value.Height > 0D && bounds.Value.Height < 2D &&
                bounds.Value.Bottom > noteInkTop && bounds.Value.Top <= (double)page.Height - frame.MarginTop)
            .Select(bounds => bounds!.Value));
    }

    private static void AssertImplicitFootnoteContent(PdfPigDocument pdf) {
        string compact = string.Concat(pdf.GetPages().Select(page => PageFootnoteCompact(page.Text)));
        string[] observed = Regex.Matches(compact, @"NOTEFIXED\d{3}").Cast<Match>().Select(match => match.Value).ToArray();
        Assert.Equal(ImplicitFootnoteParagraphs.Select(PageFootnoteCompact).ToArray(), observed);
        foreach (string token in ImplicitFootnoteBody) AssertPageFootnoteTokenOnce(pdf, token);
    }

    private static void AssertImplicitFootnoteFlow(PdfPigDocument pdf, PageFootnoteFrame frame) {
        AssertImplicitFootnoteContent(pdf);
        int priorPage = 0;
        double priorBaseline = double.PositiveInfinity;
        foreach (string token in ImplicitFootnoteParagraphs) {
            int pageNumber = PageFootnoteTokenPage(pdf, token);
            Page page = pdf.GetPage(pageNumber);
            double baseline = PageFootnoteToken(page, token)[0].StartBaseLine.Y;
            Assert.InRange(baseline, frame.MarginBottom, (double)page.Height - frame.MarginTop);
            Assert.True(pageNumber >= priorPage);
            if (pageNumber == priorPage) Assert.Equal(frame.NoteLineHeight, priorBaseline - baseline, 3);
            priorPage = pageNumber;
            priorBaseline = baseline;
        }
        for (int page = 1; page <= pdf.NumberOfPages; page++) {
            string lastNote = ImplicitFootnoteParagraphs.Last(token => PageFootnoteTokenPage(pdf, token) == page);
            Assert.InRange(PageFootnoteToken(pdf.GetPage(page), lastNote)[0].StartBaseLine.Y,
                frame.MarginBottom, frame.MarginBottom + frame.NoteLineHeight);
        }
        Assert.Equal(1, PageFootnoteTokenPage(pdf, "BODY FIXED FIRST"));
        Assert.Equal(1, PageFootnoteTokenPage(pdf, "BODY FIXED REFERENCE"));
        Assert.Equal(PageFootnoteTokenPage(pdf, ImplicitFootnoteParagraphs.Last()),
            PageFootnoteTokenPage(pdf, "FOLLOWING BODY MARKER"));
        AssertPageFootnoteBodyDoesNotOverlapNotes(pdf, ImplicitFootnoteBody, ImplicitFootnoteParagraphs);
    }

    private static void AssertImplicitFootnoteNativePartition(PdfPigDocument pdf) {
        Assert.Equal(3, pdf.NumberOfPages);
        int[] expectedCounts = { 25, 27, 3 };
        for (int page = 1; page <= pdf.NumberOfPages; page++)
            Assert.Equal(expectedCounts[page - 1], Regex.Matches(PageFootnoteCompact(pdf.GetPage(page).Text), @"NOTEFIXED\d{3}").Count);
        Assert.Equal(3, PageFootnoteTokenPage(pdf, "FOLLOWING BODY MARKER"));
    }

    private static void AssertImplicitFootnoteGlyphGeometry(PdfPigDocument expected, PdfPigDocument actual) {
        Assert.Equal(expected.NumberOfPages, actual.NumberOfPages);
        for (int page = 1; page <= expected.NumberOfPages; page++) {
            Assert.Equal(expected.GetPage(page).Width, actual.GetPage(page).Width);
            Assert.Equal(expected.GetPage(page).Height, actual.GetPage(page).Height);
        }
        foreach (string token in ImplicitFootnoteParagraphs.Concat(ImplicitFootnoteBody)) {
            int page = PageFootnoteTokenPage(expected, token);
            Assert.Equal(page, PageFootnoteTokenPage(actual, token));
            Letter[] expectedLetters = PageFootnoteToken(expected.GetPage(page), token);
            Letter[] actualLetters = PageFootnoteToken(actual.GetPage(page), token);
            Assert.Equal(expectedLetters.Select(ImplicitFootnoteGlyphGeometry).ToArray(),
                actualLetters.Select(ImplicitFootnoteGlyphGeometry).ToArray());
        }
    }

    private static ImplicitFootnoteGlyph ImplicitFootnoteGlyphGeometry(Letter letter) {
        var color = letter.Color.ToRGBValues();
        return new ImplicitFootnoteGlyph(letter.Value, Regex.Replace(letter.FontName, @"^[A-Z]{6}\+", string.Empty),
            letter.PointSize, letter.StartBaseLine.X, letter.StartBaseLine.Y, letter.EndBaseLine.X, letter.EndBaseLine.Y,
            letter.BoundingBox.Left, letter.BoundingBox.Bottom, letter.BoundingBox.Right, letter.BoundingBox.Top,
            color.Item1, color.Item2, color.Item3);
    }

    private sealed record ImplicitFootnoteGlyph(string Value, string Font, double Size,
        double StartX, double StartY, double EndX, double EndY, double Left, double Bottom, double Right, double Top,
        double Red, double Green, double Blue);
}
