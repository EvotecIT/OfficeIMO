using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordPdfPageFootnoteTests {
    [Theory]
    [InlineData("15", true)]
    [InlineData("14", false)]
    [InlineData("12", false)]
    [InlineData(null, false)]
    public void SaveAsPdf_FootnoteContinuationStoryFollowsCompatibilityMode(string? mode, bool ordinaryOnContinuation) {
        using WordDocument document = LoadFootnoteStorySelectionFixture();
        SetFootnoteStorySelectionMode(document, mode);
        var main = document._wordprocessingDocument.MainDocumentPart!;
        W.Footnote ordinary = main.FootnotesPart!.Footnotes!.Elements<W.Footnote>()
            .Single(note => note.Type?.Value == W.FootnoteEndnoteValues.Separator);
        W.Footnote continuation = main.FootnotesPart.Footnotes.Elements<W.Footnote>()
            .Single(note => note.Type?.Value == W.FootnoteEndnoteValues.ContinuationSeparator);
        // Registered identity, rather than conventional -1/0 IDs or part order,
        // selects the two stories. Keep their native-qualified content intact.
        ordinary.Id = -7;
        continuation.Id = -4;
        var properties = main.DocumentSettingsPart!.Settings!.GetFirstChild<W.FootnoteDocumentWideProperties>()!;
        properties.RemoveAllChildren<W.FootnoteSpecialReference>();
        properties.Append(new W.FootnoteSpecialReference { Id = -4 }, new W.FootnoteSpecialReference { Id = -7 });
        continuation.Remove();
        main.FootnotesPart.Footnotes.PrependChild(continuation);
        const string unusedFamily = "Unused Modern Note Face";
        if (ordinaryOnContinuation) {
            // A registered continuation story is still unused by modern layout.
            // It cannot force an unsupported-profile fallback or a font warning.
            W.Paragraph paragraph = continuation.Elements<W.Paragraph>().Single();
            paragraph.ParagraphProperties!.KeepNext = new W.KeepNext();
            foreach (W.Run run in paragraph.Elements<W.Run>()) {
                run.RunProperties ??= new W.RunProperties();
                run.RunProperties.RunFonts = new W.RunFonts { Ascii = unusedFamily, HighAnsi = unusedFamily };
            }
        }

        using var pdf = PdfPigDocument.Open(ConvertFootnoteStorySelection(document,
            ordinaryOnContinuation ? unusedFamily : null));
        Assert.True(pdf.NumberOfPages > 1);
        for (int page = 1; page <= pdf.NumberOfPages; page++) {
            bool ordinaryExpected = page == 1 || ordinaryOnContinuation;
            string expected = ordinaryExpected ? "START SEPARATOR" : "CONTINUATION SEPARATOR";
            string excluded = ordinaryExpected ? "CONTINUATION SEPARATOR" : "START SEPARATOR";
            PageFootnoteToken(pdf.GetPage(page), expected);
            Assert.DoesNotContain(PageFootnoteCompact(excluded), PageFootnoteCompact(pdf.GetPage(page).Text), StringComparison.Ordinal);
        }
        AssertFootnoteStorySelectionInventory(pdf);
    }

    [Theory]
    [InlineData("15", true)]
    [InlineData("14", false)]
    public void SaveAsPdf_RegisteredEmptyFootnoteStoryDoesNotRecreateAnImplicitRule(string mode, bool emptyOrdinary) {
        using WordDocument document = LoadFootnoteStorySelectionFixture();
        SetFootnoteStorySelectionMode(document, mode);
        W.Footnotes notes = document._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!;
        W.Footnote story = notes.Elements<W.Footnote>().Single(note => note.Type?.Value ==
            (emptyOrdinary ? W.FootnoteEndnoteValues.Separator : W.FootnoteEndnoteValues.ContinuationSeparator));
        story.RemoveAllChildren();
        story.Append(new W.Paragraph(new W.ParagraphProperties(new W.SpacingBetweenLines {
            Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto
        })));

        using var pdf = PdfPigDocument.Open(ConvertFootnoteStorySelection(document));
        Assert.True(pdf.NumberOfPages > 1);
        for (int page = 1; page <= pdf.NumberOfPages; page++) {
            string text = PageFootnoteCompact(pdf.GetPage(page).Text);
            Assert.DoesNotContain("CONTINUATIONSEPARATOR", text, StringComparison.Ordinal);
            if (!emptyOrdinary && page == 1) PageFootnoteToken(pdf.GetPage(page), "START SEPARATOR");
            else {
                Assert.DoesNotContain("STARTSEPARATOR", text, StringComparison.Ordinal);
                // This fixture has only text outside its special rules. A loaded
                // empty story must not be replaced by a default painted band.
                Assert.DoesNotContain(pdf.GetPage(page).Paths, path => path.IsFilled || path.IsStroked);
            }
        }
        AssertFootnoteStorySelectionInventory(pdf);
    }

    [Fact]
    public void SaveAsPdf_UnregisteredFootnoteStoriesDoNotSupplyContentFontsOrAdmissionRestrictions() {
        using WordDocument document = LoadFootnoteStorySelectionFixture();
        SetFootnoteStorySelectionMode(document, "14");
        var main = document._wordprocessingDocument.MainDocumentPart!;
        main.DocumentSettingsPart!.Settings!.GetFirstChild<W.FootnoteDocumentWideProperties>()!
            .RemoveAllChildren<W.FootnoteSpecialReference>();
        const string unusedFamily = "Unused Footnote Story Face";
        foreach (W.Footnote story in main.FootnotesPart!.Footnotes!.Elements<W.Footnote>().Where(note =>
            note.Type?.Value == W.FootnoteEndnoteValues.Separator ||
            note.Type?.Value == W.FootnoteEndnoteValues.ContinuationSeparator)) {
            W.Paragraph paragraph = story.Elements<W.Paragraph>().Single();
            paragraph.ParagraphProperties!.KeepNext = new W.KeepNext();
            foreach (W.Run run in paragraph.Elements<W.Run>()) {
                run.RunProperties ??= new W.RunProperties();
                run.RunProperties.RunFonts = new W.RunFonts { Ascii = unusedFamily, HighAnsi = unusedFamily };
            }
        }

        using var pdf = PdfPigDocument.Open(ConvertFootnoteStorySelection(document, unusedFamily));
        Assert.True(pdf.NumberOfPages > 1);
        string text = PageFootnoteCompact(string.Concat(pdf.GetPages().Select(page => page.Text)));
        Assert.DoesNotContain("STARTSEPARATOR", text, StringComparison.Ordinal);
        Assert.DoesNotContain("CONTINUATIONSEPARATOR", text, StringComparison.Ordinal);
        AssertFootnoteStorySelectionInventory(pdf);
    }

    private static WordDocument LoadFootnoteStorySelectionFixture() {
        byte[] input = File.ReadAllBytes(PageFootnoteFixturePath("richnote-long-rich-continuation.docx"));
        Assert.Equal(PageFootnoteFixtureHashes["richnote-long-rich-continuation"].Input, PageFootnoteSha256(input));
        using var stream = new MemoryStream(input);
        return WordDocument.Load(stream);
    }

    private static void SetFootnoteStorySelectionMode(WordDocument document, string? mode) {
        W.Settings settings = document._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!;
        W.Compatibility compatibility = settings.GetFirstChild<W.Compatibility>()!;
        foreach (W.CompatibilitySetting setting in compatibility.Elements<W.CompatibilitySetting>()
            .Where(setting => setting.Name?.Value == W.CompatSettingNameValues.CompatibilityMode).ToList()) setting.Remove();
        if (mode != null) compatibility.Append(new W.CompatibilitySetting {
            Name = W.CompatSettingNameValues.CompatibilityMode,
            Uri = "http://schemas.microsoft.com/office/word", Val = mode
        });
    }

    private static byte[] ConvertFootnoteStorySelection(WordDocument document, string? unusedFamily = null) {
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string[] before = { main.Document.OuterXml, main.FootnotesPart!.Footnotes!.OuterXml,
            main.DocumentSettingsPart!.Settings!.OuterXml, main.StyleDefinitionsPart!.Styles!.OuterXml };
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        byte[] actual = result.ToBytes();
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeFootnotePlacementApproximated");
        if (unusedFamily != null) Assert.DoesNotContain(result.Warnings,
            warning => warning.Details.Values.Any(value => value.IndexOf(unusedFamily, StringComparison.Ordinal) >= 0));
        Assert.Equal(before, new[] { main.Document.OuterXml, main.FootnotesPart.Footnotes.OuterXml,
            main.DocumentSettingsPart.Settings.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml });
        Assert.Empty(document.ValidateDocument());
        Assert.Equal(PageFootnoteFixtureHashes["richnote-long-rich-continuation"].Input,
            PageFootnoteSha256(File.ReadAllBytes(PageFootnoteFixturePath("richnote-long-rich-continuation.docx"))));
        return actual;
    }

    private static void AssertFootnoteStorySelectionInventory(PdfPigDocument pdf) {
        for (int number = 1; number <= 70; number++) AssertPageFootnoteTokenOnce(pdf, $"NOTE LONG {number:000}");
        int finalNotePage = PageFootnoteTokenPage(pdf, "NOTE LONG 070");
        for (int number = 1; number <= 6; number++) {
            string token = $"AFTER LONG {number:00}";
            AssertPageFootnoteTokenOnce(pdf, token);
            Assert.Equal(finalNotePage, PageFootnoteTokenPage(pdf, token));
        }
    }
}
