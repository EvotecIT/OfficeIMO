using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("note", false)]
    [InlineData("note", true)]
    [InlineData("separator", false)]
    [InlineData("separator", true)]
    [InlineData("continuation", false)]
    [InlineData("continuation", true)]
    public void SaveAsPdf_NoteKeepWithNextUsesSharedNoteParagraphFlow(string story, bool inherited) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        source.AddParagraph("FIRSTREFERENCE").AddFootNote("FIRSTNOTEBODY");
        source.AddParagraph("SECONDREFERENCE").AddFootNote("SECONDNOTEBODY");
        source.AddParagraph("LASTBODY");
        if (story == "continuation") source.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2010;
        SetPageFootnoteAdmissionKeepNext(source, story, inherited, directOverride: null);
        Assert.Empty(source.ValidateDocument());

        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string bodyBefore = main.Document.OuterXml;
        string notesBefore = main.FootnotesPart!.Footnotes!.OuterXml;
        string stylesBefore = main.StyleDefinitionsPart!.Styles!.OuterXml;
        string settingsBefore = main.DocumentSettingsPart!.Settings!.OuterXml;
        var result = document.ConvertToPdfWithEvidence(FootnoteDiagnosticOptions());

        for (int serialization = 0; serialization < 2; serialization++) {
            using var pdf = PdfPigDocument.Open(result.ToBytes());
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            foreach (string token in new[] { "FIRSTREFERENCE", "SECONDREFERENCE", "LASTBODY", "FIRSTNOTEBODY", "SECONDNOTEBODY" })
                Assert.Single(System.Text.RegularExpressions.Regex.Matches(text, token));
            Assert.True(text.IndexOf("LASTBODY", StringComparison.Ordinal) < text.IndexOf("FIRSTNOTEBODY", StringComparison.Ordinal));
            Assert.True(text.IndexOf("FIRSTNOTEBODY", StringComparison.Ordinal) < text.IndexOf("SECONDNOTEBODY", StringComparison.Ordinal));
            Assert.DoesNotContain(result.Warnings, item => item.Code == "native-footnote-placement-approximated");
            var page = pdf.GetPages().Single(item => item.Text.Contains("FIRSTNOTEBODY", StringComparison.Ordinal));
            var letters = page.Letters.ToArray();
            Assert.True(FindRichNoteToken(letters, "FIRSTNOTEBODY")[0].StartBaseLine.Y <
                FindRichNoteToken(letters, "LASTBODY")[0].StartBaseLine.Y,
                "Supported keep-with-next must preserve the page-note area below the body.");
        }
        Assert.Equal(bodyBefore, main.Document.OuterXml);
        Assert.Equal(notesBefore, main.FootnotesPart.Footnotes.OuterXml);
        Assert.Equal(stylesBefore, main.StyleDefinitionsPart.Styles.OuterXml);
        Assert.Equal(settingsBefore, main.DocumentSettingsPart.Settings.OuterXml);
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData("note")]
    [InlineData("separator")]
    [InlineData("continuation")]
    public void SaveAsPdf_DirectFalseKeepWithNextOverridesInheritedNoteKeepChain(string story) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        source.AddParagraph("REFERENCE").AddFootNote("NOTEBODY");
        if (story == "continuation") source.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2010;
        SetPageFootnoteAdmissionKeepNext(source, story, inherited: true, directOverride: false);
        Assert.Empty(source.ValidateDocument());
        var main = source._wordprocessingDocument.MainDocumentPart!;
        string notesBefore = main.FootnotesPart!.Footnotes!.OuterXml;
        string settingsBefore = main.DocumentSettingsPart!.Settings!.OuterXml;
        var result = source.ConvertToPdfWithEvidence(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(text, "NOTEBODY"));
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "native-footnote-placement-approximated");
        Assert.Equal(notesBefore, main.FootnotesPart.Footnotes.OuterXml);
        Assert.Equal(settingsBefore, main.DocumentSettingsPart.Settings.OuterXml);
        Assert.Empty(source.ValidateDocument());
    }

    private static void SetPageFootnoteAdmissionKeepNext(WordDocument document, string story, bool inherited, bool? directOverride) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        W.Footnotes notes = main.FootnotesPart!.Footnotes!;
        W.Footnote note = story switch {
            "note" => notes.Elements<W.Footnote>().First(item => item.Id?.Value > 0),
            "separator" => notes.Elements<W.Footnote>().Single(item => item.Type?.Value == W.FootnoteEndnoteValues.Separator),
            "continuation" => notes.Elements<W.Footnote>().Single(item => item.Type?.Value == W.FootnoteEndnoteValues.ContinuationSeparator),
            _ => throw new ArgumentOutOfRangeException(nameof(story))
        };
        if (story != "note") {
            // Admission must inspect a loaded special story, rather than an
            // unregistered stored definition replaced by the implicit default.
            W.Settings settings = main.DocumentSettingsPart!.Settings!;
            W.FootnoteDocumentWideProperties? footnotes = settings.GetFirstChild<W.FootnoteDocumentWideProperties>();
            if (footnotes == null) {
                footnotes = new W.FootnoteDocumentWideProperties();
                settings.AddChild(footnotes, true);
            }
            long id = note.Id!.Value;
            if (!footnotes.Elements<W.FootnoteSpecialReference>().Any(reference => reference.Id?.Value == id))
                footnotes.Append(new W.FootnoteSpecialReference { Id = id });
        }
        W.ParagraphProperties properties = note.Elements<W.Paragraph>().First().ParagraphProperties ??= new W.ParagraphProperties();
        if (inherited) {
            main.StyleDefinitionsPart!.Styles!.Append(new W.Style(new W.StyleName { Val = "Note keep parent" },
                new W.StyleParagraphProperties(new W.KeepNext())) { Type = W.StyleValues.Paragraph, StyleId = "NoteKeepParent" },
                new W.Style(new W.StyleName { Val = "Note keep child" }, new W.BasedOn { Val = "NoteKeepParent" }) {
                    Type = W.StyleValues.Paragraph, StyleId = "NoteKeepChild"
                });
            properties.ParagraphStyleId = new W.ParagraphStyleId { Val = "NoteKeepChild" };
        } else properties.KeepNext = new W.KeepNext();
        if (directOverride.HasValue) properties.KeepNext = new W.KeepNext { Val = directOverride.Value };
    }
}
