using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_RichNotesRetainParagraphBoundariesRunFormattingAndLinks(bool endnote) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote);
        note.RemoveAllChildren();
        note.Append(new W.Paragraph(new W.ParagraphProperties(
                new W.SpacingBetweenLines { Before = "0", After = "120", Line = "280", LineRule = W.LineSpacingRuleValues.Exact },
                new W.Indentation { Left = "360", FirstLine = "180" }),
            NoteBodyMarker(endnote), RichNoteRun("BOLDWORD", new W.Bold()),
            RichNoteRun("ITALICWORD", new W.Italic()), RichNoteRun("COLORWORD", new W.Color { Val = "C03040" }),
            new W.Hyperlink(RichNoteRun("LINKWORD")) { Anchor = "NoteFormatAnchor" }),
            new W.Paragraph(new W.ParagraphProperties(new W.SpacingBetweenLines {
                Before = "0", After = "0", Line = "280", LineRule = W.LineSpacingRuleValues.Exact })),
            new W.Paragraph(new W.ParagraphProperties(
                new W.SpacingBetweenLines { Before = "0", After = "0", Line = "280", LineRule = W.LineSpacingRuleValues.Exact },
                new W.Justification { Val = W.JustificationValues.Right }),
                RichNoteRun("LASTWORD")));
        Assert.Empty(source.ValidateDocument());
        string original = note.OuterXml;
        var result = source.ConvertToPdfWithEvidence(FootnoteDiagnosticOptions());
        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            var letters = pdf.GetPages().SelectMany(page => page.Letters).ToArray();
            Letter[] bold = FindRichNoteToken(letters, "BOLDWORD");
            Letter[] italic = FindRichNoteToken(letters, "ITALICWORD");
            Letter[] color = FindRichNoteToken(letters, "COLORWORD");
            Letter[] last = FindRichNoteToken(letters, "LASTWORD");
            Assert.All(bold, letter => Assert.Contains("Bold", letter.FontName, StringComparison.OrdinalIgnoreCase));
            Assert.All(italic, letter => Assert.True(letter.FontName.IndexOf("Italic", StringComparison.OrdinalIgnoreCase) >= 0 ||
                letter.FontName.IndexOf("Oblique", StringComparison.OrdinalIgnoreCase) >= 0, letter.FontName));
            Assert.All(color, letter => {
                var rgb = letter.Color.ToRGBValues();
                Assert.Equal(192D / 255D, rgb.Item1, 3);
                Assert.Equal(48D / 255D, rgb.Item2, 3);
                Assert.Equal(64D / 255D, rgb.Item3, 3);
            });
            Assert.True(bold[0].StartBaseLine.Y - last[0].StartBaseLine.Y >= 26D,
                "A real empty note paragraph must remain between the formatted first and final paragraphs.");
            Assert.True(last[0].StartBaseLine.X > bold[0].StartBaseLine.X + 100D,
                "The final paragraph's right alignment must survive note projection.");
            Letter[] linked = FindRichNoteToken(letters, "LINKWORD");
            // PdfPig GetHyperlinks exposes external links only; the authored
            // anchor is a named GoTo annotation, read by the shared inspector.
            PdfDocumentInfo navigation = PdfInspector.Inspect(bytes);
            PdfLinkAnnotation annotation = Assert.Single(navigation.LinkAnnotations,
                link => link.DestinationName == "NoteFormatAnchor");
            Assert.True(annotation.IsNamedDestinationLink);
            Assert.True(annotation.Width > 0D && annotation.Height > 0D);
            Assert.True(annotation.X1 <= linked.Min(letter => letter.StartBaseLine.X) + .01D);
            Assert.True(annotation.X2 >= linked.Max(letter => letter.EndBaseLine.X) - .01D);
            Assert.All(linked, letter => Assert.InRange(letter.StartBaseLine.Y, annotation.Y1 - .01D, annotation.Y2 + .01D));
            int linkPage = pdf.GetPages().Select((page, index) => (page, number: index + 1))
                .Single(item => item.page.Text.Contains("LINKWORD", StringComparison.Ordinal)).number;
            Assert.Equal(linkPage, annotation.PageNumber);
            PdfNamedDestination destination = Assert.Single(navigation.NamedDestinations,
                item => item.Name == "NoteFormatAnchor");
            int targetPage = pdf.GetPages().Select((page, index) => (page, number: index + 1))
                .Single(item => item.page.Text.Contains("LINKTARGET", StringComparison.Ordinal)).number;
            Assert.Equal(targetPage, destination.PageNumber);
        }
        Assert.Equal(original, note.OuterXml);
        Assert.Empty(source.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_NoteHiddenParagraphMarkJoinsOnlyItsImmediateSuccessor(bool endnote) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote);
        note.RemoveAllChildren();
        note.Append(new W.Paragraph(new W.ParagraphProperties(new W.ParagraphMarkRunProperties(new W.Vanish())),
                NoteBodyMarker(endnote), RichNoteRun("JOINLEFT")),
            new W.Paragraph(RichNoteRun("JOINRIGHT")), new W.Paragraph(RichNoteRun("AFTERJOIN")));
        Assert.Empty(source.ValidateDocument());
        string original = note.OuterXml;
        using var pdf = PdfPigDocument.Open(source.ToPdf(FootnoteDiagnosticOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPages().SelectMany(page => page.Letters).ToArray();
        Letter[] left = FindRichNoteToken(letters, "JOINLEFT");
        Letter[] right = FindRichNoteToken(letters, "JOINRIGHT");
        Letter[] after = FindRichNoteToken(letters, "AFTERJOIN");
        Assert.Equal(left[0].StartBaseLine.Y, right[0].StartBaseLine.Y, 3);
        Assert.True(right[0].StartBaseLine.X > left[left.Length - 1].EndBaseLine.X - .01D);
        Assert.True(after[0].StartBaseLine.Y < right[0].StartBaseLine.Y - 5D,
            "The next visible paragraph boundary must remain after the joined pair.");
        Assert.Equal(original, note.OuterXml);
    }

    [Theory]
    [InlineData("shade", false)]
    [InlineData("shade", true)]
    [InlineData("border", false)]
    [InlineData("border", true)]
    public void SaveAsPdf_FootnoteParagraphDecorationRetainsPaintAndPagePlacement(string decoration, bool inherited) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote: false);
        W.Paragraph paragraph = note.Elements<W.Paragraph>().Single();
        paragraph.RemoveAllChildren();
        OpenXmlElement paint = decoration == "shade"
            ? new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "DDEEFF" }
            : new W.ParagraphBorders(new W.TopBorder { Val = W.BorderValues.Single, Color = "006080", Size = 12U, Space = 4U },
                new W.LeftBorder { Val = W.BorderValues.Single, Color = "006080", Size = 12U, Space = 4U },
                new W.BottomBorder { Val = W.BorderValues.Single, Color = "006080", Size = 12U, Space = 4U },
                new W.RightBorder { Val = W.BorderValues.Single, Color = "006080", Size = 12U, Space = 4U });
        var properties = new W.ParagraphProperties();
        if (inherited) {
            source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
                new W.Style(new W.StyleName { Val = "Decorated note parent" }, new W.StyleParagraphProperties(paint)) {
                    Type = W.StyleValues.Paragraph, StyleId = "DecoratedNoteParent" },
                new W.Style(new W.StyleName { Val = "Decorated note child" }, new W.BasedOn { Val = "DecoratedNoteParent" }) {
                    Type = W.StyleValues.Paragraph, StyleId = "DecoratedNoteChild" });
            properties.ParagraphStyleId = new W.ParagraphStyleId { Val = "DecoratedNoteChild" };
        } else properties.Append(paint);
        paragraph.Append(properties, NoteBodyMarker(endnote: false), RichNoteRun("DECORATEDNOTE"));
        Assert.Empty(source.ValidateDocument());
        string original = note.OuterXml;
        var result = source.ConvertToPdfWithEvidence(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var page = pdf.GetPages().Single(item => item.Text.Contains("DECORATEDNOTE", StringComparison.Ordinal));
        Letter[] content = FindRichNoteToken(page.Letters.ToArray(), "DECORATEDNOTE");
        Assert.True(content[0].StartBaseLine.Y < page.Height / 2D,
            "A supported paragraph decoration must keep the note in the page-note area.");
        if (decoration == "shade") Assert.Contains(page.Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds && bounds.Width > 50D && bounds.Height > 8D &&
            bounds.Bottom < content[0].StartBaseLine.Y && bounds.Top > content[0].StartBaseLine.Y);
        else Assert.Contains(page.Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds && bounds.Height > 8D &&
            bounds.Bottom < content[0].StartBaseLine.Y && bounds.Top > content[0].StartBaseLine.Y);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "native-footnote-placement-approximated");
        Assert.Equal(original, note.OuterXml);
    }

    [Theory]
    [InlineData(false, "bare")]
    [InlineData(false, "style")]
    [InlineData(false, "normal")]
    [InlineData(false, "subscript")]
    [InlineData(true, "bare")]
    [InlineData(true, "style")]
    [InlineData(true, "normal")]
    [InlineData(true, "subscript")]
    public void SaveAsPdf_NoteReferenceUsesSourceBaselineRatherThanImplicitSuperscript(bool endnote, string profile) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        WordParagraph body = source.AddParagraph("BEFORE");
        if (endnote) body.AddEndNote("NOTEBODY"); else body.AddFootNote("NOTEBODY");
        body.AddText("AFTER");
        W.Run marker = body._paragraph.Elements<W.Run>().Single(run => run.Elements().Any(child =>
            child is W.FootnoteReference or W.EndnoteReference));
        marker.RunProperties = profile == "bare" ? null : new W.RunProperties(
            new W.RunStyle { Val = endnote ? "EndnoteReference" : "FootnoteReference" });
        if (profile is "normal" or "subscript") marker.RunProperties!.Append(new W.VerticalTextAlignment {
            Val = profile == "normal" ? W.VerticalPositionValues.Baseline : W.VerticalPositionValues.Subscript });
        Assert.Empty(source.ValidateDocument());
        string original = body._paragraph.OuterXml;
        using var pdf = PdfPigDocument.Open(source.ToPdf(FootnoteDiagnosticOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Letter[] letters = pdf.GetPage(1).Letters.ToArray();
        Letter[] before = FindRichNoteToken(letters, "BEFORE");
        string text = string.Concat(letters.Select(letter => letter.Value));
        int start = text.IndexOf("BEFORE", StringComparison.Ordinal) + "BEFORE".Length;
        Letter reference = letters[start];
        if (profile is "bare" or "normal") {
            Assert.Equal(before[0].PointSize, reference.PointSize, 3);
            Assert.Equal(before[0].StartBaseLine.Y, reference.StartBaseLine.Y, 3);
        } else {
            Assert.True(reference.PointSize < before[0].PointSize);
            if (profile == "style") Assert.True(reference.StartBaseLine.Y > before[0].StartBaseLine.Y);
            else Assert.True(reference.StartBaseLine.Y < before[0].StartBaseLine.Y);
        }
        Assert.Equal(original, body._paragraph.OuterXml);
    }

    private static OpenXmlCompositeElement AddRichNoteFixture(WordDocument source, bool endnote) {
        WordParagraph body = source.AddParagraph("BODYREFERENCE");
        if (endnote) body.AddEndNote("INITIALNOTE"); else body.AddFootNote("INITIALNOTE");
        source.AddParagraph("LINKTARGET").AddBookmark("NoteFormatAnchor");
        var main = source._wordprocessingDocument.MainDocumentPart!;
        return endnote
            ? main.EndnotesPart!.Endnotes!.Elements<W.Endnote>().Single(note => note.Id?.Value > 0)
            : main.FootnotesPart!.Footnotes!.Elements<W.Footnote>().Single(note => note.Id?.Value > 0);
    }

    private static W.Run NoteBodyMarker(bool endnote) => new(new W.RunProperties(new W.RunStyle {
        Val = endnote ? "EndnoteReference" : "FootnoteReference" }),
        endnote ? new W.EndnoteReferenceMark() : new W.FootnoteReferenceMark());

    private static W.Run RichNoteRun(string text, params OpenXmlElement[] properties) =>
        new(new W.RunProperties(properties), new W.Text(text));

    private static Letter[] FindRichNoteToken(IReadOnlyList<Letter> letters, string token) {
        string text = string.Concat(letters.Select(letter => letter.Value));
        int index = text.IndexOf(token, StringComparison.Ordinal);
        Assert.True(index >= 0, $"Missing {token} in {text}");
        Assert.Equal(index, text.LastIndexOf(token, StringComparison.Ordinal));
        return letters.Skip(index).Take(token.Length).ToArray();
    }
}
