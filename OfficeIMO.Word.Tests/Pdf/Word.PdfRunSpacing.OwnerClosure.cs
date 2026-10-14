using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "reset", false)]
    [InlineData(true, "reset", true)]
    [InlineData(false, "join", false)]
    [InlineData(true, "join", true)]
    [InlineData(false, "join-reset", false)]
    [InlineData(true, "join-reset", true)]
    [InlineData(false, "textbox", false)]
    [InlineData(true, "textbox", true)]
    [InlineData(false, "textbox-reset", false)]
    [InlineData(true, "textbox-reset", true)]
    [InlineData(false, "nested-textbox", false)]
    [InlineData(true, "nested-textbox", true)]
    [InlineData(false, "nested-textbox-prefix", false)]
    [InlineData(true, "nested-textbox-prefix", true)]
    [InlineData(false, "nested-textbox-surrounded", false)]
    [InlineData(true, "nested-textbox-surrounded", true)]
    public void HeaderFooterSpacingSurvivesAlternateStoryPaths(bool footer, string route, bool characterStyle) {
        using WordDocument document = CreateJoinedParagraphDocument();
        document.AddParagraph("BODY");
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new W.DocDefaults(new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(
            new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "24" })));
        bool reset = route.Contains("reset", StringComparison.Ordinal);
        int scale = reset ? 100 : 200;
        int tracking = reset ? 0 : 20;
        styles.Append(new W.Style(new W.StyleRunProperties(new W.Spacing { Val = 20 }, new W.CharacterScale { Val = 200L })) {
            StyleId = "ExpandedParagraph", Type = W.StyleValues.Paragraph
        });
        styles.Append(new W.Style(new W.StyleRunProperties(new W.Spacing { Val = tracking }, new W.CharacterScale { Val = scale })) {
            StyleId = "VisibleSpacing", Type = W.StyleValues.Character
        });
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordParagraph outer = story.AddParagraph();
        WordParagraph first = outer;
        if (route.Contains("textbox", StringComparison.Ordinal)) {
            WordTextBox box = outer.AddTextBox("MMMMX", WordImageTextWrapping.Square);
            first = box.Paragraphs.Single();
            if (route.StartsWith("nested-textbox", StringComparison.Ordinal)) {
                first.Text = route is "nested-textbox-prefix" or "nested-textbox-surrounded" ? "AB" : string.Empty;
                ApplyClosureSpacing(first, characterStyle, scale, tracking);
                WordParagraph envelope = first;
                first = envelope.AddTextBox("MMMMX", WordImageTextWrapping.Square).Paragraphs.Single();
                if (route == "nested-textbox-surrounded") ApplyClosureSpacing(envelope.AddText("CD"), characterStyle, scale, tracking);
            }
            if (reset) outer.SetStyleId("ExpandedParagraph");
        } else first = first.AddText(route.StartsWith("join", StringComparison.Ordinal) ? "MM" : "MMMMX");
        first.SetStyleId("Normal");
        if (reset) first.SetStyleId("ExpandedParagraph");
        ApplyClosureSpacing(first, characterStyle, scale, tracking);
        // Empty structural runs must not replace the visible text's typography.
        if (!route.Contains("textbox", StringComparison.Ordinal))
            first._paragraph.InsertAfter(new W.Run(), first._paragraph.ParagraphProperties);
        if (route.StartsWith("join", StringComparison.Ordinal)) {
            HideJoinMark(first, true);
            WordParagraph second = story.AddParagraph("MMX");
            second.SetStyleId("Normal");
            if (reset) second.SetStyleId("ExpandedParagraph");
            ApplyClosureSpacing(second, characterStyle, scale, tracking);
            second._paragraph.InsertAfter(new W.Run(), second._paragraph.ParagraphProperties);
        }
        Assert.Empty(document.ValidateDocument());
        string before = footer ? document.Footer.Default!._footer!.OuterXml : document.Header.Default!._header!.OuterXml;
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Letter[] text = pdf.GetPage(1).Letters.Where(letter => letter.Value is "M" or "X").ToArray();
        Assert.Equal("MMMMX", string.Concat(text.Select(letter => letter.Value)));
        Assert.All(text, letter => Assert.Equal(text[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        // Helvetica M advances 9.996pt at 12pt; tracking is fixed page points.
        Assert.InRange(Math.Abs(text[4].StartBaseLine.X - text[0].StartBaseLine.X - (4D * 9.996D * scale / 100D + 4D * tracking / 20D)), 0D, 0.03D);
        if (route is "nested-textbox-prefix" or "nested-textbox-surrounded") {
            Letter a = pdf.GetPage(1).Letters.Single(letter => letter.Value == "A");
            Letter b = pdf.GetPage(1).Letters.Single(letter => letter.Value == "B" && Math.Abs(letter.StartBaseLine.Y - a.StartBaseLine.Y) < 0.01D);
            Assert.InRange(Math.Abs(b.StartBaseLine.X - a.StartBaseLine.X - (8.004D * scale / 100D + tracking / 20D)), 0D, 0.03D);
        }
        Assert.Equal(before, footer ? document.Footer.Default!._footer!.OuterXml : document.Header.Default!._header!.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HeaderFooterMixedTextBoxesKeepSourceOrderAndRepeatedTextIdentity(bool footer) {
        using WordDocument document = CreateJoinedParagraphDocument(); document.AddParagraph("Body");
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordParagraph outer = story.AddParagraph("MMMMX"); outer.Style = WordParagraphStyles.Normal;
        WordParagraph firstBox = outer.AddTextBox("MMMMX", WordImageTextWrapping.Square).Paragraphs.Single();
        firstBox.Style = WordParagraphStyles.Normal; firstBox.CharacterScale = 200; firstBox.Spacing = 20;
        WordParagraph suffix = outer.AddText("MMMMX");
        WordParagraph secondBox = suffix.AddTextBox("MMMMX", WordImageTextWrapping.Square).Paragraphs.Single();
        secondBox.Style = WordParagraphStyles.Normal; secondBox.CharacterScale = 50; secondBox.Spacing = -10;
        // Imported boxes commonly end with the paragraph rather than a hard break.
        // Identical surrounding text must still keep its own run formatting.
        firstBox._paragraph.Descendants<W.Break>().ToList().ForEach(lineBreak => lineBreak.Remove());
        secondBox._paragraph.Descendants<W.Break>().ToList().ForEach(lineBreak => lineBreak.Remove());
        using var pdf = OpenJoinedParagraphPdf(document);
        Letter[] text = pdf.GetPage(1).Letters.Where(letter => letter.Value is "M" or "X").ToArray();
        Assert.Equal(string.Concat(Enumerable.Repeat("MMMMX", 4)), string.Concat(text.Select(letter => letter.Value)));
        double[] expected = { 39.984D, 83.968D, 39.984D, 17.992D };
        for (int index = 0; index < 4; index++)
            Assert.InRange(Math.Abs(text[index * 5 + 4].StartBaseLine.X - text[index * 5].StartBaseLine.X - expected[index]), 0D, 0.03D);
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void HeaderFooterPeerTextBoxListsKeepMarkersAtTheirSourcePosition(bool footer, bool pageField) {
        using WordDocument document = CreateJoinedParagraphDocument();
        var list = document.AddList(WordListStyle.Numbered);
        list.Numbering.Levels[0].LevelSuffix = WordListLevelSuffix.Nothing;
        WordParagraph template = list.AddItem("Template");
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordParagraph outer = story.AddParagraph("MMMMX");
        WordParagraph first = outer.AddTextBox("MMMMX", WordImageTextWrapping.Square).Paragraphs.Single();
        WordParagraph suffix = outer.AddText("MMMMX");
        if (pageField)
            outer._paragraph.InsertBefore(new W.SimpleField(new W.Run(new W.Text("999"))) { Instruction = " PAGE " }, suffix._run);
        WordParagraph second = suffix.AddTextBox("MMMMX", WordImageTextWrapping.Square).Paragraphs.Single();
        foreach (WordParagraph item in new[] { first, second }) {
            item._paragraph.ParagraphProperties = (W.ParagraphProperties)template._paragraph.ParagraphProperties!.CloneNode(true);
            item._paragraph.Descendants<W.Break>().ToList().ForEach(lineBreak => lineBreak.Remove());
        }
        // A plain numbered peer must not switch the whole mixed story to a
        // renderer that drops its surrounding text or page field.
        story.AddParagraph("PLAIN-PEER")._paragraph.ParagraphProperties =
            (W.ParagraphProperties)template._paragraph.ParagraphProperties!.CloneNode(true);
        template._paragraph.Remove();
        document.AddParagraph("Body");
        using var pdf = OpenJoinedParagraphPdf(document);
        string visible = string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("MMMMX1.MMMMX" + (pageField ? "1" : "") + "MMMMX2.MMMMX", visible, StringComparison.Ordinal);
        Assert.Empty(document.ValidateDocument());
    }

    private static void ApplyClosureSpacing(WordParagraph paragraph, bool characterStyle, int scale, int tracking) {
        if (characterStyle) {
            W.Run run = paragraph._paragraph.Elements<W.Run>().Last();
            run.RunProperties ??= new W.RunProperties();
            run.RunProperties.RunStyle = new W.RunStyle { Val = "VisibleSpacing" };
        } else {
            paragraph.CharacterScale = scale;
            paragraph.Spacing = tracking;
        }
    }
}
