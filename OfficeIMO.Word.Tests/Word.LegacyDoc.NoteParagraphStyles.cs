using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    public void LegacyDoc_NoteParagraphStylesUseTheSharedStyleSheet(bool endnote, bool contentControl, bool customStyle) {
        using WordDocument source = WordDocument.Create();
        WordParagraph body = source.AddParagraph("Body reference");
        body.Style = WordParagraphStyles.ListParagraph;
        body.Style = WordParagraphStyles.Normal;
        if (endnote) body.AddEndNote("Styled note"); else body.AddFootNote("Styled note");
        var main = source._wordprocessingDocument.MainDocumentPart!;
        Styles styles = main.StyleDefinitionsPart!.Styles!;
        Style target = styles.Elements<Style>().Single(style => style.StyleId?.Value == "ListParagraph");
        if (customStyle) {
            target = (Style)target.CloneNode(true);
            target.StyleId = "CustomNoteParagraph";
            target.CustomStyle = true;
            target.StyleName = new StyleName { Val = "Custom note paragraph" };
            styles.Append(target);
        }
        target.StyleParagraphProperties!.SpacingBetweenLines = new SpacingBetweenLines { After = "280" };
        target.StyleParagraphProperties.ContextualSpacing = new ContextualSpacing { Val = true };
        OpenXmlCompositeElement note = endnote
            ? main.EndnotesPart!.Endnotes!.Elements<Endnote>().Single(item => item.Id?.Value > 0)
            : main.FootnotesPart!.Footnotes!.Elements<Footnote>().Single(item => item.Id?.Value > 0);
        Paragraph paragraph = note.Elements<Paragraph>().Single();
        paragraph.ParagraphProperties!.ParagraphStyleId = new ParagraphStyleId { Val = target.StyleId!.Value };
        if (contentControl) {
            paragraph.Remove();
            note.Append(new SdtBlock(new SdtContentBlock(paragraph)));
        }
        note.Append(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = endnote ? "EndnoteText" : "FootnoteText" }),
            new Run(new Text("Default note control"))));
        string originalStyles = styles.OuterXml;
        string expectedName = target.StyleName!.Val!.Value!;
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        for (int save = 0; save < 2; save++) {
            using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
            var importedMain = reopened._wordprocessingDocument.MainDocumentPart!;
            OpenXmlCompositeElement importedNote = endnote
                ? importedMain.EndnotesPart!.Endnotes!.Elements<Endnote>().Single(item => item.Id?.Value > 0)
                : importedMain.FootnotesPart!.Footnotes!.Elements<Footnote>().Single(item => item.Id?.Value > 0);
            Paragraph actual = importedNote.Descendants<Paragraph>().Single(item => item.InnerText.Contains("Styled note"));
            string? styleId = actual.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            Style actualStyle = importedMain.StyleDefinitionsPart!.Styles!.Elements<Style>()
                .Single(item => item.StyleId?.Value == styleId);
            Assert.Equal(expectedName, actualStyle.StyleName!.Val!.Value);
            Assert.Equal("280", actualStyle.StyleParagraphProperties!.SpacingBetweenLines!.After!.Value);
            Assert.True(actualStyle.StyleParagraphProperties.ContextualSpacing!.Val!.Value);
            Paragraph defaultParagraph = importedNote.Descendants<Paragraph>().Single(item => item.InnerText.Contains("Default note control"));
            Assert.NotEqual(styleId, defaultParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            Assert.Empty(reopened.ValidateDocument());
            if (save == 0) bytes = reopened.ToBytes(WordFileFormat.Doc);
        }
        Assert.Equal(originalStyles, styles.OuterXml);
    }
}
