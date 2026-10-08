using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public void NativeListParagraphStyle_PreservesAuthoredStyleAndContextualSpacing(bool contextualSpacing, bool headerOnly) {
        using WordDocument source = WordDocument.Create();
        if (headerOnly) {
            source.AddParagraph("Body control");
            source.AddHeadersAndFooters();
            source.Header.Default.AddParagraph("First item").Style = WordParagraphStyles.ListParagraph;
        } else {
            WordList list = source.AddList(WordListStyle.Numbered);
            list.AddItem("First item");
            list.AddItem("Second item");
        }
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style style = styles.Elements<Style>().Single(item => item.StyleId?.Value == "ListParagraph");
        style.StyleParagraphProperties!.ContextualSpacing = new ContextualSpacing { Val = contextualSpacing };
        style.StyleParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines { After = "280" };
        string before = styles.OuterXml;

        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        WordParagraph paragraph = (headerOnly ? reopened.Header.Default.Paragraphs : reopened.Paragraphs)
            .First(item => item.Text == "First item");
        Style actual = reopened._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Single(item => item.StyleId?.Value == paragraph.StyleId);
        Assert.Equal("List Paragraph", actual.StyleName!.Val!.Value);
        Assert.Equal(contextualSpacing, actual.StyleParagraphProperties!.ContextualSpacing!.Val!.Value);
        Assert.Equal("280", actual.StyleParagraphProperties.SpacingBetweenLines!.After!.Value);
        Assert.Equal(before, styles.OuterXml);
        Assert.Empty(reopened.ValidateDocument());
    }
}
