using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void ParagraphText_ExplicitBreakProjectionPreservesLiteralSeparatorAndSource() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        var sourceRun = CreateFlowBreakProjectionRun();
        paragraph._paragraph.Append(sourceRun);
        var run = new WordParagraph(document, paragraph._paragraph, sourceRun);
        string sourceXml = paragraph._paragraph.OuterXml;

        string projected = run.GetTextWithNonTextBreakReplacement('\n');

        Assert.Equal("\U0001F642LITERAL\u2028BEFORE\nMIDDLE\nAFTER", projected);
        Assert.Equal("\U0001F642LITERAL\u2028BEFORE\u2028MIDDLE\u2028AFTER", run.Text);
        Assert.Equal(sourceXml, paragraph._paragraph.OuterXml);
    }

    [Fact]
    public void ParagraphText_VisibleXmlBreakProjectionPreservesLiteralSeparatorAndSource() {
        var paragraph = new W.Paragraph(CreateFlowBreakProjectionRun());
        string sourceXml = paragraph.OuterXml;

        string projected = WordParagraph.ReadVisibleText(paragraph, ' ');

        Assert.Equal("\U0001F642LITERAL\u2028BEFORE MIDDLE AFTER", projected);
        Assert.Equal("\U0001F642LITERAL\u2028BEFORE\u2028MIDDLE\u2028AFTER", WordParagraph.ReadVisibleText(paragraph));
        Assert.Equal(sourceXml, paragraph.OuterXml);
    }

    private static W.Run CreateFlowBreakProjectionRun() => new(
        new W.Text("\U0001F642LITERAL\u2028BEFORE"),
        new W.Break { Type = W.BreakValues.Page },
        new W.Text("MIDDLE"),
        new W.Break { Type = W.BreakValues.Column },
        new W.Text("AFTER"));
}
