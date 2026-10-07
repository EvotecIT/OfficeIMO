using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, true)]
    public void HeaderFieldSerializerKeepsInheritedCaseAndRunSpacing(bool footer, bool characterStyle, bool smallCaps) {
        using WordDocument document = CreateJoinedParagraphDocument();
        document.AddParagraph("BODY");
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        W.StyleRunProperties formatting = smallCaps ? new(new W.SmallCaps()) : new(new W.Caps());
        styles.Append(new W.Style(formatting) {
            Type = characterStyle ? W.StyleValues.Character : W.StyleValues.Paragraph, StyleId = "InheritedCase"
        });
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordParagraph paragraph = story.AddParagraph();
        paragraph.SetStyleId(characterStyle ? "Normal" : "InheritedCase");
        WordParagraph run = paragraph.AddText("mmmmx");
        run.FontFamily = "Arial";
        run.FontSize = 12;
        run.CharacterScale = 200;
        run.Spacing = 20;
        if (characterStyle) run._run.RunProperties!.RunStyle = new W.RunStyle { Val = "InheritedCase" };
        paragraph._paragraph.Append(new W.SimpleField(new W.Run(new W.Text("999"))) { Instruction = " PAGE " });
        Assert.Empty(document.ValidateDocument());
        using var pdf = OpenJoinedParagraphPdf(document);
        var text = pdf.GetPage(1).Letters.Where(letter => letter.Value is "M" or "X" or "m" or "x").ToArray();
        Assert.Equal("MMMMX", string.Concat(text.Select(letter => letter.Value)));
        Assert.InRange(Math.Abs(text[4].StartBaseLine.X - text[0].StartBaseLine.X - 83.968D), 0D, 0.03D);
        Assert.Contains(pdf.GetPage(1).Letters, letter => letter.Value == "1");
    }
}
