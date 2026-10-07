using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "literal", "character", false)]
    [InlineData(true, "literal", "paragraph", false)]
    [InlineData(false, "literal", "character", true)]
    [InlineData(true, "literal", "direct", false)]
    [InlineData(false, "simple", "direct", false)]
    [InlineData(true, "simple", "character", false)]
    [InlineData(false, "simple", "paragraph", false)]
    [InlineData(true, "simple", "character", true)]
    [InlineData(false, "complex", "direct", false)]
    [InlineData(true, "complex", "character", false)]
    [InlineData(false, "complex", "paragraph", false)]
    [InlineData(true, "complex", "paragraph", true)]
    public void HeaderFieldSerializationRespectsEffectiveVisibility(bool footer, string route, string hiddenSource, bool visibleOverride) {
        using WordDocument document = CreateJoinedParagraphDocument();
        document.AddParagraph("BODY");
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new W.Style(new W.StyleRunProperties(new W.Vanish())) {
            Type = hiddenSource == "paragraph" ? W.StyleValues.Paragraph : W.StyleValues.Character, StyleId = "HiddenFieldText"
        });
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordParagraph paragraph = story.AddParagraph();
        if (hiddenSource == "paragraph") paragraph.SetStyleId("HiddenFieldText");
        W.Run Visible(string text) => new(new W.RunProperties(new W.Vanish { Val = false }), new W.Text(text));
        var hiddenProperties = new W.RunProperties();
        if (hiddenSource == "direct") hiddenProperties.AddChild(new W.Vanish(), true);
        if (hiddenSource == "character") hiddenProperties.AddChild(new W.RunStyle { Val = "HiddenFieldText" }, true);
        if (visibleOverride) hiddenProperties.AddChild(new W.Vanish { Val = false }, true);
        paragraph._paragraph.Append(Visible("VISIBLE"));
        if (route == "literal") {
            paragraph._paragraph.Append(new W.Run(hiddenProperties, new W.Text("SECRET")),
                new W.SimpleField(Visible("999")) { Instruction = " PAGE " });
        } else if (route == "simple") {
            paragraph._paragraph.Append(new W.SimpleField(new W.Run(hiddenProperties, new W.Text("999"))) { Instruction = " PAGE " });
        } else {
            paragraph._paragraph.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
                new W.Run(new W.FieldCode(" PAGE ")),
                new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
                new W.Run(hiddenProperties, new W.Text("999")),
                new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
        }
        paragraph._paragraph.Append(Visible("TAIL"));
        Assert.Empty(document.ValidateDocument());
        using var pdf = OpenJoinedParagraphPdf(document);
        string text = string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value));
        string expected = route == "literal" ? "VISIBLE" + (visibleOverride ? "SECRET" : "") + "1TAIL"
            : "VISIBLE" + (visibleOverride ? "1" : "") + "TAIL";
        Assert.Contains(expected, text);
        Assert.DoesNotContain("999", text);
        if (!visibleOverride) Assert.DoesNotContain("SECRET", text);
        if (route != "literal" && !visibleOverride) Assert.DoesNotContain("1", text);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void HeaderFieldWithoutCachedTextUsesItsEffectiveVisibility(bool complex, bool hiddenParagraph) {
        using WordDocument document = CreateJoinedParagraphDocument();
        document.AddParagraph("BODY");
        WordParagraph paragraph = document.HeaderDefaultOrCreate.AddParagraph("VISIBLE");
        if (hiddenParagraph) {
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
                new W.Style(new W.StyleRunProperties(new W.Vanish())) { Type = W.StyleValues.Paragraph, StyleId = "HiddenUncachedField" });
            paragraph.SetStyleId("HiddenUncachedField");
            paragraph.Hidden = false;
        }
        if (complex) paragraph._paragraph.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode(" PAGE ")), new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
        else paragraph._paragraph.Append(new W.SimpleField { Instruction = " PAGE " });
        Assert.Empty(document.ValidateDocument());
        using var pdf = OpenJoinedParagraphPdf(document);
        string text = string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains(hiddenParagraph ? "VISIBLE" : "VISIBLE1", text);
        if (hiddenParagraph) Assert.DoesNotContain("1", text);
    }
}
