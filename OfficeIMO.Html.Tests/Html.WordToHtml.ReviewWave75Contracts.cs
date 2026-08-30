using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordToHtml {
    [Fact]
    public void CharacterStyleScriptIsEmittedOnlyByTheStructuralWrapper() {
        using WordDocument document = WordDocument.Create();
        var style = new Style { Type = StyleValues.Character, StyleId = "InheritedSuperscript" };
        style.Append(new StyleName { Val = "Inherited Superscript" });
        style.Append(new StyleRunProperties(
            new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }));
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
        document.AddParagraph().AddText("Script").SetCharacterStyleId("InheritedSuperscript");

        string html = document.ToHtml(new WordToHtmlOptions { IncludeRunClasses = true });

        string styleRule = Assert.Single(html.Split('\n'), line => line.Contains(".InheritedSuperscript {", StringComparison.Ordinal));
        Assert.DoesNotContain("vertical-align:super", styleRule, StringComparison.Ordinal);
        Assert.Equal(1, CountOccurrences(html, "<sup>"));
        Assert.Contains("<sup>Script</sup>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void DerivedCharacterStyleCanResetInheritedCapsAndSmallCaps() {
        using WordDocument document = WordDocument.Create();
        var inherited = new Style { Type = StyleValues.Character, StyleId = "InheritedCaps" };
        inherited.Append(new StyleName { Val = "Inherited Caps" });
        inherited.Append(new StyleRunProperties(new Caps(), new SmallCaps()));
        var reset = new Style { Type = StyleValues.Character, StyleId = "ResetCaps" };
        reset.Append(new StyleName { Val = "Reset Caps" });
        reset.Append(new BasedOn { Val = "InheritedCaps" });
        reset.Append(new StyleRunProperties(
            new Caps { Val = false },
            new SmallCaps { Val = false }));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(inherited);
        styles.Append(reset);
        document.AddParagraph().AddText("Mixed Case").SetCharacterStyleId("ResetCaps");

        string html = document.ToHtml(new WordToHtmlOptions { IncludeRunClasses = true });

        string styleRule = Assert.Single(html.Split('\n'), line => line.Contains(".ResetCaps {", StringComparison.Ordinal));
        Assert.Contains("font-variant:normal", styleRule, StringComparison.Ordinal);
        Assert.Contains("text-transform:none", styleRule, StringComparison.Ordinal);
    }

    private static int CountOccurrences(string value, string token) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(token, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += token.Length;
        }
        return count;
    }
}
