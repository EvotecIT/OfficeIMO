using System.Reflection;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlTextFormattingReviewWave35Tests {
    [Fact]
    public void HtmlRenderTextRetainsLegacyClrConstructorSignatures() {
        Type[] shortSignature = {
            typeof(string), typeof(double), typeof(double), typeof(double), typeof(double),
            typeof(OfficeFontInfo), typeof(OfficeColor), typeof(OfficeTextAlignment), typeof(double), typeof(int),
            typeof(string), typeof(string), typeof(string), typeof(double?), typeof(int?), typeof(bool)
        };
        Type[] positionedSignature = shortSignature.Take(15)
            .Concat(new[] { typeof(double?), typeof(bool), typeof(int?), typeof(int?) })
            .ToArray();

        Assert.NotNull(typeof(HtmlRenderText).GetConstructor(
            BindingFlags.Instance | BindingFlags.NonPublic, null, shortSignature, null));
        Assert.NotNull(typeof(HtmlRenderText).GetConstructor(
            BindingFlags.Instance | BindingFlags.NonPublic, null, positionedSignature, null));
        Type[] semanticRunSignature = {
            typeof(string), typeof(string), typeof(bool), typeof(bool), typeof(bool), typeof(bool), typeof(bool), typeof(bool),
            typeof(HtmlComputedStyle), typeof(HtmlSemanticSourceLocation), typeof(bool),
            typeof(IReadOnlyDictionary<string, string>), typeof(string)
        };
        Assert.NotNull(typeof(HtmlSemanticRun).GetConstructor(
            BindingFlags.Instance | BindingFlags.NonPublic, null, semanticRunSignature, null));
    }

    [Fact]
    public void SemanticRunsCarryAncestorDecorationPatternsThroughNestedInlineElements() {
        HtmlSemanticBlock paragraph = Assert.Single(HtmlConversionDocument.Parse("""
            <p><u style="text-decoration-style:double"><strong>Under</strong></u>
            <s style="text-decoration-style:dashed"><em>Strike</em></s></p>
            """).SemanticDocument.Sections.SelectMany(section => section.Blocks));

        HtmlSemanticRun underline = Assert.Single(paragraph.Runs, run => run.Text == "Under");
        HtmlSemanticRun strike = Assert.Single(paragraph.Runs, run => run.Text == "Strike");
        Assert.True(underline.Underline);
        Assert.Equal(OfficeTextDecorationStyle.Double, underline.UnderlineStyle);
        Assert.True(strike.Strikethrough);
        Assert.Equal(OfficeTextDecorationStyle.Dashed, strike.StrikethroughStyle);
    }

    [Fact]
    public void PowerPointSemanticBlocksStayPairedWhenLayerOrderDiffersFromDomOrder() {
        const string html = """
            <section class="officeimo-slide">
              <p data-officeimo-layer-index="1"><strong>Top layer</strong></p>
              <p data-officeimo-layer-index="0"><em>Bottom layer</em></p>
            </section>
            """;

        using PowerPointPresentation imported = HtmlConversionDocument
            .Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToPowerPointPresentationResult()
            .RequireValue();
        PowerPointTextBox[] boxes = Assert.Single(imported.Slides).TextBoxes.ToArray();

        Assert.Equal(new[] { "Bottom layer", "Top layer" }, boxes.Select(box => box.Text));
        Assert.True(Assert.Single(boxes[0].Paragraphs).Runs.Single().Italic);
        Assert.False(boxes[0].Paragraphs.Single().Runs.Single().Bold);
        Assert.True(Assert.Single(boxes[1].Paragraphs).Runs.Single().Bold);
        Assert.False(boxes[1].Paragraphs.Single().Runs.Single().Italic);
    }

    [Fact]
    public void HtmlDoubleLineThroughImportsAsNativeWordDoubleStrike() {
        using WordDocument imported = HtmlConversionDocument.Parse("""
            <p><span style="text-decoration-line:line-through;text-decoration-style:double">Double</span></p>
            """).ToWordDocumentResult().RequireValue();

        WordParagraph run = Assert.Single(imported.Paragraphs);
        Assert.True(run.DoubleStrike);
        Assert.False(run.Strike);
    }

    [Fact]
    public void WordRunExplicitCapsOffResetsInheritedCharacterStylesInHtml() {
        using WordDocument source = WordDocument.Create();
        AddCapitalizationStyle(source, "InheritedCaps", new Caps());
        AddCapitalizationStyle(source, "InheritedSmallCaps", new SmallCaps());
        WordParagraph caps = source.AddParagraph().AddText("Caps off").SetCharacterStyleId("InheritedCaps");
        caps._runProperties!.Caps = new Caps { Val = false };
        WordParagraph smallCaps = source.AddParagraph().AddText("Small caps off").SetCharacterStyleId("InheritedSmallCaps");
        smallCaps._runProperties!.SmallCaps = new SmallCaps { Val = false };

        string html = source.ToHtml(new WordToHtmlOptions { IncludeRunClasses = true });

        Assert.Contains("text-transform:none", html, StringComparison.Ordinal);
        Assert.Contains("font-variant:normal", html, StringComparison.Ordinal);
        Assert.Contains(".InheritedCaps { text-transform:uppercase", html, StringComparison.Ordinal);
        Assert.Contains(".InheritedSmallCaps { font-variant:small-caps", html, StringComparison.Ordinal);
    }

    private static void AddCapitalizationStyle(WordDocument document, string styleId, OnOffType capitalization) {
        var style = new Style { Type = StyleValues.Character, StyleId = styleId };
        style.Append(new StyleName { Val = styleId });
        var properties = new StyleRunProperties();
        properties.Append(capitalization);
        style.Append(properties);
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
    }
}
