using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordToHtml {
    [Fact]
    public void WordToHtmlPreservesCaseDistinctStyleIdsAndInheritedFormatting() {
        using var input = new MemoryStream();
        using (WordprocessingDocument package = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document, true)) {
            MainDocumentPart main = package.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "CaseBody" }),
                    new Run(new RunProperties(new RunStyle { Val = "CaseText" }), new Text("Upper"))),
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "casebody" }),
                    new Run(new RunProperties(new RunStyle { Val = "casetext" }), new Text("Lower"))),
                new Paragraph(new Run(new RunProperties(new RunStyle { Val = "htmlquote" }), new Text("Ordinary quote")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new Style(new StyleName { Val = "Red body" }, OrderedRunProperties(
                    new RunFonts { Ascii = "Georgia" }, new Color { Val = "D42137" }, new FontSize { Val = "36" },
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "D9E7F3" })) {
                    Type = StyleValues.Paragraph, StyleId = "CaseBody"
                },
                new Style(new StyleName { Val = "Blue body" }, new BasedOn { Val = "CaseBody" }, OrderedRunProperties(
                    new Color { Val = "164FBA" }, new FontSize { Val = "48" },
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "E8F1DB" })) {
                    Type = StyleValues.Paragraph, StyleId = "casebody"
                },
                new Style(new StyleName { Val = "Underlined text" }, new StyleRunProperties(new Underline { Val = UnderlineValues.Single })) {
                    Type = StyleValues.Character, StyleId = "CaseText"
                },
                new Style(new StyleName { Val = "Struck text" }, new BasedOn { Val = "CaseText" }, new StyleRunProperties(new Strike())) {
                    Type = StyleValues.Character, StyleId = "casetext"
                },
                new Style(new StyleName { Val = "HTML quote delimiter" }) { Type = StyleValues.Character, StyleId = "HtmlQuote" },
                new Style(new StyleName { Val = "Ordinary authored text" }) { Type = StyleValues.Character, StyleId = "htmlquote" });
            Assert.Empty(new OpenXmlValidator().Validate(package).Select(error => error.Description));
        }
        input.Position = 0;
        using WordDocument document = WordDocument.Load(input);
        string before = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.OuterXml;

        string output = document.ToHtml(new WordHtmlExportOptions {
            IncludeDefaultCss = false, IncludeParagraphClasses = true, IncludeRunClasses = true
        });
        HtmlConversionDocument html = HtmlConversionDocument.Parse(output);
        var computed = HtmlComputedStyleEngine.Compute(html);
        var upper = Assert.Single(html.Document.QuerySelectorAll("p.CaseBody"));
        var lower = Assert.Single(html.Document.QuerySelectorAll("p.casebody"));
        Assert.Equal("Upper", upper.TextContent);
        Assert.Equal("Lower", lower.TextContent);
        Assert.Equal("rgba(212, 33, 55, 1)", computed[upper].GetValue("color"));
        Assert.Equal("rgba(22, 79, 186, 1)", computed[lower].GetValue("color"));
        Assert.Equal(18D, computed[upper].ResolvedFontSizePoints);
        Assert.Equal(24D, computed[lower].ResolvedFontSizePoints);
        Assert.Contains("Georgia", computed[lower].GetValue("font-family"), StringComparison.Ordinal);
        var upperText = Assert.Single(upper.QuerySelectorAll("span.CaseText"));
        var lowerText = Assert.Single(lower.QuerySelectorAll("span.casetext"));
        Assert.NotNull(upperText.QuerySelector("u"));
        Assert.Null(upperText.QuerySelector("s"));
        Assert.NotNull(lowerText.QuerySelector("u"));
        Assert.NotNull(lowerText.QuerySelector("s"));
        Assert.Equal("Ordinary quote", Assert.Single(html.Document.QuerySelectorAll("span.htmlquote")).TextContent);
        Assert.Contains(lower.QuerySelectorAll("[style]"), element =>
            element.GetAttribute("style")!.Contains("background-color:#e8f1db", StringComparison.Ordinal));
        WordParagraph lowerRun = document.Paragraphs.Single(paragraph => paragraph.Text == "Lower");
        Assert.Equal("E8F1DB", WordRunFormatting.ResolveRunShadingFillColorHex(lowerRun));
        Assert.Equal(before, document.MainDocumentPartRoot.StyleDefinitionsPart.Styles.OuterXml);
    }

    private static StyleRunProperties OrderedRunProperties(params OpenXmlElement[] children) {
        var properties = new StyleRunProperties();
        foreach (OpenXmlElement child in children) properties.AddChild(child, true);
        return properties;
    }
}
