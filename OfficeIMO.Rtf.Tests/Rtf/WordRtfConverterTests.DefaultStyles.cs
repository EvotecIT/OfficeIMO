using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Theory]
    [InlineData(true, true, false)]
    [InlineData(false, true, false)]
    [InlineData(false, false, false)]
    [InlineData(false, true, true)]
    public void Word_Rtf_Bridge_Preserves_Default_Paragraph_Role_And_Style_References(
        bool tableFirst, bool hasDefault, bool omitParagraphType) {
        using var input = new MemoryStream();
        using (WordprocessingDocument package = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document, true)) {
            MainDocumentPart main = package.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Implicit body"))),
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "ExplicitBody" }),
                    new Run(new Text("Explicit body")))));
            var bodyStyle = new Style(
                new StyleName { Val = "Authored body" },
                new NextParagraphStyle { Val = "ExplicitBody" },
                new LinkedStyle { Val = "AuthoredCharacter" },
                new StyleRunProperties(new Color { Val = "D42137" }, new FontSize { Val = "36" })) {
                StyleId = "AuthoredBody", Type = StyleValues.Paragraph
            };
            if (hasDefault) bodyStyle.Default = true;
            if (omitParagraphType) bodyStyle.Type = null;
            var tableStyle = new Style(new StyleName { Val = "Authored table" }) {
                StyleId = "AuthoredTable", Type = StyleValues.Table, Default = true
            };
            var styles = new Styles();
            if (tableFirst) styles.Append(tableStyle, bodyStyle);
            else styles.Append(bodyStyle, tableStyle);
            styles.Append(new Style(
                new StyleName { Val = "Explicit body" },
                new BasedOn { Val = "AuthoredBody" },
                new StyleRunProperties(new Italic())) {
                StyleId = "ExplicitBody", Type = StyleValues.Paragraph
            });
            styles.Append(new Style(
                new StyleName { Val = "Authored character" },
                new LinkedStyle { Val = "AuthoredBody" }) {
                StyleId = "AuthoredCharacter", Type = StyleValues.Character
            });
            main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
        input.Position = 0;
        using WordDocument source = WordDocument.Load(input);
        string sourceStyles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml;

        string encoded = source.ToRtf().ToRtf(new RtfSaveOptions { IncludeGenerator = false });
        RtfDocument parsed = RtfDocument.ParseResult(encoded).Document;

        Assert.Equal(sourceStyles, source._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart!.Styles!.OuterXml);
        Assert.Equal(new[] { "Implicit body", "Explicit body" }, parsed.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
        RtfRun implicitRun = parsed.ResolveRunFormatting(parsed.Paragraphs[0], Assert.Single(parsed.Paragraphs[0].Runs));
        if (hasDefault) {
            Assert.Equal(18D, implicitRun.FontSize);
            Assert.Equal("#D42137", parsed.GetColor(Assert.IsType<int>(implicitRun.ForegroundColorIndex))!.ToString());
        } else {
            Assert.Null(implicitRun.FontSize);
            Assert.Null(implicitRun.ForegroundColorIndex);
        }

        RtfStyle body = Assert.Single(parsed.Styles, style => style.Name == "AuthoredBody");
        RtfStyle explicitStyle = Assert.Single(parsed.Styles, style => style.Name == "ExplicitBody");
        RtfStyle character = Assert.Single(parsed.Styles, style => style.Name == "AuthoredCharacter");
        Assert.Equal(explicitStyle.Id, parsed.Paragraphs[1].StyleId);
        Assert.Equal(body.Id, explicitStyle.BasedOnStyleId);
        Assert.Equal(explicitStyle.Id, body.NextStyleId);
        Assert.Equal(character.Id, body.LinkedStyleId);
        Assert.Equal(body.Id, character.LinkedStyleId);
        RtfRun explicitRun = parsed.ResolveRunFormatting(parsed.Paragraphs[1], Assert.Single(parsed.Paragraphs[1].Runs));
        Assert.Equal(18D, explicitRun.FontSize);
        Assert.Equal("#D42137", parsed.GetColor(Assert.IsType<int>(explicitRun.ForegroundColorIndex))!.ToString());
        Assert.True(explicitRun.Italic);
    }

    [Fact]
    public void Word_Rtf_Bridge_Keeps_Case_Distinct_Style_Ids_And_References() {
        using var input = new MemoryStream();
        using (WordprocessingDocument package = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document, true)) {
            MainDocumentPart main = package.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "CaseBody" }),
                    new Run(new RunProperties(new RunStyle { Val = "CaseCharacter" }), new Text("Upper"))),
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "casebody" }),
                    new Run(new RunProperties(new RunStyle { Val = "casecharacter" }), new Text("Lower")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new Style(new StyleName { Val = "Red body" },
                    new StyleRunProperties(new Color { Val = "D42137" }, new FontSize { Val = "36" })) {
                    StyleId = "CaseBody", Type = StyleValues.Paragraph, Default = true
                },
                new Style(new StyleName { Val = "Blue body" },
                    new BasedOn { Val = "CaseBody" },
                    new StyleRunProperties(new Color { Val = "1269A3" }, new FontSize { Val = "48" })) {
                    StyleId = "casebody", Type = StyleValues.Paragraph
                },
                new Style(new StyleName { Val = "Bold character" }, new StyleRunProperties(new Bold())) {
                    StyleId = "CaseCharacter", Type = StyleValues.Character
                },
                new Style(new StyleName { Val = "Italic character" }, new StyleRunProperties(new Italic())) {
                    StyleId = "casecharacter", Type = StyleValues.Character
                });
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
        input.Position = 0;
        using WordDocument source = WordDocument.Load(input);
        string sourceStyles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml;

        string encoded = source.ToRtf().ToRtf(new RtfSaveOptions { IncludeGenerator = false });
        RtfDocument parsed = RtfDocument.ParseResult(encoded).Document;

        Assert.Equal(sourceStyles, source._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart!.Styles!.OuterXml);
        Assert.Equal(new[] { "Upper", "Lower" }, parsed.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
        RtfStyle upperBody = Assert.Single(parsed.Styles, style => style.Name == "CaseBody");
        RtfStyle lowerBody = Assert.Single(parsed.Styles, style => style.Name == "casebody");
        RtfStyle upperCharacter = Assert.Single(parsed.Styles, style => style.Name == "CaseCharacter");
        RtfStyle lowerCharacter = Assert.Single(parsed.Styles, style => style.Name == "casecharacter");
        Assert.NotEqual(upperBody.Id, lowerBody.Id);
        Assert.NotEqual(upperCharacter.Id, lowerCharacter.Id);
        Assert.Equal(upperBody.Id, parsed.Paragraphs[0].StyleId);
        Assert.Equal(lowerBody.Id, parsed.Paragraphs[1].StyleId);
        Assert.Equal(upperBody.Id, lowerBody.BasedOnStyleId);
        RtfRun upper = parsed.ResolveRunFormatting(parsed.Paragraphs[0], Assert.Single(parsed.Paragraphs[0].Runs));
        RtfRun lower = parsed.ResolveRunFormatting(parsed.Paragraphs[1], Assert.Single(parsed.Paragraphs[1].Runs));
        Assert.Equal(upperCharacter.Id, upper.StyleId);
        Assert.Equal(lowerCharacter.Id, lower.StyleId);
        Assert.Equal(18D, upper.FontSize);
        Assert.Equal(24D, lower.FontSize);
        Assert.Equal("#D42137", parsed.GetColor(Assert.IsType<int>(upper.ForegroundColorIndex))!.ToString());
        Assert.Equal("#1269A3", parsed.GetColor(Assert.IsType<int>(lower.ForegroundColorIndex))!.ToString());
        Assert.True(upper.Bold);
        Assert.False(upper.Italic);
        Assert.False(lower.Bold);
        Assert.True(lower.Italic);
    }

}
