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
}
