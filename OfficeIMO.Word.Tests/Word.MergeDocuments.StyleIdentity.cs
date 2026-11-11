using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordStyleConflictMode.KeepTarget)]
    [InlineData(WordStyleConflictMode.Rename)]
    public void MergePreservesCaseDistinctStyleIdsAndDependencies(WordStyleConflictMode mode) {
        using var input = new MemoryStream();
        using (WordprocessingDocument package = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document, true)) {
            MainDocumentPart mainPart = package.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "CaseBody" }), new Run(new Text("Upper paragraph"))),
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "casebody" }), new Run(new Text("Lower paragraph")))));
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new Style(new StyleName { Val = "Upper source base" }, new StyleRunProperties(new Color { Val = "D42137" })) {
                    Type = StyleValues.Paragraph, StyleId = "CaseBase"
                },
                new Style(new StyleName { Val = "Lower source base" }, new StyleRunProperties(new Bold(), new Color { Val = "164FBA" })) {
                    Type = StyleValues.Paragraph, StyleId = "casebase"
                },
                new Style(new StyleName { Val = "Upper source body" }, new BasedOn { Val = "CaseBase" }) {
                    Type = StyleValues.Paragraph, StyleId = "CaseBody"
                },
                new Style(new StyleName { Val = "Lower source body" }, new BasedOn { Val = "casebase" }) {
                    Type = StyleValues.Paragraph, StyleId = "casebody"
                });
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
        input.Position = 0;
        using WordDocument source = WordDocument.Load(input);
        MainDocumentPart sourceMain = source._wordprocessingDocument.MainDocumentPart!;
        string sourceStylesXml = sourceMain.StyleDefinitionsPart!.Styles!.OuterXml;
        string sourceBodyXml = sourceMain.Document.Body!.OuterXml;
        using WordDocument destination = WordDocument.Create();
        destination.AddParagraph("Destination text");

        destination.AppendDocument(source, new WordDocumentMergeOptions { StyleConflictMode = mode });

        Assert.Equal(sourceStylesXml, sourceMain.StyleDefinitionsPart.Styles.OuterXml);
        Assert.Equal(sourceBodyXml, sourceMain.Document.Body.OuterXml);
        using var savedStream = new MemoryStream(destination.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        MainDocumentPart savedMain = saved.MainDocumentPart!;
        Style[] styles = savedMain.StyleDefinitionsPart!.Styles!.Elements<Style>().ToArray();
        var expected = new[] {
            (Text: "Upper paragraph", Id: "CaseBody", Parent: "CaseBase", Color: "D42137", Bold: false),
            (Text: "Lower paragraph", Id: "casebody", Parent: "casebase", Color: "164FBA", Bold: true)
        };
        foreach (var item in expected) {
            Paragraph paragraph = Assert.Single(savedMain.Document.Body!.Elements<Paragraph>(), value => value.InnerText == item.Text);
            Assert.Equal(item.Id, paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
            Style bodyStyle = Assert.Single(styles, style => style.StyleId == item.Id);
            Assert.Equal(item.Parent, bodyStyle.BasedOn!.Val!.Value);
            Style parentStyle = Assert.Single(styles, style => style.StyleId == bodyStyle.BasedOn.Val);
            Assert.Equal(item.Color, parentStyle.StyleRunProperties!.Color!.Val!.Value);
            Assert.Equal(item.Bold, parentStyle.StyleRunProperties.Bold != null);
        }
        Assert.Empty(new OpenXmlValidator().Validate(saved));
    }
}
