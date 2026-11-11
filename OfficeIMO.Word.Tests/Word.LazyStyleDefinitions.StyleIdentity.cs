using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SelectingBuiltInStylePreservesCaseDistinctAuthoredDefinition() {
        const string selectedId = "Heading1";
        const string authoredId = "heading1";
        Style authored = new(
            new StyleName { Val = "Authored lowercase style" },
            new StyleRunProperties(new Italic(), new Color { Val = "315F23" })) {
            Type = StyleValues.Paragraph,
            StyleId = authoredId
        };
        string authoredXml = authored.OuterXml;
        using var input = CreateCaseDistinctStyleDocument(authored);

        using WordDocument document = WordDocument.Load(input);
        Assert.True(document.StyleExists(authoredId));
        Assert.False(document.StyleExists(selectedId));
        WordParagraph paragraph = document.AddParagraph("Selected style");
        paragraph.Style = WordParagraphStyles.Heading1;

        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        MainDocumentPart mainPart = saved.MainDocumentPart!;
        Style[] styles = mainPart.StyleDefinitionsPart!.Styles!.Elements<Style>().ToArray();
        Assert.Equal(authoredXml, Assert.Single(styles, style => style.StyleId == authoredId).OuterXml);
        Style selected = Assert.Single(styles, style => style.StyleId == selectedId);
        Assert.Null(selected.Default);
        Assert.NotNull(selected.StyleRunProperties?.Bold);
        Assert.Equal(0, selected.StyleParagraphProperties!.OutlineLevel!.Val!.Value);
        Assert.Equal(new[] { authoredId, selectedId }, mainPart.Document.Body!.Elements<Paragraph>()
            .Select(item => item.ParagraphProperties!.ParagraphStyleId!.Val!.Value));
        Assert.Equal("Normal", Assert.Single(styles, style => style.Default?.Value == true).StyleId!.Value);
        AssertStyleReferencesResolve(mainPart, styles);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
    }

    [Fact]
    public void AddingCodeBlockPreservesCaseDistinctAuthoredStyle() {
        const string authoredId = "officeimocodeblock";
        Style authored = new(
            new StyleName { Val = "Authored code style" },
            new StyleRunProperties(new Color { Val = "315F23" })) {
            Type = StyleValues.Paragraph,
            StyleId = authoredId
        };
        string authoredXml = authored.OuterXml;
        using var input = CreateCaseDistinctStyleDocument(authored);
        using WordDocument document = WordDocument.Load(input);

        document.AddCodeBlock("return 1;");

        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        MainDocumentPart mainPart = saved.MainDocumentPart!;
        Style[] styles = mainPart.StyleDefinitionsPart!.Styles!.Elements<Style>().ToArray();
        Assert.Equal(authoredXml, Assert.Single(styles, style => style.StyleId == authoredId).OuterXml);
        Style code = Assert.Single(styles, style => style.StyleId == "OfficeIMOCodeBlock");
        Assert.False(string.IsNullOrWhiteSpace(code.StyleRunProperties!.RunFonts!.Ascii!.Value));
        Assert.Equal(new[] { authoredId, "OfficeIMOCodeBlock" }, mainPart.Document.Body!.Elements<Paragraph>()
            .Select(item => item.ParagraphProperties!.ParagraphStyleId!.Val!.Value));
        AssertStyleReferencesResolve(mainPart, styles);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
    }

    private static MemoryStream CreateCaseDistinctStyleDocument(Style authored, params Style[] additionalStyles) {
        MemoryStream input = CreateStyleIdentityDocument(new[] { authored }.Concat(additionalStyles).ToArray());
        using (WordprocessingDocument package = WordprocessingDocument.Open(input, true)) {
            package.MainDocumentPart!.Document.Body!.GetFirstChild<Paragraph>()!.ParagraphProperties =
                new ParagraphProperties(new ParagraphStyleId { Val = authored.StyleId!.Value });
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
        input.Position = 0;
        return input;
    }

    private static MemoryStream CreateStyleIdentityDocument(params Style[] additionalStyles) {
        var input = new MemoryStream();
        using (WordprocessingDocument package = WordprocessingDocument.Create(
                   input, WordprocessingDocumentType.Document, autoSave: true)) {
            MainDocumentPart mainPart = package.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Minimal document")))));
            StyleDefinitionsPart stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new Style(new StyleName { Val = "Normal" }) {
                Type = StyleValues.Paragraph,
                StyleId = "Normal",
                Default = true
            });
            stylesPart.Styles.Append(additionalStyles);
        }
        input.Position = 0;
        return input;
    }

    private static void AssertStyleReferencesResolve(MainDocumentPart mainPart, Style[] styles) {
        var ids = new HashSet<string>(styles.Select(style => style.StyleId!.Value!), StringComparer.Ordinal);
        Assert.Equal(styles.Length, ids.Count);
        Assert.All(mainPart.Document.Descendants<ParagraphStyleId>(), reference => Assert.Contains(reference.Val!.Value!, ids));
        Assert.All(styles.SelectMany(style => style.Elements<BasedOn>()), reference => Assert.Contains(reference.Val!.Value!, ids));
        Assert.All(styles.SelectMany(style => style.Elements<LinkedStyle>()), reference => Assert.Contains(reference.Val!.Value!, ids));
        Assert.All(styles.SelectMany(style => style.Elements<NextParagraphStyle>()), reference => Assert.Contains(reference.Val!.Value!, ids));
    }
}
