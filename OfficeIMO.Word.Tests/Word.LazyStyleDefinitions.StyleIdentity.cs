using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SelectingStylePreservesCaseDistinctAuthoredDefinition(bool builtIn) {
        string selectedId = builtIn ? "Heading1" : "AuthoredCaseStyle";
        string authoredId = builtIn ? "heading1" : "authoredCaseStyle";
        Style authored = new(
            new StyleName { Val = "Authored lowercase style" },
            new StyleRunProperties(new Italic(), new Color { Val = "315F23" })) {
            Type = StyleValues.Paragraph,
            StyleId = authoredId
        };
        string authoredXml = authored.OuterXml;
        using var input = CreateCaseDistinctStyleDocument(authored);

        try {
            if (!builtIn) {
                WordParagraphStyle.RegisterCustomStyle(selectedId, new WordParagraphStyleDefinition(selectedId) {
                    FontFamily = "Selected Font",
                    Bold = true
                });
            }
            using WordDocument document = WordDocument.Load(input);
            Assert.True(document.StyleExists(authoredId));
            Assert.False(document.StyleExists(selectedId));
            WordParagraph paragraph = document.AddParagraph("Selected style");
            if (builtIn) paragraph.Style = WordParagraphStyles.Heading1;
            else paragraph.SetStyleId(selectedId);

            using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
            using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
            MainDocumentPart mainPart = saved.MainDocumentPart!;
            Style[] styles = mainPart.StyleDefinitionsPart!.Styles!.Elements<Style>().ToArray();
            Assert.Equal(authoredXml, Assert.Single(styles, style => style.StyleId == authoredId).OuterXml);
            Style selected = Assert.Single(styles, style => style.StyleId == selectedId);
            Assert.Null(selected.Default);
            Assert.NotNull(selected.StyleRunProperties?.Bold);
            if (builtIn) Assert.Equal(0, selected.StyleParagraphProperties!.OutlineLevel!.Val!.Value);
            else Assert.Equal("Selected Font", selected.StyleRunProperties!.RunFonts!.Ascii!.Value);
            Assert.Equal(new[] { authoredId, selectedId }, mainPart.Document.Body!.Elements<Paragraph>()
                .Select(item => item.ParagraphProperties!.ParagraphStyleId!.Val!.Value));
            Assert.Equal("Normal", Assert.Single(styles, style => style.Default?.Value == true).StyleId!.Value);
            AssertStyleReferencesResolve(mainPart, styles);
            Assert.Empty(new OpenXmlValidator().Validate(saved));
        } finally {
            if (!builtIn) RemoveRegisteredStyles(selectedId);
        }
    }

    [Fact]
    public void SelectingRegisteredStyleAddsItsCaseDistinctRegisteredBase() {
        const string selectedId = "CaseRegisteredBase";
        const string baseId = "caseRegisteredBase";
        try {
            WordParagraphStyle.RegisterCustomStyle(selectedId, new WordParagraphStyleDefinition(selectedId) {
                Name = "Selected case style",
                BasedOnStyleId = baseId,
                Bold = true
            });
            WordParagraphStyle.RegisterCustomStyle(baseId, new WordParagraphStyleDefinition(baseId) {
                Name = "Base case style",
                FontFamily = "Inherited Font",
                Italic = true
            });
            using var input = CreateStyleIdentityDocument();
            using WordDocument document = WordDocument.Load(input);
            Assert.Single(document.Paragraphs).SetStyleId(selectedId);

            using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
            using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
            MainDocumentPart mainPart = saved.MainDocumentPart!;
            Style[] styles = mainPart.StyleDefinitionsPart!.Styles!.Elements<Style>().ToArray();
            Style selected = Assert.Single(styles, style => style.StyleId == selectedId);
            Style inherited = Assert.Single(styles, style => style.StyleId == baseId);
            Assert.Equal(baseId, selected.BasedOn!.Val!.Value);
            Assert.NotNull(selected.StyleRunProperties!.Bold);
            Assert.Equal("Inherited Font", inherited.StyleRunProperties!.RunFonts!.Ascii!.Value);
            Assert.NotNull(inherited.StyleRunProperties.Italic);
            Assert.Equal(selectedId, Assert.Single(mainPart.Document.Body!.Elements<Paragraph>())
                .ParagraphProperties!.ParagraphStyleId!.Val!.Value);
            AssertStyleReferencesResolve(mainPart, styles);
            Assert.Empty(new OpenXmlValidator().Validate(saved));
        } finally {
            RemoveRegisteredStyles(selectedId, baseId);
        }
    }

    [Fact]
    public void LoadingWithStyleOverridesPreservesCaseDistinctAuthoredStyle() {
        const string authoredId = "CaseOverrideStyle";
        const string registeredId = "caseOverrideStyle";
        Style authored = new(
            new StyleName { Val = "Authored uppercase style" },
            new StyleRunProperties(new Color { Val = "315F23" })) {
            Type = StyleValues.Paragraph,
            StyleId = authoredId
        };
        string authoredXml = authored.OuterXml;
        using var input = CreateCaseDistinctStyleDocument(authored);
        try {
            WordParagraphStyle.RegisterCustomStyle(registeredId, new WordParagraphStyleDefinition(registeredId) {
                Name = "Registered lowercase style",
                Bold = true
            });
            using WordDocument document = WordDocument.Load(input, new WordLoadOptions { OverrideStyles = true });

            using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
            using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
            MainDocumentPart mainPart = saved.MainDocumentPart!;
            Style[] styles = mainPart.StyleDefinitionsPart!.Styles!.Elements<Style>().ToArray();
            Assert.Equal(authoredXml, Assert.Single(styles, style => style.StyleId == authoredId).OuterXml);
            Assert.NotNull(Assert.Single(styles, style => style.StyleId == registeredId).StyleRunProperties!.Bold);
            Assert.Equal(authoredId, Assert.Single(mainPart.Document.Body!.Elements<Paragraph>())
                .ParagraphProperties!.ParagraphStyleId!.Val!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(saved));
        } finally {
            RemoveRegisteredStyles(registeredId);
        }
    }

    [Fact]
    public void CreatingDocumentKeepsCaseDistinctRegisteredStyles() {
        const string upperId = "CaseCreationStyle";
        const string lowerId = "caseCreationStyle";
        try {
            WordParagraphStyle.RegisterCustomStyle(upperId, new WordParagraphStyleDefinition(upperId) { Bold = true });
            WordParagraphStyle.RegisterCustomStyle(lowerId, new WordParagraphStyleDefinition(lowerId) { Italic = true });
            using WordDocument document = WordDocument.Create(new WordCreateOptions {
                DocumentType = WordDocumentType.MacroEnabledTemplate
            });
            document.AddParagraph("Uppercase style").SetStyleId(upperId);
            document.AddParagraph("Lowercase style").SetStyleId(lowerId);

            using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
            using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
            MainDocumentPart mainPart = saved.MainDocumentPart!;
            Style[] styles = mainPart.StyleDefinitionsPart!.Styles!.Elements<Style>().ToArray();
            Assert.NotNull(Assert.Single(styles, style => style.StyleId == upperId).StyleRunProperties!.Bold);
            Assert.NotNull(Assert.Single(styles, style => style.StyleId == lowerId).StyleRunProperties!.Italic);
            Assert.Equal(new[] { upperId, lowerId }, mainPart.Document.Body!.Elements<Paragraph>()
                .Select(item => item.ParagraphProperties!.ParagraphStyleId!.Val!.Value));
            AssertStyleReferencesResolve(mainPart, styles);
            Assert.Empty(new OpenXmlValidator().Validate(saved));
        } finally {
            RemoveRegisteredStyles(upperId, lowerId);
        }
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

    private static void RemoveRegisteredStyles(params string[] styleIds) {
        var field = typeof(WordParagraphStyle).GetField("_customStyles", BindingFlags.NonPublic | BindingFlags.Static)!;
        var registered = (IDictionary<string, Style>)field.GetValue(null)!;
        foreach (string styleId in styleIds) registered.Remove(styleId);
    }
}
