using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    public static TheoryData<string> WordAuthoredStyleFixtures => new() {
        "WordLayoutFixtures/native-word-report.docx",
        "Documents/sample1.docx",
        "Documents/sample3.docx",
        "Documents/Word/PremiumGaps/TemplateMailMerge/word-authored-content-control-form.docx"
    };

    [Theory]
    [MemberData(nameof(WordAuthoredStyleFixtures))]
    public void ReadOnlyStylesAreFoundWithoutLoadingStylesPartAndMatchLoadedDefinitions(string fixture) {
        byte[] package = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, fixture));
        string[] styleIds;
        using (var input = new MemoryStream(package, writable: false))
        using (WordDocument document = WordDocument.Load(input, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
            styleIds = document.Styles.Select(style => style.StyleId).ToArray();
        }

        string[] streamed = DescribeCatalogStyles(package, styleIds, loadStylesFirst: false, out bool stylesLoaded);
        Assert.False(stylesLoaded);
        Assert.NotEmpty(styleIds);
        Assert.Equal(DescribeCatalogStyles(package, styleIds, loadStylesFirst: true, out _), streamed);
    }

    [Theory]
    [InlineData("WordLayoutFixtures/native-word-report.docx")]
    [InlineData("Documents/sample1.docx")]
    public void ConversionsOfLoadedDocumentReadStylesWithoutLoadingStylesPartAndMatchLoadedDom(string fixture) {
        byte[] package = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, fixture));

        StyleCatalogConversion streamed = ConvertWithStyleCatalog(package, loadStylesFirst: false);
        StyleCatalogConversion loaded = ConvertWithStyleCatalog(package, loadStylesFirst: true);
        Assert.False(streamed.StylesLoaded);
        Assert.Equal(loaded.Inspection, streamed.Inspection);
        Assert.Equal(loaded.Markdown, streamed.Markdown);
        Assert.Equal(loaded.Html, streamed.Html);
        Assert.Equal(loaded.Pages.Count, streamed.Pages.Count);
        for (int index = 0; index < loaded.Pages.Count; index++) Assert.Equal(loaded.Pages[index], streamed.Pages[index]);
    }

    [Fact]
    public void HtmlExportElementLimitCountsAStylesPartThatWasNotLoaded() {
        string path = Path.Combine(AppContext.BaseDirectory, "WordLayoutFixtures", "native-word-report.docx");
        byte[] package = File.ReadAllBytes(path);
        long styleElements;
        using (var input = new MemoryStream(package, writable: false))
        using (WordprocessingDocument raw = WordprocessingDocument.Open(input, false)) {
            using Stream styles = raw.MainDocumentPart!.StyleDefinitionsPart!.GetStream();
            styleElements = XDocument.Load(styles).Descendants().LongCount();
        }

        // The styles part alone exceeds this budget, so the limit trips while counting it, loaded or not.
        foreach (bool loadStylesFirst in new[] { false, true }) {
            using var input = new MemoryStream(package, writable: false);
            using WordDocument document = WordDocument.Load(input, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            StyleDefinitionsPart stylesPart = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!;
            if (loadStylesFirst) _ = stylesPart.Styles;
            var exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtml(new WordHtmlExportOptions { MaxDocumentElements = styleElements - 1 }));
            Assert.Equal("word-element-limit-exceeded", exception.Code);
            Assert.Equal("/word/styles.xml", exception.LimitSource);
            Assert.Equal(loadStylesFirst, stylesPart.IsRootElementLoaded);
        }
    }

    [Fact]
    public void StylesOfANewDocumentFollowLibraryStylesAddedToItsUnloadedStylesPart() {
        using WordDocument document = WordDocument.Create();
        StyleDefinitionsPart stylesPart = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!;
        int before = document.Styles.Count;
        WordTableStyle added = Enum.GetValues(typeof(WordTableStyle)).Cast<WordTableStyle>()
            .First(style => style != WordTableStyle.OfficeIMOModern && !document.Styles.Contains(WordTableStyles.GetStyle(style).Val!.Value!));
        string styleId = WordTableStyles.GetStyle(added).Val!.Value!;
        Assert.False(stylesPart.IsRootElementLoaded);

        document.AddTable(2, 2, added);

        // The library definition is written into the still-unloaded part; the catalog must read it again.
        Assert.False(stylesPart.IsRootElementLoaded);
        Assert.True(document.Styles.Contains(styleId));
        Assert.True(document.Styles.Count > before);
    }

    private sealed record StyleCatalogConversion(string Inspection, string Markdown, string Html, IReadOnlyList<byte[]> Pages, bool StylesLoaded);

    private static StyleCatalogConversion ConvertWithStyleCatalog(byte[] package, bool loadStylesFirst) {
        using var input = new MemoryStream(package, writable: false);
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        StyleDefinitionsPart stylesPart = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!;
        if (loadStylesFirst) _ = stylesPart.Styles;
        string inspection = string.Join("\n", document.Inspect(includeImageBytes: false).Sections
            .SelectMany(section => section.Elements.OfType<WordParagraphInfo>())
            .Select(paragraph => paragraph.StyleName + "|" + paragraph.ListMarker + "|" +
                string.Join(",", paragraph.Runs.Select(run => run.CapsStyle))));
        string markdown = document.ToMarkdown().ToMarkdown();
        string html = document.ToHtml();
        IReadOnlyList<byte[]> pages = document.ToImages(OfficeImageExportFormat.Png, new WordImageExportOptions { Scale = .25D })
            .Select(page => page.Bytes).ToArray();
        return new StyleCatalogConversion(inspection, markdown, html, pages, stylesPart.IsRootElementLoaded);
    }

    private static string[] DescribeCatalogStyles(byte[] package, IEnumerable<string> styleIds, bool loadStylesFirst, out bool stylesLoaded) {
        using var input = new MemoryStream(package, writable: false);
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        StyleDefinitionsPart stylesPart = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!;
        if (loadStylesFirst) _ = stylesPart.Styles;
        var lines = new List<string> { "count=" + document.Styles.Count };
        foreach (string styleId in styleIds) {
            WordStyle style = document.Styles[styleId];
            WordStyle? byName = document.Styles.Get(style.Name);
            lines.Add(string.Join("|", style.StyleId, style.Name, style.Type, style.BasedOnStyleId, style.NextStyleId,
                style.LinkedStyleId, style.IsDefault, style.IsCustom, style.IsPrimary, style.FontFamily, style.FontSize,
                style.Bold, style.Italic, style.ParagraphAlignment, style.OutlineLevel, document.Styles.Contains(style.StyleId),
                byName?.StyleId, CanonicalStyleXml(XElement.Parse(style.Element.OuterXml))));
        }
        stylesLoaded = stylesPart.IsRootElementLoaded;
        return lines.ToArray();
    }

    // Element names, attributes and children by namespace URI, ignoring prefixes and where namespaces are declared.
    private static string CanonicalStyleXml(XElement element) {
        var builder = new StringBuilder();
        Append(element);
        return builder.ToString();

        void Append(XElement current) {
            builder.Append('<').Append(current.Name);
            foreach (XAttribute attribute in current.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
                         .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal)) {
                builder.Append(' ').Append(attribute.Name).Append("=\"").Append(attribute.Value).Append('"');
            }
            builder.Append('>');
            foreach (XNode node in current.Nodes()) {
                if (node is XElement child) Append(child);
                else if (node is XText text) builder.Append(text.Value);
            }
            builder.Append("</>");
        }
    }
}
