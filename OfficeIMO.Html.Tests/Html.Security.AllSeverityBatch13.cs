using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlAllSeverityBatch13SecurityTests {
    [Fact]
    public void HtmlToWordRejectsImageTraversalOutsideBasePath() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-html-image-" + Guid.NewGuid().ToString("N"));
        string basePath = Path.Combine(root, "content");
        string outside = Path.Combine(root, "outside.png");
        Directory.CreateDirectory(basePath);
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png"), outside);
        try {
            var options = HtmlToWordOptions.CreateTrustedDocumentProfile();
            options.BasePath = basePath;
            using WordDocument document = OfficeIMO.Html.HtmlConversionDocument
                .Parse("<img src='../outside.png' alt='blocked'>")
                .ToWordDocument(options);

            Assert.Empty(document.Images);
        } finally {
            Directory.Delete(root, true);
        }
    }

    [Fact]
    public void WordToHtmlNormalizesUntrustedStyleIdentifiersInClassesAndCss() {
        const string hostileStyle = "x} body { background:url(javascript:alert(1))";
        using WordDocument document = WordDocument.Create();
        StyleDefinitionsPart stylePart = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart
            ?? document._wordprocessingDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        stylePart.Styles ??= new Styles();
        stylePart.Styles.Append(new Style(
            new StyleRunProperties(
                new RunFonts { Ascii = "Evil</style><style>body{background:url(https://attacker.test/)}" },
                new Color { Val = "00ff00; background:url(https://attacker.test/)" })) {
            Type = StyleValues.Paragraph,
            StyleId = hostileStyle
        });
        document.AddParagraph("safe").SetStyleId(hostileStyle);
        document.AddParagraph().AddText("run-safe").SetCharacterStyleId(hostileStyle);

        string html = document.ToHtml(new WordToHtmlOptions {
            IncludeParagraphClasses = true,
            IncludeRunClasses = true
        });

        Assert.DoesNotContain(hostileStyle, html, StringComparison.Ordinal);
        Assert.DoesNotContain("attacker.test", html, StringComparison.Ordinal);
        Assert.DoesNotContain("</style><style>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("word-style-", html, StringComparison.Ordinal);
        Assert.Contains("safe", html, StringComparison.Ordinal);
        Assert.Contains("run-safe", html, StringComparison.Ordinal);
    }

    [Fact]
    public void WordToHtmlTraversesEachNestedTableOnceAndEnforcesDepth() {
        using WordDocument document = WordDocument.Create();
        WordTable outer = document.AddTable(1, 1);
        WordTable middle = outer.Rows[0].Cells[0].AddTable(1, 1);
        middle.Rows[0].Cells[0].AddTable(1, 1);

        string html = document.ToHtml();
        Assert.Equal(3, CountOccurrences(html, "<table"));
        Assert.Throws<InvalidDataException>(() =>
            document.ToHtml(new WordToHtmlOptions { MaxTableNestingDepth = 2 }));
    }

    [Fact]
    public void WordToHtmlRejectsHostileListLevelBeforeGrowingStacks() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListStyle.Bulleted);
        list.AddItem("item");
        WordParagraph paragraph = document.Paragraphs.Single(item => item.IsListItem);
        paragraph.ListItemLevel = int.MaxValue;

        Assert.Throws<InvalidDataException>(() =>
            document.ToHtml(new WordToHtmlOptions { MaxListNestingDepth = 8 }));
    }

    [Fact]
    public void HtmlStylesheetCacheIsScopedToOneConverterInstance() {
        Type converter = typeof(HtmlToWordOptions).Assembly.GetType("OfficeIMO.Word.Html.HtmlToWordConverter", throwOnError: true)!;
        FieldInfo cache = converter.GetField("_stylesheetCache", BindingFlags.Instance | BindingFlags.NonPublic)!;

        Assert.False(cache.IsStatic);
    }

    private static int CountOccurrences(string value, string token) {
        int count = 0;
        int start = 0;
        while ((start = value.IndexOf(token, start, StringComparison.OrdinalIgnoreCase)) >= 0) {
            count++;
            start += token.Length;
        }
        return count;
    }
}
