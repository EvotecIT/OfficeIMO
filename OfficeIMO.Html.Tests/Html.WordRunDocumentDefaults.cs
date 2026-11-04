using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlWordRunDocumentDefaultsTests {
    [Theory]
    [InlineData(false, false, "#ff0000")]
    [InlineData(true, false, "#ffff00")]
    [InlineData(true, true, "#ff0000")]
    public void WordToHtmlUsesDocumentDefaultBackgroundAndPreservesHighlightNone(bool highlight, bool reset, string highlightColor) {
        using WordDocument document = WordDocument.Create();
        var defaults = new RunPropertiesBaseStyle(new Shading { Val = ShadingPatternValues.Clear, Fill = "FF0000" });
        if (highlight) defaults.Append(new Highlight { Val = HighlightColorValues.Yellow });
        document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.DocDefaults =
            new DocDefaults(new RunPropertiesDefault(defaults));
        WordParagraph paragraph = document.AddParagraph("BACKGROUND");
        if (reset) paragraph.Highlight = WordHighlightColor.None;

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string html = loaded.ToHtml(new WordHtmlExportOptions { IncludeRunHighlightStyles = true });

        Assert.Contains("BACKGROUND", html, StringComparison.Ordinal);
        Assert.Contains("background-color:" + highlightColor, html, StringComparison.OrdinalIgnoreCase);
    }
}
