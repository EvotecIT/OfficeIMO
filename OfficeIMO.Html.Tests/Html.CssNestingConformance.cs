using System.Text.Json;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlCssNestingConformanceTests {
    [Theory]
    [InlineData(".x &", "<div class='a'><div class='x'><p class='b'>text</p></div></div>")]
    [InlineData("& + &", "<div class='a'><p class='b'>first</p><p class='b'>second</p></div>")]
    public void ComplexParentSelectorsRemainOneUnitWhenSubstitutingNesting(string nested, string html) {
        var document = HtmlConversionDocument.Parse("<style>.a .b { color: blue; " + nested + " { color: red; } }</style>" + html);
        var target = document.Document.QuerySelector(".b:last-child")!;
        Assert.Equal("rgba(255, 0, 0, 1)", HtmlComputedStyleEngine.Compute(document)[target].GetValue("color"));
    }

    [Fact]
    public void FileBackedNestingCorpusMatchesTheDeclaredComputedValues() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Html", "Css", "css-nesting-corpus.json");
        CssNestingCorpusCase[] corpus = JsonSerializer.Deserialize<CssNestingCorpusCase[]>(
            File.ReadAllText(path), new JsonSerializerOptions { PropertyNameCaseInsensitive = true })!;
        Assert.Equal(10, corpus.Length);

        foreach (CssNestingCorpusCase item in corpus) {
            HtmlConversionDocument document = HtmlConversionDocument.Parse(
                "<style>" + item.Css + "</style>" + item.Html);
            OfficeIMO.Html.Dom.HtmlElement target = document.Document.QuerySelector(item.TargetSelector)
                ?? throw new InvalidOperationException(item.Name + " did not produce its target element.");

            string actual = HtmlComputedStyleEngine.Compute(document)[target].GetValue(item.Property);

            Assert.True(string.Equals(item.Expected, actual, StringComparison.Ordinal),
                item.Name + " expected " + item.Property + "=" + item.Expected + " but received " + actual + ".");
        }
    }

    private sealed class CssNestingCorpusCase {
        public string Name { get; set; } = string.Empty;
        public string Css { get; set; } = string.Empty;
        public string Html { get; set; } = string.Empty;
        public string TargetSelector { get; set; } = string.Empty;
        public string Property { get; set; } = string.Empty;
        public string Expected { get; set; } = string.Empty;
    }
}
