using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlAccessibleNameTests {
    [Fact]
    public void LogicalNodeKind_AppendsSemanticKindsWithoutRenumberingExistingValues() {
        Assert.Equal(19, (int)HtmlLogicalNodeKind.TableCaption);
        Assert.Equal(20, (int)HtmlLogicalNodeKind.Code);
        Assert.Equal(21, (int)HtmlLogicalNodeKind.Quote);
        Assert.Equal(22, (int)HtmlLogicalNodeKind.Footnote);
    }

    [Fact]
    public void AccessibilitySemantics_ResolvesAriaAndHostLanguageNamesInPriorityOrder() {
        var document = HtmlDocumentParser.ParseDocument("""
<span id="chapter">Chapter</span><span id="number">four</span>
<a id="link" href="#target" aria-labelledby="chapter number"></a>
<img id="aria-image" src="cover.png" alt="Cover" aria-label="Accessible cover">
<img id="decorative" src="rule.png" alt="" title="Decorative rule">
""");

        Assert.Equal(
            "Chapter four",
            HtmlAccessibilitySemantics.GetAccessibleName(document.GetElementById("link")!, includeTextFallback: true));
        Assert.Equal(
            "Accessible cover",
            HtmlAccessibilitySemantics.GetAccessibleName(document.GetElementById("aria-image")!));
        Assert.Equal(
            string.Empty,
            HtmlAccessibilitySemantics.GetAccessibleName(document.GetElementById("decorative")!));

        var customImage = HtmlDocumentParser.ParseDocument("<media-image alt=\"Aliased image\"></media-image>");
        Assert.Equal(
            "Aliased image",
            HtmlAccessibilitySemantics.GetImageAccessibleName(customImage.Body!.FirstElementChild!));

        var svgImage = HtmlDocumentParser.ParseDocument("<svg><title>Vector diagram</title></svg>");
        Assert.Equal(
            "Vector diagram",
            HtmlAccessibilitySemantics.GetImageAccessibleName(svgImage.Body!.FirstElementChild!));
    }

    [Theory]
    [InlineData("<div role=\"heading\" aria-level=\"4\">Heading</div>", 4)]
    [InlineData("<span role=\"heading\">Heading</span>", 2)]
    [InlineData("<div role=\"heading\" aria-level=\"12\">Heading</div>", 6)]
    [InlineData("<h3>Heading</h3>", 3)]
    public void AccessibilitySemantics_ResolvesNativeAndAriaHeadingLevels(string html, int expectedLevel) {
        var document = HtmlDocumentParser.ParseDocument(html);

        Assert.True(HtmlAccessibilitySemantics.TryGetHeadingLevel(document.Body!.FirstElementChild!, out int level));
        Assert.Equal(expectedLevel, level);
    }

    [Fact]
    public void LogicalDocument_ProjectsAccessibleNamesRolesAndCapabilities() {
        HtmlLogicalDocument logical = HtmlLogicalDocumentBuilder.FromHtml("""
<main>
  <div role="heading" aria-level="3">Accessible section</div>
  <a href="#target" aria-label="Read the note"></a>
  <blockquote><p>Quoted text</p></blockquote>
  <pre data-language="csharp">Console.WriteLine(1);</pre>
  <aside id="note" epub:type="footnote" role="doc-footnote"><p>Note text</p><a epub:type="backlink" href="#target">return</a></aside>
</main>
""");

        HtmlLogicalNode heading = Find(logical.Root, HtmlLogicalNodeKind.Heading);
        HtmlLogicalNode link = Find(logical.Root, HtmlLogicalNodeKind.Link);
        HtmlLogicalNode quote = Find(logical.Root, HtmlLogicalNodeKind.Quote);
        HtmlLogicalNode code = Find(logical.Root, HtmlLogicalNodeKind.Code);
        HtmlLogicalNode footnote = Find(logical.Root, HtmlLogicalNodeKind.Footnote);

        Assert.Equal("Accessible section", heading.Text);
        Assert.Equal("Read the note", link.AccessibleName);
        Assert.Equal("Quoted text", quote.Text);
        Assert.Equal("Console.WriteLine(1);", code.Text);
        Assert.Equal("Note text", footnote.Text);
        Assert.Contains("accessibility", logical.Capabilities);
        Assert.Contains("footnotes", logical.Capabilities);
        Assert.Contains("quotes", logical.Capabilities);
        Assert.Contains("code", logical.Capabilities);
    }

    private static HtmlLogicalNode Find(HtmlLogicalNode node, HtmlLogicalNodeKind kind) {
        if (node.Kind == kind) return node;
        foreach (HtmlLogicalNode child in node.Children) {
            HtmlLogicalNode? found = FindOrNull(child, kind);
            if (found != null) return found;
        }
        throw new InvalidOperationException("Logical node was not found: " + kind);
    }

    private static HtmlLogicalNode? FindOrNull(HtmlLogicalNode node, HtmlLogicalNodeKind kind) {
        if (node.Kind == kind) return node;
        foreach (HtmlLogicalNode child in node.Children) {
            HtmlLogicalNode? found = FindOrNull(child, kind);
            if (found != null) return found;
        }
        return null;
    }
}
