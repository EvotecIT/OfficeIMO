using System.Reflection;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownHtmlApiContracts {
    [Fact]
    public void HtmlRenderingRequiresAnExplicitFragmentOrDocumentShape() {
        string[] names = typeof(MarkdownDoc)
            .GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
            .Select(static method => method.Name)
            .ToArray();

        Assert.DoesNotContain("ToHtml", names);
        Assert.Contains("ToHtmlFragment", names);
        Assert.Contains("ToHtmlDocument", names);
        Assert.Contains("ToHtmlParts", names);
    }
}
