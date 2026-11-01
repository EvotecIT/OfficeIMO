using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Tests;

public sealed class HtmlWordToHtmlLinkSafetyTests {
    private const string UnsafeLinkDiagnostic = "unsafe-link-dropped";

    [Theory]
    [InlineData("javascript:alert(1)")]
    [InlineData("vbscript:msgbox(1)")]
    [InlineData("data:text/html,<script>alert(1)</script>")]
    public void UnsafeHyperlinkKeepsItsTextWithoutTheLink(string url) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph().AddHyperLink("Click", new Uri(url));
        var options = new WordHtmlExportOptions();

        HtmlTextConversionResult result = document.ConvertToHtmlWithEvidence(options);
        string html = result.RequireValue();

        Assert.Contains("Click", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<a ", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == UnsafeLinkDiagnostic);
    }

    [Fact]
    public void WebHyperlinkIsKept() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph().AddHyperLink("Site", new Uri("https://example.com/page"));
        var options = new WordHtmlExportOptions();

        HtmlTextConversionResult result = document.ConvertToHtmlWithEvidence(options);
        string html = result.RequireValue();

        Assert.Contains("href=\"https://example.com/page\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == UnsafeLinkDiagnostic);
    }

    [Fact]
    public void UnsafeHyperlinkAroundAnEquationKeepsTheEquationWithoutTheLink() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        string relationship = document._wordprocessingDocument!.MainDocumentPart!
            .AddHyperlinkRelationship(new Uri("javascript:alert(1)"), isExternal: true).Id;
        paragraph._paragraph.Append(new Hyperlink(new M.OfficeMath(new M.Run(new M.Text("linked-only")))) { Id = relationship });
        var options = new WordHtmlExportOptions();

        HtmlTextConversionResult result = document.ConvertToHtmlWithEvidence(options);
        string html = result.RequireValue();

        Assert.Contains("linked-only", html, StringComparison.Ordinal);
        Assert.DoesNotContain("javascript:", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == UnsafeLinkDiagnostic);
    }

    [Fact]
    public void UnsafeAdditionalLinkTagIsDropped() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        var options = new WordHtmlExportOptions();
        options.AdditionalLinkTags.Add(("stylesheet", "javascript:alert(1)"));
        options.AdditionalLinkTags.Add(("stylesheet", "https://example.com/site.css"));

        HtmlTextConversionResult result = document.ConvertToHtmlWithEvidence(options);
        string html = result.RequireValue();

        Assert.DoesNotContain("javascript:", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("https://example.com/site.css", html, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == UnsafeLinkDiagnostic);
    }

    [Fact]
    public void EveryTableOfContentsListsItsHeadings() {
        using WordDocument document = WordDocument.Create();
        document.AddTableOfContent();
        document.AddParagraph("First").Style = WordParagraphStyles.Heading1;
        document.AddParagraph("Second").Style = WordParagraphStyles.Heading1;
        document.AddTableOfContent();

        string html = document.ToHtml();

        Assert.Equal(2, CountOccurrences(html, ">First</a>"));
        Assert.Equal(2, CountOccurrences(html, ">Second</a>"));
    }

    private static int CountOccurrences(string text, string value) =>
        (text.Length - text.Replace(value, string.Empty, StringComparison.Ordinal).Length) / value.Length;
}
