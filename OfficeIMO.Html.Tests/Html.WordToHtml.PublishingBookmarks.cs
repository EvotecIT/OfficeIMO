using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordToHtml {
    [Fact]
    public void PublishingRetainsAdditionalParagraphBookmarksWithoutDuplicatingThePrimaryId() {
        using var document = WordDocument.Create();
        var target = document.AddParagraph("Target text");
        target.AddBookmark("primary");
        target._paragraph.Append(new BookmarkStart { Id = "42", Name = "secondary" }, new BookmarkEnd { Id = "42" });
        document.AddParagraph().AddHyperLink("Next", "secondary");
        var result = document.ToHtmlResult(WordToHtmlOptions.CreateSemanticDocumentProfile());
        var html = OfficeIMO.Html.HtmlConversionDocument.Parse(result.Value).Document;
        Assert.Single(html.QuerySelectorAll("#primary"));
        Assert.Single(html.QuerySelectorAll("#secondary"));
        Assert.Equal("#secondary", html.QuerySelector("a")!.GetAttribute("href"));
        Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "BookmarkPositionProjected" && item.LossKind == OfficeConversionLossKind.Approximation);
    }
    [Fact]
    public void WordProducedTableAndTextBoxBookmarksRemainReachableForPublishing() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Word", "PremiumGaps", "FieldEvaluation", "word-generated-toc-table-text-box.docx");
        using var document = WordDocument.Load(path);
        var result = document.ToHtmlResult(WordToHtmlOptions.CreateSemanticDocumentProfile());
        var html = OfficeIMO.Html.HtmlConversionDocument.Parse(result.Value).Document;
        Assert.NotNull(html.QuerySelector("#_Toc233703443"));
        Assert.NotNull(html.QuerySelector("#_Toc233703444"));
        Assert.Contains("Word-generated table text-box TOC detail", html.Body!.TextContent);
        Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "BookmarkPositionProjected");
    }
}
