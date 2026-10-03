using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class MarkdownPdfAnchorContractTests {
    [Theory]
    [InlineData("<a id=\">")]
    [InlineData("<a id='>")]
    public void IncompleteAnchorCarriersRetainPlainTextWithoutBreakingExport(string html) {
        var inlines = new InlineSequence { AutoSpacing = false };
        inlines.AddRaw(new HtmlRawInline(html));
        var result = MarkdownDoc.Create().Add(new ParagraphBlock(inlines)).ToPdfDocumentResult();
        Assert.NotEmpty(result.Value.ToBytes());
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("table")]
    [InlineData("callout")]
    public void GenericBlockIdsSupportForwardLinks(string kind) {
        IMarkdownBlock block = kind switch {
            "table" => MarkdownReader.Parse("| Header |\n| --- |\n| Cell |\n").Blocks[0],
            "callout" => new CalloutBlock("note", "Anchor marker", "Body"),
            _ => new ParagraphBlock(new InlineSequence().Text("Anchor marker"))
        };
        ((MarkdownObject)block).SetAttributes(MarkdownAttributeSet.Create(elementId: "destination"));
        var document = MarkdownDoc.Create().Add(new ParagraphBlock(new InlineSequence().Link("Jump", "#destination"))).Add(block);
        var result = document.ToPdfDocumentResult();
        Assert.NotEmpty(result.Value.ToBytes());
        Assert.DoesNotContain(result.Report.Warnings, diagnostic => diagnostic.Code == "UnresolvedInternalLink");
    }

    [Fact]
    public void StyledMissingLinksInParagraphListAndTableBecomeVisibleTextWithOneDiagnostic() {
        var document = MarkdownReader.Parse("[**visible**](#missing)\n\n- [list](#missing)\n\n| H |\n| --- |\n| [cell](#missing) |\n");
        var result = document.ToPdfDocumentResult();
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(result.Value.ToBytes()).ExtractText();
        Assert.Contains("visible", text, StringComparison.Ordinal);
        Assert.Contains("list", text, StringComparison.Ordinal);
        Assert.Contains("cell", text, StringComparison.Ordinal);
        Assert.Single(result.Report.Warnings, diagnostic => diagnostic.Code == "UnresolvedInternalLink");
    }

    [Fact]
    public void InlineAnchorsInsideStyledPanelContentResolveWithoutLeakingAcrossExports() {
        var document = MarkdownReader.Parse("> [!NOTE] Panel\n> **before <a id=\"inside\"></a> after**\n\n[Jump](#inside)");
        var options = new MarkdownToPdfOptions();
        var first = document.ToPdfDocumentResult(options);
        Assert.NotEmpty(first.Value.ToBytes());
        Assert.DoesNotContain(first.Report.Warnings, diagnostic => diagnostic.Code == "UnresolvedInternalLink");
        var second = MarkdownReader.Parse("[Unresolved](#inside)").ToPdfDocumentResult(options);
        Assert.NotEmpty(second.Value.ToBytes());
        Assert.Contains(second.Report.Warnings, diagnostic => diagnostic.Code == "UnresolvedInternalLink");
    }
}
