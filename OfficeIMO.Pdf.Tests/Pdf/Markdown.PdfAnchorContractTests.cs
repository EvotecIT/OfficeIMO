using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class MarkdownPdfAnchorContractTests {
    [Theory]
    [InlineData("details")]
    [InlineData("footnote")]
    public void DirectlyRenderedSummaryAndFootnoteParagraphAnchorsAreRegistered(string kind) {
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(MarkdownDocument.Parse("**before <a id=\"direct\"></a> after**").Blocks));
        IMarkdownBlock target = kind == "details"
            ? new DetailsBlock(new SummaryBlock(paragraph.Inlines))
            : new FootnoteDefinitionBlock("note", new IMarkdownBlock[] { paragraph });
        var document = MarkdownDocument.Create().Add(new ParagraphBlock(new InlineSequence().Link("Jump", "#direct"))).Add(target);
        var result = document.ConvertToPdfWithEvidence();
        Assert.NotEmpty(result.Value.ToBytes());
        Assert.DoesNotContain(result.Report.Warnings, diagnostic => diagnostic.Code == "UnresolvedInternalLink");
    }

    [Theory]
    [InlineData("<a id=\">")]
    [InlineData("<a id='>")]
    public void IncompleteAnchorCarriersRetainPlainTextWithoutBreakingExport(string html) {
        var inlines = new InlineSequence { AutoSpacing = false };
        inlines.AddRaw(new HtmlRawInline(html));
        var result = MarkdownDocument.Create().Add(new ParagraphBlock(inlines)).ConvertToPdfWithEvidence();
        Assert.NotEmpty(result.Value.ToBytes());
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("table")]
    [InlineData("callout")]
    public void GenericBlockIdsSupportForwardLinks(string kind) {
        IMarkdownBlock block = kind switch {
            "table" => MarkdownDocument.Parse("| Header |\n| --- |\n| Cell |\n").Blocks[0],
            "callout" => new CalloutBlock("note", "Anchor marker", "Body"),
            _ => new ParagraphBlock(new InlineSequence().Text("Anchor marker"))
        };
        ((MarkdownObject)block).SetAttributes(MarkdownAttributeSet.Create(elementId: "destination"));
        var document = MarkdownDocument.Create().Add(new ParagraphBlock(new InlineSequence().Link("Jump", "#destination"))).Add(block);
        var result = document.ConvertToPdfWithEvidence();
        Assert.NotEmpty(result.Value.ToBytes());
        Assert.DoesNotContain(result.Report.Warnings, diagnostic => diagnostic.Code == "UnresolvedInternalLink");
    }

    [Fact]
    public void StyledMissingLinksInParagraphListAndTableBecomeVisibleTextWithOneDiagnostic() {
        var document = MarkdownDocument.Parse("[**visible**](#missing)\n\n- [list](#missing)\n\n| H |\n| --- |\n| [cell](#missing) |\n");
        var result = document.ConvertToPdfWithEvidence();
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(result.Value.ToBytes()).ExtractText();
        Assert.Contains("visible", text, StringComparison.Ordinal);
        Assert.Contains("list", text, StringComparison.Ordinal);
        Assert.Contains("cell", text, StringComparison.Ordinal);
        Assert.Single(result.Report.Warnings, diagnostic => diagnostic.Code == "UnresolvedInternalLink");
    }

    [Fact]
    public void InlineAnchorsInsideStyledPanelContentResolveWithoutLeakingAcrossExports() {
        var document = MarkdownDocument.Parse("> [!NOTE] Panel\n> **before <a id=\"inside\"></a> after**\n\n[Jump](#inside)");
        var options = new MarkdownPdfExportOptions();
        var first = document.ConvertToPdfWithEvidence(options);
        Assert.NotEmpty(first.Value.ToBytes());
        Assert.DoesNotContain(first.Report.Warnings, diagnostic => diagnostic.Code == "UnresolvedInternalLink");
        var second = MarkdownDocument.Parse("[Unresolved](#inside)").ConvertToPdfWithEvidence(options);
        Assert.NotEmpty(second.Value.ToBytes());
        Assert.Contains(second.Report.Warnings, diagnostic => diagnostic.Code == "UnresolvedInternalLink");
    }
}
