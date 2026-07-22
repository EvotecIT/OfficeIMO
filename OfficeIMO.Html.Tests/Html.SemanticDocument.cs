using OfficeIMO.Html;
using OfficeIMO.OneNote.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void SemanticDocument_InterpretsRichStructureStylesResourcesAndSourceLocationsOnce() {
        const string html = """
            <!doctype html>
            <html lang="en"><head><title>Semantic report</title><meta name="author" content="OfficeIMO">
            <style>.accent { font-weight: 700; text-decoration: underline; }</style></head><body>
            <main><section id="summary"><h1>Summary</h1>
            <p>Plain <a href="https://example.test"><span class="accent">linked</span></a> text.</p>
            <ol><li>First<ul><li>Nested</li></ul></li><li>Second</li></ol>
            <table aria-label="Metrics"><tr><th rowspan="2">Metric</th><th>Value</th></tr><tr><td>42</td></tr></table>
            <img src="data:image/png;base64,iVBORw0KGgo=" alt="Chart">
            </section></main></body></html>
            """;

        HtmlConversionDocument conversion = HtmlConversionDocument.Parse(html);
        HtmlSemanticDocument semantic = conversion.SemanticDocument;

        Assert.Same(semantic, conversion.SemanticDocument);
        Assert.Equal("Semantic report", semantic.Title);
        Assert.Equal("en", semantic.Language);
        Assert.Equal("OfficeIMO", semantic.Metadata["author"]);
        HtmlSemanticSection section = Assert.Single(semantic.Sections);
        Assert.Equal("Summary", section.Title);
        Assert.All(section.Blocks, block => Assert.NotNull(block.SourceLocation));
        Assert.Contains(section.Blocks, block => block.SourceLocation!.Line > 0);

        HtmlSemanticBlock paragraph = Assert.Single(section.Blocks, block => block.Kind == HtmlSemanticBlockKind.Paragraph);
        HtmlSemanticRun linked = Assert.Single(paragraph.Runs, run => run.Text.Contains("linked", StringComparison.Ordinal));
        Assert.True(linked.Bold);
        Assert.True(linked.Underline);
        Assert.Equal("https://example.test", linked.Hyperlink);

        HtmlSemanticBlock list = Assert.Single(section.Blocks, block => block.Kind == HtmlSemanticBlockKind.List);
        Assert.True(list.Ordered);
        Assert.Equal(2, list.Children.Count);
        Assert.Contains(list.Children[0].Children, block => block.Kind == HtmlSemanticBlockKind.List);

        HtmlSemanticBlock table = Assert.Single(semantic.RootTables);
        Assert.Equal("Metrics", table.Table!.Caption);
        Assert.Equal(2, table.Table.Rows.Count);
        Assert.Equal(2, table.Table.Rows[0].Cells[0].RowSpan);
        Assert.True(table.Table.Rows[0].Cells[0].IsHeader);

        HtmlSemanticResource resource = Assert.Single(semantic.Resources);
        Assert.Equal(HtmlResourceKind.Image, resource.Kind);
        Assert.Equal("Chart", resource.AlternateText);
        Assert.Equal("image/png", resource.MediaType);
    }

    [Fact]
    public void OneNoteGenericImport_ConsumesSemanticRichRunsAndNestedLists() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse("""
            <h1>Notes</h1>
            <p>Normal <strong>bold</strong> <a href="https://example.test">link</a></p>
            <ul><li>Parent<ol><li>Child</li></ol></li></ul>
            """);

        var result = source.ToOneNoteSectionResult();
        var page = Assert.Single(result.RequireValue().Pages);
        var outline = Assert.Single(page.Outlines);
        var paragraphs = outline.Children.OfType<OfficeIMO.OneNote.OneNoteParagraph>().ToList();
        Assert.Contains(paragraphs.SelectMany(paragraph => paragraph.Runs), run => run.Text == "bold" && run.Style.Bold == true);
        Assert.Contains(paragraphs.SelectMany(paragraph => paragraph.Runs), run => run.Text == "link" && run.Hyperlink == "https://example.test");
        Assert.Contains(paragraphs, paragraph => paragraph.List?.Level == 0);
        Assert.Contains(paragraphs, paragraph => paragraph.List?.Level == 1);
    }

    [Fact]
    public void AnalyzeFor_PredictsTargetLossWithSourceAndTargetProvenanceBeforeCreation() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse("""
            <style>@page { size: A4; }</style>
            <h1>Preflight</h1>
            <p><strong>Rich</strong> <a href="https://example.test">link</a></p>
            <video src="https://example.test/demo.mp4"></video>
            <form><input name="approved" type="checkbox" checked></form>
            <div data-officeimo-chart="sales"></div>
            """);

        HtmlConversionPreflight preflight = source.AnalyzeFor(HtmlConversionTarget.Markdown);

        Assert.Same(preflight, source.AnalyzeFor(HtmlConversionTarget.Markdown));
        Assert.Equal(HtmlConversionPreflightOutcome.Approximated, preflight.Get(HtmlSemanticFeature.Media).Outcome);
        Assert.True(preflight.Get(HtmlSemanticFeature.Media).IsPresent);
        Assert.Equal(HtmlConversionPreflightOutcome.Approximated, preflight.Get(HtmlSemanticFeature.Forms).Outcome);
        Assert.Equal(HtmlConversionPreflightOutcome.Omitted, preflight.Get(HtmlSemanticFeature.Charts).Outcome);
        Assert.Equal(HtmlConversionPreflightOutcome.Omitted, preflight.Get(HtmlSemanticFeature.PagedLayout).Outcome);
        Assert.True(preflight.HasPotentialLoss);
        Assert.Contains(preflight.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.ContentOmitted);
        Assert.All(preflight.Diagnostics, diagnostic => {
            Assert.False(string.IsNullOrWhiteSpace(diagnostic.Provenance.SourceAddress));
            Assert.Equal("preflight:markdown", diagnostic.Provenance.TargetAddress);
        });
    }

    [Fact]
    public void HtmlDiagnostics_AlwaysCarryAtLeastDocumentToComponentProvenance() {
        var diagnostic = new HtmlDiagnostic("OfficeIMO.Html.Test", "Example", "Example warning");
        Assert.Equal("html:document", diagnostic.Provenance.SourceAddress);
        Assert.Equal("OfficeIMO.Html.Test", diagnostic.Provenance.TargetAddress);
    }

    [Fact]
    public void HtmlDiagnostics_PreserveTheOriginalPublicClrSignatures() {
        Type[] originalParameters = {
            typeof(string), typeof(string), typeof(string), typeof(HtmlDiagnosticSeverity),
            typeof(string), typeof(string), typeof(HtmlConversionLossKind)
        };

        Assert.NotNull(typeof(HtmlDiagnostic).GetConstructor(originalParameters));
        Assert.NotNull(typeof(HtmlDiagnosticReport).GetMethod(nameof(HtmlDiagnosticReport.Add), originalParameters));
    }

    [Fact]
    public void AnalyzeFor_UsesDomEvidenceWithoutTextOrScriptFalsePositivesAndReportsExactLocation() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse("""
            <script>const sample = "data-officeimo-chart @page &lt;ins&gt;";</script>
            <p>Literal data-officeimo-formula and page-break-before text</p>
            <section class="officeimo-comments"><ol><li id="review-comment">Real comment</li></ol></section>
            <p id="page-start" style="break-before: page">Paged</p>
            """);

        HtmlConversionPreflight preflight = source.AnalyzeFor(HtmlConversionTarget.Markdown);

        Assert.False(preflight.Get(HtmlSemanticFeature.Formulas).IsPresent);
        Assert.False(preflight.Get(HtmlSemanticFeature.Charts).IsPresent);
        Assert.False(preflight.Get(HtmlSemanticFeature.Annotations).IsPresent);
        Assert.Equal(1, preflight.Get(HtmlSemanticFeature.Comments).OccurrenceCount);
        Assert.Contains("#review-comment", preflight.Get(HtmlSemanticFeature.Comments).FirstSourceLocation!.Selector, StringComparison.Ordinal);
        Assert.Equal(1, preflight.Get(HtmlSemanticFeature.PagedLayout).OccurrenceCount);
        Assert.Contains("#page-start", preflight.Get(HtmlSemanticFeature.PagedLayout).FirstSourceLocation!.Selector, StringComparison.Ordinal);
    }

    [Fact]
    public void SemanticDocument_RetainsInlineAndTableCellImagesInTheCanonicalIr() {
        HtmlSemanticDocument semantic = HtmlConversionDocument.Parse("""
            <p id="intro">Before <img src="data:image/png;base64,AA==" alt="inline" width="20"> after</p>
            <table><tr><td id="evidence">Cell <img src="data:image/png;base64,AQ==" alt="cell" height="30"></td></tr></table>
            """).SemanticDocument;

        HtmlSemanticBlock paragraph = Assert.Single(semantic.Sections.SelectMany(section => section.Blocks),
            block => block.Kind == HtmlSemanticBlockKind.Paragraph);
        HtmlSemanticResource inline = Assert.Single(paragraph.InlineResources);
        HtmlSemanticResource cell = Assert.Single(Assert.Single(semantic.RootTables).Table!.Rows[0].Cells[0].Resources);

        Assert.Equal("inline", inline.AlternateText);
        Assert.Equal(20D, inline.WidthPixels);
        Assert.Contains("#intro", inline.SourceLocation!.Selector, StringComparison.Ordinal);
        Assert.Equal("cell", cell.AlternateText);
        Assert.Equal(30D, cell.HeightPixels);
        Assert.Equal(2, semantic.Resources.Count);
        Assert.Equal(2, HtmlConversionDocument.Parse("<p><img src='data:image/png;base64,AA=='></p><table><tr><td><img src='data:image/png;base64,AQ=='></td></tr></table>")
            .AnalyzeFor(HtmlConversionTarget.Excel).Get(HtmlSemanticFeature.Images).OccurrenceCount);
    }

    [Fact]
    public void SemanticRuns_NormalizeHtmlWhitespaceAcrossStyleBoundariesAndPreservePreformattedText() {
        HtmlSemanticDocument semantic = HtmlConversionDocument.Parse(
            "<p>  Hello <strong>   brave </strong>\n world  </p><pre>  a\n b  </pre>").SemanticDocument;
        HtmlSemanticBlock paragraph = Assert.Single(semantic.Sections.SelectMany(section => section.Blocks),
            block => block.Kind == HtmlSemanticBlockKind.Paragraph);
        HtmlSemanticBlock pre = Assert.Single(semantic.Sections.SelectMany(section => section.Blocks),
            block => block.Kind == HtmlSemanticBlockKind.Code);

        Assert.Equal("Hello brave world", paragraph.Text);
        Assert.Equal(paragraph.Text, string.Concat(paragraph.Runs.Select(run => run.Text)));
        Assert.Equal("  a\n b  ", pre.Text);
        Assert.Equal(pre.Text, string.Concat(pre.Runs.Select(run => run.Text)));
    }
}
