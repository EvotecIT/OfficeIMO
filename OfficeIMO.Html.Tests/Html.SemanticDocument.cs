using OfficeIMO.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void SemanticDocument_PreservesHeadingOnlySectionTitles() {
        HtmlSemanticSection section = Assert.Single(
            HtmlConversionDocument.Parse("<h1>Only title</h1>").SemanticDocument.Sections);

        Assert.Equal("Only title", section.Title);
        Assert.Empty(section.Blocks);
    }

    [Fact]
    public void SemanticDocument_NormalizesFormControlStateUsingHtmlRules() {
        HtmlSemanticBlock form = Assert.Single(HtmlConversionDocument.Parse("""
            <form id="settings" action="/save" method="post" enctype="multipart/form-data" novalidate>
              <input name="text" required checked multiple min="1" step="2" pattern="[A-Z]+" minlength="2" maxlength="8" placeholder="Code">
              <input name="check" type="checkbox" checked readonly pattern="ignored" minlength="4" placeholder="ignored">
              <button name="save" required readonly checked multiple formaction="/draft" formmethod="get" formenctype="text/plain" formtarget="preview" formnovalidate>Save</button>
              <select name="choice"><option>One</option><option selected>Two</option></select>
              <select name="many" multiple><option selected>A</option><option selected value="b">B</option></select>
              <fieldset disabled><legend><input name="legend"></legend><input name="disabled"></fieldset>
            </form>
            """).SemanticDocument.Sections.SelectMany(section => section.Blocks));
        Dictionary<string, HtmlSemanticFormControl> controls = form.Children
            .Select(block => block.FormControl!)
            .ToDictionary(control => control.Name, StringComparer.Ordinal);

        Assert.Equal("/save", form.Form!.Action);
        Assert.Equal("post", form.Form.Method);
        Assert.Equal("multipart/form-data", form.Form.EncodingType);
        Assert.True(form.Form.NoValidate);
        Assert.Equal("text", controls["text"].Type);
        Assert.True(controls["text"].IsRequired);
        Assert.Equal("[A-Z]+", controls["text"].Pattern);
        Assert.Equal(2, controls["text"].MinimumLength);
        Assert.Equal(8, controls["text"].MaximumLength);
        Assert.Equal("Code", controls["text"].Placeholder);
        Assert.False(controls["text"].IsChecked);
        Assert.False(controls["text"].IsMultiple);
        Assert.Equal(string.Empty, controls["text"].Minimum);
        Assert.Equal(string.Empty, controls["text"].Step);
        Assert.Equal("on", controls["check"].Value);
        Assert.True(controls["check"].IsChecked);
        Assert.False(controls["check"].IsReadOnly);
        Assert.Equal(string.Empty, controls["check"].Pattern);
        Assert.Null(controls["check"].MinimumLength);
        Assert.Equal(string.Empty, controls["check"].Placeholder);
        Assert.Equal("submit", controls["save"].Type);
        Assert.False(controls["save"].IsRequired);
        Assert.False(controls["save"].IsReadOnly);
        Assert.False(controls["save"].IsChecked);
        Assert.False(controls["save"].IsMultiple);
        Assert.Equal("/draft", controls["save"].FormAction);
        Assert.Equal("get", controls["save"].FormMethod);
        Assert.Equal("text/plain", controls["save"].FormEncodingType);
        Assert.Equal("preview", controls["save"].FormTarget);
        Assert.True(controls["save"].FormNoValidate);
        Assert.Equal(new[] { "Two" }, controls["choice"].Values);
        Assert.Equal(new[] { "A", "b" }, controls["many"].Values);
        Assert.True(controls["many"].IsMultiple);
        Assert.False(controls["legend"].IsDisabled);
        Assert.True(controls["disabled"].IsDisabled);
        Assert.All(controls.Values, control => Assert.Equal("settings", control.FormOwnerId));
    }

    [Fact]
    public void SemanticDocument_ExplicitFormOwnerUsesFirstMatchingIdElement() {
        HtmlSemanticFormControl control = Assert.Single(HtmlConversionDocument.Parse("""
            <div id="owner"></div><form id="owner"></form><input name="field" form="owner">
            """).SemanticDocument.Sections.SelectMany(section => section.Blocks),
            block => block.FormControl?.Name == "field").FormControl!;

        Assert.Equal(string.Empty, control.FormOwnerId);
    }

    [Fact]
    public void SemanticDocument_NumberInputExcludesInapplicablePlaceholder() {
        HtmlSemanticFormControl control = Assert.Single(HtmlConversionDocument
            .Parse("<input type='number' name='quantity' placeholder='Count'>")
            .SemanticDocument.Sections.SelectMany(section => section.Blocks)).FormControl!;

        Assert.Equal(string.Empty, control.Placeholder);
    }

    [Fact]
    public void SemanticDocument_FileInputDoesNotFabricateASelectedFile() {
        HtmlSemanticFormControl control = Assert.Single(HtmlConversionDocument
            .Parse("<input type='file' name='attachment' value='report.pdf'>")
            .SemanticDocument.Sections.SelectMany(section => section.Blocks)).FormControl!;

        Assert.Empty(control.Values);
        Assert.Equal(string.Empty, control.Value);
    }

    [Theory]
    [InlineData("<input type='range'>", "50")]
    [InlineData("<input type='range' value='invalid'>", "50")]
    [InlineData("<input type='range' min='10' max='20'>", "15")]
    [InlineData("<input type='range' min='10' max='20' value='4'>", "10")]
    [InlineData("<input type='range' min='10' max='20' value='24'>", "20")]
    [InlineData("<input type='range' min='0' max='10' step='3' value='8'>", "8")]
    [InlineData("<input type='range' min='0' max='10' value='005.0'>", "005.0")]
    [InlineData("<input type='range' min='-1e308' max='1e308'>", "0")]
    public void SemanticDocument_RangeInputReportsItsSanitizedCurrentValue(string html, string expected) {
        HtmlSemanticFormControl control = Assert.Single(HtmlConversionDocument
            .Parse(html)
            .SemanticDocument.Sections.SelectMany(section => section.Blocks)).FormControl!;

        Assert.Equal(expected, control.Value);
    }

    [Fact]
    public void SemanticDocument_WhitespacePaddedFormKeywordsUseTheirInvalidValueDefaults() {
        HtmlSemanticBlock form = Assert.Single(HtmlConversionDocument.Parse("""
            <form method=" post " enctype=" multipart/form-data ">
              <input name="choice" type=" checkbox " checked>
              <input name="unicode" type="checKbox" checked>
              <button name="save" type=" reset " formmethod=" post " formenctype=" text/plain ">Save</button>
            </form>
            """).SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal("get", form.Form!.Method);
        Assert.Equal("application/x-www-form-urlencoded", form.Form.EncodingType);
        HtmlSemanticFormControl input = Assert.Single(form.Children, child => child.FormControl?.Name == "choice").FormControl!;
        HtmlSemanticFormControl unicode = Assert.Single(form.Children, child => child.FormControl?.Name == "unicode").FormControl!;
        HtmlSemanticFormControl button = Assert.Single(form.Children, child => child.FormControl?.Name == "save").FormControl!;
        Assert.Equal("text", input.Type);
        Assert.False(input.IsChecked);
        Assert.Equal("text", unicode.Type);
        Assert.False(unicode.IsChecked);
        Assert.Equal("submit", button.Type);
        Assert.Equal("get", button.FormMethod);
        Assert.Equal("application/x-www-form-urlencoded", button.FormEncodingType);
    }

    [Fact]
    public void SemanticDocument_PreservesOrderedListDirectionAndItemOrdinals() {
        HtmlSemanticBlock list = Assert.Single(HtmlConversionDocument
            .Parse("<ol start='3' reversed><li>A</li><li value='10'>B</li><li>C</li></ol>")
            .SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal(HtmlSemanticListKind.Ordered, list.List!.Kind);
        Assert.Equal(3, list.List.Start);
        Assert.True(list.List.IsReversed);
        Assert.Equal(new int?[] { 3, 10, 9 }, list.Children.Select(item => item.ListItem!.Ordinal));
        Assert.Equal(new int?[] { null, 10, null }, list.Children.Select(item => item.ListItem!.ExplicitOrdinal));
        Assert.Equal("3. A\n10. B\n9. C", list.Text);
    }

    [Fact]
    public void SemanticDocument_ListTextRecursivelyIncludesNestedItems() {
        HtmlSemanticBlock list = Assert.Single(HtmlConversionDocument.Parse("""
            <ul><li>Parent<ol><li>Child<ul><li>Grandchild</li></ul></li></ol></li><li>Sibling</li></ul>
            """).SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal("• Parent\n  1. Child\n    • Grandchild\n• Sibling", list.Text);
    }

    [Fact]
    public void SemanticDocument_DefinitionListIncludesDivWrappedGroups() {
        HtmlSemanticBlock list = Assert.Single(HtmlConversionDocument.Parse("""
            <dl><div><dt>Term</dt><dd>Description</dd></div><dt>Loose</dt><dd>Tail</dd></dl>
            """).SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal(new[] {
            HtmlSemanticListItemKind.Term,
            HtmlSemanticListItemKind.Description,
            HtmlSemanticListItemKind.Term,
            HtmlSemanticListItemKind.Description
        }, list.Children.Select(item => item.ListItem!.Kind));
        Assert.Equal("Term\n  Description\nLoose\n  Tail", list.Text);
    }

    [Fact]
    public void SemanticDocument_ListOrdinalsSaturateInsteadOfOverflowing() {
        HtmlSemanticBlock list = Assert.Single(HtmlConversionDocument
            .Parse("<ol start='2147483647'><li>A</li><li>B</li></ol>")
            .SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal(new int?[] { int.MaxValue, int.MaxValue },
            list.Children.Select(item => item.ListItem?.Ordinal));
    }

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
    public void OneNoteGenericImportRetainsNestedAndFollowingItemsAfterAnEmptyParent() {
        OfficeIMO.OneNote.OneNotePage page = Assert.Single(HtmlConversionDocument
            .Parse("<ul><li><ol><li>Nested</li></ol></li><li>Following</li></ul>")
            .ToOneNoteSectionResult()
            .RequireValue()
            .Pages);
        OfficeIMO.OneNote.OneNoteParagraph[] paragraphs = Assert.Single(page.Outlines).Children
            .OfType<OfficeIMO.OneNote.OneNoteParagraph>()
            .ToArray();

        Assert.Equal(new[] { "Nested", "Following" },
            paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));
        Assert.Equal(new int?[] { 1, 0 }, paragraphs.Select(paragraph => paragraph.List?.Level));
    }

    [Fact]
    public void OneNoteGenericImportPreservesEffectiveOrderedListOrdinals() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<ol start='3' reversed><li>A</li><li value='10'>B</li><li>C</li></ol>");

        var page = Assert.Single(source.ToOneNoteSectionResult().RequireValue().Pages);
        var paragraphs = Assert.Single(page.Outlines).Children
            .OfType<OfficeIMO.OneNote.OneNoteParagraph>()
            .ToArray();

        Assert.Equal(new int?[] { 3, 10, 9 }, paragraphs.Select(paragraph => paragraph.List?.DisplayIndex));
        Assert.All(paragraphs, paragraph => Assert.True(paragraph.List?.Ordered));
        Assert.All(paragraphs, paragraph => Assert.True(paragraph.List?.Restart));
    }

    [Fact]
    public void OneNoteGenericImportUsesNativeContinuationForOrdinaryOrderedLists() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<ol><li>A</li><li>B</li></ol>");

        OfficeIMO.OneNote.OneNoteSection section = source.ToOneNoteSectionResult().RequireValue();
        var paragraphs = Assert.Single(Assert.Single(section.Pages).Outlines).Children
            .OfType<OfficeIMO.OneNote.OneNoteParagraph>()
            .ToArray();
        Assert.Equal(new int?[] { 1, null }, paragraphs.Select(paragraph => paragraph.List?.DisplayIndex));
        Assert.Equal(new bool?[] { true, false }, paragraphs.Select(paragraph => paragraph.List?.Restart));

        byte[] bytes = OfficeIMO.OneNote.OneNoteSectionWriter.Write(section);
        OfficeIMO.OneNote.OneNoteSection reloaded = OfficeIMO.OneNote.OneNoteSectionReader.Read(new MemoryStream(bytes));
        var reloadedParagraphs = Assert.Single(Assert.Single(reloaded.Pages).Outlines).Children
            .OfType<OfficeIMO.OneNote.OneNoteParagraph>()
            .ToArray();
        Assert.Equal(new int?[] { 1, null }, reloadedParagraphs.Select(paragraph => paragraph.List?.DisplayIndex));
        Assert.Equal(new bool?[] { true, false }, reloadedParagraphs.Select(paragraph => paragraph.List?.Restart));
    }

    [Fact]
    public void OneNoteGenericImportRestartsSeparateListsAndRendersEffectiveOrdinals() {
        OfficeIMO.OneNote.OneNotePage page = Assert.Single(HtmlConversionDocument
            .Parse("<ol><li>A</li><li>B</li></ol><p>Break</p><ol><li>C</li></ol>")
            .ToOneNoteSectionResult()
            .RequireValue()
            .Pages);
        OfficeIMO.OneNote.OneNoteParagraph[] listParagraphs = Assert.Single(page.Outlines).Children
            .OfType<OfficeIMO.OneNote.OneNoteParagraph>()
            .Where(paragraph => paragraph.List?.Ordered == true)
            .ToArray();

        Assert.Equal(new int?[] { 1, null, 1 }, listParagraphs.Select(paragraph => paragraph.List?.DisplayIndex));
        Assert.Equal(new bool?[] { true, false, true }, listParagraphs.Select(paragraph => paragraph.List?.Restart));

        string[] rendered = page
            .ToDrawing(new OfficeIMO.OneNote.OneNotePageRenderingOptions { IncludeTitle = false })
            .Elements
            .OfType<OfficeIMO.Drawing.OfficeDrawingRichText>()
            .Select(item => string.Concat(item.Runs.Select(run => run.Text)))
            .Where(text => text.EndsWith("A", StringComparison.Ordinal)
                || text.EndsWith("B", StringComparison.Ordinal)
                || text.EndsWith("C", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(new[] { "1. A", "2. B", "1. C" }, rendered);
    }

    [Fact]
    public void OneNoteGenericImportClampsNonpositiveListOrdinalsWithDiagnostic() {
        HtmlToOneNoteSectionResult result = HtmlConversionDocument
            .Parse("<ol start='-2'><li>A</li><li>B</li><li>C</li><li>D</li><li>E</li></ol>")
            .ToOneNoteSectionResult();
        OfficeIMO.OneNote.OneNoteParagraph[] paragraphs =
            Assert.Single(Assert.Single(result.Value.Pages).Outlines).Children
                .OfType<OfficeIMO.OneNote.OneNoteParagraph>()
                .ToArray();

        Assert.Equal(new int?[] { 1, 1, 1, 1, null },
            paragraphs.Select(paragraph => paragraph.List?.DisplayIndex));
        Assert.Equal(new bool?[] { true, true, true, true, false },
            paragraphs.Select(paragraph => paragraph.List?.Restart));
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated);
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
            typeof(string), typeof(string), typeof(OfficeConversionLossKind)
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
