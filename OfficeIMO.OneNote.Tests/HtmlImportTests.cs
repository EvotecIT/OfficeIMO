using OfficeIMO.Html;
using OfficeIMO.OneNote.Html;

namespace OfficeIMO.OneNote.Tests;

public sealed class HtmlImportTests {
    [Fact]
    public void GenericHtml_PreservesTextInsideOrdinaryContainers() {
        HtmlToOneNoteSectionResult result = HtmlConversionDocument.Parse("<div>Hello from a generic container</div>")
            .ToOneNoteSectionResult();

        OneNoteParagraph paragraph = Assert.Single(Assert.Single(Assert.Single(result.Value.Pages).Outlines).Children.OfType<OneNoteParagraph>());
        Assert.Equal("Hello from a generic container", string.Concat(paragraph.Runs.Select(run => run.Text)));
    }

    [Fact]
    public void RejectedOversizedParagraphDoesNotConsumeShapeBudget() {
        var options = new HtmlToOneNoteOptions {
            Limits = new HtmlImportLimits {
                MaxMetadataCharacters = 5,
                MaxShapes = 1
            }
        };

        HtmlToOneNoteSectionResult result = HtmlConversionDocument.Parse("<p>too long</p><p>valid</p>")
            .ToOneNoteSectionResult(options);

        OneNoteOutline outline = Assert.Single(Assert.Single(result.Value.Pages).Outlines);
        OneNoteParagraph paragraph = Assert.Single(outline.Children.OfType<OneNoteParagraph>());
        Assert.Equal("valid", string.Concat(paragraph.Runs.Select(run => run.Text)));
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded);
        Assert.DoesNotContain(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void EmptyTableDoesNotConsumeShapeBudget() {
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxShapes = 1;

        HtmlToOneNoteSectionResult result = HtmlConversionDocument.Parse("<table></table><p>valid</p>")
            .ToOneNoteSectionResult(new HtmlToOneNoteOptions { Limits = limits });

        OneNoteOutline outline = Assert.Single(Assert.Single(result.Value.Pages).Outlines);
        OneNoteParagraph paragraph = Assert.Single(outline.Children.OfType<OneNoteParagraph>());
        Assert.Equal("valid", string.Concat(paragraph.Runs.Select(run => run.Text)));
        Assert.Empty(outline.Children.OfType<OneNoteTable>());
        Assert.DoesNotContain(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void EmptyTableDoesNotConsumeTableBudget() {
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxShapes = 1;
        limits.MaxTables = 1;

        HtmlToOneNoteSectionResult result = HtmlConversionDocument.Parse("<table></table><table><tr><td></td></tr></table>")
            .ToOneNoteSectionResult(new HtmlToOneNoteOptions { Limits = limits });

        OneNoteOutline outline = Assert.Single(Assert.Single(result.Value.Pages).Outlines);
        Assert.Single(outline.Children.OfType<OneNoteTable>());
        Assert.Equal(1, result.Tables);
        Assert.DoesNotContain(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void HtmlImportBuildsTypedPagesRunsListsTablesAndImages() {
        const string html = """
            <section aria-label="Project">
              <h2>Project</h2>
              <p>Hello <strong>team</strong>.</p>
              <ul><li>First</li><li>Second</li></ul>
              <table><tr><th>Owner</th><td>Ada</td></tr></table>
              <img alt="dot" src="data:image/png;base64,AQID">
            </section>
            """;

        HtmlToOneNoteSectionResult result = HtmlConversionDocument.Parse(html).ToOneNoteSectionResult();

        Assert.True(result.Succeeded);
        Assert.Equal(1, result.Pages);
        Assert.Equal(1, result.Tables);
        Assert.Equal(1, result.Images);
        Assert.Equal("Project", result.Value.Pages[0].Title);
        OneNoteOutline outline = Assert.Single(result.Value.Pages[0].Outlines);
        Assert.Contains(outline.Children.OfType<OneNoteParagraph>(), paragraph => paragraph.Runs.Any(run => run.Style.Bold == true && run.Text == "team"));
        Assert.Equal(2, outline.Children.OfType<OneNoteParagraph>().Count(paragraph => paragraph.List != null));
    }

    [Fact]
    public void OneNoteHtmlExportExposesTheSharedTextResultContract() {
        var section = new OneNoteSection { Name = "Notes" };
        section.Pages.Add(new OneNotePage { Title = "Page" });

        HtmlTextConversionResult result = section.ToHtmlDocumentResult();

        Assert.True(result.Succeeded);
        Assert.Contains("Page", result.Value);
    }

    [Fact]
    public void OneNoteHtmlExportGroupsConsecutiveListItemsIntoOneList() {
        var section = new OneNoteSection { Name = "Lists" };
        var page = new OneNotePage { Title = "Page" };
        foreach (string text in new[] { "First", "Second" }) {
            var paragraph = new OneNoteParagraph {
                List = new OneNoteListInfo { Ordered = true, Level = 0 }
            };
            paragraph.Runs.Add(new OneNoteTextRun { Text = text });
            page.DirectContent.Add(paragraph);
        }
        section.Pages.Add(page);

        string html = section.ToHtmlDocumentResult().Value;

        Assert.Equal(1, html.Split(new[] { "<ol data-level=\"0\">" }, StringSplitOptions.None).Length - 1);
        Assert.Equal(2, html.Split(new[] { "<li>" }, StringSplitOptions.None).Length - 1);
        Assert.Contains("<ol data-level=\"0\"><li>First</li><li>Second</li></ol>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void OneNoteHtmlExportClosesParagraphBeforeRenderingChildBlocks() {
        var section = new OneNoteSection { Name = "Structure" };
        var page = new OneNotePage { Title = "Page" };
        var parent = new OneNoteParagraph();
        parent.Runs.Add(new OneNoteTextRun { Text = "Parent" });
        var child = new OneNoteParagraph();
        child.Runs.Add(new OneNoteTextRun { Text = "Child" });
        parent.Children.Add(child);
        page.DirectContent.Add(parent);
        section.Pages.Add(page);

        string html = section.ToHtmlDocumentResult().Value;

        Assert.Contains("<p>Parent</p><p>Child</p>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<p>Parent<p>", html, StringComparison.Ordinal);
    }
}
