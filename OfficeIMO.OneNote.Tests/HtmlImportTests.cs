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
    public void OneNoteHtmlExportNestsConsecutiveListLevelsUnderTheirParentItem() {
        var section = new OneNoteSection { Name = "Nested lists" };
        var page = new OneNotePage { Title = "Page" };
        foreach ((string text, int level, bool ordered) in new[] {
            ("Parent", 0, false),
            ("Child one", 1, true),
            ("Child two", 1, true),
            ("Sibling", 0, false)
        }) {
            var paragraph = new OneNoteParagraph {
                List = new OneNoteListInfo { Ordered = ordered, Level = level }
            };
            paragraph.Runs.Add(new OneNoteTextRun { Text = text });
            page.DirectContent.Add(paragraph);
        }
        section.Pages.Add(page);

        string html = section.ToHtmlDocumentResult().Value;

        Assert.Contains(
            "<ul data-level=\"0\"><li>Parent<ol data-level=\"1\"><li>Child one</li><li>Child two</li></ol></li><li>Sibling</li></ul>",
            html,
            StringComparison.Ordinal);
    }

    [Fact]
    public void OneNoteSemanticHtmlRoundTripPreservesArgbAlpha() {
        var section = new OneNoteSection { Name = "Alpha" };
        var page = new OneNotePage { Title = "Page" };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun {
            Text = "Styled",
            Style = {
                ColorArgb = 0x80336699U,
                HighlightColorArgb = 0x40FFF2CCU
            }
        });
        paragraph.Runs.Add(new OneNoteTextRun {
            Text = "Transparent",
            Style = { ColorArgb = 0x00336699U }
        });
        var table = new OneNoteTable();
        var row = new OneNoteTableRow();
        var cell = new OneNoteTableCell { ShadingColorArgb = 0xA0112233U };
        cell.Content.Add(new OneNoteParagraph { Runs = { new OneNoteTextRun { Text = "Cell" } } });
        row.Cells.Add(cell);
        table.Rows.Add(row);
        page.DirectContent.Add(paragraph);
        page.DirectContent.Add(table);
        section.Pages.Add(page);

        string html = section.ToHtmlDocumentResult().Value;
        Assert.Contains("#33669980", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#FFF2CC40", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#112233A0", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#33669900", html, StringComparison.OrdinalIgnoreCase);

        OneNoteSection imported = HtmlConversionDocument.Parse(
                html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToOneNoteSectionResult()
            .RequireValue();
        OneNoteOutline outline = Assert.Single(Assert.Single(imported.Pages).Outlines);
        OneNoteTextRun run = Assert.Single(
            outline.Children.OfType<OneNoteParagraph>().SelectMany(item => item.Runs),
            item => item.Text == "Styled");
        OneNoteTableCell importedCell = Assert.Single(Assert.Single(
            outline.Children.OfType<OneNoteTable>().Single().Rows).Cells);

        Assert.Equal(0x80336699U, run.Style.ColorArgb);
        Assert.Equal(0x40FFF2CCU, run.Style.HighlightColorArgb);
        OneNoteTextRun transparent = Assert.Single(
            outline.Children.OfType<OneNoteParagraph>().SelectMany(item => item.Runs),
            item => item.Text == "Transparent");
        Assert.Equal(0x00336699U, transparent.Style.ColorArgb);
        Assert.Equal(0xA0112233U, importedCell.ShadingColorArgb);
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

    [Fact]
    public void OneNoteHtmlRoundTripReconstructsStructuredInlineMath() {
        var section = new OneNoteSection { Name = "Math" };
        var page = new OneNotePage { Title = "Page" };
        var paragraph = new OneNoteParagraph();
        OfficeIMO.Drawing.OfficeMathExpression expected = OfficeIMO.Drawing.OfficeMath.Fraction(
            OfficeIMO.Drawing.OfficeMath.Identifier("x"),
            OfficeIMO.Drawing.OfficeMath.Number("2"));
        paragraph.AddMath(expected);
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        string html = section.ToHtmlDocumentResult().Value;
        HtmlToOneNoteSectionResult imported = HtmlConversionDocument.Parse(html).ToOneNoteSectionResult();

        OneNoteTextRun run = Assert.Single(Assert.Single(imported.Value.Pages).Outlines
            .SelectMany(outline => outline.Children)
            .OfType<OneNoteParagraph>()
            .SelectMany(item => item.Runs), item => item.MathExpression != null);
        Assert.True(run.Style.IsMath);
        Assert.Equal(expected, run.MathExpression);
    }
}
