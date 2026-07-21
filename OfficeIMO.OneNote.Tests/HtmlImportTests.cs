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
}
