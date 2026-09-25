using OfficeIMO.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlOfficeAdapters {
    [Fact]
    public void OneNote_ReportsOnlyApplicableUnloadedStylesheetLinks() {
        const string html = "<html><head>"
            + "<link rel='stylesheet' href='https://example.test/active.css'>"
            + "<link rel='stylesheet' disabled href='https://example.test/disabled.css'>"
            + "<link rel='alternate stylesheet' title='alternate' href='https://example.test/alternate.css'>"
            + "<link rel='stylesheet' type='text/less' href='https://example.test/not-css.css'>"
            + "<link rel='stylesheet' media='print' href='https://example.test/print.css'>"
            + "<style>p { color: #336699; }</style>"
            + "</head><body><p>Styled text</p></body></html>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);

        HtmlToOneNoteSectionResult section = source.ToOneNoteSectionResult();
        HtmlDiagnostic diagnostic = Assert.Single(section.Report.Diagnostics,
            item => item.Code == "HtmlStylesheetLinkSkipped");
        Assert.Equal("https://example.test/active.css", diagnostic.Source);
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.True(section.Report.HasLoss);

        OneNoteParagraph paragraph = Assert.Single(Assert.Single(Assert.Single(section.Value.Pages).Outlines)
            .Children.OfType<OneNoteParagraph>());
        Assert.Equal(0xFF336699U, Assert.Single(paragraph.Runs).Style.ColorArgb);

        HtmlToOneNoteNotebookResult notebook = source.ToOneNoteNotebookResult();
        Assert.Equal(diagnostic.Source, Assert.Single(notebook.Report.Diagnostics,
            item => item.Code == "HtmlStylesheetLinkSkipped").Source);
    }

    [Fact]
    public void OneNote_AppliedInlineStylesheetDoesNotReportExternalResourceLoss() {
        const string html = "<style>p { color: #336699; }</style><p>Styled text</p>";

        HtmlToOneNoteSectionResult result = HtmlConversionDocument.Parse(html).ToOneNoteSectionResult();

        Assert.False(result.Report.HasLoss);
        OneNoteParagraph paragraph = Assert.Single(Assert.Single(Assert.Single(result.Value.Pages).Outlines)
            .Children.OfType<OneNoteParagraph>());
        Assert.Equal(0xFF336699U, Assert.Single(paragraph.Runs).Style.ColorArgb);
    }
}
