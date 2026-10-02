using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlOneNoteTableCaptions {
    private const string RegulatoryTableCaption = "List of National Secondary Drinking Water Regulations";
    private const string RegulatoryTableLink = "https://example.org/standards";
    private const string RegulatoryTableHtml = "<main><h1>Drinking water</h1><table><caption><a href='"
        + RegulatoryTableLink + "'>" + RegulatoryTableCaption + "</a></caption><tr><th>Contaminant</th><th>Standard</th></tr>"
        + "<tr><td>Fluoride</td><td>2.0 mg/L</td></tr></table></main>";

    [Fact]
    public void EditableTableCaptionSurvivesOneNoteRoundTrip() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(RegulatoryTableHtml);

        HtmlToOneNoteSectionResult oneNoteResult = source.ToOneNoteSectionResult();
        OneNoteSection reopenedOneNote = OneNoteSectionReader.Read(
            new MemoryStream(OneNoteSectionWriter.Write(oneNoteResult.RequireValue())));
        Assert.Contains(RegulatoryTableCaption, reopenedOneNote.ToHtmlDocument(), StringComparison.Ordinal);
        Assert.Single(reopenedOneNote.Pages.SelectMany(page => page.Outlines)
            .SelectMany(outline => outline.Children).OfType<OneNoteTable>());
        Assert.Contains(reopenedOneNote.Pages.SelectMany(page => page.Outlines)
            .SelectMany(outline => outline.Children).OfType<OneNoteParagraph>()
            .SelectMany(paragraph => paragraph.Runs), run =>
                run.Text == RegulatoryTableCaption && run.Hyperlink == RegulatoryTableLink);
        Assert.Contains(oneNoteResult.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);

    }
}
