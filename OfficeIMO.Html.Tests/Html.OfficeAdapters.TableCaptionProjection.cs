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

public partial class HtmlOfficeAdapters {
    private const string RegulatoryTableCaption = "List of National Secondary Drinking Water Regulations";
    private const string RegulatoryTableLink = "https://example.org/standards";
    private const string RegulatoryTableHtml = "<main><h1>Drinking water</h1><table><caption><a href='"
        + RegulatoryTableLink + "'>" + RegulatoryTableCaption + "</a></caption><tr><th>Contaminant</th><th>Standard</th></tr>"
        + "<tr><td>Fluoride</td><td>2.0 mg/L</td></tr></table></main>";

    [Fact]
    public void EditableTableCaptionSurvivesOneNotePowerPointAndMarkdownRoundTrips() {
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

        HtmlToPowerPointResult powerPointResult = source.ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using (PowerPointPresentation presentation = powerPointResult.RequireValue()) {
            using var artifact = new MemoryStream();
            presentation.Save(artifact);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(
                new MemoryStream(artifact.ToArray()),
                new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            Assert.Contains(reopened.Slides.SelectMany(slide => slide.TextBoxes),
                box => box.Text == RegulatoryTableCaption);
            Assert.Contains(reopened.Slides.SelectMany(slide => slide.TextBoxes)
                .SelectMany(box => box.Paragraphs).SelectMany(paragraph => paragraph.Runs), run =>
                    run.Text == RegulatoryTableCaption
                    && run.Hyperlink == new Uri(RegulatoryTableLink));
            Assert.Single(reopened.Slides.SelectMany(slide => slide.Tables));
        }
        Assert.Contains(powerPointResult.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);

        HtmlToMarkdownResult markdownResult = source.ToMarkdownDocumentResult();
        string markdown = markdownResult.RequireValue().ToMarkdown();
        MarkdownDoc reopenedMarkdown = MarkdownDoc.Parse(markdown);
        Assert.Contains(RegulatoryTableCaption, reopenedMarkdown.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains(RegulatoryTableLink, reopenedMarkdown.ToMarkdown(), StringComparison.Ordinal);
        Assert.Single(reopenedMarkdown.Blocks.OfType<TableBlock>());
        Assert.Contains(markdownResult.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void ExcelRetainsFullLinkedCaptionAfterNativeTableCells() {
        HtmlToExcelResult result = HtmlConversionDocument.Parse(RegulatoryTableHtml).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.RequireValue();
        using var artifact = new MemoryStream();
        workbook.Save(artifact);
        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(artifact.ToArray()));

        ExcelSheet table = Assert.Single(reopened.Sheets);
        Assert.Equal("Contaminant", table.CellAt(1, 1).GetValue<string>());
        Assert.Equal("Fluoride", table.CellAt(2, 1).GetValue<string>());
        Assert.Equal(RegulatoryTableCaption, table.CellAt(4, 1).GetValue<string>());
        Assert.Equal(RegulatoryTableLink, table.GetHyperlinks()["A4"].Target);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation
            && diagnostic.Detail?.Contains("originalLength=" + RegulatoryTableCaption.Length,
                StringComparison.Ordinal) == true);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentOmitted
            && diagnostic.Message.Contains("caption", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void WordAndRtfRetainLinkedTableCaptionAfterReload() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(RegulatoryTableHtml);

        HtmlToWordResult wordResult = source.ToWordDocumentResult();
        using var wordArtifact = new MemoryStream();
        using (WordDocument word = wordResult.RequireValue()) word.Save(wordArtifact);
        using WordDocument reopenedWord = WordDocument.Load(new MemoryStream(wordArtifact.ToArray()));
        string wordHtml = reopenedWord.ToHtmlResult().RequireValue();
        Assert.Contains(RegulatoryTableCaption, wordHtml, StringComparison.Ordinal);
        Assert.Contains(RegulatoryTableLink, wordHtml, StringComparison.Ordinal);

        HtmlToRtfResult rtfResult = source.ToRtfDocumentResult();
        RtfDocument reopenedRtf = RtfDocument.Load(rtfResult.RequireValue().ToBytes());
        string rtfHtml = reopenedRtf.ToHtml();
        Assert.Contains(RegulatoryTableCaption, rtfHtml, StringComparison.Ordinal);
        Assert.Contains(RegulatoryTableLink, rtfHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointKeepsLargeTableCaptionWithFirstPaginatedRow() {
        string rows = string.Concat(Enumerable.Range(1, 12).Select(index =>
            "<tr><td>Row " + index + "</td><td>Standard " + index + "</td></tr>"));
        string html = "<main><h1>Drinking water</h1><table><caption>"
            + RegulatoryTableCaption + "</caption>" + rows + "</table></main>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.RequireValue();
        PowerPointSlide captionSlide = Assert.Single(presentation.Slides,
            slide => slide.TextBoxes.Any(box => box.Text == RegulatoryTableCaption));
        Assert.Contains(captionSlide.TextBoxes, box => box.Text.Contains("Row 1, cell 1", StringComparison.Ordinal));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.Message.Contains("table and its caption", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void PowerPointCaptionMetadataLimitDoesNotOmitTheTable() {
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxMetadataCharacters = 12;
        const string html = "<main><h1>Water</h1><table><caption>Caption exceeds metadata limit</caption>"
            + "<tr><th>Term</th><th>Value</th></tr><tr><td>pH</td><td>7</td></tr></table></main>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using PowerPointPresentation presentation = result.RequireValue();
        Assert.Single(presentation.Slides.SelectMany(slide => slide.Tables));
        Assert.DoesNotContain(presentation.Slides.SelectMany(slide => slide.TextBoxes),
            box => box.Text.Contains("Caption exceeds", StringComparison.Ordinal));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded
            && diagnostic.Message.Contains("table caption", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void PowerPointAuthoredTableGeometryStaysNativeWithCaption() {
        const string html = "<main><h1>Water</h1><table data-officeimo-top='90' data-officeimo-height='440'>"
            + "<caption>" + RegulatoryTableCaption + "</caption>"
            + "<tr><th>Term</th><th>Value</th></tr><tr><td>pH</td><td>7</td></tr></table></main>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.RequireValue();
        Assert.Single(presentation.Slides.SelectMany(slide => slide.Tables));
        Assert.Contains(presentation.Slides.SelectMany(slide => slide.TextBoxes),
            box => box.Text == RegulatoryTableCaption);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Message.Contains("split into editable text", StringComparison.OrdinalIgnoreCase));
    }
}
