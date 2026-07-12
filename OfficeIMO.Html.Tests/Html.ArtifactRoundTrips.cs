using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlArtifactRoundTrips {
    private const string Marker = "OfficeIMO artifact marker";

    [Fact]
    public void GenericHtml_ProducesReadableWordRtfMarkdownPdfAndSvgArtifactsFromOnePreparedDocument() {
        HtmlConversionDocument source = HtmlConversionDocumentBuilder.Build(
            "<h1>Conversion proof</h1><p>" + Marker + "</p><ul><li>First</li><li>Second</li></ul>");

        HtmlToWordResult wordResult = source.ToWordDocumentResult();
        using var docx = new MemoryStream();
        wordResult.Document.Save(docx);
        using WordDocument reopenedWord = WordDocument.Load(new MemoryStream(docx.ToArray()), new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly });

        HtmlToRtfResult rtfResult = source.ToRtfDocumentResult();
        string rtf = rtfResult.Document.ToRtf();
        RtfReadResult reopenedRtf = RtfDocument.Read(rtf);

        string markdown = source.ToMarkdown();
        byte[] pdf = source.ToPdf();
        string svg = source.ToSvg();

        Assert.NotEmpty(reopenedWord.Find(Marker, StringComparison.Ordinal));
        Assert.Contains(Marker, string.Join("\n", reopenedRtf.Document.Paragraphs.Select(paragraph => paragraph.ToPlainText())), StringComparison.Ordinal);
        Assert.Contains(Marker, markdown, StringComparison.Ordinal);
        Assert.Contains(Marker, PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(Marker, svg, StringComparison.Ordinal);
        Assert.True(pdf.Length > 100);
        Assert.StartsWith("{\\rtf", rtf, StringComparison.Ordinal);
        Assert.True(wordResult.Succeeded);
        Assert.True(rtfResult.Succeeded);
    }

    [Fact]
    public void SemanticHtml_ProducesReopenableXlsxAndPptxArtifacts() {
        using ExcelDocument sourceWorkbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sourceSheet = sourceWorkbook.AddWorkSheet("Evidence");
        sourceSheet.CellValue(1, 1, "Label");
        sourceSheet.CellValue(2, 1, Marker);
        sourceSheet.MergeRange("A2:B2");
        HtmlToExcelResult excelResult = sourceWorkbook.ToHtml().ToExcelDocumentResult();
        using var xlsx = new MemoryStream();
        excelResult.Workbook.Save(xlsx);
        using ExcelDocument reopenedWorkbook = ExcelDocument.Load(new MemoryStream(xlsx.ToArray()), readOnly: true);

        using PowerPointPresentation sourcePresentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide sourceSlide = sourcePresentation.AddSlide();
        PowerPointTable sourceTable = sourceSlide.AddTablePoints(2, 2, 40, 60, 300, 120);
        sourceTable.GetCell(0, 0).Text = Marker;
        sourceTable.MergeCells(0, 0, 0, 1);
        HtmlToPowerPointResult powerPointResult = sourcePresentation.ToHtml().ToPowerPointPresentationResult();
        using var pptx = new MemoryStream();
        powerPointResult.Presentation.Save(pptx);
        using PowerPointPresentation reopenedPresentation = PowerPointPresentation.Open(
            new MemoryStream(pptx.ToArray()),
            new PowerPointStreamOpenOptions { Mode = PowerPointOpenMode.ReadOnly });

        ExcelSheet reopenedSheet = Assert.Single(reopenedWorkbook.Sheets);
        Assert.True(reopenedSheet.TryGetCellText(2, 1, out string excelText));
        Assert.Equal(Marker, excelText);
        Assert.Equal("A2:B2", Assert.Single(reopenedSheet.GetMergedRanges()).A1Range);
        PowerPointTable reopenedTable = Assert.Single(Assert.Single(reopenedPresentation.Slides).Tables);
        Assert.Equal(Marker, reopenedTable.GetCell(0, 0).Text);
        Assert.Equal((1, 2), reopenedTable.GetCell(0, 0).Merge);
        Assert.True(xlsx.Length > 100);
        Assert.True(pptx.Length > 100);
        Assert.True(excelResult.Succeeded);
        Assert.True(powerPointResult.Succeeded);
    }
}
