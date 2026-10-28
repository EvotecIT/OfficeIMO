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
        HtmlConversionDocument source = OfficeIMO.Html.HtmlConversionDocument.Parse(
            "<h1>Conversion proof</h1><p>" + Marker + "</p><ul><li>First</li><li>Second</li></ul>");

        HtmlToWordResult wordResult = source.ConvertToWordWithEvidence();
        using var docx = new MemoryStream();
        wordResult.Value.Save(docx);
        using WordDocument reopenedWord = WordDocument.Load(new MemoryStream(docx.ToArray()), new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });

        HtmlToRtfResult rtfResult = source.ToRtfDocumentResult();
        string rtf = rtfResult.Value.ToRtf();
        RtfReadResult reopenedRtf = RtfDocument.Read(rtf);

        string markdown = source.ToMarkdown().ToMarkdown();
        byte[] pdf = (source.ToPdf()).ToBytes();
        string svg = source.ToSvg();

        Assert.NotEmpty(reopenedWord.Find(Marker, StringComparison.Ordinal));
        Assert.Contains(Marker, string.Join("\n", reopenedRtf.Document.Paragraphs.Select(paragraph => paragraph.ToPlainText())), StringComparison.Ordinal);
        Assert.Contains(Marker, markdown, StringComparison.Ordinal);
        Assert.Contains(Marker, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        // SVG can retain logical text on outlined groups or painted text elements.
        string svgText = ReadSvgSemanticText(System.Xml.Linq.XDocument.Parse(svg).Root!);
        Assert.Contains(string.Concat(Marker.Where(character => !char.IsWhiteSpace(character))),
            string.Concat(svgText.Where(character => !char.IsWhiteSpace(character))), StringComparison.Ordinal);
        Assert.True(pdf.Length > 100);
        Assert.StartsWith("{\\rtf", rtf, StringComparison.Ordinal);
        Assert.True(wordResult.Succeeded);
        Assert.True(rtfResult.Succeeded);
    }

    private static string ReadSvgSemanticText(System.Xml.Linq.XElement element) {
        string? label = element.Attribute("aria-label")?.Value;
        if (!string.IsNullOrEmpty(label)
            && (element.Name.LocalName == "text" || element.Attribute("role")?.Value == "img")) return label;
        if (element.Name.LocalName == "text") return element.Value;
        return string.Concat(element.Elements().Select(ReadSvgSemanticText));
    }

    [Fact]
    public void SemanticHtml_ProducesReopenableXlsxAndPptxArtifacts() {
        using ExcelDocument sourceWorkbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sourceSheet = sourceWorkbook.AddWorksheet("Evidence");
        sourceSheet.CellValue(1, 1, "Label");
        sourceSheet.CellValue(2, 1, Marker);
        sourceSheet.MergeRange("A2:B2");
        HtmlToExcelResult excelResult = OfficeIMO.Html.HtmlConversionDocument.Parse(sourceWorkbook.ToHtml()).ToExcelDocumentResult();
        using var xlsx = new MemoryStream();
        excelResult.Value.Save(xlsx);
        using ExcelDocument reopenedWorkbook = ExcelDocument.Load(new MemoryStream(xlsx.ToArray()), new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });

        using PowerPointDocument sourcePresentation = PowerPointDocument.Create(new MemoryStream());
        PowerPointSlide sourceSlide = sourcePresentation.AddSlide();
        PowerPointTable sourceTable = sourceSlide.AddTable(2,2,global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(60),global::OfficeIMO.OfficeLength.Points(300),global::OfficeIMO.OfficeLength.Points(120));
        sourceTable.GetCell(0, 0).Text = Marker;
        sourceTable.MergeCells(0, 0, 0, 1);
        HtmlToPowerPointResult powerPointResult = OfficeIMO.Html.HtmlConversionDocument.Parse(sourcePresentation.ToHtml()).ToPowerPointDocumentResult();
        using var pptx = new MemoryStream();
        powerPointResult.Value.Save(pptx);
        using PowerPointDocument reopenedPresentation = PowerPointDocument.Load(
            new MemoryStream(pptx.ToArray()),
            new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });

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
