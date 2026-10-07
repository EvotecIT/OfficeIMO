using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutPaintProjectionTests {
    [Theory]
    [InlineData("float:left", false)]
    [InlineData("float:left", true)]
    [InlineData("display:flex", false)]
    [InlineData("display:flex", true)]
    [InlineData("display:grid", false)]
    [InlineData("display:grid", true)]
    public void OneEditableBoxRetainsCompleteTextAcrossPaintLayers(string placement, bool reorderedPaint) {
        string positioning = reorderedPaint ? " style='position:relative;z-index:1'" : string.Empty;
        string html = "<div style='" + placement + ";width:300px;height:80px'>"
            + "<span>Lead <span" + positioning + ">Alpha</span> Tail</span></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(HtmlConversionDocument.Parse(html));

        HtmlRenderLayoutRegion region = Assert.Single(projection.Regions);
        Assert.Equal("Lead Alpha Tail", region.SourceText);
        Assert.DoesNotContain("Lead Alpha Tail", projection.RemainingDocument.Body!.TextContent);
        Assert.DoesNotContain(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionFragmented);
        Assert.Contains("Lead", projection.RenderedDocument.Text);
        Assert.Contains("Alpha", projection.RenderedDocument.Text);
        Assert.Contains("Tail", projection.RenderedDocument.Text);
    }

    [Theory]
    [InlineData("word")]
    [InlineData("excel")]
    [InlineData("powerpoint")]
    [InlineData("rtf")]
    public void NativeTargetsReopenOneSourceOrderedRegion(string target) {
        const string html = "<div style='float:left;width:300px;height:80px'>"
            + "<span>Lead <span style='position:relative;z-index:1'>Alpha</span> Tail</span></div>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        using var stream = new MemoryStream();
        switch (target) {
            case "word": {
                var result = source.ToWordDocumentResult();
                using WordDocument document = result.Value;
                document.Save(stream);
                using WordDocument reopened = WordDocument.Load(new MemoryStream(stream.ToArray()),
                    new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
                WordTextBox box = Assert.Single(reopened.TextBoxes);
                Assert.Equal("Lead Alpha Tail", string.Join("", box.Paragraphs.Select(paragraph => paragraph.Text)).TrimEnd('\r', '\n'));
                Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected);
                break;
            }
            case "excel": {
                var result = source.ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
                using ExcelDocument document = result.Value;
                document.Save(stream);
                using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(stream.ToArray()),
                    new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
                var cells = reopened.Sheets.SelectMany(sheet => sheet.EnumerateCells())
                    .Where(cell => cell.Value?.ToString() == "Lead Alpha Tail").ToList();
                Assert.Single(cells);
                Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected);
                break;
            }
            case "powerpoint": {
                var result = source.ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
                using PowerPointPresentation document = result.Value;
                document.Save(stream);
                using PowerPointPresentation reopened = PowerPointPresentation.Load(new MemoryStream(stream.ToArray()),
                    new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
                PowerPointTextBox box = Assert.Single(reopened.Slides.SelectMany(slide => slide.TextBoxes),
                    candidate => candidate.Text.Contains("Lead", StringComparison.Ordinal));
                Assert.Equal("Lead Alpha Tail", box.Text);
                Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected);
                break;
            }
            case "rtf": {
                var result = source.ToRtfDocumentResult();
                RtfReadResult reopened = RtfDocument.Read(result.Value.ToRtf());
                RtfParagraph frame = Assert.Single(reopened.Document.Paragraphs, paragraph => paragraph.Frame.HasAnyValue);
                Assert.Equal("Lead Alpha Tail", frame.ToPlainText());
                Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected);
                break;
            }
        }
    }

}
