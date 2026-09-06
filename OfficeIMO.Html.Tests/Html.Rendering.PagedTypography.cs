using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(5)]
    [InlineData(40)]
    public void HtmlPdfTypography_UsesResolvedNestedScriptScaleAndNumericOffset(int offset) {
        string html = "<p style='font:20px/30px Helvetica;margin:0'>NORMAL " +
            "<sup><sup>NESTED</sup></sup> <span style='vertical-align:" + offset + "px'>RAISED</span></p>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var options = new HtmlToPdfOptions();
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        HtmlRenderText nested = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "NESTED");
        PdfCore.PdfTextSpan[] spans = PdfCore.PdfReadDocument.Open(document.ToPdfBytes(options)).Pages[0].GetTextSpans().ToArray();
        PdfCore.PdfTextSpan nestedPdf = Assert.Single(spans, span => span.Text == "NESTED");
        Assert.Equal(nested.Font.Size * nested.BaselineScale * .75D, nestedPdf.FontSize, 3);
        PdfCore.PdfTextSpan normalPdf = Assert.Single(spans, span => span.Text.Trim() == "NORMAL");
        PdfCore.PdfTextSpan raisedPdf = Assert.Single(spans, span => span.Text == "RAISED");
        Assert.Equal(offset * .75D, raisedPdf.Y - normalPdf.Y, 3);
    }

    [Fact]
    public void HtmlPagedTypography_PreservesMixedRunsAndRowsAcrossEveryExportedPage() {
        var rows = new StringBuilder();
        for (int row = 0; row < 18; row++) {
            rows.Append("<tr><td>Row").Append(row.ToString("D2"))
                .Append("</td><td>Polish: Zażółć. Combining: e\u0301. <strong>Bold</strong> <em>italic</em></td></tr>");
        }
        string html = """
            <style>
            @page { size: 360px 260px; margin: 30px 20px;
              @top-center { content: "Typography packet"; font-size: 10px; }
              @bottom-right { content: "Page " counter(page); font-size: 9px; }
            }
            body { margin: 0; font: 12px/16px Arial, sans-serif; }
            p { margin: 0 0 8px; } table { width: 100%; border-collapse: collapse; }
            td, th { border: 1px solid #789; padding: 4px; text-align: left; }
            th { background: #def; } tr { break-inside: avoid; }
            .summary { break-inside: avoid; border-top: 2px solid #258; padding-top: 6px; }
            </style>
            <p>Mixed baseline: normal <sup>2</sup> and <sub>n</sub>.</p>
            <p dir="rtl">שלום 123 ABC</p>
            <table><thead><tr><th>Record</th><th>Description</th></tr></thead><tbody>
            """ + rows + "</tbody></table><section class='summary'><p>SummaryStart</p><p>SummaryEnd</p></section>";
        var options = new HtmlToPdfOptions();
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        Assert.True(rendered.Pages.Count >= 3);
        for (int row = 0; row < 18; row++) {
            string marker = "Row" + row.ToString("D2");
            Assert.Equal(1, rendered.Text.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
            HtmlRenderPage page = Assert.Single(rendered.Pages, item => item.Visuals.OfType<HtmlRenderText>().Any(text => text.Text == marker));
            Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Record");
        }
        HtmlRenderPage summaryPage = Assert.Single(rendered.Pages, page => page.Visuals.OfType<HtmlRenderText>().Any(text => text.Text == "SummaryStart"));
        Assert.Contains(summaryPage.Visuals.OfType<HtmlRenderText>(), text => text.Text == "SummaryEnd");
        Assert.All(rendered.Pages, page => {
            Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Typography packet");
            Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Page " + page.PageNumber);
        });
        byte[] pdf = document.ToPdfBytes(options);
        Assert.Equal(rendered.Pages.Count, PdfCore.PdfInspector.Inspect(pdf).PageCount);
        string extracted = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("Zażółć", extracted, StringComparison.Ordinal);
        Assert.Contains("שלום", extracted, StringComparison.Ordinal);
        Assert.Contains("e\u0301", extracted, StringComparison.Ordinal);
        IReadOnlyList<OfficeImageExportResult> png = document.ExportImages(OfficeImageExportFormat.Png, options);
        IReadOnlyList<OfficeImageExportResult> svg = document.ExportImages(OfficeImageExportFormat.Svg, options);
        Assert.Equal(rendered.Pages.Count, png.Count);
        Assert.Equal(rendered.Pages.Count, svg.Count);
        for (int page = 0; page < rendered.Pages.Count; page++) {
            Assert.Equal(360, png[page].Width);
            Assert.Equal(260, png[page].Height);
            Assert.Equal(png[page].Width, svg[page].Width);
            Assert.Equal(png[page].Height, svg[page].Height);
        }
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_TYPOGRAPHY_ARTIFACTS");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output);
            File.WriteAllText(Path.Combine(output, "source.html"), html);
            File.WriteAllBytes(Path.Combine(output, "typography.pdf"), pdf);
            for (int page = 0; page < png.Count; page++) {
                File.WriteAllBytes(Path.Combine(output, $"page-{page + 1:D2}.png"), png[page].Bytes);
                File.WriteAllBytes(Path.Combine(output, $"page-{page + 1:D2}.svg"), svg[page].Bytes);
            }
        }
    }
}
