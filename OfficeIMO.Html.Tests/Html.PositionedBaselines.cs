using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPositionedBaselineTests {
    [Fact]
    public void SavedPdfAndSvgUseTheSameAuthoredTextBaselines() {
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, Margins = HtmlRenderMargins.All(0), DefaultFontFamily = "Pinned" };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSerif4-Regular.otf")));
        var document = HtmlConversionDocument.Parse("<p style='margin:16px;font:20px/32px Pinned'>A<sub>2</sub>B<sup>3</sup>C</p>");
        string svg = System.Text.Encoding.UTF8.GetString(document.ExportImages(OfficeImageExportFormat.Svg, options)[0].Bytes);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        HtmlRenderText[] text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Equal(new[] { "A", "2", "B", "3", "C" }, text.Select(item => item.Text).ToArray());
        HtmlRenderText subscript = Assert.Single(text, item => item.Text == "2");
        HtmlRenderText superscript = Assert.Single(text, item => item.Text == "3");
        Assert.True(subscript.Font.Size < text[0].Font.Size);
        Assert.True(superscript.Font.Size < text[0].Font.Size);
        double normal = Baseline(text[0]);
        Assert.Equal(normal, Baseline(text[2]), 6);
        Assert.Equal(normal, Baseline(text[4]), 6);
        Assert.Equal(normal + 20D * 0.15D, Baseline(subscript), 6);
        Assert.Equal(normal - 20D * 0.30D, Baseline(superscript), 6);

        // Retain nonempty source and font-specific vertical contour proof,
        // including reduced script faces and authored baseline displacements.
        Assert.Equal("A2B3C", SvgPaintedText.Read(svg));
        HtmlRenderingTests.AssertPdfBackedSvgRetainsBaselineInk(rendered.Pages[0], svg);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToPdfBytes(new HtmlToPdfOptions(options)));
        var page = pdf.GetPage(1);
        Assert.Equal("A2B3C", string.Concat(page.Letters.Select(letter => letter.Value)));
        foreach (HtmlRenderText source in text) {
            var letter = Assert.Single(page.Letters, letter => letter.Value == source.Text);
            double actual = ((double)page.Height - letter.StartBaseLine.Y) / 0.75D;
            Assert.InRange(actual - Baseline(source), -0.02D, 0.02D);
        }

        static double Baseline(HtmlRenderText source) => source.Y + source.Font.Size + source.BaselineOffset;
    }

    [Fact]
    public void TopLevelInlineContentDoesNotReportAChangeOfPageGeometry() {
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, PageSize = new OfficePageSize(5, 4), Margins = HtmlRenderMargins.All(0) };
        var result = HtmlConversionDocument.Parse("<span>First inline line</span><br><svg width='100' height='30'><text x='0' y='20'>SVG label</text></svg>")
            .ExportImages(OfficeImageExportFormat.Png, options);
        Assert.Single(result);
        Assert.DoesNotContain(result[0].Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
    }
}
