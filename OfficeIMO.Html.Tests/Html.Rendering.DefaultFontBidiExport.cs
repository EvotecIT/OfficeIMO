using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Pdf;
using System.Xml.Linq;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("", "Left<span style='color:red'> </span>שלום<span style='color:blue'> </span>42")]
    [InlineData("letter-spacing:1px", "Left שלום 42")]
    [InlineData("word-spacing:2px", "Left שלום 42")]
    public void HtmlDefaultFont_BidiInklessSlicesKeepOneLogicalSvgLabel(string declarations, string content) {
        string html = "<p style='margin:0;font:16px BidiFixture;" + declarations + "'>" + content + "</p>";
        var options = new HtmlToPdfOptions { AllowSystemFontFallback = false, Margins = HtmlRenderMargins.All(0D) };
        options.Fonts.Add("BidiFixture", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(true,
            "Left שלום 42".Select(character => (int)character).Distinct().ToArray()));
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        byte[] pdf = source.ToPdfBytes(options);
        PdfCore.PdfReadDocument reopened = PdfCore.PdfReadDocument.Open(pdf);
        Assert.Equal("Leftשלום42", string.Concat(reopened.ExtractText().Where(character => !char.IsWhiteSpace(character))));
        var svg = XDocument.Parse(source.ToSvg(options));
        string label = Assert.Single(svg.Descendants().Select(element => (string?)element.Attribute("aria-label")),
            value => value != null && value.Contains("Left", StringComparison.Ordinal) && value.Contains("שלום", StringComparison.Ordinal))!;
        Assert.Equal("Leftשלום42", string.Concat(label.Where(character => !char.IsWhiteSpace(character))));
        HtmlRenderResult raw = HtmlRenderEngine.Execute(source, HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png, options));
        raw.Document.RequireNoLoss();
        Assert.Equal(raw.ExportImages()[0].Bytes, reopened.Pages[0].ExportImage(OfficeImageExportFormat.Png,
            new PdfCore.PdfImageExportOptions { Scale = 96D / 72D, BackgroundColor = OfficeColor.White }).Bytes);
    }
}
