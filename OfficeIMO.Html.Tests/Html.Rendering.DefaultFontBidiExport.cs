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
        Assert.Equal("Left שלום 42", reopened.ExtractText().Trim());
        var svg = XDocument.Parse(source.ToSvg(options));
        string label = Assert.Single(svg.Descendants().Select(element => (string?)element.Attribute("aria-label")),
            value => value != null && value.Contains("Left", StringComparison.Ordinal) && value.Contains("שלום", StringComparison.Ordinal))!;
        Assert.Equal("Left שלום 42", label);
        HtmlRenderResult raw = HtmlRenderEngine.Execute(source, HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png, options));
        raw.Document.RequireNoLoss();
        Assert.Equal(raw.ExportImages()[0].Bytes, reopened.Pages[0].ExportImage(OfficeImageExportFormat.Png,
            new PdfCore.PdfImageExportOptions { Scale = 96D / 72D, BackgroundColor = OfficeColor.White }).Bytes);
    }
    [Fact]
    public void HtmlDefaultFont_ResolvedBidiBudgetExhaustionKeepsDiagnosedStaticText() {
        const string text = "Left שלום 42";
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false, MaxOutlinedTextPathCommands = 1
        };
        options.Fonts.Add("BidiFixture", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(true,
            text.Select(character => (int)character).Distinct().ToArray()));
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<p style='font:16px BidiFixture;word-spacing:2px'>" + text + "</p>");
        PdfCore.PdfDocumentConversionResult result = source.ToPdfDocumentResult(options);

        Assert.Equal(text, PdfCore.PdfReadDocument.Open(result.ToBytes()).ExtractText().Trim());
        Assert.Contains(result.Report.Warnings, warning =>
            warning.Code == HtmlPdfDiagnosticCodes.FontOutlineBudgetApproximated
                && warning.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
    }

}
