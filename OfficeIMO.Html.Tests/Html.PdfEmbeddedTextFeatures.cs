using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlPdfEmbeddedTextFeatureTests {
    [Theory]
    [InlineData(false, "font-kerning:none")]
    [InlineData(false, "font-feature-settings:'liga' 1")]
    [InlineData(false, "font-variant-numeric:tabular-nums")]
    [InlineData(true, "font-kerning:none")]
    [InlineData(true, "font-feature-settings:'liga' 1")]
    [InlineData(true, "font-variant-numeric:tabular-nums")]
    public void StaticFontFeaturesRemainEmbeddedUnderTheOutlineCommandLimit(bool cff, string features) {
        byte[] font = cff
            ? File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSansPro-Regular.otf"))
            : ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i');
        string html = "<style>@font-face{font-family:FeatureFont;src:url('data:font/"
            + (cff ? "otf" : "ttf") + ";base64," + Convert.ToBase64String(font)
            + "')}p{margin:0;font:32px FeatureFont;" + features + "}</style><p>fi</p>";
        var options = new HtmlToPdfOptions { MaxOutlinedTextPathCommands = 1 };

        PdfCore.PdfDocumentConversionResult result = HtmlConversionDocument.Parse(html).ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();

        Assert.Equal("fi", PdfCore.PdfReadDocument.Open(pdf).ExtractText().Trim());
        Assert.Contains(cff ? "/FontFile3 " : "/FontFile2 ", Encoding.GetEncoding(28591).GetString(pdf), StringComparison.Ordinal);
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == HtmlPdfDiagnosticCodes.FontProgramOutlined);
    }

    [Theory]
    [InlineData("font-weight:bold")]
    [InlineData("font-style:italic")]
    [InlineData("font-weight:bold;font-style:italic")]
    public void SyntheticStylesWithFeaturesKeepTheirBoundedOutlineRoute(string style) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i');
        string html = "<style>@font-face{font-family:FeatureFont;src:url('data:font/ttf;base64,"
            + Convert.ToBase64String(font) + "')}p{margin:0;font:32px FeatureFont;"
            + "font-feature-settings:'liga' 1;" + style + "}</style><p>fi</p>";
        var options = new HtmlToPdfOptions { MaxOutlinedTextPathCommands = 1 };

        PdfCore.PdfDocumentConversionResult limited = HtmlConversionDocument.Parse(html).ToPdfDocumentResult(options);
        Assert.Equal("fi", PdfCore.PdfReadDocument.Open(limited.ToBytes()).ExtractText().Trim());
        Assert.Contains(limited.Report.Warnings, warning =>
            warning.Code == HtmlPdfDiagnosticCodes.FontOutlineBudgetApproximated
            && warning.LossKind == OfficeConversionLossKind.Approximation);
        Assert.True(limited.HasLoss);
    }

    [Theory]
    [InlineData("liga", "normal")]
    [InlineData("liga", "dark")]
    [InlineData("ss01", "normal")]
    [InlineData("ss01", "dark")]
    public void SubstitutedColorGlyphsRetainPalettePaintAndOutlineLimits(string feature, string palette) {
        byte[] font = ManagedTextShapingTestAssets.CreateColorLigatureFont('f', 'i', feature);
        string html = "<p style=\"margin:0;font:32px ColorLigature;font-feature-settings:'"
            + feature + "' 1;font-palette:" + palette + "\">fi</p>";
        var options = new HtmlToPdfOptions();
        options.Fonts.Add("ColorLigature", font);

        PdfCore.PdfDocumentConversionResult result = HtmlConversionDocument.Parse(html).ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        OfficeColor[] paints = PdfCore.PdfDocument.Load(pdf).Render.Drawing(1).Shapes
            .Select(shape => shape.Shape.FillColor ?? OfficeColor.Transparent).ToArray();

        Assert.Contains(palette == "dark" ? OfficeColor.Yellow : OfficeColor.Red, paints);
        Assert.Contains(palette == "dark" ? OfficeColor.FromRgb(0, 128, 0) : OfficeColor.Blue, paints);
        Assert.Equal("fi", PdfCore.PdfReadDocument.Open(pdf).ExtractText().Trim());

        options.MaxOutlinedTextPathCommands = 1;
        PdfCore.PdfDocumentConversionResult limited = HtmlConversionDocument.Parse(html).ToPdfDocumentResult(options);
        Assert.Equal("fi", PdfCore.PdfReadDocument.Open(limited.ToBytes()).ExtractText().Trim());
        Assert.Contains(limited.Report.Warnings, warning =>
            warning.Code == HtmlPdfDiagnosticCodes.FontOutlineBudgetApproximated
            && warning.LossKind == OfficeConversionLossKind.Approximation);
        Assert.True(limited.HasLoss);
    }
}
