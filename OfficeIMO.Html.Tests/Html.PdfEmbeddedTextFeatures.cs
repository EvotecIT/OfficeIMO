using System;
using System.IO;
using System.Text;
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

        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
            HtmlConversionDocument.Parse(html).ToPdfDocumentResult(options));

        Assert.Contains("point budget", error.Message, StringComparison.Ordinal);
    }
}
