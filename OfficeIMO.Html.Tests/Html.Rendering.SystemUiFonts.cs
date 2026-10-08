using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlSystemUiFontTests {
    private const string MacFont = "/System/Library/Fonts/SFNS.ttf";

    [Theory]
    [InlineData("system-ui")]
    [InlineData("-apple-system")]
    [InlineData("BlinkMacSystemFont")]
    public void MacAliasLayoutAndPdfMatchAnExplicitScopedUiFace(string family) {
        if (!File.Exists(MacFont)) return;
        HtmlConversionDocument source = Source(family);
        var options = new HtmlToPdfOptions { ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreateTrustedHost() };
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(source, options);
        Assert.True(rendered.Fonts.TryResolveFaceForText("Animation", family, OfficeFontStyle.Regular, 16D, out var face));
        Assert.Equal(72.5625D, face!.Program.Measure("Animation", 16D), 9);
        Assert.False(face.CanEmbedAsStaticPdfFont);
        var reference = new HtmlToPdfOptions { AllowSystemFontFallback = false,
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreateTrustedHost() };
        Assert.True(reference.Fonts.TryAdd(family, File.ReadAllBytes(MacFont)));
        Assert.Equal(source.ToPdfBytes(reference), source.ToPdfBytes(options));
        Assert.Empty(options.Fonts.Faces);
    }

    [Fact]
    public void ExplicitInstalledFamilyPrecedesUiAliasInItsCssList() {
        if (!File.Exists(MacFont)) return;
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(Source("Helvetica Neue, system-ui"));
        Assert.Contains(rendered.Fonts.Faces, face => face.FamilyName == "Helvetica Neue");
        Assert.DoesNotContain(rendered.Fonts.Faces, face => face.FamilyName == "system-ui");
    }

    [Fact]
    public void CallerSuppliedAliasAndDisabledSystemFallbackKeepTheirScopes() {
        var options = new HtmlRenderOptions { AllowSystemFontFallback = false };
        Assert.Empty(HtmlRenderEngine.Render(Source("system-ui"), options).Fonts.Faces);
        if (!File.Exists(MacFont)) return;
        options.Fonts.Add("system-ui", File.ReadAllBytes(MacFont));
        Assert.Single(HtmlRenderEngine.Render(Source("system-ui"), options).Fonts.Faces);
    }

    [Fact]
    public void DecodedFontBytesAndLaterInlineResourcesShareAnOperationBudget() {
        var session = new HtmlResourceSession(maxResourceBytes: 100, maxTotalResourceBytes: 100);
        session.AcceptDecodedFontBytes(80);
        Assert.True(session.TryAcceptInline(HtmlResourceKind.Font, "data:font/ttf;base64,AQ==",
            new HtmlResolvedResource(new byte[20], "font/ttf"), out _, out _));
        Assert.False(session.TryAcceptInline(HtmlResourceKind.Font, "data:font/ttf;base64,Ag==",
            new HtmlResolvedResource(new byte[1], "font/ttf"), out string code, out _));
        Assert.Equal(HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded, code);
        Assert.Equal(80, session.DecodedFontBytes);
        Assert.Equal(20, session.AcceptedResourceBytes);
        Assert.Throws<InvalidOperationException>(() => session.AcceptDecodedFontBytes(1));
    }

    [Fact]
    public void ParentAndFrameDecodedFontsShareTheSameResourceBudget() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        string rule = "<style>@font-face{font-family:Probe;src:url('data:font/ttf;base64,"
            + Convert.ToBase64String(font) + "')}p{font-family:Probe}</style><p>A</p>";
        string child = rule.Replace("&", "&amp;").Replace("\"", "&quot;");
        HtmlRenderDocument result = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            rule + "<iframe srcdoc=\"" + child + "\"></iframe>"), new HtmlRenderOptions {
                MaxResourceBytes = font.Length + 1L,
                MaxTotalResourceBytes = font.Length * 3L + 1L
            });
        Assert.Single(result.Fonts.Faces);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded);
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, false)]
    public void PdfResourcePolicyCanDisableInstalledUiFaces(bool systemFonts, bool documentFonts) {
        var options = new HtmlToPdfOptions { ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreateTrustedHost() };
        options.ResourcePolicy.AllowSystemFontEmbedding = systemFonts;
        options.ResourcePolicy.AllowDocumentFontEmbedding = documentFonts;
        HtmlPdfRenderRequestResult result = Source("system-ui").RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));
        Assert.Empty(result.RenderResult.Document.Fonts.Faces);
        Assert.DoesNotContain(result.RenderResult.Diagnostics, item => item.Code == "InstalledFontResolved");
        Assert.Contains("Animation", OfficeIMO.Pdf.PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    private static HtmlConversionDocument Source(string family) => HtmlConversionDocument.Parse(
        "<style>p{font-family:" + family + ";font-size:16px;font-weight:400}</style><p>Animation</p>");
}
