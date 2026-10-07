using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlFontUsageTests {
    [Theory]
    [InlineData("Good", "", false)]
    [InlineData("Missing,Good", "", true)]
    [InlineData("Good,Missing", "", false)]
    [InlineData("Good", "<p style='display:none;font-family:Missing'>Hidden</p>", false)]
    [InlineData("Good", "<style>p::before{content:'Generated';font-family:Missing}</style>", true)]
    [InlineData("Good", "<button style='font-family:Missing'>Label</button>", true)]
    [InlineData("Good", "<svg width='120' height='40'><text x='0' y='20' font-family='Missing'>SVG text</text></svg>", true)]
    [InlineData("Good", "<svg width='120' height='40'><text display='none' font-family='Missing'>Hidden SVG</text></svg>", false)]
    [InlineData("Good", "<svg width='120' height='40'><text visibility='hidden' font-family='Missing'>Hidden SVG</text></svg>", false)]
    [InlineData("Good", "<svg width='120' height='40'><text visibility='hidden'><tspan visibility='visible' font-family='Missing'>Visible SVG</tspan></text></svg>", true)]
    public async Task UnavailableFaceLossDependsOnTextRequest(string families, string additionalHtml, bool expectedLoss) {
        HtmlRenderDocument rendered = await RenderAsync("p{font-family:" + families + "}", "<p>Visible text</p>" + additionalHtml);
        Assert.Equal(expectedLoss, rendered.HasLoss);
        HtmlDiagnostic unavailable = Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable);
        Assert.Equal(expectedLoss ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.None, unavailable.LossKind);
        Assert.Equal(expectedLoss ? HtmlDiagnosticSeverity.Warning : HtmlDiagnosticSeverity.Info, unavailable.Severity);
    }

    [Theory]
    [InlineData(400, false)]
    [InlineData(700, true)]
    public async Task UnavailableWeightIsAttributedOnlyWhenPreferred(int weight, bool expectedLoss) {
        HtmlRenderDocument rendered = await RenderAsync("p{font-family:Good;font-weight:" + weight + "}", "<p>Visible text</p>",
            "@font-face{font-family:Good;font-weight:700;src:url('https://font.example/missing-bold.ttf')}", missingFamily: null);
        Assert.Equal(expectedLoss, rendered.HasLoss);
    }

    [Theory]
    [InlineData("Latin text", false)]
    [InlineData("\u03a9", true)]
    public async Task UnavailableUnicodeSubsetDoesNotAffectUnrelatedText(string text, bool expectedLoss) {
        HtmlRenderDocument rendered = await RenderAsync("p{font-family:Missing,Good}", "<p>" + text + "</p>",
            "@font-face{font-family:Missing;unicode-range:U+0370-03FF;src:url('https://font.example/missing.ttf')}", missingFamily: null);
        Assert.Equal(expectedLoss, rendered.HasLoss);
    }

    [Theory]
    [InlineData("Hello\u200e")]
    [InlineData("\u200d")]
    [InlineData("\ufe0f")]
    public async Task ControlOnlyGraphemesDoNotRequestUnavailableSubset(string text) {
        var rendered = await RenderAsync("p{font-family:Missing,Good}", "<p>" + text + "</p>",
            "@font-face{font-family:Missing;unicode-range:U+0370-03FF;src:url('https://font.example/missing.ttf')}", missingFamily: null);
        var unavailable = Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable);
        Assert.Equal(HtmlDiagnosticSeverity.Info, unavailable.Severity);
        Assert.Equal(OfficeConversionLossKind.None, unavailable.LossKind);
    }

    [Fact]
    public async Task CallerFaceCanSatisfyUnavailableCssFace() {
        var options = CreateOptions();
        options.Fonts.Add("Missing", ReadFont());
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(HtmlConversionDocument.Parse(Style("p{font-family:Missing}") + "<p>Visible text</p>"), options);
        Assert.False(rendered.HasLoss);
    }

    [Theory]
    [InlineData("Good", false)]
    [InlineData("Broken,Good", true)]
    public async Task InvalidFontProgramIsLossBearingOnlyWhenRequested(string family, bool expectedLoss) {
        HtmlRenderDocument rendered = await RenderAsync("p{font-family:" + family + "}", "<p>Visible text</p>",
            "@font-face{font-family:Broken;src:url('data:font/ttf;base64,AAECAw==')}", missingFamily: null);
        Assert.Equal(expectedLoss, rendered.HasLoss);
        HtmlDiagnostic unsupported = Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.FontFormatUnsupported);
        Assert.Equal(expectedLoss ? HtmlDiagnosticSeverity.Warning : HtmlDiagnosticSeverity.Info, unsupported.Severity);
    }

    [Fact]
    public async Task InvalidEarlierSourceDoesNotDegradeUsableFace() {
        HtmlRenderDocument rendered = await RenderAsync("p{font-family:Fallback}", "<p>Visible text</p>",
            "@font-face{font-family:Fallback;src:url('data:font/ttf;base64,AAECAw=='),url('https://font.example/good.ttf')}", missingFamily: null);
        Assert.False(rendered.HasLoss);
        Assert.Equal(HtmlDiagnosticSeverity.Info, Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.FontFormatUnsupported).Severity);
    }

    [Fact]
    public async Task FrameFontUseDoesNotPromoteUnusedParentDeclaration() {
        string child = System.Net.WebUtility.HtmlEncode("<style>@font-face{font-family:Missing;src:url('https://font.example/missing.ttf')}p{font-family:Missing}</style><p>Frame text</p>");
        HtmlRenderDocument rendered = await RenderAsync("p{font-family:Good}", "<p>Parent text</p><iframe srcdoc=\"" + child + "\"></iframe>");
        HtmlDiagnostic[] unavailable = rendered.Diagnostics.Where(x => x.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable).ToArray();
        Assert.Equal(2, unavailable.Length);
        Assert.Equal(HtmlDiagnosticSeverity.Info, unavailable[0].Severity);
        Assert.Equal(HtmlDiagnosticSeverity.Warning, unavailable[1].Severity);
    }

    [Theory]
    [InlineData("Good,Missing", false)]
    [InlineData("Missing,Good", true)]
    public async Task PdfReportPreservesFontUseAttribution(string families, bool expectedLoss) {
        var options = new HtmlToPdfOptions {
            ResourceResolver = CreateOptions().ResourceResolver,
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreateTrustedHost()
        };
        var result = await HtmlConversionDocument.Parse(Style("p{font-family:" + families + "}") + "<p>Visible text</p>")
            .ToPdfDocumentResultAsync(options);
        var unavailable = Assert.Single(result.Report.Warnings, x => x.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable);
        Assert.Equal(expectedLoss ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.None, unavailable.LossKind);
        Assert.Equal(expectedLoss, result.Report.HasLoss);
        Assert.Contains("Visible text", OfficeIMO.Pdf.PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task PageMarginTextAttributesUnavailableHeaderFace() {
        var options = CreateOptions();
        options.Mode = HtmlRenderMode.Paged;
        options.Margins = HtmlRenderMargins.All(40);
        var rendered = await HtmlRenderEngine.RenderAsync(HtmlConversionDocument.Parse(Style("p{font-family:Good}@page{@top-center{content:'Header';font-family:Missing}}") + "<p>Body</p>"), options);
        Assert.Equal(HtmlDiagnosticSeverity.Warning, Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable).Severity);
    }

    [Theory]
    [InlineData("value='\u03a9'", "input{font-family:Missing,Good}")]
    [InlineData("placeholder='\u03a9'", "input{font-family:Good}input::placeholder{font-family:Missing,Good}")]
    public async Task InputTextAttributesUnavailableUnicodeSubset(string attributes, string css) {
        var rendered = await RenderAsync(css, "<input " + attributes + ">",
            "@font-face{font-family:Missing;unicode-range:U+0370-03FF;src:url('https://font.example/missing.ttf')}", missingFamily: null);
        Assert.Equal(HtmlDiagnosticSeverity.Warning, Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable).Severity);
    }

    [Fact]
    public async Task CombiningTextRequiresOneCoveringFaceRatherThanIndependentScalars() {
        var source = HtmlConversionDocument.Parse("<style>"
            + "@font-face{font-family:Good;unicode-range:U+0065;src:url('https://font.example/good.ttf')}"
            + "@font-face{font-family:Good;unicode-range:U+0301;src:url('https://font.example/good.ttf')}"
            + "@font-face{font-family:Missing;src:url('https://font.example/missing.ttf')}"
            + "p{font-family:Good,Missing}</style><p>e\u0301</p>");
        var rendered = await HtmlRenderEngine.RenderAsync(source, CreateOptions());
        Assert.Equal(HtmlDiagnosticSeverity.Warning, Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable).Severity);
    }

    [Fact]
    public async Task QuotedFamilyContainingCommaSatisfiesTextBeforeMissingFallback() {
        var rendered = await RenderAsync("p{font-family:'Good, Family',Missing}", "<p>Visible text</p>",
            "@font-face{font-family:'Good, Family';src:url('https://font.example/good.ttf')}");
        Assert.False(rendered.HasLoss);
        Assert.Equal(HtmlDiagnosticSeverity.Info, Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable).Severity);
    }

    private static async Task<HtmlRenderDocument> RenderAsync(string css, string body, string extraFaces = "", string? missingFamily = "Missing") =>
        await HtmlRenderEngine.RenderAsync(HtmlConversionDocument.Parse(Style(css, extraFaces, missingFamily) + body), CreateOptions());

    private static string Style(string css, string extraFaces = "", string? missingFamily = "Missing") =>
        "<style>@font-face{font-family:Good;src:url('https://font.example/good.ttf')}"
        + (missingFamily == null ? "" : "@font-face{font-family:" + missingFamily + ";src:url('https://font.example/missing.ttf')}")
        + extraFaces + css + "</style>";

    private static HtmlRenderOptions CreateOptions() {
        byte[] font = ReadFont();
        return new HtmlRenderOptions {
            ResourceResolver = (request, cancellationToken) => Task.FromResult<HtmlResolvedResource?>(
                request.Uri.AbsolutePath == "/good.ttf" ? new HtmlResolvedResource(font, "font/ttf") : null)
        };
    }

    private static byte[] ReadFont() => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "RobotoFlex.ttf"));
}
