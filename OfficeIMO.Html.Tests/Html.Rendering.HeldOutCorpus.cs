using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    public static IEnumerable<object[]> HeldOutRenderingCorpusCaseIds => HtmlRenderingHeldOutCorpus.Load().Cases
        .Select(item => new object[] { item.Id });

    [Theory]
    [MemberData(nameof(HeldOutRenderingCorpusCaseIds))]
    public void HeldOutRenderingCorpus_ExercisesEverySelectedIntentAndOutputFamily(string caseId) {
        HtmlRenderingHeldOutCase scenario = HtmlRenderingHeldOutCorpus.Load().Cases.Single(item => item.Id == caseId);
        HtmlConversionDocument source = scenario.LoadDocument();

        HtmlRenderOptions screenOptions = CreateHeldOutScreenOptions();
        ValidateImageResult(source, scenario, HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png, screenOptions);
        ValidateImageResult(source, scenario, HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Svg, screenOptions);

        HtmlRenderOptions printOptions = CreateHeldOutPrintOptions();
        ValidatePdfResult(source, scenario, HtmlRenderIntentProfile.PrintPaged, printOptions, scenario.Manifest.ExpectedPrintPageCount);
        ValidateImageResult(source, scenario, HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png, printOptions);
        ValidateImageResult(source, scenario, HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Svg, printOptions);

        HtmlRenderOptions snapshotOptions = CreateHeldOutSnapshotOptions();
        ValidatePdfResult(source, scenario, HtmlRenderIntentProfile.ScreenSnapshotPaged, snapshotOptions, expectedPageCount: null);
        ValidateImageResult(source, scenario, HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Png, snapshotOptions);
        ValidateImageResult(source, scenario, HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Svg, snapshotOptions);
    }

    private static void ValidatePdfResult(
        HtmlConversionDocument source,
        HtmlRenderingHeldOutCase scenario,
        HtmlRenderIntentProfile profile,
        HtmlRenderOptions options,
        int? expectedPageCount) {
        HtmlRenderRequest request = HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, options);
        HtmlPdfRenderRequestResult result = source.RenderToPdfResult(request);
        byte[] bytes = result.ToBytes();

        Assert.True(bytes.Length > 100);
        Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(bytes, 0, 5));
        Assert.True(result.RenderResult.Request.MatchesNamedProfile);
        if (expectedPageCount.HasValue) Assert.Equal(expectedPageCount.Value, result.RenderResult.Document.Pages.Count);
        ValidateHeldOutDocument(result.RenderResult.Document, scenario);
    }

    private static void ValidateImageResult(
        HtmlConversionDocument source,
        HtmlRenderingHeldOutCase scenario,
        HtmlRenderIntentProfile profile,
        HtmlRenderEncoder encoder,
        HtmlRenderOptions options) {
        HtmlRenderRequest request = HtmlRenderRequest.Create(profile, encoder, options);
        HtmlRenderResult result = HtmlRenderEngine.Execute(source, request);
        IReadOnlyList<OfficeImageExportResult> images = result.ExportImages();

        Assert.True(result.Request.MatchesNamedProfile);
        Assert.Equal(result.Document.Pages.Count, images.Count);
        Assert.All(images, image => {
            Assert.True(image.Bytes.Length > 100);
            if (encoder == HtmlRenderEncoder.Png) {
                Assert.Equal(new byte[] { 137, 80, 78, 71 }, image.Bytes.Take(4));
            } else {
                string prefix = System.Text.Encoding.UTF8.GetString(image.Bytes, 0, Math.Min(image.Bytes.Length, 200));
                Assert.Contains("<svg", prefix, StringComparison.OrdinalIgnoreCase);
            }
        });
        if (profile == HtmlRenderIntentProfile.PrintPaged) {
            Assert.Equal(scenario.Manifest.ExpectedPrintPageCount, result.Document.Pages.Count);
        }
        ValidateHeldOutDocument(result.Document, scenario);
    }

    private static void ValidateHeldOutDocument(HtmlRenderDocument document, HtmlRenderingHeldOutCase scenario) {
        string text = NormalizeCorpusWhitespace(document.Text);
        Assert.All(scenario.Manifest.TextMarkers, marker =>
            Assert.Contains(NormalizeCorpusWhitespace(marker), text, StringComparison.Ordinal));
        Assert.DoesNotContain(document.Diagnostics, diagnostic => diagnostic.Severity == HtmlDiagnosticSeverity.Error);
    }

    private static HtmlRenderOptions CreateHeldOutScreenOptions() => new() {
        Mode = HtmlRenderMode.Continuous,
        ViewportWidth = 768D,
        ViewportHeight = 900D,
        Margins = HtmlRenderMargins.All(0D),
        Scale = 1D,
        BackgroundColor = OfficeColor.White,
        UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
        FidelityPolicy = HtmlRenderFidelityPolicy.AllowDiagnosedLoss
    };

    private static HtmlRenderOptions CreateHeldOutPrintOptions() => new() {
        Mode = HtmlRenderMode.Paged,
        ViewportWidth = 768D,
        ViewportHeight = 900D,
        PageSize = new OfficePageSize(8.27D, 11.69D),
        Margins = HtmlRenderMargins.All(40D),
        Scale = 1D,
        BackgroundColor = OfficeColor.White,
        UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
        FidelityPolicy = HtmlRenderFidelityPolicy.AllowDiagnosedLoss
    };

    private static HtmlRenderOptions CreateHeldOutSnapshotOptions() => new() {
        Mode = HtmlRenderMode.Continuous,
        ViewportWidth = 768D,
        ViewportHeight = 900D,
        PageSize = new OfficePageSize(8D, 9.375D),
        Margins = HtmlRenderMargins.All(0D),
        HonorCssPageRules = false,
        Scale = 1D,
        BackgroundColor = OfficeColor.White,
        UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
        FidelityPolicy = HtmlRenderFidelityPolicy.AllowDiagnosedLoss
    };
}
