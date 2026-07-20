using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void RasterCanvas_UsesOneCachedProviderRunForMeasurementAndPainting() {
        byte[] fontData = CreateMinimalTrueTypeFont(CreateFormat12Cmap('A'));
        var fonts = new OfficeFontFaceCollection().Add("Shaping Demo", fontData);
        var provider = new RasterMappingTextShapingProvider(
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: 800, offsetX: 25, offsetY: 10));
        var canvas = new OfficeRasterCanvas(
            new OfficeRasterImage(120, 40, OfficeColor.White),
            font: null,
            fonts: fonts,
            textShapingProvider: provider,
            textShapingLanguage: "ar-SA");

        Assert.Equal(800D, canvas.MeasureText("A", 1000D, "Shaping Demo"));
        Assert.Equal(800D, canvas.MeasureText("A", 1000D, "Shaping Demo"));
        canvas.DrawTextLine(
            "A",
            0D,
            0D,
            24D,
            OfficeColor.Black,
            alignment: OfficeTextAlignment.Left,
            fontFamily: "Shaping Demo");

        OfficeTextShapingRequest request = Assert.Single(provider.Requests);
        Assert.Equal("A", request.Text);
        Assert.Equal("ar-SA", request.Language);
        Assert.Equal(OfficeTextDirection.LeftToRight, request.Direction);
        Assert.Equal(1000, request.UnitsPerEm);
        Assert.Equal(fontData, request.FontData);
    }

    [Fact]
    public void RasterCanvas_NormalizesNegativeProviderAdvanceForLeftBasedLayout() {
        var fonts = new OfficeFontFaceCollection()
            .Add("RTL Demo", CreateMinimalTrueTypeFont(CreateFormat12Cmap(0x05D0)));
        var provider = new RasterMappingTextShapingProvider(
            new OfficeShapedGlyph(1, "\u05D0", 0, advanceWidth: -700));
        var canvas = new OfficeRasterCanvas(
            new OfficeRasterImage(80, 30),
            font: null,
            fonts: fonts,
            textShapingProvider: provider);

        Assert.Equal(700D, canvas.MeasureText("\u05D0", 1000D, "RTL Demo"));
        Assert.Equal(OfficeTextDirection.RightToLeft, Assert.Single(provider.Requests).Direction);
    }

    [Fact]
    public void RasterCanvas_PaintsProviderSelectedGlyphContours() {
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var fonts = new OfficeFontFaceCollection().Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A'));
        var image = new OfficeRasterImage(80, 40, OfficeColor.White);
        var canvas = new OfficeRasterCanvas(
            image,
            font: null,
            fonts: fonts,
            textShapingProvider: provider);

        canvas.DrawTextLine(
            "A",
            2D,
            2D,
            32D,
            OfficeColor.Black,
            fontFamily: ManagedTextShapingTestAssets.FamilyName);

        Assert.Contains(image.GetPixels(), channel => channel == 0);
        Assert.Single(provider.Requests);
    }

    [Fact]
    public void RasterCanvas_RejectsProviderGlyphsOutsideTheSelectedFont() {
        var fonts = new OfficeFontFaceCollection()
            .Add("Shaping Demo", CreateMinimalTrueTypeFont(CreateFormat12Cmap('A')));
        var provider = new RasterMappingTextShapingProvider(
            new OfficeShapedGlyph(99, "A", 0, advanceWidth: 500));
        var canvas = new OfficeRasterCanvas(
            new OfficeRasterImage(80, 30),
            font: null,
            fonts: fonts,
            textShapingProvider: provider);

        ArgumentException exception = Assert.Throws<ArgumentException>(
            () => canvas.MeasureText("A", 12D, "Shaping Demo"));
        Assert.Contains("outside the selected font glyph range", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RasterCanvas_ObservesCancellationBeforeHostTextShaping() {
        var fonts = new OfficeFontFaceCollection().Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A'));
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var canvas = new OfficeRasterCanvas(
            new OfficeRasterImage(80, 30),
            font: null,
            fonts: fonts,
            textShapingProvider: provider,
            cancellationToken: cancellation.Token);

        Assert.Throws<OperationCanceledException>(
            () => canvas.MeasureText("A", 12D, ManagedTextShapingTestAssets.FamilyName));
        Assert.Empty(provider.Requests);
    }

    [Fact]
    public void RasterCanvas_IdentifiesTheSelectedTrueTypeCollectionFace() {
        byte[] collection = ManagedTextShapingTestAssets.CreateFontCollection('B');
        OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoad(collection, 1);
        Assert.NotNull(font);
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var canvas = new OfficeRasterCanvas(
            new OfficeRasterImage(80, 30),
            font!,
            fonts: null,
            textShapingProvider: provider);

        canvas.MeasureText("B", 12D);

        OfficeTextShapingRequest request = Assert.Single(provider.Requests);
        Assert.Equal(1, request.FontCollectionIndex);
        Assert.Equal(collection, request.FontData);
    }

    [Fact]
    public void DrawingRasterRenderer_CarriesShapingIntoNestedEffectScenes() {
        byte[] fontData = CreateMinimalTrueTypeFont(CreateFormat12Cmap('A'));
        var nested = new OfficeDrawing(40D, 20D)
            .AddFont("Shaping Demo", fontData)
            .AddText("A", 0D, 0D, 40D, 20D, new OfficeFontInfo("Shaping Demo", 12D));
        var drawing = new OfficeDrawing(40D, 20D)
            .AddEffectDrawing(nested, OfficeTransform.Identity);
        var provider = new RasterMappingTextShapingProvider(
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: 600));
        using var cancellation = new CancellationTokenSource();

        OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
            TextShapingProvider = provider,
            TextShapingLanguage = "en-US",
            CancellationToken = cancellation.Token
        });

        Assert.Contains(provider.Requests, request =>
            request.Text == "A" &&
            request.Language == "en-US" &&
            request.CancellationToken == cancellation.Token);
    }

    [Fact]
    public void SharedImageBuilders_ConfigureAndCloneTextShaping() {
        var provider = new RasterMappingTextShapingProvider(
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: 500));
        var singleOptions = new TestImageExportOptions();
        var batchOptions = new TestImageExportOptions();

        new TestImageExportBuilder(singleOptions).WithTextShaping(provider, " ar-SA ");
        new TestImageExportBatchBuilder(batchOptions, "page").WithTextShaping(provider, " pl-PL ");

        Assert.Same(provider, singleOptions.TextShapingProvider);
        Assert.Equal("ar-SA", singleOptions.TextShapingLanguage);
        Assert.Same(provider, batchOptions.TextShapingProvider);
        Assert.Equal("pl-PL", batchOptions.TextShapingLanguage);
    }

    [Fact]
    public void ManagedFallback_ContextualizesArabicAndProducesVisualGlyphOrder() {
        OfficeManagedTextFallback fallback = OfficeManagedTextShaper.Resolve(
            "اب",
            OfficeTrueTypeFont.TryLoad(ManagedTextShapingTestAssets.CreateFont(
                0x0627,
                0x0628,
                0xFE8D,
                0xFE8F))!);

        Assert.True(fallback.Used);
        Assert.False(fallback.Incomplete);
        Assert.Equal("\uFE8F\uFE8D", fallback.Text);
        Assert.Equal("123 \uFE8F\uFE8D", OfficeManagedTextShaper.ToVisualOrder(
            OfficeArabicTextShaper.Shape("اب 123")));
    }

    [Fact]
    public void ManagedFallback_PreservesArabicIndicDigitOrder() {
        Assert.Equal(
            "١٢٣ با",
            OfficeManagedTextShaper.ToVisualOrder("اب ١٢٣"));
    }

    [Fact]
    public void RasterCanvas_ReportsIncompleteFallbackForIndicShaping() {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var fonts = new OfficeFontFaceCollection().Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(0x0915, 0x093F));
        var canvas = new OfficeRasterCanvas(
            new OfficeRasterImage(120, 40, OfficeColor.White),
            font: null,
            fonts: fonts,
            diagnosticSink: diagnostics,
            diagnosticSource: "managed Devanagari test");

        canvas.MeasureText("कि", 18D, ManagedTextShapingTestAssets.FamilyName);

        OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics);
        Assert.Equal(OfficeImageExportDiagnosticCodes.TextShapingFallback, diagnostic.Code);
        Assert.Contains("cannot provide complete", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal(OfficeImageExportLossKind.Approximation, diagnostic.LossKind);
    }

    [Fact]
    public void RasterCanvas_ReportsOneManagedFallbackApproximation() {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var fonts = new OfficeFontFaceCollection().Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(
                0x0627,
                0x0628,
                0xFE8D,
                0xFE8F));
        var canvas = new OfficeRasterCanvas(
            new OfficeRasterImage(120, 40, OfficeColor.White),
            font: null,
            fonts: fonts,
            diagnosticSink: diagnostics,
            diagnosticSource: "managed Arabic test");

        canvas.MeasureText("اب", 18D, ManagedTextShapingTestAssets.FamilyName);
        canvas.DrawTextLine(
            "اب",
            0D,
            0D,
            18D,
            OfficeColor.Black,
            fontFamily: ManagedTextShapingTestAssets.FamilyName);

        OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics);
        Assert.Equal(OfficeImageExportDiagnosticCodes.TextShapingFallback, diagnostic.Code);
        Assert.Equal(OfficeImageExportLossKind.Approximation, diagnostic.LossKind);
        Assert.Equal("managed Arabic test", diagnostic.Source);
    }

    [Fact]
    public void RasterCanvas_ReportsIncompleteFallbackForUnsupportedJoiningScript() {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var fonts = new OfficeFontFaceCollection().Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(0x0710, 0x0720, 0x072B));
        var canvas = new OfficeRasterCanvas(
            new OfficeRasterImage(120, 40, OfficeColor.White),
            font: null,
            fonts: fonts,
            diagnosticSink: diagnostics);

        canvas.MeasureText("ܫܠܐ", 18D, ManagedTextShapingTestAssets.FamilyName);

        OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics);
        Assert.Equal(OfficeImageExportDiagnosticCodes.TextShapingFallback, diagnostic.Code);
        Assert.Contains("cannot provide complete", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RasterCanvas_HostProviderSuppressesManagedFallbackDiagnostic() {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var fonts = new OfficeFontFaceCollection().Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(0x0627, 0x0628));
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var canvas = new OfficeRasterCanvas(
            new OfficeRasterImage(120, 40, OfficeColor.White),
            font: null,
            fonts: fonts,
            textShapingProvider: provider,
            diagnosticSink: diagnostics);

        canvas.MeasureText("اب", 18D, ManagedTextShapingTestAssets.FamilyName);

        Assert.Empty(diagnostics);
        Assert.Equal("اب", Assert.Single(provider.Requests).Text);
    }

    private sealed class RasterMappingTextShapingProvider : IOfficeTextShapingProvider {
        private readonly OfficeTextShapingResult _result;

        internal RasterMappingTextShapingProvider(params OfficeShapedGlyph[] glyphs) {
            _result = new OfficeTextShapingResult(glyphs);
        }

        internal List<OfficeTextShapingRequest> Requests { get; } = new();

        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Requests.Add(request);
            return _result;
        }
    }
}
