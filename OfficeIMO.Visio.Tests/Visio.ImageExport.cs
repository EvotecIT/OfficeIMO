using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioImageExport {
    [Fact]
    public void EmbeddedSvgPreviewRejectsExcessiveNestingWithoutRecursingToTheProcessStack() {
        const int depth = 512;
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                     string.Concat(Enumerable.Repeat("<g>", depth)) +
                     "<rect width='10' height='10' fill='red'/>" +
                     string.Concat(Enumerable.Repeat("</g>", depth)) +
                     "</svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        bool rendered = VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg),
            imageResolver: null,
            outlineFont: null,
            fonts: null,
            textShapingProvider: null,
            textShapingLanguage: null,
            diagnosticSink: diagnostics,
            diagnosticSource: "nested.svg",
            cancellationToken: default,
            out OfficeRasterImage? image);

        Assert.False(rendered);
        Assert.Null(image);
        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void EmbeddedSvgPreviewBoundsAuxiliaryRecursiveWalkers() {
        const int depth = 512;
        string nestedText = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><text x='1' y='5'>" +
                            string.Concat(Enumerable.Repeat("<tspan>", depth)) + "x" +
                            string.Concat(Enumerable.Repeat("</tspan>", depth)) + "</text></svg>";
        string nestedClip = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><defs><clipPath id='c'>" +
                            string.Concat(Enumerable.Repeat("<g>", depth)) + "<rect width='10' height='10'/>" +
                            string.Concat(Enumerable.Repeat("</g>", depth)) +
                            "</clipPath></defs><rect width='10' height='10' clip-path='url(#c)'/></svg>";
        var gradient = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink' width='10' height='10'><defs><linearGradient id='g0' x1='0'><stop offset='0' stop-color='red'/></linearGradient>");
        for (int index = 1; index <= depth; index++) {
            gradient.Append("<linearGradient id='g").Append(index).Append("' xlink:href='#g").Append(index - 1).Append("'");
            if (index == depth) {
                gradient.Append("><stop offset='0' stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient>");
            } else {
                gradient.Append("/>");
            }
        }
        gradient.Append("</defs><rect width='10' height='10' fill='url(#g").Append(depth).Append(")'/></svg>");

        foreach ((string Name, string Svg) scenario in new[] {
            ("text", nestedText),
            ("clip", nestedClip),
            ("gradient", gradient.ToString())
        }) {
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            bool rendered = VisioSvgPreviewRasterizer.TryRasterize(
                Encoding.UTF8.GetBytes(scenario.Svg), null, null, null, null, null,
                diagnostics, "recursive.svg", default, out OfficeRasterImage? image);

            Assert.False(rendered, scenario.Name);
            Assert.Null(image);
            Assert.Contains(diagnostics, diagnostic =>
                diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
                diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Fact]
    public void EmbeddedSvgPreviewBoundsAggregateCssSelectorEvaluation() {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><style>");
        for (int index = 0; index < 400; index++) {
            svg.Append("[data-effect-").Append(index).Append("] { filter:url(#f); }");
        }
        svg.Append("</style>");
        for (int index = 0; index < 300; index++) {
            svg.Append("<rect width='1' height='1'/>");
        }
        svg.Append("</svg>");
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.False(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg.ToString()), null, null, null, null, null,
            diagnostics, "selector-budget.svg", default, out OfficeRasterImage? image));
        Assert.Null(image);
        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void EmbeddedSvgPreviewCountsClipPathNodesAgainstTheElementBudget() {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><defs><clipPath id='c'>");
        for (int index = 0; index <= 100000; index++) svg.Append("<g/>");
        svg.Append("</clipPath></defs><rect width='10' height='10' clip-path='url(#c)'/></svg>");
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.False(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg.ToString()), null, null, null, null, null,
            diagnostics, "clip-budget.svg", default, out OfficeRasterImage? image));
        Assert.Null(image);
        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void EmbeddedSvgPreviewReportsUnsupportedVisualEffects() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                           "<rect width='10' height='10' fill='red' filter='url(#blur)'/>" +
                           "</svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        bool rendered = VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg),
            imageResolver: null,
            outlineFont: null,
            fonts: null,
            textShapingProvider: null,
            textShapingLanguage: null,
            diagnosticSink: diagnostics,
            diagnosticSource: "effect.svg",
            cancellationToken: default,
            out OfficeRasterImage? image);

        Assert.True(rendered);
        Assert.NotNull(image);
        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Theory]
    [InlineData("style='filter:url(#blur)'")]
    [InlineData("mask='url(#mask)'")]
    [InlineData("class='blur'")]
    public void EmbeddedSvgPreviewReportsTextSpanVisualEffects(string textSpanAttributes) {
        string style = textSpanAttributes.Contains("class=", StringComparison.Ordinal)
            ? "<style>.blur { filter:url(#blur); }</style>"
            : string.Empty;
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" + style +
                     "<text x='1' y='5'><tspan " + textSpanAttributes + ">x</tspan></text></svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "tspan-effect.svg", default, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void EmbeddedSvgPreviewCountsSiblingTextSpansAgainstTheElementBudget() {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><text x='1' y='5'>");
        for (int index = 0; index < 100001; index++) svg.Append("<tspan>x</tspan>");
        svg.Append("</text></svg>");
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.False(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg.ToString()), null, null, null, null, null,
            diagnostics, "wide-text.svg", default, out OfficeRasterImage? image));
        Assert.Null(image);
        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void EmbeddedSvgPreviewReportsUnsupportedTextChildren() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                           "<text><textPath href='#path'>unsupported</textPath></text></svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.False(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "text-path.svg", default, out _));
        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void EmbeddedSvgPreviewTreatsTextAnchorsAsSupportedContainers() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                           "<text x='1' y='5'><a href='https://example.test/'><tspan>x</tspan></a></text></svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "text-anchor.svg", default, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.DoesNotContain(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss);
    }

    [Theory]
    [InlineData("", "style='mask:url(#mask)'")]
    [InlineData("<style>.blur { filter: url(#blur); }</style>", "class='blur'")]
    [InlineData("<style>svg rect { filter: url(#blur); }</style>", "")]
    [InlineData("<style>svg > rect { mask: url(#mask); }</style>", "")]
    [InlineData("<style>* { filter: url(#blur); }</style>", "")]
    [InlineData("<style>[data-effect] { mask: url(#mask); }</style>", "data-effect='true'")]
    [InlineData("<style>rect:last-child { filter: url(#blur); }</style>", "")]
    public void EmbeddedSvgPreviewReportsCssVisualEffects(string styleDefinition, string rectangleAttributes) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                     styleDefinition + "<rect width='10' height='10' fill='red' " + rectangleAttributes + "/></svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "css-effect.svg", default, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void EmbeddedSvgPreviewDoesNotReportEffectsFromUnmatchedCssSelectors() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                           "<style>.unused { filter: url(#blur); } [data-unused] { mask: url(#mask); }</style>" +
                           "<rect width='10' height='10' fill='red'/></svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "unused-css-effect.svg", default, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.DoesNotContain(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Theory]
    [InlineData(".blur", "class='Blur'")]
    [InlineData("RECT", "")]
    [InlineData("[DATA-effect]", "data-effect='true'")]
    public void EmbeddedSvgPreviewMatchesXmlSelectorsCaseSensitively(
        string selector,
        string rectangleAttributes) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                     "<style>" + selector + " { filter: url(#blur); }</style>" +
                     "<rect width='10' height='10' fill='red' " + rectangleAttributes + "/></svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "case-sensitive-selector.svg", default, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.DoesNotContain(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Theory]
    [InlineData("<style>.blur { filter:url(#blur); } .blur { filter:none; }</style>", "class='blur'")]
    [InlineData("<style>.blur { filter:url(#blur); }</style>", "class='blur' style='filter:none'")]
    [InlineData("<style>.blur { filter:none; }</style>", "class='blur' filter='url(#blur)'")]
    [InlineData("<style>.blur { filter:url(#blur) !important; }</style>", "class='blur' style='filter:none !important'")]
    [InlineData("", "style='filter:initial'")]
    [InlineData("", "style='mask:unset'")]
    [InlineData("", "style='filter:revert'")]
    [InlineData("", "style='mask:revert-layer'")]
    [InlineData("", "style='filter:inherit'")]
    public void EmbeddedSvgPreviewHonorsCssEffectOverrides(string styleDefinition, string rectangleAttributes) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                     styleDefinition + "<rect width='10' height='10' fill='red' " + rectangleAttributes + "/></svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "css-override.svg", default, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.DoesNotContain(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Theory]
    [InlineData("<style>.blur { filter:url(#blur)!important;filter:none; }</style>", "class='blur'")]
    [InlineData("", "style='filter:url(#blur)!important;filter:none'")]
    public void EmbeddedSvgPreviewPreservesImportantEffectsAcrossDuplicateDeclarations(
        string styleDefinition,
        string rectangleAttributes) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                     styleDefinition + "<rect width='10' height='10' fill='red' " + rectangleAttributes + "/></svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "css-important.svg", default, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void EmbeddedSvgPreviewIgnoresNonvisualMetadata() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                           "<metadata><producer>OfficeIMO</producer></metadata>" +
                           "<rect width='10' height='10' fill='red'/></svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "metadata.svg", default, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.DoesNotContain(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss);
    }

    [Fact]
    public void EmbeddedSvgPreviewReportsLossWhenOnlyUnsupportedContentIsVisible() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                           "<foreignObject width='10' height='10'/></svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.False(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "unsupported-only.svg", default, out _));
        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void EmbeddedSvgPreviewTreatsAnchorContainersAsSupportedGroups() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                           "<a href='https://example.test/'><rect width='10' height='10' fill='red'/></a></svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "anchor.svg", default, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.DoesNotContain(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss);
    }

    [Fact]
    public void RetainedSvgApiUsesCanonicalDimensionValidation() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Bounds").Size(2, 1);
        page.AddRectangle(1, 0.5, 1, 0.5, "Bounds");

        string highResolutionSvg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 65536D });
        Assert.Contains("<svg", highResolutionSvg, StringComparison.Ordinal);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = double.MaxValue }));
    }

    [Fact]
    public async System.Threading.Tasks.Task SaveAsSvgAsync_CancellationDoesNotMutateCallerOptions() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Reusable options").Size(2, 1);
        page.AddRectangle(1, 0.5, 1.5, 0.6, "Reusable");
        var options = new VisioSvgSaveOptions();
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();
        using var output = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            page.SaveAsSvgAsync(output, options, cancellation.Token));

        string svg = page.ToSvg(options);
        Assert.Contains("Reusable", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void DirectImageExportEnforcesRenderTimeout() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Deadline").Size(2, 1);
        var options = new VisioImageExportOptions { RenderTimeout = TimeSpan.FromTicks(1) };

        Assert.Throws<OfficeImageExportTimeoutException>(() =>
            page.ExportImage(OfficeImageExportFormat.Svg, options));
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void PageExport_EncodesEverySharedImageFormat(OfficeImageExportFormat format) {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Formats").Size(2, 1);
        page.AddRectangle(1, 0.5, 1.2, 0.5, "Format marker");
        var options = new VisioImageExportOptions { Scale = 0.5D, Supersampling = 1 };

        OfficeImageExportResult result = page.ExportImage(format, options);

        Assert.Equal(format, result.Format);
        Assert.Equal(96, result.Width);
        Assert.Equal(48, result.Height);
        if (format == OfficeImageExportFormat.Svg) {
            Assert.Contains("<svg", Encoding.UTF8.GetString(result.Bytes), StringComparison.Ordinal);
        } else {
            Assert.Equal(format.GetMimeType(), OfficeImageReader.Identify(result.Bytes).MimeType);
        }
    }

    [Fact]
    public void PageFitWithinAndConnectorOptionsApplyToRasterAndSvg() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Bounded").Size(4, 2);
        page.AddRectangle(2, 1, 2, 1, "Bounded");

        OfficeImageExportResult png = page.ToImage()
            .WithScale(2D)
            .FitWithin(240, 240)
            .IncludeConnectorLabels(false)
            .AsPng()
            .Export();
        OfficeImageExportResult svg = page.ToImage()
            .WithScale(2D)
            .FitWithin(240, 240)
            .ResolveConnectorLabelOverlaps(false)
            .AsSvg()
            .Export();

        Assert.Equal((240, 120), (png.Width, png.Height));
        Assert.Equal((png.Width, png.Height), (svg.Width, svg.Height));
    }

    [Fact]
    public void DocumentFluentBatchExport_SelectsPagesAndSavesPortableResults() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        document.AddPage("First").Size(2, 1).AddRectangle(1, 0.5, 1, 0.5, "One");
        document.AddPage("Second").Size(2, 1).AddEllipse(1, 0.5, 1, 0.5, "Two");
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        try {
            IReadOnlyList<OfficeImageExportResult> results = document.ToImages()
                .FromPage(1)
                .TakePages(1)
                .AtDpi(48D)
                .AsWebp()
                .Save(folder);

            OfficeImageExportResult result = Assert.Single(results);
            Assert.Equal("Second", result.Name);
            Assert.Equal("Visio page 2", result.Source);
            Assert.Equal("image/webp", OfficeImageReader.Identify(result.Bytes).MimeType);
            Assert.True(File.Exists(Path.Combine(folder, "Second.webp")));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void DocumentAndPageConvenienceMethods_ProduceRequestedRasterFormats() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Convenience").Size(2, 1);
        page.AddDiamond(1, 0.5, 1, 0.6, "Choice");
        var options = new VisioImageExportOptions { Scale = 0.5D, Supersampling = 1 };

        Assert.Equal("image/jpeg", OfficeImageReader.Identify(page.ToJpeg(options)).MimeType);
        Assert.Equal("image/tiff", OfficeImageReader.Identify(document.ToTiff(options)).MimeType);
        Assert.Equal("image/webp", OfficeImageReader.Identify(page.ToWebp(options)).MimeType);
        Assert.Equal(OfficeImageExportFormat.Svg, document.ToImage().AtDpi(48D).AsSvg().Export().Format);
    }

    [Fact]
    public void RasterExport_ReducesScaleWithAVisibleDiagnosticInsteadOfOverAllocating() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Large").Size(100, 100);
        page.AddRectangle(50, 50, 80, 80, "Large");
        var options = new VisioImageExportOptions {
            Scale = 1D,
            Supersampling = 1,
            MaximumRasterPixels = 10_000L
        };

        OfficeImageExportResult result = page.ExportImage(OfficeImageExportFormat.Png, options);

        Assert.True((long)result.Width * result.Height <= options.MaximumRasterPixels);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.RasterScaleReduced);
    }

    [Fact]
    public void ImageExportOptions_CloneRasterSettingsAcrossDocumentSelection() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        document.AddPage("Options").Size(2, 1).AddRectangle(1, 0.5, 1, 0.5, "Options");
        var options = new VisioImageExportOptions { Scale = 0.5D, Supersampling = 1 };
        options.RasterEncoding.Jpeg.Quality = 67;
        options.RasterEncoding.Jpeg.Progressive = true;

        OfficeImageExportResult result = document.ExportImage(OfficeImageExportFormat.Jpeg, options);

        Assert.Equal("image/jpeg", OfficeImageReader.Identify(result.Bytes).MimeType);
        Assert.Contains((byte)0xC2, result.Bytes);
    }

    [Fact]
    public void SvgExport_RejectsOverflowingScaleBeforeRendering() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Overflow").Size(2, 1);
        var options = new VisioImageExportOptions { Scale = double.MaxValue };

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            page.ExportImage(OfficeImageExportFormat.Svg, options));
    }

    [Fact]
    public void SharedFontsDriveVisioRasterAndAreEmbeddedInSvgWithoutSubstitution() {
        OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadDefault(out string? fontPath);
        if (font == null || string.IsNullOrWhiteSpace(fontPath)) return;

        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Typography").Size(3, 1.5);
        VisioShape shape = page.AddRectangle(1.5, 0.75, 2.4, 0.7, "Scoped Visio text");
        shape.TextStyle = new VisioTextStyle {
            FontFamily = "Scoped Visio",
            Size = 18D
        };
        var options = new VisioImageExportOptions {
            Supersampling = 1
        };
        options.Fonts.Add("Scoped Visio", File.ReadAllBytes(fontPath));

        OfficeImageExportResult png = page.ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult svg = page.ExportImage(OfficeImageExportFormat.Svg, options);
        string svgText = Encoding.UTF8.GetString(svg.Bytes);

        Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.FontSubstituted);
        Assert.DoesNotContain(svg.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.FontSubstituted);
        Assert.Contains("@font-face", svgText, StringComparison.Ordinal);
        Assert.Contains("Scoped Visio", svgText, StringComparison.Ordinal);
        Assert.Contains(Convert.ToBase64String(File.ReadAllBytes(fontPath)), svgText, StringComparison.Ordinal);
    }

    [Fact]
    public void DirectVisioExportAppliesFontLossPolicyBeforeReturning() {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Policy").Size(2, 1);
        VisioShape shape = page.AddRectangle(1, 0.5, 1.5, 0.6, "Policy");
        shape.TextStyle = new VisioTextStyle { FontFamily = "OfficeIMO Definitely Missing" };
        var options = new VisioImageExportOptions {
            Supersampling = 1,
            Policy = new OfficeImageExportPolicy { RequireNoLoss = true }
        };

        OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(
            () => page.ExportImage(OfficeImageExportFormat.Png, options));

        Assert.Contains(exception.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.FontSubstituted);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg)]
    public void PackagePreview_UsesCallerCodecThroughCanonicalExport(OfficeImageExportFormat format) {
        using MemoryStream package = new();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("CallerCodec").Size(2, 1);
        AddCustomPreviewShape(page);
        var codec = new SolidImageCodec(OfficeColor.FromRgb(12, 90, 180));
        var options = new VisioImageExportOptions {
            Scale = 0.5D,
            Supersampling = 1,
            ImageCodec = codec
        };

        OfficeImageExportResult result = page.ExportImage(format, options);

        Assert.Equal(1, codec.DecodeCalls);
        Assert.Contains(
            result.Diagnostics,
            diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodedByCallerCodec);
        Assert.DoesNotContain(
            result.Diagnostics,
            diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
        if (format == OfficeImageExportFormat.Svg) {
            Assert.Contains("data:image/png;base64,", Encoding.UTF8.GetString(result.Bytes), StringComparison.Ordinal);
        }
    }

    private static void AddCustomPreviewShape(VisioPage page) {
        VisioMaster master = new("custom-preview", "CustomPreview", new VisioShape("master-shape", 0, 0, 1, 1, string.Empty));
        master.RawMasterRelationships.Add(new VisioAssets.MasterRelationshipContent {
            Id = "rIdImage",
            Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            Target = "../media/preview.custom",
            ContentType = "application/octet-stream",
            Extension = ".custom",
            Data = new byte[] { 1, 2, 3, 4 }
        });
        VisioShape shape = page.AddRectangle(1, 0.5, 1, 0.7, string.Empty);
        shape.Master = master;
        shape.NameU = master.NameU;
        shape.SetUserCell("OfficeIMO.StencilPreviewImageRelationshipId", "rIdImage", "STR");
        shape.SetUserCell("OfficeIMO.StencilPreviewImageTarget", "../media/preview.custom", "STR");
    }

    private sealed class SolidImageCodec : IOfficeRasterImageCodec {
        private readonly OfficeColor _color;

        internal SolidImageCodec(OfficeColor color) {
            _color = color;
        }

        internal int DecodeCalls { get; private set; }

        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            DecodeCalls++;
            image = new OfficeRasterImage(2, 2, _color);
            return true;
        }
    }
}
