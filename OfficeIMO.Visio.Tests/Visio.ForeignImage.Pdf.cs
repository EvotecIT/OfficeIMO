using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class VisioForeignImageRenderingTests {
    [Fact]
    public void IndependentLegacyBitmapSurvivesPdfPreviewProjectionAndPackageRoundTrip() {
        using var input = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "angularjs-simple-scope.vdx"));
        VisioDocument document = VisioDocument.LoadLegacyXml(input).Value;
        foreach (VisioDocument candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            var options = new VisioDocumentProjectionOptions {
                IncludePngPreviewAssets = true,
                PngOptions = new VisioPngSaveOptions { RenderText = false, RenderStencilArtwork = false, Supersampling = 1 }
            };
            OfficeDocumentModel model = candidate.ToOfficeDocumentModel("independent.vdx", options);
            byte[] expected = Assert.Single(model.Assets).PayloadBytes!;
            PdfDocumentConversionResult result = candidate.ToPdfDocumentResult(new VisioToPdfOptions {
                SourceName = "independent.vdx", VisioOptions = options
            });
            byte[] pdf = result.ToBytes();
            PdfReadDocument read = PdfReadDocument.Open(pdf);
            PdfExtractedImage embedded = Assert.Single(read.Pages.SelectMany(page => page.GetImages()));
            Assert.True(OfficeRasterImageDecoder.TryDecode(expected, out OfficeRasterImage? expectedImage));
            Assert.True(OfficeRasterImageDecoder.TryDecode(embedded.Bytes, out OfficeRasterImage? actualImage));
            Assert.Equal(expectedImage!.Width, actualImage!.Width);
            Assert.Equal(expectedImage.Height, actualImage.Height);
            for (int y = 0; y < expectedImage.Height; y++)
                for (int x = 0; x < expectedImage.Width; x++)
                    Assert.Equal(expectedImage.GetPixel(x, y), actualImage.GetPixel(x, y));
            PdfConversionWarning normalization = Assert.Single(result.Warnings, warning => warning.Code == "VISIO_BITMAP_SIZE_NORMALIZED");
            Assert.Equal(OfficeConversionLossKind.None, normalization.LossKind);
            Assert.Equal("1", normalization.Details["sourcePage"]);
            Assert.Contains("Scope", read.ExtractText());
            Assert.Contains(result.Warnings, warning => warning.Code == "pdf-projection-visio-preview-embedded");
            Assert.False(result.Report.HasLoss);
            result.RequireNoLoss();
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PreviewOmissionsRemainLocatedAndLossBearingThroughNeutralReaderAndPdfModels(bool svg) {
        using var input = Fixture();
        var xml = XDocument.Load(input);
        xml.Descendants(Ns + "ForeignData").Single().SetAttributeValue("ForeignType", "Object");
        var document = Load(xml);
        var options = new VisioDocumentProjectionOptions {
            IncludeSvgPreviewAssets = svg, IncludePngPreviewAssets = !svg,
            SvgOptions = new VisioSvgSaveOptions { RenderText = false },
            PngOptions = new VisioPngSaveOptions { RenderText = false, PixelsPerInch = 20, Supersampling = 1 }
        };

        OfficeDocumentModel model = document.ToOfficeDocumentModel("opaque.vdx", options);
        OfficeDocumentModelDiagnostic diagnostic = Assert.Single(model.Diagnostics,
            item => item.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
        var reader = document.ToOfficeDocumentReadResult("opaque.vdx", visioOptions: new ReaderVisioOptions {
            IncludeSvgPreviewAssets = svg, IncludePngPreviewAssets = !svg,
            SvgOptions = options.SvgOptions, PngOptions = options.PngOptions
        });
        var readerDiagnostic = Assert.Single(reader.Diagnostics, item => item.Code == diagnostic.Code);
        PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(new VisioToPdfOptions {
            SourceName = "opaque.vdx", VisioOptions = options
        });
        conversion.ToBytes();

        Assert.Equal("opaque.vdx", diagnostic.Location!.Path);
        Assert.Equal(1, diagnostic.Location.Page);
        Assert.Equal(svg ? "preview-svg" : "preview-png", diagnostic.Location.SourceBlockKind);
        Assert.Equal("Omission", diagnostic.Attributes["lossKind"]);
        Assert.Equal("Omission", readerDiagnostic.Attributes["lossKind"]);
        PdfConversionWarning warning = Assert.Single(conversion.Warnings, item => item.Code == diagnostic.Code);
        Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
        Assert.True(conversion.Report.HasLoss);
        Assert.Equal("opaque.vdx", warning.Details["sourcePath"]);
        if (svg) Assert.Contains(conversion.Warnings, item => item.Code == "pdf-projection-asset-listed-not-embedded");
    }

    [Fact]
    public void PreviewRasterReductionReportsTheActualAssetLimitWithoutMutatingOptions() {
        using var input = Fixture();
        var document = VisioDocument.LoadLegacyXml(input).Value;
        var pngOptions = new VisioPngSaveOptions {
            RenderText = false, Supersampling = 1, MaximumRasterPixels = 4096
        };
        var options = new VisioDocumentProjectionOptions { IncludePngPreviewAssets = true, PngOptions = pngOptions };

        OfficeDocumentModel model = document.ToOfficeDocumentModel(options: options);
        Assert.True(OfficeImageReader.TryIdentify(Assert.Single(model.Assets).PayloadBytes!, ".png", out OfficeImageInfo info));
        Assert.True((long)info.Width * info.Height <= 4096);
        Assert.Contains(model.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.RasterScaleReduced && item.Attributes["lossKind"] == "Approximation");
        var result = document.ToPdfDocumentResult(new VisioToPdfOptions { VisioOptions = options });
        result.ToBytes();
        Assert.Contains(result.Warnings, item => item.Code == OfficeImageExportDiagnosticCodes.RasterScaleReduced && item.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Equal(96, pngOptions.PixelsPerInch);
        Assert.Equal(4096, pngOptions.MaximumRasterPixels);
        Assert.False(pngOptions.CancellationToken.IsCancellationRequested);
    }

    [Fact]
    public void OperationCancellationReachesSvgPreviewDecoding() {
        var document = WithImage(Convert.FromBase64String(AnimatedWebp));
        using var cancellation = new CancellationTokenSource();
        var codec = new CancelingPreviewCodec(cancellation);
        var svgOptions = new VisioSvgSaveOptions { ImageCodec = codec, RenderText = false };

        Assert.Throws<OperationCanceledException>(() => document.ToOfficeDocumentModel(
            options: new VisioDocumentProjectionOptions { IncludeSvgPreviewAssets = true, SvgOptions = svgOptions },
            cancellationToken: cancellation.Token));

        Assert.True(codec.WasCalled);
        Assert.False(svgOptions.CancellationToken.IsCancellationRequested);
    }

    private sealed class CancelingPreviewCodec(CancellationTokenSource cancellation) : IOfficeRasterImageCodec {
        public bool WasCalled { get; private set; }
        public bool TryDecode(byte[] bytes, string? contentType, out OfficeRasterImage? image) {
            WasCalled = true;
            cancellation.Cancel();
            image = null;
            return false;
        }
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    public void PreviewFontDiagnosticsDescribeOnlyRenderedShapeTextAndConnectorLabels(bool svg, bool renderText, bool renderLabels) {
        using var input = Fixture();
        var document = VisioDocument.LoadLegacyXml(input).Value;
        var page = document.Pages[0];
        page.Shapes[0].Text = "Shape text";
        page.Shapes[0].TextStyle = new VisioTextStyle { FontFamily = "OfficeIMO Missing Shape Font" };
        page.Connectors.Add(new VisioConnector("2", new OfficePoint(0.5, 0.5), new OfficePoint(3.5, 0.5)) {
            Label = "Connector label",
            TextStyle = new VisioTextStyle { FontFamily = "OfficeIMO Missing Label Font" }
        });
        var options = new VisioDocumentProjectionOptions {
            IncludeSvgPreviewAssets = svg, IncludePngPreviewAssets = !svg,
            SvgOptions = new VisioSvgSaveOptions { RenderText = renderText, RenderConnectorLabels = renderLabels },
            PngOptions = new VisioPngSaveOptions {
                RenderText = renderText, RenderConnectorLabels = renderLabels, PixelsPerInch = 20, Supersampling = 1
            }
        };

        OfficeDocumentModel model = document.ToOfficeDocumentModel(options: options);
        var fontDiagnostics = model.Diagnostics.Where(item => item.Code == OfficeImageExportDiagnosticCodes.FontSubstituted).ToArray();

        Assert.Equal(renderText, fontDiagnostics.Any(item => item.Message.Contains("OfficeIMO Missing Shape Font")));
        Assert.Equal(renderLabels, fontDiagnostics.Any(item => item.Message.Contains("OfficeIMO Missing Label Font")));
        if (!svg && !renderText && !renderLabels) {
            var result = document.ToPdfDocumentResult(new VisioToPdfOptions { VisioOptions = options });
            result.RequireNoLoss().ToBytes();
            Assert.False(result.Report.HasLoss);
        }
    }
}
