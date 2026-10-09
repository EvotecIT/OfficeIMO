using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Visio.Pdf;
using Xunit;

namespace OfficeIMO.Visio.Tests;

public sealed class VisioDrawingReviewRegressionTests {
    [Theory]
    [InlineData("edge")]
    [InlineData("rotated")]
    [InlineData("route")]
    [InlineData("arrow")]
    public void EdgeCrossingArtworkKeepsItsVisiblePartInPdf(string kind) {
        VisioDocument source = VisioDocument.Create(); source.UseMastersByDefault = false;
        VisioPage page = source.AddPage("Edge", 3, 2);
        if (kind is "edge" or "rotated") {
            var shape = page.AddRectangle(0, 1, 1, 1);
            shape.FillColor = OfficeColor.Red;
            if (kind == "rotated") shape.Angle = .5;
        } else {
            var connector = page.AddConnector("route", new OfficePoint(kind == "route" ? -.5 : 1, 0), new OfficePoint(2, 0));
            connector.LineWeight = .05;
            if (kind == "arrow") connector.EndArrow = EndArrow.Triangle;
        }
        foreach (VisioDocument document in Reopened(source)) {
            byte[] before = document.ToLegacyXmlResult().Value;
            OfficeDrawing drawing = document.ToDrawings().Value[0];
            Assert.Contains("path", OfficeDrawingSvgExporter.ToSvg(drawing, 1, OfficeSvgSizeUnit.Point));
            byte[] png = drawing.ExportImage(OfficeImageExportFormat.Png, new OfficeImageExportOptions { BackgroundColor = OfficeColor.White }).Bytes;
            Assert.True(OfficeRasterImageDecoder.TryDecode(png, out OfficeRasterImage? raster));
            int ink = 0;
            for (int y = 0; y < raster!.Height; y++) for (int x = 0; x < raster.Width; x++) {
                OfficeColor color = raster.GetPixel(x, y);
                if (Math.Min(color.R, Math.Min(color.G, color.B)) < 180) ink++;
            }
            Assert.True(ink > 70, $"The visible {kind} artwork was lost during raster projection.");
            byte[] pdf = document.ToPdfBytes(new VisioToPdfOptions { Mode = VisioPdfProjectionMode.DiagramPages });
            Assert.Equal((216D, 144D), Assert.Single(PdfReadDocument.Open(pdf).Pages).GetPageSize());
            Assert.Equal(before, document.ToLegacyXmlResult().Value);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfFontOverridesUseTheSameResourcesForReportedClippingAndPaintedText(bool widePdf) {
        byte[] narrow = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] wide = WithWideAdvances(narrow);
        VisioDocument document = VisioDocument.Create(); document.UseMastersByDefault = false;
        var shape = document.AddPage("Font override").AddRectangle(2, 2, .5, .16, "AAAAAAAAAA");
        shape.TextStyle = new VisioTextStyle { Size = 6, FontFamily = "Arial", LeftMargin = 0, RightMargin = 0, TopMargin = 0, BottomMargin = 0 };
        var drawingOptions = new VisioDrawingOptions(); drawingOptions.Fonts.Add("Arial", widePdf ? narrow : wide);
        var pdfOptions = new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Arial", widePdf ? wide : narrow));
        var result = document.ToPdfDocumentResult(new VisioToPdfOptions {
            Mode = VisioPdfProjectionMode.DiagramPages, DrawingOptions = drawingOptions, PdfOptions = pdfOptions
        });
        string text = PdfReadDocument.Open(result.ToBytes()).ExtractText();
        Assert.Equal(!widePdf, text.Contains("AAAAAAAAAA"));
        Assert.Equal(widePdf, result.FidelityDiagnostics.Any(d => d.Code == "VISIO_DRAWING_TEXT_CLIPPED" && d.LossKind == OfficeConversionLossKind.Omission));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfTextPreflightLetsTheEffectiveShapingProviderObserveCancellation(bool pdfProvider) {
        using var cancellation = new CancellationTokenSource();
        var provider = new CancellingProvider(cancellation);
        var document = VisioDocument.Create(); document.UseMastersByDefault = false;
        document.AddPage("Cancellation").AddRectangle(2, 2, 1, 1, "A").TextStyle = new VisioTextStyle { FontFamily = "Arial" };
        var drawingOptions = new VisioDrawingOptions { TextShapingProvider = pdfProvider ? null : provider };
        drawingOptions.Fonts.Add("Arial", ManagedTextShapingTestAssets.CreateFont('A'));
        var options = new VisioToPdfOptions {
            Mode = VisioPdfProjectionMode.DiagramPages, DrawingOptions = drawingOptions,
            PdfOptions = new PdfOptions { TextShapingProvider = pdfProvider ? provider : null }
        };
        Assert.ThrowsAny<OperationCanceledException>(() => document.ToPdfDocumentResult(options, cancellation.Token));
        Assert.Equal(cancellation.Token, provider.ObservedToken);
        Assert.False(provider.ReturnedAfterCancellation);
    }

    private sealed class CancellingProvider(CancellationTokenSource source) : IOfficeTextShapingProvider {
        internal CancellationToken ObservedToken { get; private set; }
        internal bool ReturnedAfterCancellation { get; private set; }
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            ObservedToken = request.CancellationToken;
            source.Cancel();
            request.CancellationToken.ThrowIfCancellationRequested();
            ReturnedAfterCancellation = true;
            return null;
        }
    }

    private static IEnumerable<VisioDocument> Reopened(VisioDocument source) {
        yield return source;
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(source.ToLegacyXmlResult().Value)).Value;
        yield return VisioDocument.Load(new MemoryStream(source.ToBytes()));
    }

    private static byte[] WithWideAdvances(byte[] source) {
        byte[] bytes = (byte[])source.Clone();
        int U16(int offset) => bytes[offset] * 256 + bytes[offset + 1];
        int U32(int offset) => checked((int)((uint)bytes[offset] << 24 | (uint)bytes[offset + 1] << 16 | (uint)bytes[offset + 2] << 8 | bytes[offset + 3]));
        int metrics = 0, count = 0;
        for (int i = 0; i < U16(4); i++) {
            int entry = 12 + i * 16;
            string name = Encoding.ASCII.GetString(bytes, entry, 4);
            if (name == "hmtx") metrics = U32(entry + 8);
            if (name == "hhea") count = U16(U32(entry + 8) + 34);
        }
        Assert.True(metrics > 0 && count > 0);
        for (int i = 0; i < count; i++) { bytes[metrics + i * 4] = 7; bytes[metrics + i * 4 + 1] = 208; }
        return bytes;
    }
}
