using System.Collections.ObjectModel;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterPngRenderCancellationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PngRenderingObservesCancellationAfterTheFinalDrawingDiagnostic(bool cancel) {
        byte[] source = OfficePngWriter.Encode(new OfficeRasterImage(3, 3, OfficeColor.Red));
        var drawing = new OfficeDrawing(2, 2);
        drawing.AddImage(source, "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 2, 2)));
        using var cancellation = new CancellationTokenSource();
        var diagnostics = new CancellingDiagnostics(cancel ? cancellation : null);
        var options = new OfficeDrawingRasterRenderOptions {
            Background = OfficeColor.White,
            MaximumRasterPixels = 4,
            DiagnosticSink = diagnostics,
            CancellationToken = cancellation.Token
        };

        if (cancel) {
            OperationCanceledException exception = Assert.Throws<OperationCanceledException>(() =>
                OfficeDrawingRasterRenderer.ToPng(drawing, options));
            Assert.Equal(cancellation.Token, exception.CancellationToken);
        } else {
            byte[] png = OfficeDrawingRasterRenderer.ToPng(drawing, options);
            Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? image));
            Assert.Equal((2, 2), (image!.Width, image.Height));
            Assert.Equal(OfficeColor.White, image.GetPixel(1, 1));
        }

        Assert.Equal(OfficeImageExportDiagnosticCodes.SourceImageDecodeOmitted,
            Assert.Single(diagnostics).Code);
    }

    private sealed class CancellingDiagnostics : Collection<OfficeImageExportDiagnostic> {
        private readonly CancellationTokenSource? _cancellation;

        internal CancellingDiagnostics(CancellationTokenSource? cancellation) => _cancellation = cancellation;

        protected override void InsertItem(int index, OfficeImageExportDiagnostic item) {
            base.InsertItem(index, item);
            // Hosts can abort a requested export when its fidelity policy rejects an omission.
            _cancellation?.Cancel();
        }
    }
}
