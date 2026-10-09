using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfDrawingTextCancellationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DrawingTextMeasurementPassesSaveCancellationToTheEffectiveFontShaper(bool namedFont) {
        using var cancellation = new CancellationTokenSource();
        var provider = new CancellingProvider(cancellation);
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        var options = new PdfOptions { TextShapingProvider = provider };
        if (namedFont) options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Arial", font));
        else options.EmbedStandardFont(PdfStandardFont.Helvetica, font, "Cancellation font");
        var drawing = new OfficeDrawing(72, 72)
            .AddRichText(new[] { new OfficeRichTextRun("A", 12, OfficeColor.Black, fontFamily: "Arial") }, 0, 0, 72, 72);
        PdfDocument pdf = PdfDocument.Create(options).Drawing(drawing);

        Assert.ThrowsAny<System.OperationCanceledException>(() => pdf.ToBytes(cancellation.Token));
        Assert.Equal(cancellation.Token, provider.ObservedToken);
    }

    private sealed class CancellingProvider(CancellationTokenSource source) : IOfficeTextShapingProvider {
        internal CancellationToken ObservedToken { get; private set; }
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            if (request.CancellationToken.CanBeCanceled) {
                ObservedToken = request.CancellationToken;
                source.Cancel();
                request.CancellationToken.ThrowIfCancellationRequested();
            }
            return null;
        }
    }
}
