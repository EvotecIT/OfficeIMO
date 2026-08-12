using System.Collections.Generic;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioImageExportClippingTests {
    [Theory]
    [InlineData("<text x='0' y='5'>x</text>")]
    [InlineData("<rect width='5' height='10'/><use href='#shape'/>")]
    public void EmbeddedSvgPreviewReportsClipPathGeometryItCannotRepresent(string clipContent) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                     "<defs><rect id='shape' width='5' height='10'/><clipPath id='clip'>" +
                     clipContent + "</clipPath></defs>" +
                     "<rect width='10' height='10' fill='red' clip-path='url(#clip)'/></svg>";
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "unsupported-clip.svg", default, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }
}
