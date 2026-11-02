using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSvgRasterAdmissionTests {
    [Theory]
    [InlineData("stroke:blue;stroke-width:1")]
    [InlineData("fill:url(#paint)")]
    [InlineData("font-size:12px")]
    public void RasterAdmissionPreservesSupportedNativeStylesheets(string declarations) {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8'>" +
            "<defs><linearGradient id='paint'><stop stop-color='red'/><stop offset='1' stop-color='red'/></linearGradient></defs>" +
            "<style>.p{" + declarations + "}</style><rect class='p' x='1' y='1' width='6' height='6' fill='red'/></svg>");
        Assert.True(OfficeSvgDrawingReader.TryRead(svg, out _, out int unsupported));
        Assert.Equal(0, unsupported);

        Assert.True(PdfDocument.TryPrepareImageBytes(svg, out var bytes, out _, out _, out var reason), reason);
        Assert.NotEmpty(bytes);
    }

    [Theory]
    [InlineData(8, true)]
    [InlineData(256, true)]
    [InlineData(257, false)]
    public void ImagePreparationAppliesTheCumulativeSvgRepaintLimit(int rectangles, bool accepted) {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8'>" +
            string.Concat(Enumerable.Repeat("<rect width='8' height='8' fill='red'/>", rectangles)) + "</svg>");
        Assert.Equal(accepted, OfficeSvgDrawingReader.IsWithinSafetyLimits(svg));

        bool result = PdfDocument.TryPrepareImageBytes(svg, out byte[] bytes, out var info, out bool transcoded, out var reason);

        Assert.Equal(accepted, result);
        if (accepted) {
            Assert.NotEmpty(bytes);
            Assert.NotNull(info);
            Assert.True(transcoded);
            Assert.Null(reason);
            Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? image));
            Assert.NotNull(image);
            Assert.Equal(OfficeColor.Red, image.GetPixel(4, 4));
        } else {
            Assert.Empty(bytes);
            Assert.Null(info);
            Assert.False(transcoded);
            Assert.False(string.IsNullOrWhiteSpace(reason));
        }
    }

    [Theory]
    [InlineData("opacity", 8, true)]
    [InlineData("opacity", 257, false)]
    [InlineData("image", 8, true)]
    [InlineData("image", 257, false)]
    [InlineData("stylesheet", 8, true)]
    [InlineData("stylesheet", 257, false)]
    [InlineData("blend", 8, true)]
    [InlineData("blend", 257, false)]
    public void NestedPaintAndStylesShareTheImagePreparationBudget(string kind, int repetitions, bool accepted) {
        string element = kind == "blend" ? "<g style='mix-blend-mode:multiply'><rect width='1' height='1' fill='red'/></g>" :
            kind == "opacity" ? "<g opacity='.5'><rect width='8' height='8' fill='red'/></g>" :
            kind == "stylesheet" ? "<rect class='p' width='8' height='8'/>" :
            "<image width='8' height='8' href='data:image/png;base64," +
            System.Convert.ToBase64String(OfficePngWriter.Encode(new OfficeRasterImage(8, 8, OfficeColor.Red))) + "'/>";
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8'>" +
            "<style>.p{fill:red}</style>" + string.Concat(Enumerable.Repeat(element, repetitions)) + "</svg>");
        Assert.True(OfficeSvgDrawingReader.TryRead(svg, out _, out int unsupported));
        Assert.Equal(0, unsupported);

        bool result = PdfDocument.TryPrepareImageBytes(svg, out byte[] bytes, out _, out _, out string? reason);
        Assert.Equal(accepted, result);
        if (accepted) Assert.NotEmpty(bytes);
        else Assert.Contains("Cumulative SVG", reason);
    }

    [Theory]
    [InlineData(8, true)]
    [InlineData(257, false)]
    public void ForegroundSvgRasterFallbackAppliesTheSameWorkLimit(int rectangles, bool accepted) {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8'>" +
            string.Concat(Enumerable.Repeat("<rect width='8' height='8' fill='red'/>", rectangles)) + "</svg>");
        byte[] Export() => PdfDocument.Create(new PdfCreateOptions { PageWidth = OfficeLength.Points(24), PageHeight = OfficeLength.Points(24),
            MarginLeft = OfficeLength.Points(0), MarginTop = OfficeLength.Points(0), MarginRight = OfficeLength.Points(0), MarginBottom = OfficeLength.Points(0) })
            .Canvas(canvas => canvas.ForegroundImage(svg, 0, 0, 8, 8)).ToBytes(new PdfSaveOptions { AllowLoss = true });
        if (accepted) Assert.NotEmpty(Export());
        else Assert.Contains("Cumulative SVG", Assert.Throws<System.NotSupportedException>(() => Export()).Message);
    }
}
