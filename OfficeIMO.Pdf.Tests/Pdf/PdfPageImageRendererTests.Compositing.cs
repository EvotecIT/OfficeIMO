using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData("false")]
    [InlineData("true")]
    public void RenderPages_CompositesTransparencyGroupThroughInvocationBlendAndOpacity(string isolated) {
        // The group's own Normal state resets blending only inside the group; the completed group
        // still multiplies onto the red backdrop at 80% opacity: 0.2 * 1 at 0.8 over 1 gives 0.36.
        string form = BuildStreamObject(5,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Group << /S /Transparency /I " + isolated +
            " /K false /CS /DeviceRGB >> /Resources << /ExtGState << /GS0 7 0 R >> >>",
            "/GS0 gs 0.2 g 0 0 120 200 re f");
        byte[] pdf = BuildSingleStreamPdf("1 0 0 rg 0 0 240 200 re f /GS1 gs /Fm0 Do",
            "<< /ExtGState << /GS1 6 0 R >> /XObject << /Fm0 5 0 R >> >>", form,
            "6 0 obj\n<< /Type /ExtGState /BM /Multiply /ca 0.8 /CA 0.8 >>\nendobj",
            "7 0 obj\n<< /Type /ExtGState /BM /Normal /ca 1 /CA 1 >>\nendobj");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeRasterImageDecoder.Decode(result.Bytes!);

        OfficeColor grouped = raster.GetPixel(60, 100);
        Assert.InRange(grouped.R, 90, 94);
        Assert.Equal(0, grouped.G);
        Assert.Equal(0, grouped.B);
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), raster.GetPixel(180, 100));
    }
}
