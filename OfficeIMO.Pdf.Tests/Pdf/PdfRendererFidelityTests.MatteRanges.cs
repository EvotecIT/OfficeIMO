using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfRendererFidelityTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void LabMatteUsesColorSpaceRangesBeforeUnpremultiplication(bool customRange, bool explicitDecode) {
        string ranges = customRange ? "/Range [-25 75 -100 100]" : "";
        string colorSpace = "[/Lab << /WhitePoint [.9505 1 1.089] " + ranges + " >>]";
        string decode = explicitDecode ? "/Decode [0 100 -25 75 -100 100]" : "";
        // L*=75 over a white L*=100 matte at alpha .5 represents neutral L*=50.
        // Neutral L*=50 is approximately sRGB119; compositing over white yields187.
        var objects = new[] {
            (ImageDictionary(1, 1, colorSpace, 8, 3, "/SMask 6 0 R " + decode), new byte[] { 191, (byte)(customRange ? 64 : 128), 128 }),
            Mask(1, 1, 8, new byte[] { 128 }, "/Matte [100 0 0]")
        };
        OfficeRasterImage image = Render(BuildPdf("q 80 0 0 80 10 10 cm /Im1 Do Q", "<< /XObject << /Im1 5 0 R >> >>", objects: objects));
        AssertColor(image, 50, 50, 187, 187, 187, tolerance: 6);
    }
}
