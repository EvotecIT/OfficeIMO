using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class OfficeColorSpaceConverterTests {
    [Fact]
    public void CalibratedRgb_AppliesGammaAndColumnMajorMatrix() {
        OfficeColor color = OfficeColorSpaceConverter.FromCalibratedRgb(
            0.5D,
            0.25D,
            0.75D,
            0.95047D,
            1D,
            1.08883D,
            gamma: new[] { 2D, 2D, 2D },
            matrix: new[] {
                0.4124564D, 0.2126729D, 0.0193339D,
                0.3575761D, 0.7151522D, 0.119192D,
                0.1804375D, 0.072175D, 0.9503041D
            });

        Assert.InRange(color.R, 136, 138);
        Assert.InRange(color.G, 70, 72);
        Assert.InRange(color.B, 197, 199);
    }

    [Fact]
    public void LabAndCmyk_ProduceStableSrgbPrimaries() {
        OfficeColor labRed = OfficeColorSpaceConverter.FromLab(53.24D, 80.09D, 67.2D);
        OfficeColor cmykRed = OfficeColorSpaceConverter.FromCmyk(0D, 1D, 1D, 0D);

        Assert.InRange(labRed.R, 245, 255);
        Assert.InRange(labRed.G, 0, 15);
        Assert.InRange(labRed.B, 0, 15);
        Assert.Equal(OfficeColor.Red, cmykRed);
    }
}
