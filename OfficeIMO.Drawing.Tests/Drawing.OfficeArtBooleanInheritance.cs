using OfficeIMO.Drawing.Binary;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(0x0004)]
    [InlineData(0x0147)]
    [InlineData(0x0181)]
    [InlineData(0x0107)]
    [InlineData(0x0108)]
    public void OfficeArtBooleanInheritance_MalformedEffectiveScalarsDoNotReviveOlderValues(int id) {
        var properties = new[] { new OfficeArtProperty(0, (ushort)id, 65536),
            new OfficeArtProperty(1, (ushort)(id | 0x8000), 1, 1, complexData: new byte[] { 0 }) };
        switch (id) {
            case 0x0004: Assert.Null(OfficeArtShapeTransform.Decode(0, properties).RotationDegrees); break;
            case 0x0147: Assert.Null(OfficeArtShapeGeometry.Decode(properties).AdjustmentValues[0]); break;
            case 0x0181: Assert.Null(OfficeArtShapeStyle.Decode(properties).FillColor); break;
            case 0x0107: Assert.Null(OfficeArtPictureProperties.Decode(properties).TransparentColor); break;
            case 0x0108: Assert.Null(OfficeArtPictureProperties.Decode(properties).ContrastRaw); break;
        }
    }

    [Fact]
    public void OfficeArtBooleanInheritance_MalformedEffectiveBooleanGroupDoesNotReviveOlderBits() {
        var properties = new[] { new OfficeArtProperty(0, 0x01BF, 0x00100010),
            new OfficeArtProperty(1, 0x81BF, 1, 1, complexData: new byte[] { 0 }) };
        Assert.Null(OfficeArtShapeStyle.Decode(properties).FillEnabled);
    }

    [Fact]
    public void OfficeArtBooleanInheritance_FlipOverridesKeepIndependentBitsAndAllowExplicitClearing() {
        var properties = new[] {
            new OfficeArtProperty(0, 0x033F, 0x01000000U),
            new OfficeArtProperty(0, 0x033F, 0x02000200U)
        };
        OfficeArtShapeTransform transform = OfficeArtShapeTransform.Decode(1U << 6, properties);
        Assert.True(transform.FlipHorizontal);
        Assert.True(transform.FlipVertical);
        transform = OfficeArtShapeTransform.Decode(1U << 6, properties.Concat(new[] {
            new OfficeArtProperty(0, 0x033F, 0x00000200U),
            new OfficeArtProperty(0, 0x033F, 0x00000100U)
        }).ToArray());
        Assert.False(transform.FlipHorizontal);
        Assert.False(transform.FlipVertical);
    }

    [Fact]
    public void OfficeArtBooleanInheritance_ProtectionKeepsIndependentLocksAndAllowsExplicitClearing() {
        var properties = new[] {
            new OfficeArtProperty(0, 0x007F, (1U << 7) | (1U << 23)),
            new OfficeArtProperty(0, 0x007F, (1U << 8) | (1U << 24))
        };
        Assert.True(OfficeArtShapeProtection.Decode(properties).LockRotation);
        Assert.True(OfficeArtShapeProtection.Decode(properties).LockAspectRatio);
        Assert.False(OfficeArtShapeProtection.Decode(properties.Concat(new[] {
            new OfficeArtProperty(0, 0x007F, 1U << 7)
        }).ToArray()).LockRotation);
    }
}
