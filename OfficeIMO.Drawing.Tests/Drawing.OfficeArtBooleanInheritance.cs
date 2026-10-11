using OfficeIMO.Drawing.Binary;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
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
