using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class DrawingColorTransforms {
    [Fact]
    public void DrawingTintAndShadeUseTheInputColorRatio() {
        OfficeColor color = OfficeColor.FromRgb(51, 102, 153);

        Assert.Equal(OfficeColor.FromRgb(235, 240, 245), OfficeColorTransforms.Tint(color, 0.1D));
        Assert.Equal(OfficeColor.FromRgb(5, 10, 15), OfficeColorTransforms.Shade(color, 0.1D));
        Assert.Equal(OfficeColor.FromRgb(153, 178, 204), OfficeColorTransforms.Tint(color, 0.5D));
    }

    [Fact]
    public void AlphaTransformsPreserveRgbAndClampTheirResult() {
        OfficeColor color = OfficeColor.FromRgba(20, 40, 60, 128);

        Assert.Equal(OfficeColor.FromRgba(20, 40, 60, 64), OfficeColorTransforms.WithAlpha(color, 0.25D));
        Assert.Equal(OfficeColor.FromRgba(20, 40, 60, 255), OfficeColorTransforms.ModulateAlpha(color, 3D));
        Assert.Equal(OfficeColor.FromRgba(20, 40, 60, 0), OfficeColorTransforms.OffsetAlpha(color, -1D));
    }

    [Fact]
    public void LuminanceTransformsOperateInHslSpace() {
        OfficeColor color = OfficeColor.FromRgb(32, 64, 96);

        Assert.Equal(OfficeColor.FromRgb(16, 32, 48), OfficeColorTransforms.ModulateLuminance(color, 0.5D));
        Assert.Equal(OfficeColor.FromRgb(64, 128, 191), OfficeColorTransforms.OffsetLuminance(color, 0.25D));
    }

    [Fact]
    public void SpreadsheetTintAdjustsHslLuminance() {
        OfficeColor color = OfficeColor.FromRgb(79, 129, 189);

        Assert.Equal(OfficeColor.FromRgb(149, 179, 215), OfficeColorTransforms.SpreadsheetTint(color, 0.4D));
        Assert.Equal(OfficeColor.FromRgb(79, 29, 28), OfficeColorTransforms.SpreadsheetTint(OfficeColor.FromRgb(192, 80, 77), -0.6D));
    }

    [Fact]
    public void InvalidTransformArgumentsAreRejected() {
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeColorTransforms.Tint(OfficeColor.Black, 1.1D));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeColorTransforms.SpreadsheetTint(OfficeColor.Black, -1.1D));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeColorTransforms.ModulateAlpha(OfficeColor.Black, double.NaN));
    }
}
