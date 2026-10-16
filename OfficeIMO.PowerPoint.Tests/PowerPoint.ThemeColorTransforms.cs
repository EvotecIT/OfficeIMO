using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using Xunit;

namespace OfficeIMO.Tests;

public partial class PowerPoint {
    [Theory]
    [InlineData(true, 204, 210, 216)]
    [InlineData(false, 10, 61, 84)]
    public void ThemeColorResolver_TintAndShadeUseLinearRgbAndPreserveAlpha(bool tint, byte red, byte green, byte blue) {
        var transforms = new RgbColorModelHex { Val = "156082" };
        transforms.Append(tint
            ? (DocumentFormat.OpenXml.OpenXmlElement)new Tint { Val = 40000 }
            : new Shade { Val = 40000 });
        transforms.Append(new AlphaModulation { Val = 50000 });

        OfficeColor actual = OfficeOpenXmlThemeColorResolver.ApplyTransforms(
            OfficeColor.FromRgba(21, 96, 130, 128), transforms);

        Assert.Equal(OfficeColor.FromRgba(red, green, blue, 64), actual);
    }

    [Fact]
    public void ThemeColorResolver_KeepsTintShadeAndAlphaTransformsInSourceOrder() {
        var tintThenShade = new RgbColorModelHex(new Tint { Val = 40000 }, new Shade { Val = 50000 },
            new Alpha { Val = 50000 }, new AlphaModulation { Val = 50000 }) { Val = "156082" };
        var shadeThenTint = new RgbColorModelHex(new Shade { Val = 50000 }, new Tint { Val = 40000 },
            new AlphaModulation { Val = 50000 }, new Alpha { Val = 50000 }) { Val = "156082" };

        Assert.Equal(OfficeColor.FromRgba(149, 154, 158, 64), OfficeOpenXmlThemeColorResolver.ResolveColor(tintThenShade, null));
        Assert.Equal(OfficeColor.FromRgba(204, 207, 210, 128), OfficeOpenXmlThemeColorResolver.ResolveColor(shadeThenTint, null));
    }

    [Fact]
    public void ThemeColorResolver_DistinguishesComplementFromChannelInverse() {
        var complement = new RgbColorModelHex(new Complement()) { Val = "336699" };
        var inverse = new RgbColorModelHex(new Inverse()) { Val = "336699" };

        Assert.Equal(OfficeColor.FromRgb(153, 102, 51), OfficeOpenXmlThemeColorResolver.ResolveColor(complement, null));
        Assert.Equal(OfficeColor.FromRgb(204, 153, 102), OfficeOpenXmlThemeColorResolver.ResolveColor(inverse, null));
    }

    [Fact]
    public void ThemeColorResolver_UsesEveryDrawingColorChoiceAsPlaceholder() {
        var themeFill = new SolidFill(new SchemeColor {
            Val = SchemeColorValues.PhColor
        });
        var references = new DocumentFormat.OpenXml.OpenXmlElement[] {
            new FillReference(new RgbColorModelHex { Val = "FF0000" }),
            new FillReference(new RgbColorModelPercentage {
                RedPortion = 100000,
                GreenPortion = 0,
                BluePortion = 0
            }),
            new FillReference(new HslColor {
                HueValue = 0,
                SatValue = 100000,
                LumValue = 50000
            })
        };

        Assert.All(references, reference => Assert.Equal(
            OfficeColor.FromRgb(255, 0, 0),
            OfficeOpenXmlThemeColorResolver.ResolveColor(
                themeFill, null, reference)));
    }

    [Fact]
    public void ThemeColorResolver_ConvertsLinearScRgbToSrgbChannels() {
        var color = new RgbColorModelPercentage {
            RedPortion = 50000,
            GreenPortion = 25000,
            BluePortion = 0
        };

        Assert.Equal(OfficeColor.FromRgb(188, 137, 0),
            OfficeOpenXmlThemeColorResolver.ResolveColor(color, null));
    }

    [Fact]
    public void ThemeColorResolver_ResolvesHslThemeSchemeEntries() {
        var scheme = new ColorScheme(
            new Accent1Color(new HslColor {
                HueValue = 0,
                SatValue = 100000,
                LumValue = 50000
            })) { Name = "HSL theme" };
        var fill = new SolidFill(new SchemeColor {
            Val = SchemeColorValues.Accent1
        });

        Assert.Equal(OfficeColor.FromRgb(255, 0, 0),
            OfficeOpenXmlThemeColorResolver.ResolveColor(fill, scheme));
    }
}
