using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using Xunit;

namespace OfficeIMO.Tests;

public partial class PowerPoint {
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
}
