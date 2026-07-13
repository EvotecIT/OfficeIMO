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
}
