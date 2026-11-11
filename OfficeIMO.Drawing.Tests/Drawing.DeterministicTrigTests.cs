using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingDeterministicTrigTests {
    /// <summary>
    /// Round joins and caps take their angles from this helper. The expected bits were produced on macOS arm64; CI runs the
    /// same assertions on Linux and Windows x64, where Math.Sin/Cos/Atan2 can differ in the last bit.
    /// </summary>
    [Fact]
    public void RoundJoinAnglesHaveTheSameBitsOnEveryPlatformAndStayAccurate() {
        Assert.Equal(4605754516372524270L, BitConverter.DoubleToInt64Bits(OfficeDeterministicTrig.Sin(1D)));
        Assert.Equal(-4617980757132604547L, BitConverter.DoubleToInt64Bits(OfficeDeterministicTrig.Cos(2.5D)));
        Assert.Equal(-4614362743452165700L, BitConverter.DoubleToInt64Bits(OfficeDeterministicTrig.Atan2(-3D, 0.5D)));
        Assert.Equal(4614116101887326496L, BitConverter.DoubleToInt64Bits(OfficeDeterministicTrig.Atan2(0.25D, -4D)));

        for (double angle = -7D; angle <= 7D; angle += 0.0007D) {
            Assert.InRange(OfficeDeterministicTrig.Sin(angle) - Math.Sin(angle), -2e-16, 2e-16);
            Assert.InRange(OfficeDeterministicTrig.Cos(angle) - Math.Cos(angle), -2e-16, 2e-16);
            double y = Math.Sin(angle), x = Math.Cos(angle);
            Assert.InRange(OfficeDeterministicTrig.Atan2(y, x) - Math.Atan2(y, x), -1e-15, 1e-15);
        }
    }
}
