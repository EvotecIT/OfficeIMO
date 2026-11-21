using System.Text;
using OfficeIMO.Toml;
using Xunit;

namespace OfficeIMO.Toml.Tests;

public sealed class TomlRoundTripNumberTextTests {
    // TOML floats use the shortest round-trip digits; .NET Framework's "R" wrote this value with 17 digits.
    [Fact]
    public void FloatsUseTheShortestRoundTripDigits() {
        var text = new StringBuilder();
        TomlFormatter.AppendFloat(text, 0.4694715627858908D);

        Assert.Equal("0.4694715627858908", text.ToString());
    }
}
