using OfficeIMO.Yaml;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

public sealed class YamlRoundTripNumberTextTests {
    // YAML floats use the shortest round-trip digits; .NET Framework's "R" wrote this value with 17 digits.
    [Fact]
    public void FloatsUseTheShortestRoundTripDigits() {
        Assert.Equal("0.4694715627858908", YamlScalarFormat.Format(0.4694715627858908D));
    }
}
