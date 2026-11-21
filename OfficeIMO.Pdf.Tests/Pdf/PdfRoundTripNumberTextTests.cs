using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

// PDF numbers written at full precision use the shortest round-trip digits on every framework; .NET Framework's "R"
// wrote this value with 17 digits.
public sealed class PdfRoundTripNumberTextTests {
    [Fact]
    public void PreciseNumbersUseTheShortestRoundTripDigits() {
        Assert.Equal("0.4694715627858908", PdfSyntaxEscaper.Number(0.4694715627858908D));
        Assert.Equal("0.4694715627858908", PdfNumberFormatter.Precise(0.4694715627858908D));
    }
}
