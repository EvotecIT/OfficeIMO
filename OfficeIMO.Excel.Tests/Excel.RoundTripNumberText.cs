using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelRoundTripNumberTextTests {
    // A copied value keeps negative zero; .NET Framework's "R" wrote it as 0.
    [Fact]
    public void ExactNumberTextKeepsNegativeZero() {
        Assert.Equal("-0", InvariantNumberText.GetExact(-0D));
    }
}
