using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfComplianceProfileCompatibilityTests {
    [Theory]
    [InlineData(PdfComplianceProfile.None, 0)]
    [InlineData(PdfComplianceProfile.PdfA2B, 1)]
    [InlineData(PdfComplianceProfile.PdfA2U, 2)]
    [InlineData(PdfComplianceProfile.PdfA2A, 3)]
    [InlineData(PdfComplianceProfile.PdfA3B, 4)]
    [InlineData(PdfComplianceProfile.PdfA3U, 5)]
    [InlineData(PdfComplianceProfile.PdfA3A, 6)]
    [InlineData(PdfComplianceProfile.PdfUa1, 7)]
    [InlineData(PdfComplianceProfile.FacturX, 8)]
    [InlineData(PdfComplianceProfile.Zugferd, 9)]
    [InlineData(PdfComplianceProfile.PdfA4, 10)]
    [InlineData(PdfComplianceProfile.PdfA4E, 11)]
    [InlineData(PdfComplianceProfile.PdfA4F, 12)]
    [InlineData(PdfComplianceProfile.PdfUa2, 13)]
    public void NumericValuesRemainStableForPackageConsumers(PdfComplianceProfile profile, int expectedValue) {
        Assert.Equal(expectedValue, (int)profile);
    }
}
