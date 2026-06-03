using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    [Fact]
    public void OptionsOverloadUsesRequestedProfileAndNoneHasNoRequirements() {
        var options = new PdfOptions {
            ComplianceProfile = PdfComplianceProfile.PdfA2B
        };

        PdfComplianceReadinessReport requestedReport = PdfComplianceAnalyzer.Assess(options);
        PdfComplianceReadinessReport noneReport = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.None, new PdfOptions());

        Assert.Equal(PdfComplianceProfile.PdfA2B, requestedReport.Profile);
        AssertRequirement(requestedReport, "pdfa-identification", PdfComplianceRequirementStatus.Missing);
        Assert.Equal(PdfComplianceProfile.None, noneReport.Profile);
        Assert.True(noneReport.IsReady);
        Assert.Empty(noneReport.Requirements);
        Assert.Null(noneReport.FindRequirement("pdfa-identification"));
    }


}
