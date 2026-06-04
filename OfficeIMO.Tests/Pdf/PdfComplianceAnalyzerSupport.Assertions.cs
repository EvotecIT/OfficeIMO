using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    private static PdfComplianceRequirement AssertRequirement(PdfComplianceReadinessReport report, string id, PdfComplianceRequirementStatus status) {
        PdfComplianceRequirement requirement = Assert.Single(report.Requirements, requirement => requirement.Id == id);
        Assert.True(requirement.Status == status, "Requirement " + id + " expected " + status + " but was " + requirement.Status + ": " + requirement.Diagnostic);
        Assert.False(string.IsNullOrWhiteSpace(requirement.DisplayName));
        Assert.False(string.IsNullOrWhiteSpace(requirement.Diagnostic));
        return requirement;
    }


}
