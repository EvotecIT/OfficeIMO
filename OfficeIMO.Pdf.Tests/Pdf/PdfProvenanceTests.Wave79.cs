using OfficeIMO.Provenance;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void InspectionAndRemovalReportDecodedBytesAgainstTheSharedBudget() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxExpandedContainerBytes = 1024 * 1024;

        OfficeProvenanceReport inspection = PdfProvenance.Inspect(pdf, options.Limits);
        OfficeProvenanceRemovalResult removal = PdfProvenance.Remove(pdf, options);

        Assert.True(inspection.ExpandedInspectionBytes > 0);
        Assert.Equal(inspection.ExpandedInspectionBytes, removal.Before.ExpandedInspectionBytes);
        Assert.True(
            removal.Before.ExpandedInspectionBytes + removal.After.ExpandedInspectionBytes <=
            options.Limits.MaxExpandedContainerBytes);
    }
}
