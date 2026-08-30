using OfficeIMO.Pdf;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Fact]
    public void PdfXReadinessTreatsUndecodableOutputProfilesAsFailedEvidence() {
        byte[] pdf = BuildUnsupportedFilteredOutputIntentProfilePdf();
        ReplaceAscii(pdf, "/GTS_PDFA1", "/GTS_PDFX ");

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfX4, pdf);

        Assert.Equal(
            PdfComplianceRequirementStatus.Missing,
            report.FindRequirement("readback-pdfx-output-intent")!.Status);
    }

    private static void ReplaceAscii(byte[] bytes, string oldValue, string newValue) {
        Assert.Equal(oldValue.Length, newValue.Length);
        int offset = bytes.AsSpan().IndexOf(Encoding.ASCII.GetBytes(oldValue));
        Assert.True(offset >= 0, "Expected PDF token was not found.");
        Encoding.ASCII.GetBytes(newValue, 0, newValue.Length, bytes, offset);
    }
}
