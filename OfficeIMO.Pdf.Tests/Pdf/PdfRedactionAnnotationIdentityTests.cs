using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRedactionAnnotationIdentityTests {
    [Theory]
    [InlineData("/Open false", "/Open true")]
    [InlineData("/M (D:20260902120000Z)", "/M (D:20260902120100Z)")]
    [InlineData("/LL 2 /LLE 1 /CO [0 0]", "/LL 4 /LLE 3 /CO [2 1]")]
    [InlineData("/A << /S /JavaScript /JS (first) >>", "/A << /S /JavaScript /JS (second) >>")]
    [InlineData("/OfficeIMOState << /Mode /First >>", "/OfficeIMOState << /Mode /Second >>")]
    public void AppliedPlanVerificationRejectsChangedUnredactedAnnotationDictionaryGraph(
        string sourceEntries,
        string rewrittenEntries) {
        byte[] source = BuildTextAnnotationIdentityPdf(sourceEntries);
        PdfRedactionPlan plan = CreatePlan(source);

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            BuildTextAnnotationIdentityPdf(rewrittenEntries),
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void AppliedPlanVerificationAcceptsObjectRenumberedUnredactedAnnotationGraph() {
        byte[] source = BuildTextAnnotationIdentityPdf("/Open false");
        PdfRedactionPlan plan = CreatePlan(source);

        byte[] redacted = PdfRedactionApplier.Apply(source, plan);
        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.NotEqual(
            Assert.Single(PdfInspector.Inspect(source).Annotations).ObjectNumber,
            Assert.Single(PdfInspector.Inspect(redacted).Annotations).ObjectNumber);
        Assert.True(report.IsVerified, string.Join("; ", report.Issues.Select(static issue => issue.Message)));
    }

    [Fact]
    public void AppliedPlanVerificationRejectsChangedIndirectAnnotationActionPayload() {
        byte[] source = BuildTextAnnotationIdentityPdf(
            "/A 51 0 R",
            "<< /S /JavaScript /JS (first) >>");
        PdfRedactionPlan plan = CreatePlan(source);

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            BuildTextAnnotationIdentityPdf(
                "/A 51 0 R",
                "<< /S /JavaScript /JS (second) >>"),
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void AppliedPlanVerificationAcceptsPreservedAnnotationAppearanceGraph() {
        const string appearance = "q 1 0 0 rg 0 0 40 40 re f Q";
        byte[] source = BuildTextAnnotationIdentityPdf(
            "/AP << /N 51 0 R >>",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 40 40] /Length {appearance.Length} >>\nstream\n{appearance}\nendstream");
        PdfRedactionPlan plan = CreatePlan(source);

        byte[] redacted = PdfRedactionApplier.Apply(source, plan);
        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.True(report.IsVerified, string.Join("; ", report.Issues.Select(static issue => issue.Message)));
    }

    private static PdfRedactionPlan CreatePlan(byte[] source) =>
        PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 10D, 10D, 10D, 10D, "reviewed blank area")
        ]);

    private static byte[] BuildTextAnnotationIdentityPdf(string annotationEntries, string? additionalObject = null) {
        var lines = new List<string> {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Annots [50 0 R] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", "", "endstream", "endobj",
            "50 0 obj", $"<< /Type /Annot /Subtype /Text /Rect [100 100 140 140] /P 3 0 R /Contents (annotation identity) {annotationEntries} >>", "endobj"
        };
        if (additionalObject != null) {
            lines.Add("51 0 obj");
            lines.Add(additionalObject);
            lines.Add("endobj");
        }
        lines.Add("trailer");
        lines.Add(additionalObject == null ? "<< /Root 1 0 R /Size 51 >>" : "<< /Root 1 0 R /Size 52 >>");
        lines.Add("%%EOF");
        return Encoding.ASCII.GetBytes(string.Join("\n", lines));
    }
}
