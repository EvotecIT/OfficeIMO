using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public class PdfRedactionRenderingIdentityTests {
    [Fact]
    public void AppliedPlanVerificationRejectsChangedFormLocalFontProgramGraph() {
        AssertPlanIdentityChanged(
            BuildFormEmbeddedFontIdentityPdf("source-form-font-program"),
            BuildFormEmbeddedFontIdentityPdf("changed-form-font-program"));
    }

    [Fact]
    public void AppliedPlanVerificationRejectsChangedNestedFormLocalFontProgramGraph() {
        AssertPlanIdentityChanged(
            BuildNestedFormEmbeddedFontIdentityPdf("source-nested-form-font-program"),
            BuildNestedFormEmbeddedFontIdentityPdf("changed-nested-form-font-program"));
    }

    [Theory]
    [InlineData("[6 2] 0 d", "[60 2] 0 d")]
    [InlineData("[6 2] 0 d", "[6 2] 1 d")]
    public void AppliedPlanVerificationRejectsChangedExactUnredactedDashPattern(
        string sourceDash,
        string rewrittenDash) {
        AssertPlanIdentityChanged(
            BuildVectorStyleIdentityPdf(sourceDash),
            BuildVectorStyleIdentityPdf(rewrittenDash));
    }

    [Fact]
    public void AppliedPlanVerificationRejectsChangedDashInheritedByFormXObject() {
        AssertPlanIdentityChanged(
            BuildFormInheritedDashIdentityPdf("[6 2] 0 d"),
            BuildFormInheritedDashIdentityPdf("[60 2] 1 d"));
    }

    private static void AssertPlanIdentityChanged(byte[] source, byte[] rewritten) {
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 150D, 20D, 10D, 10D, "reviewed blank area")
        ]);

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            rewritten,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    private static byte[] BuildVectorStyleIdentityPdf(string styleOperators) {
        string content = $"q {styleOperators} 1 0 0 RG 4 w 20 20 80 60 re S Q";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture)} >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
    }

    private static byte[] BuildFormInheritedDashIdentityPdf(string dashOperator) {
        string pageContent = $"q {dashOperator} /Fm1 Do Q";
        const string formContent = "1 0 0 RG 4 w 20 20 80 60 re S";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /XObject << /Fm1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture)} >>", "stream", pageContent, "endstream", "endobj",
            "5 0 obj", $"<< /Type /XObject /Subtype /Form /BBox [0 0 200 200] /Length {Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture)} >>", "stream", formContent, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));
    }

    private static byte[] BuildFormEmbeddedFontIdentityPdf(string fontProgram) {
        const string pageContent = "/Fm1 Do";
        const string formContent = "BT /F1 12 Tf 20 100 Td (Visible form text) Tj ET";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /XObject << /Fm1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture)} >>", "stream", pageContent, "endstream", "endobj",
            "5 0 obj", $"<< /Type /XObject /Subtype /Form /BBox [0 0 200 200] /Resources << /Font << /F1 6 0 R >> >> /Length {Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture)} >>", "stream", formContent, "endstream", "endobj",
            "6 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /FontDescriptor 7 0 R >>", "endobj",
            "7 0 obj", "<< /Type /FontDescriptor /FontName /Helvetica /Flags 32 /FontBBox [0 -200 1000 900] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 700 /StemV 80 /FontFile 8 0 R >>", "endobj",
            "8 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(fontProgram).ToString(CultureInfo.InvariantCulture)} >>", "stream", fontProgram, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 9 >>", "%%EOF"
        }));
    }

    private static byte[] BuildNestedFormEmbeddedFontIdentityPdf(string fontProgram) {
        const string pageContent = "/Outer Do";
        const string outerFormContent = "/Inner Do";
        const string innerFormContent = "BT /F1 12 Tf 20 100 Td (Visible nested form text) Tj ET";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /XObject << /Outer 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture)} >>", "stream", pageContent, "endstream", "endobj",
            "5 0 obj", $"<< /Type /XObject /Subtype /Form /BBox [0 0 200 200] /Resources << /XObject << /Inner 6 0 R >> >> /Length {Encoding.ASCII.GetByteCount(outerFormContent).ToString(CultureInfo.InvariantCulture)} >>", "stream", outerFormContent, "endstream", "endobj",
            "6 0 obj", $"<< /Type /XObject /Subtype /Form /BBox [0 0 200 200] /Resources << /Font << /F1 7 0 R >> >> /Length {Encoding.ASCII.GetByteCount(innerFormContent).ToString(CultureInfo.InvariantCulture)} >>", "stream", innerFormContent, "endstream", "endobj",
            "7 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /FontDescriptor 8 0 R >>", "endobj",
            "8 0 obj", "<< /Type /FontDescriptor /FontName /Helvetica /Flags 32 /FontBBox [0 -200 1000 900] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 700 /StemV 80 /FontFile 9 0 R >>", "endobj",
            "9 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(fontProgram).ToString(CultureInfo.InvariantCulture)} >>", "stream", fontProgram, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 10 >>", "%%EOF"
        }));
    }
}
