using OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfRedactionProofTestSupport {
    public static PdfRedactionProofResult BuildAndVerifyRedactionRemovalProof() {
        byte[] source = BuildRedactionRemovalProofPdf();
        PdfRedactionArea area = FindAreaForText(source, "Sensitive payroll token PAY-SECRET-2026");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area });
        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        PdfRedactionVerificationOptions options = CreateVerificationOptions();
        PdfRedactionVerificationReport verification = PdfRedactionVerification.AssertVerified(redacted, options);

        return new PdfRedactionProofResult(source, redacted, area, plan, verification);
    }

    public static PdfRedactionVerificationOptions CreateVerificationOptions() {
        return new PdfRedactionVerificationOptions()
            .RequireRemovedText("Sensitive payroll token", "PAY-SECRET-2026")
            .RequireRetainedText("Visible compliance marker", "Public summary marker");
    }

    private static byte[] BuildRedactionRemovalProofPdf() {
        return PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Meta(
                title: "PDF Redaction Removal Gate",
                author: "OfficeIMO",
                subject: "Redaction removal proof",
                keywords: "pdf,redaction,removal")
            .Paragraph(paragraph => paragraph.Text("Visible compliance marker"))
            .Paragraph(paragraph => paragraph.Text("Sensitive payroll token PAY-SECRET-2026"))
            .Paragraph(paragraph => paragraph.Text("Public summary marker"))
            .ToBytes();
    }

    private static PdfRedactionArea FindAreaForText(byte[] pdf, string text) {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        foreach (PdfLogicalTextBlock block in logical.TextBlocks) {
            if (block.Text.IndexOf(text, StringComparison.Ordinal) >= 0) {
                double width = Math.Max(1D, block.XEnd - block.XStart);
                double height = Math.Max(12D, block.FontSize + 8D);
                return new PdfRedactionArea(block.PageNumber, block.XStart - 2D, block.BaselineY - block.FontSize - 2D, width + 4D, height, "sensitive-token");
            }
        }

        throw new InvalidOperationException("Unable to find redaction proof marker: " + text);
    }
}

internal sealed class PdfRedactionProofResult {
    public PdfRedactionProofResult(
        byte[] source,
        byte[] redacted,
        PdfRedactionArea area,
        PdfRedactionPlan plan,
        PdfRedactionVerificationReport verification) {
        Source = source;
        Redacted = redacted;
        Area = area;
        Plan = plan;
        Verification = verification;
    }

    public byte[] Source { get; }

    public byte[] Redacted { get; }

    public PdfRedactionArea Area { get; }

    public PdfRedactionPlan Plan { get; }

    public PdfRedactionVerificationReport Verification { get; }
}
