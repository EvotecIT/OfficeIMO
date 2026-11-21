using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfAWorkflowTests {
    [Fact]
    public async Task ConversionPublishesAPdfACopyWithItsChangesAndARefusalPublishesNothing() {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-pdfa-workflow-" + Guid.NewGuid().ToString("N"))).FullName;
        try {
            string source = Path.Combine(root, "source.pdf"), converted = Path.Combine(root, "archive.pdf"), refused = Path.Combine(root, "refused.pdf");
            byte[] original = PdfWithAttachment();
            File.WriteAllBytes(source, original);

            OfficeWorkflowResult result = await OfficeWorkflow.ConvertToPdfA(source, new PdfAConversionOptions { Profile = PdfComplianceProfile.PdfA2B })
                .To(converted).RunAsync();
            OfficeWorkflowResult refusal = await OfficeWorkflow.ConvertToPdfA(source,
                    new PdfAConversionOptions { Profile = PdfComplianceProfile.PdfA2B, Attachments = PdfAAttachmentPolicy.Refuse })
                .To(refused).RunAsync();

            Assert.True(result.Succeeded, result.Summary);
            Assert.Equal(2, PdfReadDocument.Open(File.ReadAllBytes(converted)).XmpMetadata?.PdfAPart);
            Assert.Equal("PDF/A-2b", result.ConversionEvidence!.Facts["profile"]);
            OfficeWorkflowDiagnostic removed = Assert.Single(result.Diagnostics, item => item.Code == "PdfA.EmbeddedFileRemoved" &&
                item.Details.TryGetValue("objectNumber", out string? number) && number == "5");
            Assert.Equal(OfficeWorkflowDiagnosticSeverity.Warning, removed.Severity);

            Assert.Equal(OfficeWorkflowStatus.Failed, refusal.Status);
            Assert.Equal(OfficeWorkflowFailureKind.UnsupportedInput, refusal.FailureKind);
            OfficeWorkflowDiagnostic violation = Assert.Single(refusal.Diagnostics, item => item.Code == "PdfA.EmbeddedFileNotAllowed");
            Assert.Equal(OfficeWorkflowDiagnosticSeverity.Error, violation.Severity);
            Assert.False(File.Exists(refused));
            Assert.Equal(["archive.pdf", "source.pdf"], Directory.GetFiles(root).Select(path => Path.GetFileName(path)!).Order().ToArray());
            Assert.Equal(original, File.ReadAllBytes(source));
        } finally {
            Directory.Delete(root, true);
        }
    }

    // A vector-only page with one embedded file, so the result does not depend on installed fonts.
    private static byte[] PdfWithAttachment() {
        string[] objects = [
            "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles << /Names [(data.xml) 5 0 R] >> >> >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << >> /Contents 4 0 R >>",
            "<< /Length 21 >>\nstream\n0 g 20 20 50 50 re f\n\nendstream",
            "<< /Type /Filespec /F (data.xml) /EF << /F 6 0 R >> >>",
            "<< /Type /EmbeddedFile /Length 7 >>\nstream\n<data/>\nendstream"
        ];
        var pdf = new StringBuilder("%PDF-1.4\n");
        var offsets = new List<int>();
        for (int index = 0; index < objects.Length; index++) {
            offsets.Add(pdf.Length);
            pdf.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = pdf.Length;
        pdf.Append("xref\n0 ").Append(objects.Length + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) pdf.Append(offset.ToString("0000000000", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        pdf.Append("trailer\n<< /Size ").Append(objects.Length + 1).Append(" /Root 1 0 R >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }
}
