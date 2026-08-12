using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAcroFormReviewRegressionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AppendOnlyFill_RejectsPushButtons(bool useTryFill) {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Push button append guard")).ToBytes();
        byte[] authored = PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "calculate",
            Kind = PdfFormFieldCreationKind.PushButton,
            Caption = "Calculate"
        })).ToBytes();
        var values = new Dictionary<string, string> { ["calculate"] = "Off" };

        if (useTryFill) {
            PdfOperationResult<PdfDocument> result = PdfDocument.Open(authored).Forms.TryFill(values);
            Assert.False(result.Succeeded);
            Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Contains("Push-button", StringComparison.Ordinal));
        } else {
            Assert.Throws<ArgumentException>(() => PdfDocument.Open(authored).Forms.AppendRevision(values));
        }
    }

    [Fact]
    public void Create_RejectsChildBelowInheritedTerminalFieldWithWidgetKids() {
        PdfDocument document = PdfDocument.Open(BuildInheritedTerminalFieldPdf());

        ArgumentException exception = Assert.Throws<ArgumentException>(() => document.Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "section.existing.child",
            Value = "new"
        })));

        Assert.Contains("terminal field", exception.Message, StringComparison.OrdinalIgnoreCase);
        PdfFormField existing = Assert.Single(document.Inspect().FormFields);
        Assert.Equal("section.existing", existing.Name);
        Assert.Equal("before", existing.Value);
    }

    private static byte[] BuildInheritedTerminalFieldPdf() {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [8 0 R] >>", "endobj",
            "5 0 obj", "<< /Fields [6 0 R] >>", "endobj",
            "6 0 obj", "<< /FT /Tx /T (section) /Kids [7 0 R] >>", "endobj",
            "7 0 obj", "<< /Parent 6 0 R /T (existing) /V (before) /Kids [8 0 R] >>", "endobj",
            "8 0 obj", "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 20 160 48] /P 3 0 R >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 9 >>", "%%EOF"
        }));
    }
}
