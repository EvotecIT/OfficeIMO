using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfIncrementalUpdaterTests {
    [Fact]
    public void UpdateMetadata_AppendsIncrementalRevisionAndPreservesContent() {
        byte[] original = PdfDocument.Create()
            .Meta(title: "Original title", author: "Original author")
            .Paragraph(paragraph => paragraph.Text("Incremental update body text"))
            .ToBytes();

        byte[] updated = PdfIncrementalUpdater.UpdateMetadata(original, title: "Updated title");
        PdfDocumentInfo info = PdfInspector.Inspect(updated);

        Assert.True(updated.Length > original.Length);
        Assert.Equal("Updated title", info.Metadata.Title);
        Assert.Equal("Original author", info.Metadata.Author);
        Assert.Contains("Incremental update body text", PdfTextExtractor.ExtractAllText(updated), StringComparison.Ordinal);
        Assert.True(info.Security.HasIncrementalUpdates);
        Assert.True(info.Security.HasPreviousRevision);
        Assert.True(info.Security.RevisionCount >= 2);
        Assert.Contains(info.Security.Revisions, revision => revision.HasPreviousRevision);
    }

    [Fact]
    public void UpdateMetadata_PreservesNonZeroRootGenerationInAppendedTrailer() {
        byte[] original = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 2 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] >>",
            "endobj",
            "trailer",
            "<< /Root 1 2 R /Size 4 >>",
            "startxref",
            "123",
            "%%EOF"
        }));

        byte[] updated = PdfIncrementalUpdater.UpdateMetadata(original, title: "Updated title");
        string raw = PdfEncoding.Latin1GetString(updated);

        Assert.Contains("/Root 1 2 R", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/Root 1 0 R /Info", raw, StringComparison.Ordinal);
        Assert.Equal("Updated title", PdfInspector.Inspect(updated).Metadata.Title);
    }

    [Fact]
    public void UpdateFormFields_PreservesNonZeroInfoGenerationInAppendedTrailer() {
        byte[] original = BuildGeneratedFormPdfWithInfoGeneration();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(original, new Dictionary<string, string> {
            ["Name"] = "Grace"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = false,
            KeepNeedAppearances = true
        });
        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(original).Length);

        Assert.Contains("/Info 8 2 R", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void UpdateFormFields_PreservesInheritedFontGenerationInAppearanceResources() {
        byte[] original = BuildGeneratedFormPdfWithInfoGeneration();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(original, new Dictionary<string, string> {
            ["Name"] = "Grace"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });
        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(original).Length);

        Assert.Contains("/Helv 4 2 R", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("/Helv 4 0 R", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void UpdateFormFields_AppendsIndirectFieldsArrayContainingDirectField() {
        byte[] original = BuildIndirectFieldsArrayWithDirectFieldPdf();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(original, new Dictionary<string, string> {
            ["Name"] = "Grace"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = false,
            KeepNeedAppearances = true
        });
        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(original).Length);

        Assert.Contains("8 0 obj", appended, StringComparison.Ordinal);
        Assert.Contains("/V <4772616365>", appended, StringComparison.Ordinal);
    }

    private static byte[] BuildGeneratedFormPdfWithInfoGeneration() {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 7 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [5 0 R] /Contents 6 0 R >>",
            "endobj",
            "4 2 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Widget /FT /Tx /T (Name) /V (Ada) /Rect [50 50 180 70] /F 4 >>",
            "endobj",
            "6 0 obj",
            "<< /Length 44 >>",
            "stream",
            "BT /F1 12 Tf 72 720 Td (Form field) Tj ET",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Fields [5 0 R] /DR << /Font << /Helv 4 2 R >> >> >>",
            "endobj",
            "8 2 obj",
            "<< /Producer (OfficeIMO fixture) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Info 8 2 R /Size 9 >>",
            "startxref",
            "123",
            "%%EOF"
        }));
    }

    private static byte[] BuildIndirectFieldsArrayWithDirectFieldPdf() {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 7 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] >>",
            "endobj",
            "7 0 obj",
            "<< /Fields 8 0 R >>",
            "endobj",
            "8 0 obj",
            "[<< /FT /Tx /T (Name) /V (Ada) /Rect [50 50 180 70] /F 4 >>]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "startxref",
            "123",
            "%%EOF"
        }));
    }
}
