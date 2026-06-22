using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSignatureValidatorTests {
    [Fact]
    public void Validate_ReportsUnsignedPdfWithoutErrors() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Unsigned document"))
            .ToBytes();

        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(pdf);

        Assert.False(report.HasSignatures);
        Assert.True(report.IsStructurallyValid);
        Assert.False(report.CryptographicTrustVerified);
        Assert.Contains(report.Findings, finding => finding.Code == "NoSignatures");
    }

    [Fact]
    public void Validate_ReportsSignatureStructureAndPreservationMarkers() {
        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(BuildSignedIncrementalPdf());

        Assert.True(report.HasSignatures);
        Assert.True(report.IsStructurallyValid);
        Assert.True(report.HasWarnings);
        Assert.True(report.RequiresAppendOnlyMutation);
        Assert.True(report.HasLongTermValidationEvidence);
        Assert.False(report.CryptographicTrustVerified);
        Assert.False(report.DigestVerified);
        Assert.False(report.CertificateChainVerified);
        Assert.False(report.RevocationChecked);
        Assert.False(report.TimestampValidationPerformed);

        PdfSignatureValidationResult result = Assert.Single(report.Signatures);
        Assert.Equal("Approval", result.Signature.FieldName);
        Assert.Equal(new long[] { 0, 10, 20, 30 }, result.Signature.ByteRangeValues);
        Assert.True(result.Signature.HasRecognizedSubFilter);
        Assert.True(result.Signature.UsesDetachedCmsSubFilter);
        Assert.True(result.HasCompleteByteRangeShape);
        Assert.True(result.ByteRangeSegmentsAreOrdered);
        Assert.False(result.ByteRangeCoversEndOfFile);
        Assert.Equal(40, result.ByteRangeCoveredBytes);
        Assert.Equal(10, result.ByteRangeGapStart);
        Assert.Equal(10, result.ByteRangeGapLength);
        Assert.True(result.UnsignedByteCount > 0);
        Assert.True(result.ByteRangeCoverageRatio > 0);
        Assert.Contains(report.Findings, finding => finding.Code == "SignatureByteRangeDoesNotCoverEof");
        Assert.Contains(report.Findings, finding => finding.Code == "SignatureDetachedCmsSubFilter");
        Assert.Contains(report.Findings, finding => finding.Code == "SignatureByteRangeCoverage");
        Assert.Contains(report.Findings, finding => finding.Code == "DocMDPDetected");
        Assert.Contains(report.Findings, finding => finding.Code == "LongTermValidationEvidenceDetected");
        Assert.Contains(report.Findings, finding => finding.Code == "CryptographicTrustNotVerified");
    }

    [Fact]
    public void Validate_FlagsMalformedSignatureByteRange() {
        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(BuildMalformedSignaturePdf());

        Assert.True(report.HasSignatures);
        Assert.False(report.IsStructurallyValid);
        Assert.Contains(report.Findings, finding => finding.Code == "SignatureUnsupportedByteRangeShape");
        Assert.Contains(report.Findings, finding => finding.Code == "SignatureMissingContents");
    }

    private static byte[] BuildSignedIncrementalPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 7 0 R /Perms << /DocMDP 6 0 R /UR3 6 0 R >> /DSS 9 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /FT /Sig /T (Approval) /V 6 0 R /Subtype /Widget /Rect [10 10 120 40] >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /Name (Alice) /ByteRange [0 10 20 30] /Contents <001122> /Reference [<< /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /V /1.2 /P 2 >> >>] >>",
            "endobj",
            "7 0 obj",
            "<< /Fields [5 0 R] /SigFlags 3 >>",
            "endobj",
            "8 0 obj",
            "<< /Producer (OfficeIMO signed fixture) >>",
            "endobj",
            "9 0 obj",
            "<< /Certs [10 0 R] /OCSPs [11 0 R] /CRLs [12 0 R] /VRI << /ABCDEF << /Cert [10 0 R] /OCSP [11 0 R] /CRL [12 0 R] /TS 13 0 R >> >> >>",
            "endobj",
            "10 0 obj",
            "<< /Type /EmbeddedFile /Length 0 >>",
            "endobj",
            "11 0 obj",
            "<< /Type /EmbeddedFile /Length 0 >>",
            "endobj",
            "12 0 obj",
            "<< /Type /EmbeddedFile /Length 0 >>",
            "endobj",
            "13 0 obj",
            "<< /Type /TimestampEvidence /Length 0 >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Info 8 0 R /ID [(abc) (def)] /Size 14 /Prev 100 >>",
            "startxref",
            "100",
            "%%EOF",
            "startxref",
            "200",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildMalformedSignaturePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
            "endobj",
            "4 0 obj",
            "<< /FT /Sig /T (Broken) /V 6 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Fields [4 0 R] /SigFlags 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Sig /Filter /Adobe.PPKLite /ByteRange [0 10 20] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }
}
