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
        Assert.Equal("Unsigned", report.ProofStatus);
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
        Assert.True(report.HasOfflineLongTermValidationReadiness);
        Assert.Equal("LtvEvidenceReady", report.ProofStatus);
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
        Assert.Equal("StructuralIssues", report.ProofStatus);
        Assert.Contains(report.Findings, finding => finding.Code == "SignatureUnsupportedByteRangeShape");
        Assert.Contains(report.Findings, finding => finding.Code == "SignatureMissingContents");
    }

    [Fact]
    public void PrepareExternalSignature_AppendsSignablePlaceholderAndDigest() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("External signing draft"))
            .ToBytes();

        PdfAppendOnlyMutationReport appendOnly = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(pdf);
        Assert.True(appendOnly.CanPrepareExternalSignature);
        Assert.Contains("SignaturePrepare", appendOnly.SupportedActions);

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            pdf,
            new PdfExternalSignatureOptions {
                FieldName = "Approval",
                Name = "Alice",
                Reason = "Approval",
                ReservedSignatureContentsBytes = 512,
                SigningTime = new DateTimeOffset(2026, 6, 22, 12, 0, 0, TimeSpan.Zero)
            });

        Assert.True(preparation.PreparedPdf.Length > pdf.Length);
        Assert.Equal("Approval", preparation.FieldName);
        Assert.Equal("adbe.pkcs7.detached", preparation.SubFilter);
        Assert.Equal(4, preparation.ByteRangeValues.Count);
        Assert.Equal(0, preparation.ByteRangeValues[0]);
        Assert.Equal(preparation.ContentsHexOffset - 1, preparation.ByteRangeValues[1]);
        Assert.Equal(preparation.ContentsHexOffset + preparation.ContentsHexLength + 1, preparation.ByteRangeValues[2]);
        Assert.Equal(preparation.PreparedPdf.LongLength, preparation.ByteRangeValues[2] + preparation.ByteRangeValues[3]);
        Assert.Equal(32, preparation.ComputeSha256Digest().Length);

        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(preparation.PreparedPdf);

        Assert.True(report.HasSignatures);
        Assert.True(report.IsStructurallyValid);
        PdfSignatureValidationResult result = Assert.Single(report.Signatures);
        Assert.Equal("Approval", result.Signature.FieldName);
        Assert.True(result.ByteRangeCoversEndOfFile);
        Assert.True(result.ByteRangeGapMatchesContents);
        Assert.True(result.Signature.HasContents);
        Assert.True(result.Signature.ContentsSizeBytes >= 512);
        Assert.Contains(report.Findings, finding => finding.Code == "SignatureDetachedCmsSubFilter");
        Assert.Contains(report.Findings, finding => finding.Code == "AcroFormAppendOnly");
    }

    [Fact]
    public void PrepareExternalSignature_EmitsDocTimeStampTypeForDocumentTimestamps() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Timestamp signing draft"))
            .ToBytes();

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            pdf,
            new PdfExternalSignatureOptions {
                FieldName = "Timestamp",
                SubFilter = PdfExternalSignatureSubFilter.DocumentTimestamp,
                ReservedSignatureContentsBytes = 256
            });

        string preparedText = Encoding.ASCII.GetString(preparation.PreparedPdf);
        PdfSignatureValidationResult signature = Assert.Single(PdfSignatureValidator.Validate(preparation.PreparedPdf).Signatures);

        Assert.Equal("ETSI.RFC3161", preparation.SubFilter);
        Assert.Contains("/Type /DocTimeStamp", preparedText, StringComparison.Ordinal);
        Assert.Contains("/SubFilter /ETSI.RFC3161", preparedText, StringComparison.Ordinal);
        Assert.True(signature.Signature.IsDocumentTimestamp);
    }

    [Fact]
    public void ApplyExternalSignature_PatchesReservedContentsWithoutChangingByteRange() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("External signature injection"))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            pdf,
            new PdfExternalSignatureOptions {
                FieldName = "Approval",
                ReservedSignatureContentsBytes = 256
            });

        byte[] signature = { 0x30, 0x82, 0x01, 0x0A, 0xAA, 0x55 };
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, signature);

        Assert.Equal(preparation.PreparedPdf.Length, signed.Length);
        Assert.Contains("3082010AAA55", Encoding.ASCII.GetString(signed));
        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(signed);
        PdfSignatureValidationResult result = Assert.Single(report.Signatures);
        Assert.Equal(preparation.ByteRangeValues, result.Signature.ByteRangeValues);
        Assert.True(result.Signature.ContentsSizeBytes >= signature.Length);
        Assert.DoesNotContain(report.Findings, finding => finding.Code == "SignatureEmptyContents");
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
