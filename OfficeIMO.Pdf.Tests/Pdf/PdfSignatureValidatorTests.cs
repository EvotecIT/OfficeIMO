using System.IO.Compression;
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
    public void Validate_RejectsSignatureMarkersWithoutReadableSignatureDictionary() {
        byte[] pdf = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 4 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
            "endobj",
            "4 0 obj",
            "<< /Fields [5 0 R] /SigFlags 1 >>",
            "endobj",
            "5 0 obj",
            "<< /FT /Sig /T (Malformed) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "startxref",
            "123",
            "%%EOF"
        }));

        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(pdf);

        Assert.True(report.HasSignatures);
        Assert.Equal(0, report.SignatureCount);
        Assert.True(report.HasErrors);
        Assert.False(report.IsStructurallyValid);
        Assert.Contains(report.Findings, finding => finding.Code == "UnreadableSignature");
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
        Assert.Equal(new long[] { 0, 8, 16, 30 }, result.Signature.ByteRangeValues);
        Assert.True(result.Signature.HasRecognizedSubFilter);
        Assert.True(result.Signature.UsesDetachedCmsSubFilter);
        Assert.True(result.HasCompleteByteRangeShape);
        Assert.True(result.ByteRangeSegmentsAreOrdered);
        Assert.False(result.ByteRangeCoversEndOfFile);
        Assert.Equal(38, result.ByteRangeCoveredBytes);
        Assert.Equal(8, result.ByteRangeGapStart);
        Assert.Equal(8, result.ByteRangeGapLength);
        Assert.True(result.ByteRangeGapMatchesContents);
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
    public void Validate_ComparesByteRangeGapWithEncodedContentsPlaceholder() {
        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(BuildSignedPdfWithEncodedContentsGap());

        PdfSignatureValidationResult result = Assert.Single(report.Signatures);
        Assert.Equal(8, result.ByteRangeGapLength);
        Assert.Equal(3, result.Signature.ContentsSizeBytes);
        Assert.Equal(8, result.Signature.ContentsEncodedSizeBytes);
        Assert.True(result.ByteRangeGapMatchesContents);
    }

    [Fact]
    public void Validate_ComparesByteRangeGapWithActualContentsTokenSpan() {
        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(BuildSignedPdfWithWhitespaceContentsGap());

        PdfSignatureValidationResult result = Assert.Single(report.Signatures);
        Assert.Equal(10, result.ByteRangeGapLength);
        Assert.Equal(3, result.Signature.ContentsSizeBytes);
        Assert.Equal(10, result.Signature.ContentsEncodedSizeBytes);
        Assert.True(result.ByteRangeGapMatchesContents);
        Assert.DoesNotContain(report.Findings, finding => finding.Code == "SignatureByteRangeContentsGapMismatch");
    }

    [Fact]
    public void Validate_DetectsContentsGapMismatchInCompactPdfSyntax() {
        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(BuildCompactMismatchedSignatureGapPdf());

        PdfSignatureValidationResult result = Assert.Single(report.Signatures);
        Assert.Equal(10, result.ByteRangeGapLength);
        Assert.Equal(8, result.Signature.ContentsEncodedSizeBytes);
        Assert.False(result.ByteRangeGapMatchesContents);
        Assert.Contains(report.Findings, finding => finding.Code == "SignatureByteRangeContentsGapMismatch");
    }

    [Fact]
    public void SecurityScan_HandlesManyDefaultSizedSignatureContentsWithinParserBudget() {
        const int signatureCount = 24;
        const int reservedSignatureContentsBytes = 32768;
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits {
                MaxObjectParsingTime = TimeSpan.FromSeconds(15)
            }
        };

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(
            BuildManySignaturePdf(signatureCount, reservedSignatureContentsBytes),
            options);

        Assert.Equal(signatureCount, security.Signatures.Count);
        Assert.All(
            security.Signatures,
            signature => Assert.Equal(
                (reservedSignatureContentsBytes * 2) + 2,
                signature.ContentsEncodedSizeBytes));
    }

    [Fact]
    public void Validate_DoesNotFabricateEncodedLengthForUnterminatedLiteralContents() {
        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(
            BuildUnterminatedSignatureContentsPdf("(001122"));

        PdfSignatureValidationResult result = Assert.Single(report.Signatures);
        Assert.Null(result.Signature.ContentsEncodedSizeBytes);
        Assert.Null(result.ByteRangeGapMatchesContents);
    }

    [Fact]
    public void Validate_DoesNotTreatNestedClosingDelimiterAsLiteralTermination() {
        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(
            BuildUnterminatedSignatureContentsPdf("(001(122"));

        PdfSignatureValidationResult result = Assert.Single(report.Signatures);
        Assert.Null(result.Signature.ContentsEncodedSizeBytes);
        Assert.Null(result.ByteRangeGapMatchesContents);
    }

    [Fact]
    public void Parser_DoesNotFabricateEncodedLengthForUnterminatedHexContents() {
        byte[] pdf = BuildIndirectUnterminatedHexSignatureContentsPdf();

        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfStringObj contents = Assert.IsType<PdfStringObj>(objects[4].Value);
        PdfSignatureValidationResult result = Assert.Single(PdfSignatureValidator.Validate(pdf).Signatures);

        Assert.Null(contents.EncodedTokenLength);
        Assert.Null(result.Signature.ContentsEncodedSizeBytes);
        Assert.Null(result.ByteRangeGapMatchesContents);
    }

    [Fact]
    public void Validate_DoesNotTreatDecodedObjectStreamLengthAsRawContentsSpan() {
        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(
            BuildCompressedObjectStreamSignaturePdf());

        PdfSignatureValidationResult result = Assert.Single(report.Signatures);
        Assert.Equal(10, result.ByteRangeGapLength);
        Assert.Null(result.Signature.ContentsEncodedSizeBytes);
        Assert.Null(result.ByteRangeGapMatchesContents);
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
    public void Validate_FlagsByteRangeGapThatDoesNotMatchContentsHexLiteral() {
        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(BuildMismatchedSignatureGapPdf());

        Assert.True(report.HasSignatures);
        Assert.False(report.IsStructurallyValid);
        PdfSignatureValidationResult result = Assert.Single(report.Signatures);
        Assert.False(result.ByteRangeGapMatchesContents);
        Assert.Contains(report.Findings, finding => finding.Code == "SignatureByteRangeContentsGapMismatch");
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
    public void PrepareExternalSignature_IgnoresEarlierMatchingContentsPlaceholders() {
        byte[] pdf = BuildPdfWithEarlierContentsPlaceholder(reservedSignatureContentsBytes: 256);

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            pdf,
            new PdfExternalSignatureOptions {
                FieldName = "Approval",
                ReservedSignatureContentsBytes = 256
            });

        Assert.True(preparation.ContentsHexOffset > pdf.Length);
        Assert.Equal(preparation.ContentsHexOffset - 1, preparation.ByteRangeValues[1]);
        Assert.True(PdfSignatureValidator.Validate(preparation.PreparedPdf).IsStructurallyValid);
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

    [Fact]
    public void ApplyExternalSignature_IgnoresEarlierNonSignatureContentsPlaceholders() {
        const int reservedSignatureContentsBytes = 256;
        string zeros = new string('0', reservedSignatureContentsBytes * 2);
        byte[] pdf = BuildPdfWithEarlierContentsPlaceholder(reservedSignatureContentsBytes);
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            pdf,
            new PdfExternalSignatureOptions {
                FieldName = "Approval",
                ReservedSignatureContentsBytes = reservedSignatureContentsBytes
            });
        string preparedText = Encoding.ASCII.GetString(preparation.PreparedPdf);
        int earlierContentsOffset = preparedText.IndexOf("/Contents <" + zeros + ">", StringComparison.Ordinal);

        byte[] signature = { 0x30, 0x82, 0x01, 0x0A, 0xAA, 0x55 };
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation.PreparedPdf, signature);
        string signedText = Encoding.ASCII.GetString(signed);

        Assert.True(earlierContentsOffset >= 0);
        Assert.True(earlierContentsOffset < preparation.ContentsHexOffset);
        Assert.Equal(zeros, signedText.Substring(earlierContentsOffset + "/Contents <".Length, zeros.Length));
        Assert.StartsWith("3082010AAA55", signedText.Substring(preparation.ContentsHexOffset, signature.Length * 2), StringComparison.Ordinal);

        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(signed);
        PdfSignatureValidationResult result = Assert.Single(report.Signatures);
        Assert.True(report.IsStructurallyValid);
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
            "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /Name (Alice) /ByteRange [0 8 16 30] /Contents <001122> /Reference [<< /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /V /1.2 /P 2 >> >>] >>",
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

    private static byte[] BuildPdfWithEarlierContentsPlaceholder(int reservedSignatureContentsBytes) {
        string zeros = new string('0', reservedSignatureContentsBytes * 2);
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [4 0 R] /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Annot /Subtype /Text /Rect [10 10 30 30] /Contents <" + zeros + "> >>",
            "endobj",
            "5 0 obj",
            "<< /Length 34 >>",
            "stream",
            "BT /F1 12 Tf 72 720 Td (Text) Tj ET",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildMismatchedSignatureGapPdf() {
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
            "<< /FT /Sig /T (Mismatch) /V 6 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Fields [4 0 R] /SigFlags 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [0 10 20 30] /Contents <001122> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCompactMismatchedSignatureGapPdf() {
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
            "<< /FT /Sig /T (Compact) /V 6 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Fields [4 0 R] /SigFlags 1 >>",
            "endobj",
            "6 0 obj <</Type/Sig/Filter/Adobe.PPKLite/SubFilter/adbe.pkcs7.detached/ByteRange[0 10 20 30]/Contents<001122> >> endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildManySignaturePdf(int signatureCount, int reservedSignatureContentsBytes) {
        var pdf = new StringBuilder()
            .AppendLine("%PDF-1.7")
            .AppendLine("1 0 obj")
            .AppendLine("<< /Type /Catalog /Pages 2 0 R >>")
            .AppendLine("endobj")
            .AppendLine("2 0 obj")
            .AppendLine("<< /Type /Pages /Count 0 /Kids [] >>")
            .AppendLine("endobj");
        string contents = new string('0', reservedSignatureContentsBytes * 2);
        for (int i = 0; i < signatureCount; i++) {
            int objectNumber = i + 3;
            pdf.Append(objectNumber)
                .AppendLine(" 0 obj")
                .Append("<< /Type /Sig /ByteRange [0 1 3 1] /Contents <")
                .Append(contents)
                .AppendLine("> >>")
                .AppendLine("endobj");
        }

        pdf.AppendLine("trailer")
            .Append("<< /Root 1 0 R /Size ")
            .Append(signatureCount + 3)
            .AppendLine(" >>")
            .AppendLine("startxref")
            .AppendLine("0")
            .AppendLine("%%EOF");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }

    private static byte[] BuildUnterminatedSignatureContentsPdf(string contentsToken) {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Sig /ByteRange [0 10 20 30] /Contents " + contentsToken + " >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 4 >>",
            "startxref",
            "0",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildIndirectUnterminatedHexSignatureContentsPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Sig /ByteRange [0 10 20 30] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<001122",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "startxref",
            "0",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCompressedObjectStreamSignaturePdf() {
        const string header = "4 0\n";
        const string signature = "<< /Type /Sig /ByteRange [0 1 11 1] /Contents <00112233> >>";
        byte[] decoded = Encoding.ASCII.GetBytes(header + signature);
        byte[] compressed;
        using (var output = new MemoryStream()) {
            using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
                deflate.Write(decoded, 0, decoded.Length);
            }

            compressed = output.ToArray();
        }

        var encoded = new StringBuilder(compressed.Length * 2 + 1);
        foreach (byte value in compressed) {
            encoded.Append(value.ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
        }
        encoded.Append('>');

        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [] >>",
            "endobj",
            "5 0 obj",
            "<< /Type /ObjStm /N 1 /First " +
                Encoding.ASCII.GetByteCount(header).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " /Filter [/ASCIIHexDecode /FlateDecode] /Length " +
                encoded.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " >>",
            "stream",
            encoded.ToString(),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "startxref",
            "0",
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

    private static byte[] BuildSignedPdfWithEncodedContentsGap() {
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
            "<< /FT /Sig /T (Approval) /V 6 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Fields [4 0 R] /SigFlags 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [0 10 18 30] /Contents <001122> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSignedPdfWithWhitespaceContentsGap() {
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
            "<< /FT /Sig /T (Approval) /V 6 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Fields [4 0 R] /SigFlags 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [0 10 20 30] /Contents <00 11\n22> >>",
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
