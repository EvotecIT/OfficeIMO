using System.Collections.Generic;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Fact]
    public void Preflight_BlocksEncryptedPdfButReportsSecuritySettings() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildEncryptedPdf());

        Assert.False(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasEncryption);
        Assert.Null(report.DocumentInfo);
        Assert.True(report.HasSecurityDiagnostics);
        AssertReadBlocker(report, PdfReadBlockerKind.Encryption, "PDF encryption dictionary is missing /O.");
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.Encryption, "Encrypted input requires operation-specific planning. Authenticated unsigned PDFs support proven page, metadata, sanitization, and simple form rewrites when the required permissions are authorized; security changes require owner authorization.");

        PdfDocumentSecurityInfo security = report.Probe.Security;
        Assert.True(security.HasEncryption);
        Assert.True(security.HasReadableEncryptionSettings);
        Assert.Equal(5, security.EncryptObjectNumber);
        Assert.Equal("Standard", security.EncryptionFilter);
        Assert.Equal("adbe.pkcs7.s5", security.EncryptionSubFilter);
        Assert.Equal(2, security.EncryptionVersion);
        Assert.Equal(3, security.EncryptionRevision);
        Assert.Equal(128, security.EncryptionLengthBits);
        Assert.Equal(3900, security.EncryptionPermissions);
        Assert.False(security.EncryptMetadata);
        Assert.True(security.AllowsPrinting);
        Assert.True(security.AllowsModification);
        Assert.True(security.AllowsCopying);
        Assert.True(security.AllowsAnnotationChanges);
        Assert.True(security.AllowsFormFilling);
        Assert.True(security.AllowsAccessibilityExtraction);
        Assert.True(security.AllowsDocumentAssembly);
        Assert.True(security.AllowsHighQualityPrinting);
        Assert.Equal(1, security.RootObjectNumber);
        Assert.Equal(6, security.InfoObjectNumber);
        Assert.True(security.HasTrailerId);
        Assert.Equal(1, security.StartXrefCount);
        Assert.Equal(321, security.LastStartXrefOffset);
        Assert.Equal(new[] { 321 }, security.StartXrefOffsets);
        Assert.Empty(security.PreviousXrefOffsets);
        PdfDocumentRevisionInfo encryptedRevision = Assert.Single(security.Revisions);
        Assert.Equal(1, encryptedRevision.RevisionNumber);
        Assert.Equal(321, encryptedRevision.StartXrefOffset);
        Assert.Null(encryptedRevision.PreviousXrefOffset);
        Assert.False(security.RequiresAppendOnlyMutation);
        Assert.False(report.RequiresAppendOnlyMutation);
        Assert.False(report.CanAppendOnlyMutate);
        Assert.Empty(report.AppendOnlyMutationDiagnostics);
        Assert.Contains("PDF encryption was detected using /Filter /Standard and /SubFilter /adbe.pkcs7.s5 (R=3, 128-bit). No password authorization was established.", report.SecurityDiagnostics);
        Assert.Contains("Raw encryption permissions /P=3900 were detected and are enforced for user-password operations unless the caller explicitly ignores restrictions.", report.SecurityDiagnostics);
        Assert.Contains("PDF encryption was detected using /Filter /Standard and /SubFilter /adbe.pkcs7.s5 (R=3, 128-bit). No password authorization was established.", report.GetCapabilityDiagnostics(PdfPreflightCapability.ExtractText));
        Assert.Contains("Raw encryption permissions /P=3900 were detected and are enforced for user-password operations unless the caller explicitly ignores restrictions.", report.GetCapabilityDiagnostics(PdfPreflightCapability.ManipulatePages));
    }

    [Fact]
    public void PdfDocumentPreflight_ReportsEncryptedInputWithoutOpeningIt() {
        string path = System.IO.Path.GetTempFileName();
        try {
            System.IO.File.WriteAllBytes(path, BuildEncryptedPdf());

            PdfDocumentPreflight report = PdfDocument.Preflight(path);

            Assert.False(report.CanRead);
            Assert.True(report.Probe.HasEncryption);
            Assert.Contains(report.ReadBlockers, blocker => blocker.Kind == PdfReadBlockerKind.Encryption);
        } finally {
            System.IO.File.Delete(path);
        }
    }

    [Fact]
    public void Preflight_DoesNotTreatResourceEncryptNameAsDocumentEncryption() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildPdfWithEncryptNamedXObject());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.False(report.Probe.HasEncryption);
        Assert.NotNull(report.DocumentInfo);
        Assert.False(report.DocumentInfo!.Security.HasEncryption);
        Assert.Null(report.DocumentInfo.Security.EncryptObjectNumber);
        Assert.Null(report.DocumentInfo.Security.EncryptionFilter);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
    }

    [Fact]
    public void Preflight_ExposesAppendOnlyPolicyForUnsignedPdf() {
        byte[] pdf = PdfDocument.Create()
            .Meta(title: "Append-only policy")
            .Paragraph(paragraph => paragraph.Text("Unsigned append-only policy proof."))
            .ToBytes();

        PdfDocumentPreflight report = PdfInspector.Preflight(pdf);
        PdfAppendOnlyMutationReport policy = report.AppendOnlyMutationReport;
        PdfValidationResult validation = PdfValidator.Validate(pdf);

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.False(report.RequiresAppendOnlyMutation);
        Assert.False(report.CanAppendOnlyMutate);
        Assert.True(report.CanAppendMetadataRevision);
        Assert.True(report.CanPrepareExternalSignatureRevision);
        Assert.True(policy.CanAppendMetadata);
        Assert.True(policy.CanPrepareExternalSignature);
        Assert.True(policy.CanAppendAnnotations);
        Assert.Contains("Metadata", policy.SupportedActions);
        Assert.Contains("SignaturePrepare", policy.SupportedActions);
        Assert.Contains("Annotations", policy.SupportedActions);
        Assert.Empty(policy.Blockers);
        Assert.Empty(policy.Warnings);
        Assert.Contains("full rewrites may also be possible", policy.Summary, System.StringComparison.Ordinal);
        Assert.Same(policy, report.AppendOnlyMutationReport);
        Assert.True(validation.CanAppendMetadataRevision);
        Assert.True(validation.CanPrepareExternalSignatureRevision);
        Assert.Equal(policy.SupportedActions, validation.AppendOnlyMutationReport.SupportedActions);
    }

    [Fact]
    public void Inspect_ReadsSignatureAndRevisionSecurityState() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildSignedIncrementalPdf());

        Assert.True(info.HasSecurityState);
        Assert.True(info.HasSignatures);
        PdfDocumentSecurityInfo security = info.Security;
        Assert.False(security.HasEncryption);
        Assert.True(security.HasSignatures);
        Assert.Equal(1, security.SignatureFieldCount);
        Assert.Equal(new[] { 5 }, security.SignatureFieldObjectNumbers);
        Assert.Equal(new[] { "Approval" }, security.SignatureFieldNames);
        Assert.Equal(1, security.SignatureValueCount);
        PdfSignatureInfo signature = Assert.Single(security.Signatures);
        Assert.Equal(1, security.SignatureCount);
        Assert.Equal(6, signature.ObjectNumber);
        Assert.Equal(5, signature.FieldObjectNumber);
        Assert.Equal("Approval", signature.FieldName);
        Assert.True(signature.HasFieldLock);
        Assert.NotNull(signature.FieldLock);
        Assert.Equal("Include", signature.FieldLock.Action);
        Assert.True(signature.FieldLock.LocksIncludedFields);
        Assert.False(signature.FieldLock.LocksAllFields);
        Assert.Equal(new[] { "Total", "Approver" }, signature.FieldLock.Fields);
        Assert.True(signature.HasSeedValue);
        Assert.NotNull(signature.SeedValue);
        Assert.Equal("Adobe.PPKLite", signature.SeedValue.Filter);
        Assert.Equal(new[] { "adbe.pkcs7.detached" }, signature.SeedValue.SubFilters);
        Assert.Equal(new[] { "SHA256", "SHA512" }, signature.SeedValue.DigestMethods);
        Assert.Equal(new[] { "Approval", "Final" }, signature.SeedValue.Reasons);
        Assert.Equal(3, signature.SeedValue.Flags);
        Assert.True(signature.SeedValue.AddRevInfo);
        Assert.Equal(2, signature.SeedValue.MDPPermissionLevel);
        Assert.Equal("Adobe.PPKLite", signature.Filter);
        Assert.Equal("adbe.pkcs7.detached", signature.SubFilter);
        Assert.Equal("Alice", signature.SignerName);
        Assert.Equal("Warsaw", signature.Location);
        Assert.Equal("Approval", signature.Reason);
        Assert.Equal("alice@example.test", signature.ContactInfo);
        Assert.Equal("D:20260607120000+02'00'", signature.SigningTimeRaw);
        Assert.True(signature.HasByteRange);
        Assert.Equal(4, signature.ByteRangeValueCount);
        Assert.Equal(2, signature.ByteRangeSegmentCount);
        Assert.True(signature.HasContents);
        Assert.Equal(3, signature.ContentsSizeBytes);
        Assert.Equal(1, signature.ReferenceCount);
        Assert.True(security.HasByteRange);
        Assert.Equal(4, security.ByteRangeValueCount);
        Assert.Equal(2, security.ByteRangeSegmentCount);
        Assert.Equal(3, security.AcroFormSignatureFlags);
        Assert.True(security.AcroFormSignaturesExist);
        Assert.True(security.AcroFormAppendOnly);
        Assert.True(security.HasDocMDPPermissions);
        Assert.Equal(6, security.DocMDPSignatureObjectNumber);
        Assert.Equal("DocMDP", security.DocMDPTransformMethod);
        Assert.Equal("1.2", security.DocMDPTransformVersion);
        Assert.Equal(2, security.DocMDPPermissionLevel);
        Assert.True(security.HasUsageRights);
        Assert.Equal(new[] { 6 }, security.UsageRightsObjectNumbers);
        Assert.True(security.HasDocumentSecurityStore);
        Assert.True(security.HasLongTermValidationEvidence);
        Assert.Equal(9, security.DocumentSecurityStore.ObjectNumber);
        Assert.Equal(1, security.DocumentSecurityStore.VriEntryCount);
        Assert.Equal(new[] { "ABCDEF" }, security.DocumentSecurityStore.VriKeys);
        Assert.Equal(new[] { 10, 11 }, security.DocumentSecurityStore.CertificateObjectNumbers);
        Assert.Equal(new[] { 12 }, security.DocumentSecurityStore.OcspObjectNumbers);
        Assert.Equal(new[] { 13 }, security.DocumentSecurityStore.CrlObjectNumbers);
        Assert.Equal(new[] { 10 }, security.DocumentSecurityStore.VriCertificateObjectNumbers);
        Assert.Equal(new[] { 12 }, security.DocumentSecurityStore.VriOcspObjectNumbers);
        Assert.Equal(new[] { 13 }, security.DocumentSecurityStore.VriCrlObjectNumbers);
        Assert.Equal(new[] { 14 }, security.DocumentSecurityStore.TimestampObjectNumbers);
        Assert.Equal(4, security.DocumentSecurityStore.TopLevelEvidenceObjectCount);
        Assert.Equal(4, security.DocumentSecurityStore.VriEvidenceObjectCount);
        Assert.Equal(1, security.RootObjectNumber);
        Assert.Equal(8, security.InfoObjectNumber);
        Assert.True(security.HasTrailerId);
        Assert.True(security.HasIncrementalUpdates);
        Assert.True(security.HasPreviousRevision);
        Assert.Equal(2, security.StartXrefCount);
        Assert.Equal(200, security.LastStartXrefOffset);
        Assert.Equal(new[] { 100, 200 }, security.StartXrefOffsets);
        Assert.Equal(new[] { 100 }, security.PreviousXrefOffsets);
        Assert.Equal(2, security.RevisionCount);
        Assert.Equal(1, security.Revisions[0].RevisionNumber);
        Assert.Equal(100, security.Revisions[0].StartXrefOffset);
        Assert.Null(security.Revisions[0].PreviousXrefOffset);
        Assert.Equal(2, security.Revisions[1].RevisionNumber);
        Assert.Equal(200, security.Revisions[1].StartXrefOffset);
        Assert.Equal(100, security.Revisions[1].PreviousXrefOffset);
        Assert.True(security.RequiresAppendOnlyMutation);
        Assert.True(security.BlocksOfficeIMOAppendOnlyMutation);
    }

    [Fact]
    public void Preflight_ReportsSignatureAndRevisionSecurityDiagnostics() {
        byte[] pdf = BuildSignedIncrementalPdf();
        PdfDocumentPreflight report = PdfInspector.Preflight(pdf);

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.RequiresAppendOnlyMutation);
        Assert.False(report.CanAppendOnlyMutate);
        Assert.False(report.CanAppendMetadataRevision);
        Assert.False(report.CanAppendFormFieldRevision);
        Assert.False(report.CanPrepareExternalSignatureRevision);
        Assert.False(report.Can(PdfPreflightCapability.AppendMetadataRevision));
        Assert.False(report.Can(PdfPreflightCapability.AppendFormFieldRevision));
        Assert.False(report.Can(PdfPreflightCapability.PrepareExternalSignatureRevision));
        Assert.True(report.HasSecurityDiagnostics);
        Assert.Contains("PDF signature markers were detected in 1 signature field (Approval); rewrite would invalidate signatures unless append-only signature preservation is implemented.", report.SecurityDiagnostics);
        Assert.Contains("Signature /ByteRange markers were detected with 2 segments.", report.SecurityDiagnostics);
        Assert.Contains("AcroForm /SigFlags indicates append-only updates are expected.", report.SecurityDiagnostics);
        Assert.Contains("Catalog /Perms contains DocMDP permissions; rewrite requires certification-signature preservation semantics.", report.SecurityDiagnostics);
        Assert.Contains("DocMDP certification permission level /P=2 was detected.", report.SecurityDiagnostics);
        Assert.Contains("Catalog /Perms contains usage-rights entries; rewrite may invalidate viewer-extended rights.", report.SecurityDiagnostics);
        Assert.Contains("Document Security Store (/DSS) was detected with 1 VRI entry; signature validation evidence must be preserved during mutation.", report.SecurityDiagnostics);
        Assert.Contains("Incremental update markers were detected (2 startxref sections); safe mutation requires append-only revision preservation.", report.SecurityDiagnostics);
        Assert.Contains("Usage-rights entries must be preserved before append-only mutation.", report.AppendOnlyMutationDiagnostics);

        PdfAppendOnlyMutationReport policy = report.AppendOnlyMutationReport;
        Assert.False(policy.CanAppendAny);
        Assert.False(policy.CanAppendMetadata);
        Assert.False(policy.CanAppendFormFields);
        Assert.False(policy.CanPrepareExternalSignature);
        Assert.True(policy.BlocksAllAppendOnlyMutation);
        Assert.Contains("Metadata", policy.BlockedActions);
        Assert.Contains("FormFill", policy.BlockedActions);
        Assert.Contains("SignaturePrepare", policy.BlockedActions);
        Assert.Contains("Attachments", policy.BlockedActions);
        Assert.Contains("UsageRights", policy.Blockers);
        Assert.Contains("Signed", policy.Blockers);
        Assert.Contains("DocMDPAllowsFormFill", policy.Warnings);
        Assert.Contains("SignedDocMDPFormFill", policy.Warnings);
        Assert.Contains("ExistingIncrementalRevisions", policy.Warnings);
        Assert.Contains("AcroFormAppendOnly", policy.Warnings);
        Assert.Contains("Append-only mutation is blocked for this input", policy.Summary, System.StringComparison.Ordinal);

        PdfAppendOnlyMutationReport directPolicy = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(report.Probe.Security);
        Assert.Equal(policy.SupportedActions, directPolicy.SupportedActions);
        Assert.Equal(policy.BlockedActions, directPolicy.BlockedActions);
        Assert.Equal(policy.Blockers, directPolicy.Blockers);
        Assert.Equal(policy.Warnings, directPolicy.Warnings);

        IReadOnlyList<string> appendMetadataDiagnostics = report.GetCapabilityDiagnostics(PdfPreflightCapability.AppendMetadataRevision);
        Assert.Contains("PDF append-only metadata revision is not available for this PDF.", appendMetadataDiagnostics);
        Assert.Contains("Append-only blocker: UsageRights.", appendMetadataDiagnostics);
        Assert.Contains("Append-only blocker: Signed.", appendMetadataDiagnostics);
        Assert.Contains("Append-only warning: ExistingIncrementalRevisions.", appendMetadataDiagnostics);

        IReadOnlyList<string> appendSignatureDiagnostics = report.GetCapabilityDiagnostics(PdfPreflightCapability.PrepareExternalSignatureRevision);
        Assert.Contains("PDF append-only external-signature preparation is not available for this PDF.", appendSignatureDiagnostics);
        Assert.Contains("Append-only blocker: UsageRights.", appendSignatureDiagnostics);
        Assert.Contains("Append-only warning: AcroFormAppendOnly.", appendSignatureDiagnostics);

        PdfValidationResult validation = PdfValidator.Validate(pdf);
        Assert.False(validation.CanAppendMetadataRevision);
        Assert.False(validation.CanAppendFormFieldRevision);
        Assert.False(validation.CanPrepareExternalSignatureRevision);
        Assert.False(validation.Can(PdfPreflightCapability.AppendMetadataRevision));
        Assert.False(validation.Can(PdfPreflightCapability.AppendFormFieldRevision));
        Assert.False(validation.Can(PdfPreflightCapability.PrepareExternalSignatureRevision));
        Assert.Equal(policy.Blockers, validation.AppendOnlyMutationReport.Blockers);

        IReadOnlyList<string> pageDiagnostics = report.GetCapabilityDiagnostics(PdfPreflightCapability.ManipulatePages);
        Assert.Contains("Signed PDF files are not supported for rewriting by OfficeIMO.Pdf yet.", pageDiagnostics);
        Assert.Contains("PDF signature markers were detected in 1 signature field (Approval); rewrite would invalidate signatures unless append-only signature preservation is implemented.", pageDiagnostics);
        Assert.Contains("Incremental update markers were detected (2 startxref sections); safe mutation requires append-only revision preservation.", pageDiagnostics);

        IReadOnlyList<string> formDiagnostics = report.GetCapabilityDiagnostics(PdfPreflightCapability.FillSimpleFormFields);
        Assert.Contains("Signed PDF files are not supported for form filling or flattening by OfficeIMO.Pdf yet.", formDiagnostics);
        Assert.Contains("PDF signature markers were detected in 1 signature field (Approval); rewrite would invalidate signatures unless append-only signature preservation is implemented.", formDiagnostics);
        Assert.Contains("Signature /ByteRange markers were detected with 2 segments.", formDiagnostics);
    }

    [Fact]
    public void LogicalDocument_Load_ReadsSignatureAndRevisionSecurityState() {
        PdfLogicalDocument document = PdfLogicalDocument.Load(BuildSignedIncrementalPdf());

        Assert.True(document.HasSecurityState);
        Assert.True(document.Security.HasSignatures);
        Assert.True(document.Security.HasIncrementalUpdates);
        Assert.True(document.Security.RequiresAppendOnlyMutation);
        Assert.Equal(new[] { "Approval" }, document.Security.SignatureFieldNames);
        Assert.Equal(new[] { 100, 200 }, document.Security.StartXrefOffsets);
        Assert.Equal(new[] { 100 }, document.Security.PreviousXrefOffsets);
        Assert.Equal(2, document.Security.ByteRangeSegmentCount);
        PdfSignatureInfo signature = Assert.Single(document.Security.Signatures);
        Assert.Equal("Alice", signature.SignerName);
        Assert.Equal(new[] { "Total", "Approver" }, signature.FieldLock?.Fields);
        Assert.Equal(new[] { "SHA256", "SHA512" }, signature.SeedValue?.DigestMethods);
        Assert.True(document.Security.HasLongTermValidationEvidence);
        Assert.Equal(new[] { "ABCDEF" }, document.Security.DocumentSecurityStore.VriKeys);
        Assert.Equal(new[] { 14 }, document.Security.DocumentSecurityStore.TimestampObjectNumbers);
        Assert.Equal("DocMDP", document.Security.DocMDPTransformMethod);
        Assert.Equal(2, document.Security.DocMDPPermissionLevel);
    }

    private static byte[] BuildEncryptedPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Filter /Standard /SubFilter /adbe.pkcs7.s5 /V 2 /R 3 /Length 128 /P 3900 /EncryptMetadata false >>",
            "endobj",
            "6 0 obj",
            "<< /Producer (OfficeIMO security fixture) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Info 6 0 R /Encrypt 5 0 R /ID [(abc) (def)] /Size 7 >>",
            "startxref",
            "321",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithEncryptNamedXObject() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /XObject << /Encrypt 5 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 15 >>",
            "stream",
            "q /Encrypt Do Q",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
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
            "<< /FT /Sig /T (Approval) /V 6 0 R /Subtype /Widget /Rect [10 10 120 40] /Lock << /Type /SigFieldLock /Action /Include /Fields [(Total) (Approver)] >> /SV << /Filter /Adobe.PPKLite /SubFilter [/adbe.pkcs7.detached] /DigestMethod [/SHA256 /SHA512] /Reasons [(Approval) (Final)] /Ff 3 /AddRevInfo true /MDP << /P 2 >> >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /Name (Alice) /Location (Warsaw) /Reason (Approval) /ContactInfo (alice@example.test) /M (D:20260607120000+02'00') /ByteRange [0 10 20 30] /Contents <001122> /Reference [<< /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /V /1.2 /P 2 >> >>] >>",
            "endobj",
            "7 0 obj",
            "<< /Fields [5 0 R] /SigFlags 3 >>",
            "endobj",
            "8 0 obj",
            "<< /Producer (OfficeIMO signed fixture) >>",
            "endobj",
            "9 0 obj",
            "<< /Certs [10 0 R 11 0 R] /OCSPs [12 0 R] /CRLs [13 0 R] /VRI << /ABCDEF << /Cert [10 0 R] /OCSP [12 0 R] /CRL [13 0 R] /TS 14 0 R >> >> >>",
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
            "<< /Type /EmbeddedFile /Length 0 >>",
            "endobj",
            "14 0 obj",
            "<< /Type /TimestampEvidence /Length 0 >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Info 8 0 R /ID [(abc) (def)] /Size 15 /Prev 100 >>",
            "startxref",
            "100",
            "%%EOF",
            "startxref",
            "200",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }
}
