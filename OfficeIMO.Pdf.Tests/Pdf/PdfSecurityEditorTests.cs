using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSecurityEditorTests {
    [Fact]
    public void EncryptExistingPdfDefaultsToAes256AndPreservesDocumentContent() {
        byte[] source = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        var encryption = new PdfStandardEncryptionOptions("new-open") {
            OwnerPassword = "new-owner",
            AllowedPermissions = PdfStandardPermissions.Print | PdfStandardPermissions.CopyContents
        };

        PdfSecurityMutationResult result = PdfSecurityEditor.Encrypt(source, encryption);
        PdfDocumentInfo output = PdfInspector.Inspect(result.Pdf, new PdfReadOptions { Password = "new-open" });

        Assert.Equal(PdfSecurityMutationKind.Encrypt, result.Kind);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, result.MutationPlan.ExecutionMode);
        Assert.True(result.PreservationReport.IsPreserved, result.PreservationReport.Summary);
        Assert.True(result.IsEncrypted);
        Assert.Equal(6, output.Security.EncryptionRevision);
        Assert.Equal(256, output.Security.EncryptionLengthBits);
        Assert.Equal(PdfStandardPermissions.Print | PdfStandardPermissions.CopyContents, output.Security.AllowedStandardPermissions);
        Assert.Equal(result.PreservationReport.Original.PageCount, output.PageCount);
        Assert.Equal(result.PreservationReport.Original.Attachments.Count, output.Attachments.Count);
        Assert.Equal(result.PreservationReport.Original.Outlines.Count, output.Outlines.Count);
        Assert.Throws<PdfInvalidPasswordException>(() => PdfInspector.Inspect(result.Pdf, new PdfReadOptions { Password = "wrong" }));
    }

    [Fact]
    public void DecryptRequiresOwnerPasswordAndProducesReadablePlaintext() {
        byte[] encrypted = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Meta(title: "Owner-authorized source")
            .Paragraph(paragraph => paragraph.Text("Owner-only decryption proof"))
            .ToBytes();

        PdfMutationPlan userPlan = PdfMutationPlanner.Plan(
            encrypted,
            PdfMutationOperation.ChangeEncryption,
            new PdfReadOptions { Password = "open" });
        PdfOperationResult<PdfSecurityMutationResult> refused = PdfDocument.Open(encrypted).TryDecrypt("open");
        PdfSecurityMutationResult result = PdfSecurityEditor.Decrypt(encrypted, "owner");

        Assert.Equal(PdfPasswordAuthenticationRole.User, userPlan.Preflight.Probe.Security.PasswordAuthenticationRole);
        Assert.Equal(PdfMutationExecutionMode.Blocked, userPlan.ExecutionMode);
        Assert.Contains("FullRewrite.Encryption.OwnerAuthorizationRequired", userPlan.BlockerCodes);
        Assert.False(refused.Succeeded);
        Assert.False(result.IsEncrypted);
        Assert.Equal(PdfPasswordAuthenticationRole.Owner, result.SourceSecurity.PasswordAuthenticationRole);
        Assert.True(result.PreservationReport.IsPreserved, result.PreservationReport.Summary);
        Assert.Equal("Owner-authorized source", PdfInspector.Inspect(result.Pdf).Metadata.Title);
        Assert.Contains("Owner-only decryption proof", PdfTextExtractor.ExtractAllText(result.Pdf), StringComparison.Ordinal);
    }

    [Fact]
    public void ReencryptChangesPasswordAlgorithmAndPermissions() {
        byte[] encrypted = PdfDocument.Create(new PdfOptions().SetEncryption("old-open", "old-owner"))
            .Paragraph(paragraph => paragraph.Text("Re-encryption proof"))
            .ToBytes();
        var replacement = new PdfStandardEncryptionOptions("new-open") {
            OwnerPassword = "new-owner",
            Algorithm = PdfStandardEncryptionAlgorithm.Aes128,
            AllowedPermissions = PdfStandardPermissions.FillForms
        };

        PdfSecurityMutationResult result = PdfDocument.Open(encrypted).Reencrypt("old-owner", replacement);
        PdfDocumentInfo userInfo = PdfInspector.Inspect(result.Pdf, new PdfReadOptions { Password = "new-open" });
        PdfDocumentInfo ownerInfo = PdfInspector.Inspect(result.Pdf, new PdfReadOptions { Password = "new-owner" });

        Assert.Equal(PdfSecurityMutationKind.Reencrypt, result.Kind);
        Assert.True(result.PreservationReport.IsPreserved, result.PreservationReport.Summary);
        Assert.Equal(4, userInfo.Security.EncryptionRevision);
        Assert.Equal(128, userInfo.Security.EncryptionLengthBits);
        Assert.Equal(PdfStandardPermissions.FillForms, userInfo.Security.AllowedStandardPermissions);
        Assert.Equal(PdfPasswordAuthenticationRole.User, userInfo.Security.PasswordAuthenticationRole);
        Assert.Equal(PdfPasswordAuthenticationRole.Owner, ownerInfo.Security.PasswordAuthenticationRole);
        Assert.Throws<PdfInvalidPasswordException>(() => PdfInspector.Inspect(result.Pdf, new PdfReadOptions { Password = "old-open" }));
        Assert.Throws<PdfPermissionDeniedException>(() =>
            PdfTextExtractor.ExtractAllText(result.Pdf, (PdfTextLayoutOptions?)null, new PdfReadOptions { Password = "new-open" }));
        Assert.Contains(
            "Re-encryption proof",
            PdfTextExtractor.ExtractAllText(result.Pdf, (PdfTextLayoutOptions?)null, new PdfReadOptions { Password = "new-owner" }),
            StringComparison.Ordinal);
    }

    [Fact]
    public void SecurityRewriteRemainsBlockedForSignedDocuments() {
        byte[] signed = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();

        PdfMutationPlan plan = PdfMutationPlanner.Plan(signed, PdfMutationOperation.ChangeEncryption);

        Assert.Equal(PdfMutationExecutionMode.Blocked, plan.ExecutionMode);
        Assert.Contains("FullRewrite.SignaturePreservation", plan.BlockerCodes);
    }

    [Fact]
    public void EncryptNormalizesSupportedXrefObjectStreamAndGenerationInputs() {
        byte[][] sources = {
            PdfRewritePreservationTestSupport.BuildSourceStructurePreservationProofPdf(),
            PdfExternalDocumentCompatibilityTests.BuildHybridClassicXrefPdfWithXRefStmAndTrailingStaleDuplicatePage(),
            PdfExternalDocumentCompatibilityTests.BuildIncrementalClassicXrefPdfWithWrongGenerationReplacementPage()
        };

        for (int i = 0; i < sources.Length; i++) {
            PdfSecurityMutationResult result = PdfSecurityEditor.Encrypt(
                sources[i],
                new PdfStandardEncryptionOptions("corpus-open") { OwnerPassword = "corpus-owner" });
            PdfDocumentInfo output = PdfInspector.Inspect(result.Pdf, new PdfReadOptions { Password = "corpus-open" });

            Assert.True(result.PreservationReport.IsPreserved, "Case " + i + ": " + result.PreservationReport.Summary);
            Assert.True(output.Security.HasEncryption);
            Assert.Equal(result.PreservationReport.Original.PageCount, output.PageCount);
            Assert.False(output.Security.HasXrefStreams);
            Assert.False(output.Security.HasObjectStreams);
            Assert.False(output.Security.HasIncrementalUpdates);
        }
    }
}
