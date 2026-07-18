using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    [Fact]
    public void ProofReportRequiresVeraPdfBeforeClaimingPdfA() {
        var options = new PdfOptions()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B);

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.PdfA3B,
            options,
            externalValidations: null,
            generatedStandardFonts: Array.Empty<PdfStandardFont>());

        Assert.True(proof.IsInternallyReady);
        Assert.False(proof.HasRequiredExternalValidation);
        Assert.False(proof.CanClaimConformance);
        Assert.Contains(PdfExternalValidatorKind.VeraPdf, proof.RequiredExternalValidators);
        Assert.Contains(PdfExternalValidatorKind.VeraPdf, proof.MissingExternalValidators);
    }

    [Fact]
    public void ProofReportAllowsPdfAClaimWhenRequiredValidatorPasses() {
        var options = new PdfOptions()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B);
        byte[] artifact = PdfDocument.Create(options).ToBytes();
        PdfExternalValidationResult veraPdf = PdfExternalValidationResult.PassedForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "1.30.2",
            "PDF/A-3b profile accepted.",
            artifact,
            "PDF/A-3b");

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.PdfA3B,
            options,
            artifact,
            new[] { veraPdf },
            generatedStandardFonts: Array.Empty<PdfStandardFont>());

        Assert.True(proof.IsInternallyReady);
        Assert.True(proof.HasRequiredExternalValidation);
        Assert.True(proof.CanClaimConformance);
        Assert.Empty(proof.MissingExternalValidators);
        Assert.Empty(proof.FailedExternalValidations);
        Assert.Empty(proof.MissingRequirements);
        Assert.Empty(proof.UnsupportedRequirements);
        Assert.Empty(proof.BlockingRequirements);
        Assert.Equal(
            proof.Readiness.Requirements.Count,
            proof.Readiness.Requirements.Select(static requirement => requirement.Id).Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void ProofReportCountsUnprofiledValidatorResultForRequestedProof() {
        var options = new PdfOptions()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B);
        byte[] artifact = PdfDocument.Create(options).ToBytes();
        PdfExternalValidationResult veraPdf = PdfExternalValidationResult.PassedForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "1.30.2",
            "PDF/A profile accepted.",
            artifact);

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.PdfA3B,
            options,
            artifact,
            new[] { veraPdf },
            generatedStandardFonts: Array.Empty<PdfStandardFont>());

        Assert.True(proof.IsInternallyReady);
        Assert.True(proof.HasRequiredExternalValidation);
        Assert.True(proof.CanClaimConformance);
        Assert.Equal(1, proof.PassedExternalValidationCount);
        Assert.Empty(proof.MissingExternalValidators);
    }

    [Fact]
    public void DocumentProofUsesGeneratedEvidenceBeforeAcceptingExternalPdfAValidation() {
        PdfDocument document = PdfDocument.Create(new PdfOptions()
                .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B))
            .Paragraph(paragraph => paragraph.Text("Generated text must still have embedded font proof."));
        byte[] artifact = document.ToBytes();
        PdfExternalValidationResult veraPdf = PdfExternalValidationResult.PassedForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "1.30.2",
            "PDF/A-3b profile accepted.",
            artifact,
            "PDF/A-3b");

        PdfComplianceProofReport proof = document.AssessComplianceProof(PdfComplianceProfile.PdfA3B, artifact, new[] { veraPdf });

        Assert.False(proof.IsInternallyReady);
        Assert.True(proof.HasRequiredExternalValidation);
        Assert.False(proof.CanClaimConformance);
        Assert.Equal("InternalGaps", proof.ProofStatus);
        AssertRequirement(proof.Readiness, "embedded-font-coverage", PdfComplianceRequirementStatus.Missing);
    }

    [Fact]
    public void DocumentProofWithoutArtifactCannotBeClaimable() {
        PdfExternalValidationResult veraPdf = PdfExternalValidationResult.Passed(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "PDF/A-3b profile accepted.",
            "PDF/A-3b");

        PdfComplianceProofReport proof = PdfDocument.Create(new PdfOptions()
                .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B)
                .RequireCompliance(PdfComplianceProfile.PdfA3B))
            .AssessComplianceProof(new[] { veraPdf });

        Assert.Equal(PdfComplianceProfile.PdfA3B, proof.Profile);
        Assert.True(proof.IsInternallyReady);
        Assert.False(proof.HasArtifactEvidence);
        Assert.False(proof.HasRequiredExternalValidation);
        Assert.False(proof.CanClaimConformance);
        Assert.Equal("MissingArtifactEvidence", proof.ProofStatus);
    }

    [Fact]
    public void ProofReportRejectsValidatorPassBoundToDifferentArtifactWithSameLength() {
        var options = new PdfOptions()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B);
        byte[] validatedArtifact = PdfDocument.Create(options).ToBytes();
        byte[] requestedArtifact = (byte[])validatedArtifact.Clone();
        requestedArtifact[requestedArtifact.Length - 1] ^= 0x01;
        var validatedAtUtc = new System.DateTimeOffset(2026, 7, 15, 0, 0, 0, System.TimeSpan.Zero);
        PdfExternalValidationResult veraPdf = PdfExternalValidationResult.PassedForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "1.30.2",
            "PDF/A-3b profile accepted.",
            validatedArtifact,
            "PDF/A-3b",
            warnings: new[] { "Informational validator warning." },
            validatedAtUtc: validatedAtUtc);

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.PdfA3B,
            options,
            requestedArtifact,
            new[] { veraPdf },
            generatedStandardFonts: Array.Empty<PdfStandardFont>());

        Assert.True(proof.IsInternallyReady);
        Assert.True(proof.HasArtifactEvidence);
        Assert.False(proof.HasRequiredExternalValidation);
        Assert.False(proof.CanClaimConformance);
        Assert.Equal("MissingExternalValidation", proof.ProofStatus);
        Assert.Contains(PdfExternalValidatorKind.VeraPdf, proof.MissingExternalValidators);
        Assert.True(veraPdf.HasArtifactBinding);
        Assert.Equal("1.30.2", veraPdf.ValidatorVersion);
        Assert.Equal(validatedArtifact.LongLength, veraPdf.ArtifactSizeBytes);
        Assert.NotEqual(proof.ArtifactSha256, veraPdf.ArtifactSha256);
        Assert.Equal(validatedAtUtc, veraPdf.ValidatedAtUtc);
        Assert.Equal(new[] { "Informational validator warning." }, veraPdf.Warnings);
    }

    [Fact]
    public void ReadinessAcceptsOpenTypeCffEmbeddedFontsForPdfAFontCoverage() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        var options = new PdfOptions()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B)
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath!), "OfficeIMO Source Serif CFF");

        PdfComplianceReadinessReport readiness = PdfComplianceAnalyzer.Assess(
            PdfComplianceProfile.PdfA3B,
            options,
            new[] { PdfStandardFont.Helvetica });

        PdfComplianceRequirement requirement = AssertRequirement(readiness, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
        Assert.Contains("OpenType/CFF", requirement.Diagnostic, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadinessReportsMalformedOpenTypeCffFontsAsInvalidPdfAFontCoverage() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var options = new PdfOptions()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B)
            .EmbedStandardFont(PdfStandardFont.Helvetica, CreateOverflowingCmapOpenTypeCffFont(File.ReadAllBytes(fontPath!)), "OfficeIMO Malformed CFF");

        PdfComplianceReadinessReport readiness = PdfComplianceAnalyzer.Assess(
            PdfComplianceProfile.PdfA3B,
            options,
            new[] { PdfStandardFont.Helvetica });

        PdfComplianceRequirement requirement = AssertRequirement(readiness, "embedded-font-coverage", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("Replace invalid embedded TrueType or OpenType/CFF mappings", requirement.Diagnostic, StringComparison.Ordinal);
        Assert.Contains("Helvetica", requirement.Diagnostic, StringComparison.Ordinal);
    }

    [Fact]
    public void ProofReportRejectsValidatorPassForDifferentProfile() {
        var options = new PdfOptions()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B);
        byte[] artifact = PdfDocument.Create(options).ToBytes();
        PdfExternalValidationResult veraPdf = PdfExternalValidationResult.PassedForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "1.30.2",
            "PDF/A-2b profile accepted.",
            artifact,
            "PDF/A-2b");

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.PdfA3B,
            options,
            artifact,
            new[] { veraPdf },
            generatedStandardFonts: Array.Empty<PdfStandardFont>());

        Assert.True(proof.IsInternallyReady);
        Assert.False(proof.HasRequiredExternalValidation);
        Assert.False(proof.CanClaimConformance);
        Assert.Contains(PdfExternalValidatorKind.VeraPdf, proof.MissingExternalValidators);
    }

    [Fact]
    public void ProofReportFindExternalValidationReturnsRequestedProfileResult() {
        var options = new PdfOptions()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B);
        byte[] artifact = PdfDocument.Create(options).ToBytes();
        PdfExternalValidationResult otherProfile = PdfExternalValidationResult.PassedForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "1.30.2",
            "PDF/A-2b profile accepted.",
            artifact,
            "PDF/A-2b");
        PdfExternalValidationResult requestedProfile = PdfExternalValidationResult.PassedForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "1.30.2",
            "PDF/A-3b profile accepted.",
            artifact,
            "PDF/A-3b");

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.PdfA3B,
            options,
            artifact,
            new[] { otherProfile, requestedProfile },
            generatedStandardFonts: Array.Empty<PdfStandardFont>());

        Assert.Same(requestedProfile, proof.FindExternalValidation(PdfExternalValidatorKind.VeraPdf));
    }

    [Fact]
    public void ProofReportBlocksClaimWhenRequiredValidatorFailedEvenIfLaterResultPassed() {
        var options = new PdfOptions()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B);
        byte[] artifact = PdfDocument.Create(options).ToBytes();
        PdfExternalValidationResult failed = PdfExternalValidationResult.FromExitCodeForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            1,
            "veraPDF",
            "1.30.2",
            "Output intent failed policy validation.",
            artifact,
            "PDF/A-3b",
            successExitCode: 0);
        PdfExternalValidationResult passed = PdfExternalValidationResult.PassedForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "1.30.2",
            "A later run passed.",
            artifact,
            "PDF/A-3b");

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.PdfA3B,
            options,
            artifact,
            new[] { failed, passed },
            generatedStandardFonts: Array.Empty<PdfStandardFont>());

        Assert.True(proof.IsInternallyReady);
        Assert.False(proof.CanClaimConformance);
        Assert.Contains(PdfExternalValidatorKind.VeraPdf, proof.MissingExternalValidators);
        PdfExternalValidationResult result = Assert.Single(proof.FailedExternalValidations);
        Assert.Equal(PdfExternalValidationStatus.Failed, result.Status);
        Assert.Equal(1, result.ExitCode);
    }

    [Fact]
    public void ProofReportExposesBlockingValidatorProofRowWhenFailureAndPassAreSupplied() {
        var options = new PdfOptions()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B);
        byte[] artifact = PdfDocument.Create(options).ToBytes();
        PdfExternalValidationResult failed = PdfExternalValidationResult.FromExitCodeForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            1,
            "veraPDF",
            "1.30.2",
            "Output intent failed policy validation.",
            artifact,
            "PDF/A-3b",
            successExitCode: 0);
        PdfExternalValidationResult passed = PdfExternalValidationResult.PassedForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "1.30.2",
            "A later run passed.",
            artifact,
            "PDF/A-3b");

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.PdfA3B,
            options,
            artifact,
            new[] { failed, passed },
            generatedStandardFonts: Array.Empty<PdfStandardFont>());

        PdfExternalValidatorProof row = Assert.Single(proof.ExternalValidatorProofs);
        Assert.NotNull(proof.FindExternalValidatorProof(PdfExternalValidatorKind.VeraPdf));
        Assert.Equal(PdfExternalValidatorKind.VeraPdf, row.ValidatorKind);
        Assert.Equal(PdfExternalValidatorProofStatus.Failed, row.Status);
        Assert.False(row.IsSatisfied);
        Assert.True(row.HasBlockingValidation);
        Assert.True(row.BlocksConformanceClaim);
        Assert.Same(failed, row.PrimaryValidation);
        Assert.Same(failed, row.BlockingValidation);
        Assert.Same(passed, row.PassingValidation);
        Assert.Equal("veraPDF", row.ValidatorName);
        Assert.Equal("Output intent failed policy validation.", row.Diagnostic);
        Assert.Equal(1, row.ExitCode);
    }

    [Fact]
    public void ProofReportRequiresVeraPdfAndMustangForElectronicInvoiceProfiles() {
        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.FacturX,
            new PdfOptions(),
            externalValidations: Array.Empty<PdfExternalValidationResult>(),
            generatedStandardFonts: Array.Empty<PdfStandardFont>());

        Assert.False(proof.IsInternallyReady);
        Assert.False(proof.CanClaimConformance);
        Assert.Contains(PdfExternalValidatorKind.VeraPdf, proof.RequiredExternalValidators);
        Assert.Contains(PdfExternalValidatorKind.Mustang, proof.RequiredExternalValidators);
        Assert.Contains(PdfExternalValidatorKind.VeraPdf, proof.MissingExternalValidators);
        Assert.Contains(PdfExternalValidatorKind.Mustang, proof.MissingExternalValidators);
    }

    [Fact]
    public void ProofReportExposesPerValidatorRowsForElectronicInvoiceProfiles() {
        byte[] artifact = PdfDocument.Create(new PdfOptions().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B)).ToBytes();
        PdfExternalValidationResult veraPdf = PdfExternalValidationResult.PassedForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "1.30.2",
            "PDF/A-3b carrier accepted.",
            artifact,
            "PDF/A-3b");
        PdfExternalValidationResult mustang = PdfExternalValidationResult.NotRunForArtifact(
            PdfExternalValidatorKind.Mustang,
            "Mustang",
            "2.20.0",
            "Mustang is not configured on this runner.",
            artifact,
            "Factur-X");

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.FacturX,
            new PdfOptions(),
            artifact,
            externalValidations: new[] { veraPdf, mustang },
            generatedStandardFonts: Array.Empty<PdfStandardFont>());

        Assert.Equal(new[] { PdfExternalValidatorKind.VeraPdf, PdfExternalValidatorKind.Mustang }, proof.ExternalValidatorProofs.Select(row => row.ValidatorKind).ToArray());
        PdfExternalValidatorProof veraPdfRow = proof.FindExternalValidatorProof(PdfExternalValidatorKind.VeraPdf)!;
        PdfExternalValidatorProof mustangRow = proof.FindExternalValidatorProof(PdfExternalValidatorKind.Mustang)!;

        Assert.Equal(PdfExternalValidatorProofStatus.Passed, veraPdfRow.Status);
        Assert.True(veraPdfRow.IsSatisfied);
        Assert.False(veraPdfRow.BlocksConformanceClaim);
        Assert.Same(veraPdf, veraPdfRow.PrimaryValidation);

        Assert.Equal(PdfExternalValidatorProofStatus.NotRun, mustangRow.Status);
        Assert.True(mustangRow.IsMissing);
        Assert.True(mustangRow.BlocksConformanceClaim);
        Assert.Same(mustang, mustangRow.PrimaryValidation);
        Assert.Equal("Mustang is not configured on this runner.", mustangRow.Diagnostic);
        Assert.Contains(PdfExternalValidatorKind.Mustang, proof.MissingExternalValidators);
        Assert.False(proof.CanClaimConformance);
    }

    [Fact]
    public void ProofReportRequiresPdfUaValidatorForPdfUaProfiles() {
        PdfComplianceReadinessReport readiness = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfUa1, new PdfOptions());
        PdfComplianceReadinessReport ua2Readiness = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfUa2, new PdfOptions());

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(readiness);
        PdfComplianceProofReport ua2Proof = PdfComplianceAnalyzer.AssessProof(ua2Readiness);

        Assert.Contains(PdfExternalValidatorKind.PdfUaValidator, proof.RequiredExternalValidators);
        Assert.Contains(PdfExternalValidatorKind.PdfUaValidator, proof.MissingExternalValidators);
        Assert.Contains(PdfExternalValidatorKind.PdfUaValidator, ua2Proof.RequiredExternalValidators);
        Assert.Contains(PdfExternalValidatorKind.PdfUaValidator, ua2Proof.MissingExternalValidators);
        AssertRequirement(readiness, "pdfua-validation", PdfComplianceRequirementStatus.Unsupported);
    }

    private static byte[] CreateOverflowingCmapOpenTypeCffFont(byte[] fontData) {
        byte[] corrupted = fontData.ToArray();
        int cmapOffset = FindOpenTypeTableOffset(corrupted, "cmap");
        int encodingRecordOffset = cmapOffset + 4;
        WriteUInt32(corrupted, encodingRecordOffset + 4, 0x80000000U);
        return corrupted;
    }

    private static int FindOpenTypeTableOffset(byte[] data, string tag) {
        int tableCount = ReadUInt16(data, 4);
        for (int i = 0; i < tableCount; i++) {
            int recordOffset = 12 + i * 16;
            if (Encoding.ASCII.GetString(data, recordOffset, 4) == tag) {
                return (int)ReadUInt32(data, recordOffset + 8);
            }
        }

        throw new InvalidOperationException("Required OpenType table '" + tag + "' was not found.");
    }

    private static ushort ReadUInt16(byte[] data, int offset) =>
        (ushort)((data[offset] << 8) | data[offset + 1]);

    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24) |
        ((uint)data[offset + 1] << 16) |
        ((uint)data[offset + 2] << 8) |
        data[offset + 3];

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}
