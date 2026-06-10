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
        PdfExternalValidationResult veraPdf = PdfExternalValidationResult.Passed(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "PDF/A-3b profile accepted.",
            "PDF/A-3b");

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.PdfA3B,
            options,
            new[] { veraPdf },
            generatedStandardFonts: Array.Empty<PdfStandardFont>());

        Assert.True(proof.IsInternallyReady);
        Assert.True(proof.HasRequiredExternalValidation);
        Assert.True(proof.CanClaimConformance);
        Assert.Empty(proof.MissingExternalValidators);
        Assert.Empty(proof.FailedExternalValidations);
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
        PdfExternalValidationResult veraPdf = PdfExternalValidationResult.Passed(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "PDF/A-2b profile accepted.",
            "PDF/A-2b");

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.PdfA3B,
            options,
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
        PdfExternalValidationResult otherProfile = PdfExternalValidationResult.Passed(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "PDF/A-2b profile accepted.",
            "PDF/A-2b");
        PdfExternalValidationResult requestedProfile = PdfExternalValidationResult.Passed(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "PDF/A-3b profile accepted.",
            "PDF/A-3b");

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.PdfA3B,
            options,
            new[] { otherProfile, requestedProfile },
            generatedStandardFonts: Array.Empty<PdfStandardFont>());

        Assert.Same(requestedProfile, proof.FindExternalValidation(PdfExternalValidatorKind.VeraPdf));
    }

    [Fact]
    public void ProofReportBlocksClaimWhenRequiredValidatorFailedEvenIfLaterResultPassed() {
        var options = new PdfOptions()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B);
        PdfExternalValidationResult failed = PdfExternalValidationResult.Failed(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "Output intent failed policy validation.",
            "PDF/A-3b",
            exitCode: 1);
        PdfExternalValidationResult passed = PdfExternalValidationResult.Passed(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "A later run passed.",
            "PDF/A-3b");

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.PdfA3B,
            options,
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
    public void ProofReportRequiresPdfUaValidatorForPdfUaProfiles() {
        PdfComplianceReadinessReport readiness = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfUa1, new PdfOptions());

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(readiness);

        Assert.Contains(PdfExternalValidatorKind.PdfUaValidator, proof.RequiredExternalValidators);
        Assert.Contains(PdfExternalValidatorKind.PdfUaValidator, proof.MissingExternalValidators);
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
