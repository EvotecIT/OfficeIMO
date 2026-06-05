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
}
