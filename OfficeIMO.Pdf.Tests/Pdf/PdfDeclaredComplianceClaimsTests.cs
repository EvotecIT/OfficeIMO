using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDeclaredComplianceClaimsTests {
    [Fact]
    public void DeclaredClaims_DoNotTreatMetadataAsConformanceProof() {
        PdfComplianceArtifact artifact = PdfDocument.Create(
                new PdfOptions().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B))
            .Meta(title: "Declared compliance gate")
            .CreateComplianceArtifact(PdfComplianceProfile.PdfA3B);
        byte[] bytes = artifact.ToBytes();

        PdfDeclaredComplianceClaimsReport report = PdfDocument.Open(bytes).AssessDeclaredComplianceClaims();
        PdfDeclaredComplianceClaim claim = Assert.Single(report.Claims);

        Assert.Equal(PdfArtifactFingerprint.ComputeSha256(bytes), report.ArtifactSha256);
        Assert.Equal(bytes.LongLength, report.ArtifactSizeBytes);
        Assert.Equal(PdfDeclaredComplianceStandard.PdfA, claim.Standard);
        Assert.Equal("PDF/A-3b", claim.Declaration);
        Assert.Equal(PdfComplianceProfile.PdfA3B, claim.Profile);
        Assert.Equal(PdfDeclaredComplianceClaimStatus.MissingExternalValidation, claim.Status);
        Assert.False(claim.CanClaimConformance);
        Assert.False(report.CanClaimAllDeclaredConformance);
    }

    [Fact]
    public void DeclaredClaims_AcceptArtifactBoundExternalValidationOnlyAfterInternalReadiness() {
        PdfComplianceArtifact artifact = PdfDocument.Create(
                new PdfOptions().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B))
            .Meta(title: "Exact declared compliance proof")
            .CreateComplianceArtifact(PdfComplianceProfile.PdfA3B);
        byte[] bytes = artifact.ToBytes();
        PdfExternalValidationResult validation = PdfExternalValidationResult.PassedForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "1.30.2",
            "PDF/A-3b profile accepted.",
            bytes,
            "PDF/A-3b");

        PdfDeclaredComplianceClaimsReport report = PdfDocument.Open(bytes).AssessDeclaredComplianceClaims(new[] { validation });
        PdfDeclaredComplianceClaim claim = Assert.Single(report.Claims);

        Assert.Equal(PdfDeclaredComplianceClaimStatus.Claimable, claim.Status);
        Assert.True(claim.CanClaimConformance);
        Assert.True(report.CanClaimAllDeclaredConformance);
        Assert.True(claim.Proof!.HasRequiredExternalValidation);
    }

    [Fact]
    public void DeclaredClaims_ReportUnsupportedPdfAOneWithoutInventingAProfile() {
        string path = Path.Combine(
            AppContext.BaseDirectory,
            "Pdf",
            "Fixtures",
            "Interoperability",
            "openpreserve-pdfa1b-text.pdf");

        PdfDeclaredComplianceClaim claim = Assert.Single(PdfDocument.Open(path).AssessDeclaredComplianceClaims().Claims);

        Assert.Equal("PDF/A-1b", claim.Declaration);
        Assert.Null(claim.Profile);
        Assert.False(claim.IsRecognized);
        Assert.Equal(PdfDeclaredComplianceClaimStatus.UnsupportedProfile, claim.Status);
        Assert.Contains("not mapped", claim.Diagnostic, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeclaredClaims_ReturnEmptyReportForOrdinaryPdf() {
        byte[] bytes = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Ordinary PDF.")).ToBytes();

        PdfDeclaredComplianceClaimsReport report = PdfDocument.Open(bytes).AssessDeclaredComplianceClaims();

        Assert.False(report.HasClaims);
        Assert.Empty(report.Claims);
        Assert.False(report.CanClaimAllDeclaredConformance);
    }
}
