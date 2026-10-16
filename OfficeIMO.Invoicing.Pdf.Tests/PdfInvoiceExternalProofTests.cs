using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Invoicing.Pdf.Tests;

/// <summary>Exercises the real invoice renderer, rather than a hand-built invoice-shaped PDF.</summary>
public sealed class PdfInvoiceExternalProofTests {
    [Theory]
    [InlineData("invoice-pdfa4f", PdfComplianceProfile.PdfA4F, false, 4, "F")]
    [InlineData("invoice-pdfa3b-ua1", PdfComplianceProfile.PdfA3B, true, 3, "B")]
    public void HybridInvoicePreservesRequestedProfilesAndExactXml(string name, PdfComplianceProfile archival,
        bool accessible, int part, string conformance) {
        Invoice invoice = InvoiceFixture.Create();
        var contract = new InvoiceXmlOptions(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii, InvoiceProfile.En16931);
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, contract);
        byte[] xml = snapshot.ToXmlBytes();
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        byte[] font = File.ReadAllBytes(fontPath!);
        var options = new PdfOptions { IncludeStandardFontToUnicodeMaps = true };
        options.UsePdfA(archival)
            .EmbedStandardFont(PdfStandardFont.Helvetica, font, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "OfficeIMO Source Serif");
        if (accessible) options.UsePdfUa();

        byte[] pdf = snapshot.ToPdfBytes(options);
        PdfDocumentInfo info = PdfDocument.Load(pdf).Inspect();
        PdfXmpMetadataInfo xmp = Assert.IsType<PdfXmpMetadataInfo>(info.XmpMetadata);
        Assert.Equal(part == 4 ? "2.0" : "1.7", info.HeaderVersion);
        Assert.Equal(part, xmp.PdfAPart);
        Assert.Equal(conformance, xmp.PdfAConformance);
        Assert.Equal(accessible ? (int?)1 : null, xmp.PdfUaPart);
        PdfDocument saved = PdfDocument.Load(pdf);
        Assert.Equal(xml, Assert.Single(saved.Attachments.Extract()).Bytes);
        PdfComplianceReadinessReport readiness = saved.AssessCompliance(PdfComplianceProfile.FacturX);
        Assert.Empty(readiness.MissingRequirements);
        foreach (string requirementId in new[] { "readback-pdf-file-version", "readback-pdfa-identification",
            "readback-einvoice-xmp", "readback-associated-invoice-file" }) {
            PdfComplianceRequirement requirement = Assert.IsType<PdfComplianceRequirement>(readiness.FindRequirement(requirementId));
            Assert.True(requirement.Status == PdfComplianceRequirementStatus.Satisfied, requirement.Diagnostic);
        }
        PdfComplianceProofReport proof = saved.AssessComplianceProof(PdfComplianceProfile.FacturX);
        Assert.True(proof.IsInternallyReady);
        Assert.False(proof.CanClaimConformance); // External proof is collected independently on these exact bytes.
        string text = PdfReadDocument.Open(pdf).ExtractUnwrappedText();
        Assert.Contains(invoice.Number, text);
        Assert.Contains("Consulting", text);
        Assert.Contains("119.00 EUR", text);

        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        output = Path.Combine(output, "standards");
        Directory.CreateDirectory(output);
        File.WriteAllBytes(Path.Combine(output, name + ".pdf"), pdf);
        File.WriteAllBytes(Path.Combine(output, name + ".xml"), xml);
    }
    [Theory]
    [InlineData(PdfComplianceProfile.PdfA4F, "PDF/A-4f", "PDF/A-3b")]
    [InlineData(PdfComplianceProfile.PdfA3B, "PDF/A-3b", "PDF/A-4f")]
    public void InvoiceProofMatchesTheBackingArchivalProfileForPassedAndFailedResults(
        PdfComplianceProfile archival, string expectedProfile, string wrongProfile) {
        var contract = new InvoiceXmlOptions(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii, InvoiceProfile.En16931);
        var options = new PdfOptions().UsePdfA(archival);
        byte[] font = File.ReadAllBytes(PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!);
        options.EmbedStandardFont(PdfStandardFont.Helvetica, font, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "OfficeIMO Source Serif");
        byte[] pdf = PdfInvoiceDocument.Create(InvoiceFixture.Create(), contract).ToPdfBytes(options);
        PdfDocument saved = PdfDocument.Load(pdf);
        PdfExternalValidationResult Result(PdfExternalValidatorKind kind, string profile, bool success) =>
            PdfExternalValidationResult.FromExitCodeForArtifact(kind, success ? 0 : 1,
                "Test evidence", "fixture", "Profile matching regression", pdf, profile);
        PdfExternalValidationResult mustang = Result(PdfExternalValidatorKind.Mustang, "Factur-X", true);
        foreach (PdfComplianceProfile invoiceProfile in new[] { PdfComplianceProfile.FacturX, PdfComplianceProfile.Zugferd }) {
            PdfExternalValidationResult invoiceValidator = invoiceProfile == PdfComplianceProfile.FacturX
                ? mustang : Result(PdfExternalValidatorKind.Mustang, "ZUGFeRD", true);
            var wrongPass = saved.AssessComplianceProof(invoiceProfile, new[] {
                invoiceValidator, Result(PdfExternalValidatorKind.VeraPdf, wrongProfile, true) });
            Assert.False(wrongPass.HasRequiredExternalValidation);
            var rightPassAndWrongFailure = saved.AssessComplianceProof(invoiceProfile, new[] {
                invoiceValidator, Result(PdfExternalValidatorKind.VeraPdf, expectedProfile, true),
                Result(PdfExternalValidatorKind.VeraPdf, wrongProfile, false) });
            Assert.True(rightPassAndWrongFailure.CanClaimConformance);
            var matchingFailure = saved.AssessComplianceProof(invoiceProfile, new[] {
                invoiceValidator, Result(PdfExternalValidatorKind.VeraPdf, expectedProfile, true),
                Result(PdfExternalValidatorKind.VeraPdf, expectedProfile, false) });
            Assert.False(matchingFailure.CanClaimConformance);
        }
    }

}
