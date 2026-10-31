using System.Text;
using System.Xml.Linq;
using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Invoicing.Pdf.Tests;

/// <summary>Exercises the real invoice renderer, rather than a hand-built invoice-shaped PDF.</summary>
public sealed class PdfInvoiceExternalProofTests {
    [Theory]
    [InlineData("invoice-minimum-pdfa3b", InvoiceProfile.Minimum, PdfComplianceProfile.PdfA3B, false, "B")]
    [InlineData("invoice-basic-wl-pdfa3b", InvoiceProfile.BasicWithoutLines, PdfComplianceProfile.PdfA3B, false, "B")]
    [InlineData("invoice-basic-pdfa3b", InvoiceProfile.Basic, PdfComplianceProfile.PdfA3B, false, "B")]
    [InlineData("invoice-en16931-pdfa3a", InvoiceProfile.En16931, PdfComplianceProfile.PdfA3A, false, "A")]
    [InlineData("invoice-extended-pdfa3u", InvoiceProfile.Extended, PdfComplianceProfile.PdfA3U, false, "U")]
    [InlineData("invoice-pdfa3b-ua1", InvoiceProfile.En16931, PdfComplianceProfile.PdfA3B, true, "B")]
    public void HybridInvoicePreservesRequestedProfilesAndExactXml(string name, InvoiceProfile profile, PdfComplianceProfile archival,
        bool accessible, string conformance) {
        Invoice invoice = InvoiceFixture.Create();
        var contract = new InvoiceXmlOptions(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii, profile,
            profile is InvoiceProfile.En16931 or InvoiceProfile.Extended
                ? InvoiceProjectionPolicy.RejectDataLoss : InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, contract);
        byte[] xml = snapshot.ToXmlBytes();
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        byte[] font = File.ReadAllBytes(fontPath!);
        var options = new PdfCreateOptions { IncludeStandardFontToUnicodeMaps = true };
        options.UsePdfA(archival)
            .EmbedStandardFont(PdfStandardFont.Helvetica, font, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "OfficeIMO Source Serif");
        if (accessible) options.UsePdfUa();

        byte[] pdf = snapshot.ToPdfBytes(options);
        PdfDocumentInfo info = PdfDocument.Load(new System.IO.MemoryStream(pdf)).Inspect();
        PdfXmpMetadataInfo xmp = Assert.IsType<PdfXmpMetadataInfo>(info.XmpMetadata);
        Assert.Equal("1.7", info.HeaderVersion);
        Assert.Equal(3, xmp.PdfAPart);
        Assert.Equal(conformance, xmp.PdfAConformance);
        Assert.Equal(accessible ? (int?)1 : null, xmp.PdfUaPart);
        PdfDocument saved = PdfDocument.Load(new System.IO.MemoryStream(pdf));
        Assert.Equal(xml, Assert.Single(saved.Attachments.Extract()).Bytes);
        PdfComplianceReadinessReport readiness = saved.AssessCompliance(PdfComplianceProfile.FacturX);
        Assert.Empty(readiness.MissingRequirements);
        Assert.Empty(saved.AssessCompliance(archival).MissingRequirements);
        Assert.Equal(profile, InvoiceProfileDeclaration.Read(xml).Profile);
        Assert.Equal(InvoiceProfiles.GetXmpConformanceLevel(profile), xmp.ElectronicInvoiceConformanceLevel);
        foreach (string requirementId in new[] { "readback-pdf-file-version", "readback-pdfa-identification",
            "readback-einvoice-xmp", "readback-associated-invoice-file" }) {
            PdfComplianceRequirement requirement = Assert.IsType<PdfComplianceRequirement>(readiness.FindRequirement(requirementId));
            Assert.True(requirement.Status == PdfComplianceRequirementStatus.Satisfied, requirement.Diagnostic);
        }
        PdfComplianceProofReport proof = saved.AssessComplianceProof(PdfComplianceProfile.FacturX);
        Assert.True(proof.IsInternallyReady);
        Assert.False(proof.CanClaimConformance); // External proof is collected independently on these exact bytes.
        string text = PdfDocument.Load(new System.IO.MemoryStream(pdf)).ExtractUnwrappedText();
        Assert.Contains(invoice.Number, text);
        Assert.Equal(profile is not (InvoiceProfile.Minimum or InvoiceProfile.BasicWithoutLines),
            text.IndexOf("Consulting", StringComparison.Ordinal) >= 0);
        Assert.Contains("119.00 EUR", text);

        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        output = Path.Combine(output, "standards");
        Directory.CreateDirectory(output);
        File.WriteAllBytes(Path.Combine(output, name + ".pdf"), pdf);
        File.WriteAllBytes(Path.Combine(output, name + ".xml"), xml);
    }
    [Theory]
    [InlineData(InvoiceProfile.Minimum, true)]
    [InlineData(InvoiceProfile.Minimum, false)]
    [InlineData(InvoiceProfile.BasicWithoutLines, true)]
    [InlineData(InvoiceProfile.BasicWithoutLines, false)]
    [InlineData(InvoiceProfile.Basic, true)]
    [InlineData(InvoiceProfile.Basic, false)]
    public void StrictLowerProfileGateRejectsMissingGrandTotalOrInconsistentTaxBasis(
        InvoiceProfile profile, bool missingGrandTotal) {
        var contract = new InvoiceXmlOptions(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii, profile, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(
            PdfInvoiceDocument.Create(InvoiceFixture.Create(), contract).ToXmlBytes()));
        XNamespace ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";
        XElement totals = Assert.Single(xml.Descendants(ram + "SpecifiedTradeSettlementHeaderMonetarySummation"));
        string field = missingGrandTotal ? "GrandTotalAmount" : "TaxBasisTotalAmount";
        XElement amount = Assert.Single(totals.Elements(ram + field));
        if (missingGrandTotal) amount.Remove();
        else amount.Value = "101.00";

        byte[] font = File.ReadAllBytes(PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!);
        var options = new PdfCreateOptions().UseFacturX(Encoding.UTF8.GetBytes(xml.ToString()),
                textFallbacks: PdfTextFallbackFeatures.None)
            .UsePdfA(PdfComplianceProfile.PdfA3B)
            .SetEmbeddedFileModificationDate("factur-x.xml", DateTimeOffset.UtcNow)
            .EmbedStandardFont(PdfStandardFont.Helvetica, font, "OfficeIMO Source Serif");
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => {
            PdfDocument document = PdfDocument.Create(options);
            document.Content.Paragraph(p => p.Text("Reduced profile strict gate regression"));
            document.ToBytes();
        });
        Assert.Contains("einvoice-xml-reduced-profile", error.Message);
        Assert.Contains(field, error.Message);
    }

    [Fact]
    public void StrictMinimumGateAcceptsOmittedTaxTotalAndIndependentAmountDue() {
        var contract = new InvoiceXmlOptions(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii, InvoiceProfile.Minimum, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(
            PdfInvoiceDocument.Create(InvoiceFixture.Create(), contract).ToXmlBytes()));
        XNamespace ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";
        XElement totals = Assert.Single(xml.Descendants(ram + "SpecifiedTradeSettlementHeaderMonetarySummation"));
        Assert.Single(totals.Elements(ram + "TaxTotalAmount")).Remove();
        Assert.Single(totals.Elements(ram + "DuePayableAmount")).Value = "99.00";
        Assert.Empty(totals.Elements(ram + "TaxTotalAmount"));
        Assert.NotEqual(Assert.Single(totals.Elements(ram + "GrandTotalAmount")).Value,
            Assert.Single(totals.Elements(ram + "DuePayableAmount")).Value);

        byte[] font = File.ReadAllBytes(PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!);
        var options = new PdfCreateOptions().UseFacturX(Encoding.UTF8.GetBytes(xml.ToString()),
                textFallbacks: PdfTextFallbackFeatures.None)
            .UsePdfA(PdfComplianceProfile.PdfA3B)
            .SetEmbeddedFileModificationDate("factur-x.xml", DateTimeOffset.UtcNow)
            .EmbedStandardFont(PdfStandardFont.Helvetica, font, "OfficeIMO Source Serif");
        PdfDocument document = PdfDocument.Create(options);
        document.Content.Paragraph(p => p.Text("MINIMUM optional totals regression"));
        byte[] pdf = document.ToBytes();
        PdfDocument saved = PdfDocument.Load(new System.IO.MemoryStream(pdf));
        Assert.Empty(saved.AssessCompliance(PdfComplianceProfile.FacturX).MissingRequirements);
        Assert.True(saved.AssessComplianceProof(PdfComplianceProfile.FacturX).IsInternallyReady);
    }

    [Theory]
    [InlineData(PdfComplianceProfile.PdfA3A, "PDF/A-3a", "PDF/A-3b")]
    [InlineData(PdfComplianceProfile.PdfA3B, "PDF/A-3b", "PDF/A-4f")]
    [InlineData(PdfComplianceProfile.PdfA3U, "PDF/A-3u", "PDF/A-3b")]
    public void InvoiceProofMatchesTheBackingArchivalProfileForPassedAndFailedResults(
        PdfComplianceProfile archival, string expectedProfile, string wrongProfile) {
        var contract = new InvoiceXmlOptions(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii, InvoiceProfile.En16931);
        var options = new PdfCreateOptions().UsePdfA(archival);
        byte[] font = File.ReadAllBytes(PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!);
        options.EmbedStandardFont(PdfStandardFont.Helvetica, font, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "OfficeIMO Source Serif");
        byte[] pdf = PdfInvoiceDocument.Create(InvoiceFixture.Create(), contract).ToPdfBytes(options);
        PdfDocument saved = PdfDocument.Load(new System.IO.MemoryStream(pdf));
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
