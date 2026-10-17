using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Invoicing.Pdf.Tests;

public sealed class PdfInvoiceContainerTests {
    [Theory]
    [InlineData(PdfComplianceProfile.PdfA2A)]
    [InlineData(PdfComplianceProfile.PdfA2B)]
    [InlineData(PdfComplianceProfile.PdfA2U)]
    [InlineData(PdfComplianceProfile.PdfA4)]
    [InlineData(PdfComplianceProfile.PdfA4E)]
    [InlineData(PdfComplianceProfile.PdfA4F)]
    [InlineData(PdfComplianceProfile.PdfUa2)]
    [InlineData(PdfComplianceProfile.PdfX1A2003)]
    [InlineData(PdfComplianceProfile.PdfX4)]
    public void UnsupportedProfileOnlyRequestCannotBeReplacedByInvoiceGroundwork(PdfComplianceProfile profile) {
        PdfInvoiceDocument snapshot = Snapshot();
        foreach (bool useHelper in new[] { false, true }) {
            PdfOptions options = useHelper ? new PdfOptions().RequireCompliance(profile)
                : new PdfOptions { ComplianceProfile = profile };
            Assert.Null(options.PdfAIdentification);
            Assert.Null(options.PdfUaIdentification);
            NotSupportedException error = Assert.Throws<NotSupportedException>(() => snapshot.ToPdfBytes(options));
            foreach (string supported in new[] { "PDF/A-3a", "PDF/A-3b", "PDF/A-3u", "PDF/UA-1" })
                Assert.Contains(supported, error.Message);
            Assert.Contains(profile.ToString(), error.Message);
            Assert.Contains("ToPresentationPdfBytes", error.Message);
            Assert.Equal(profile, options.ComplianceProfile);
            Assert.Null(options.PdfAIdentification);
            Assert.Null(options.ElectronicInvoiceMetadata);
        }
    }

    [Theory]
    [InlineData(PdfComplianceProfile.PdfA3A, "A", null)]
    [InlineData(PdfComplianceProfile.PdfA3B, "B", null)]
    [InlineData(PdfComplianceProfile.PdfA3U, "U", null)]
    [InlineData(PdfComplianceProfile.PdfUa1, "B", 1)]
    public void SupportedProfileOnlyRequestPreservesItsConformance(
        PdfComplianceProfile profile, string conformance, int? uaPart) {
        PdfOptions options = Options().RequireCompliance(profile);
        Assert.Null(options.PdfAIdentification);
        Assert.Null(options.PdfUaIdentification);
        byte[] pdf = Snapshot().ToPdfBytes(options);
        PdfDocument saved = PdfDocument.Load(pdf);
        Assert.Equal(3, saved.Inspect().XmpMetadata!.PdfAPart);
        Assert.Equal(conformance, saved.Inspect().XmpMetadata!.PdfAConformance);
        Assert.Equal(uaPart, saved.Inspect().XmpMetadata!.PdfUaPart);
        Assert.Empty(saved.AssessCompliance(profile).MissingRequirements);
        Assert.True(saved.AssessComplianceProof(PdfComplianceProfile.FacturX).IsInternallyReady);
        Assert.Equal(profile, options.ComplianceProfile);
        Assert.Null(options.PdfAIdentification);
        Assert.Null(options.PdfUaIdentification);
    }

    [Fact]
    public void HybridUa2IsRejectedBeforeRenderingAndPresentationUa2RemainsAvailable() {
        PdfInvoiceDocument snapshot = Snapshot();
        PdfOptions options = Options().UsePdfUa(PdfComplianceProfile.PdfUa2);
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => snapshot.ToPdfBytes(options));
        foreach (string supported in new[] { "PDF/A-3a", "PDF/A-3b", "PDF/A-3u", "PDF/UA-1" })
            Assert.Contains(supported, error.Message);
        Assert.Contains("PDF/UA-2", error.Message);
        Assert.Contains("ToPresentationPdfBytes", error.Message);
        Assert.Null(options.ElectronicInvoiceMetadata);

        byte[] presentation = snapshot.ToPresentationPdfBytes(options);
        PdfDocument saved = PdfDocument.Load(presentation);
        Assert.Equal(2, saved.Inspect().XmpMetadata!.PdfUaPart);
        Assert.Empty(saved.Attachments.Extract());
        Assert.Empty(saved.AssessCompliance(PdfComplianceProfile.PdfUa2).MissingRequirements);
    }

    [Theory]
    [InlineData(PdfComplianceProfile.PdfA2B, 2, "B")]
    [InlineData(PdfComplianceProfile.PdfA4, 4, null)]
    [InlineData(PdfComplianceProfile.PdfA4E, 4, "E")]
    [InlineData(PdfComplianceProfile.PdfA4F, 4, "F")]
    public void UnsupportedHybridContainerIsRejectedWithoutChangingPresentationOptions(
        PdfComplianceProfile archival, int part, string? conformance) {
        PdfInvoiceDocument snapshot = Snapshot();
        PdfOptions options = Options().UsePdfA(archival);
        if (archival == PdfComplianceProfile.PdfA4F) {
            options.AddEmbeddedFile("evidence.txt", new byte[] { 65 }, "text/plain", PdfAssociatedFileRelationship.Data)
                .SetEmbeddedFileModificationDate("evidence.txt", DateTimeOffset.UtcNow);
        }

        NotSupportedException error = Assert.Throws<NotSupportedException>(() => snapshot.ToPdfBytes(options));
        foreach (string supported in new[] { "PDF/A-3a", "PDF/A-3b", "PDF/A-3u", "PDF/UA-1" })
            Assert.Contains(supported, error.Message);
        Assert.Equal(part, options.PdfAIdentification!.Part);
        Assert.Null(options.ElectronicInvoiceMetadata);

        byte[] presentation = snapshot.ToPresentationPdfBytes(options);
        PdfDocument saved = PdfDocument.Load(presentation);
        Assert.Equal(part, saved.Inspect().XmpMetadata!.PdfAPart);
        Assert.Equal(conformance, saved.Inspect().XmpMetadata!.PdfAConformance);
        Assert.DoesNotContain(saved.Attachments.Extract(), file => file.FileName == "factur-x.xml");
        Assert.Empty(saved.AssessCompliance(archival).MissingRequirements);
    }

    [Fact]
    public void SavedA4fInvoiceCannotClaimFacturXOrZugferdEvenWithPassingExternalEvidence() {
        // Metadata groundwork can represent an existing unsupported invoice container.
        // It must remain assessable without promoting general A-4f attachment support
        // to a valid Factur-X/ZUGFeRD invoice carrier.
        PdfOptions options = Options().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA4F)
            .AddFacturXInvoiceXml(Snapshot().ToXmlBytes(), relationship: PdfAssociatedFileRelationship.Alternative)
            .SetEmbeddedFileModificationDate("factur-x.xml", DateTimeOffset.UtcNow);
        PdfDocument document = PdfDocument.Create(options);
        document.Content.Paragraph(p => p.Text("Invoice container regression"));
        foreach (PdfComplianceProfile profile in new[] { PdfComplianceProfile.FacturX, PdfComplianceProfile.Zugferd }) {
            PdfComplianceReadinessReport generated = document.AssessCompliance(profile);
            Assert.Contains(generated.MissingRequirements, requirement => requirement.Id == "pdfa-identification");
        }
        byte[] pdf = document.ToBytes();
        PdfDocument saved = PdfDocument.Load(pdf);
        Assert.Empty(saved.AssessCompliance(PdfComplianceProfile.PdfA4F).MissingRequirements);
        foreach (PdfComplianceProfile profile in new[] { PdfComplianceProfile.FacturX, PdfComplianceProfile.Zugferd }) {
            PdfExternalValidationResult Result(PdfExternalValidatorKind kind, string name) =>
                PdfExternalValidationResult.FromExitCodeForArtifact(kind, 0, "Synthetic test evidence", "fixture",
                    "Invoice carrier policy regression", pdf, name);
            PdfComplianceProofReport proof = saved.AssessComplianceProof(profile, new[] {
                Result(PdfExternalValidatorKind.VeraPdf, "PDF/A-4f"),
                Result(PdfExternalValidatorKind.Mustang, profile == PdfComplianceProfile.FacturX ? "Factur-X" : "ZUGFeRD")
            });
            Assert.False(proof.IsInternallyReady);
            Assert.False(proof.HasRequiredExternalValidation);
            Assert.False(proof.CanClaimConformance);
            Assert.Contains(proof.Readiness.MissingRequirements, requirement => requirement.Id == "readback-pdfa-identification");
        }

        options.RequireCompliance(PdfComplianceProfile.PdfA4F);
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => {
            PdfDocument strictDocument = PdfDocument.Create(options);
            strictDocument.Content.Paragraph(p => p.Text("Strict invoice carrier regression"));
            strictDocument.ToBytes();
        });
        Assert.Contains("PDF/A-3a", error.Message);
    }

    [Fact]
    public void SavedInvoiceUa2DeclarationBlocksInvoiceProof() {
        PdfOptions options = Options().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B)
            .ConfigurePdfUaGroundwork(PdfComplianceProfile.PdfUa2)
            .AddFacturXInvoiceXml(Snapshot().ToXmlBytes(), relationship: PdfAssociatedFileRelationship.Alternative)
            .SetEmbeddedFileModificationDate("factur-x.xml", DateTimeOffset.UtcNow);
        // Represent inconsistent declarations from an existing document without
        // letting its PDF 2.0 version be the only reason invoice proof fails.
        options.FileVersion = PdfFileVersion.Pdf17;
        PdfDocument document = PdfDocument.Create(options);
        document.Content.Paragraph(p => p.Text("Invoice UA declaration regression"));
        byte[] pdf = document.ToBytes();
        PdfDocument saved = PdfDocument.Load(pdf);
        foreach (PdfComplianceProfile profile in new[] { PdfComplianceProfile.FacturX, PdfComplianceProfile.Zugferd }) {
            Assert.Contains(document.AssessCompliance(profile).MissingRequirements,
                requirement => requirement.Id == "einvoice-pdfua-profile");
            PdfComplianceProofReport proof = saved.AssessComplianceProof(profile);
            Assert.Contains(proof.Readiness.MissingRequirements,
                requirement => requirement.Id == "readback-einvoice-pdfua-profile");
            Assert.False(proof.IsInternallyReady);
            Assert.False(proof.CanClaimConformance);
        }
    }

    private static PdfInvoiceDocument Snapshot() => PdfInvoiceDocument.Create(InvoiceFixture.Create(),
        new InvoiceXmlOptions(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii, InvoiceProfile.En16931));

    private static PdfOptions Options() {
        byte[] font = File.ReadAllBytes(PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!);
        return new PdfOptions { IncludeStandardFontToUnicodeMaps = true }
            .EmbedStandardFont(PdfStandardFont.Helvetica, font, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "OfficeIMO Source Serif");
    }
}
