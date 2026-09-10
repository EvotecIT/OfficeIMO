using OfficeIMO.Internal.Invoicing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public class PdfInvoiceProfileConsistencyTests {
    [Theory]
    [InlineData((int)InvoiceProfile.Basic, "BASIC")]
    [InlineData((int)InvoiceProfile.BasicWithoutLines, "BASIC WL")]
    [InlineData((int)InvoiceProfile.En16931, "EN 16931")]
    [InlineData((int)InvoiceProfile.Extended, "EXTENDED")]
    [InlineData((int)InvoiceProfile.ExtendedCtcFr, "EXTENDED-CTC-FR")]
    public void AttachmentHelperDerivesCanonicalXmpFromXml(int profile, string expected) {
        byte[] xml = Cii((InvoiceProfile)profile);
        var options = new PdfOptions().UseFacturX(xml, textFallbacks: PdfTextFallbackFeatures.None);
        Assert.Equal(expected, options.ElectronicInvoiceMetadata!.ConformanceLevel);
        byte[] pdf = PdfDocument.Create(options).Paragraph(p => p.Text("Invoice profile test")).ToBytes();
        Assert.Equal(xml, PdfAttachmentExtractor.ExtractAttachments(pdf).Single().Bytes);
        var report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.FacturX, pdf);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied,
            report.Requirements.Single(r => r.Id == "readback-einvoice-profile-consistency").Status);
    }

    [Fact]
    public void ExplicitMismatchIsRejectedBeforeMutatingOptions() {
        var options = new PdfOptions();
        Assert.Throws<ArgumentException>(() => options.UseFacturX(Cii(InvoiceProfile.En16931), "BASIC"));
        Assert.Empty(options.EmbeddedFiles);
        Assert.Null(options.ElectronicInvoiceMetadata);
        var invoice = PdfCiiInvoiceDocument.Load(Cii(InvoiceProfile.Basic));
        Assert.Throws<ArgumentException>(() => options.UseFacturXDocument(invoice, "EN 16931"));
        Assert.Equal("BASIC", options.UseFacturXDocument(invoice).ElectronicInvoiceMetadata!.ConformanceLevel);
    }

    [Fact]
    public void GenericSettersCannotBypassProfileAgreement() {
        var options = new PdfOptions().AddEmbeddedFile("factur-x.xml", Cii(InvoiceProfile.En16931), "application/xml", PdfAssociatedFileRelationship.Data)
            .SetElectronicInvoiceMetadata("BASIC");
        var report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, options);
        Assert.Equal(PdfComplianceRequirementStatus.Missing,
            report.Requirements.Single(r => r.Id == "einvoice-profile-consistency").Status);
        Assert.Throws<ArgumentException>(() => PdfDocument.Create(options).ToBytes());
    }

    [Fact]
    public void SavedArtifactReadbackDetectsMetadataTampering() {
        byte[] generated = PdfDocument.Create(new PdfOptions().UseFacturX(Cii(InvoiceProfile.En16931), textFallbacks: PdfTextFallbackFeatures.None))
            .Paragraph(p => p.Text("Invoice profile test")).ToBytes();
        string original = PdfEncoding.Latin1GetString(generated);
        // Keep stream and xref offsets unchanged; only change the invoice profile value.
        string tampered = original.Replace("<fx:ConformanceLevel>EN 16931</fx:ConformanceLevel>", "<fx:ConformanceLevel>BASIC   </fx:ConformanceLevel>");
        Assert.NotEqual(original, tampered);
        byte[] pdf = PdfEncoding.Latin1GetBytes(tampered);
        var report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.FacturX, pdf);
        var requirement = report.Requirements.Single(r => r.Id == "readback-einvoice-profile-consistency");
        Assert.Equal(PdfComplianceRequirementStatus.Missing, requirement.Status);
        Assert.Contains("does not match", requirement.Diagnostic);
    }

    [Fact]
    public void DuplicateCanonicalAttachmentsCannotPassReadiness() {
        var options = new PdfOptions().UseFacturX(Cii(InvoiceProfile.En16931), textFallbacks: PdfTextFallbackFeatures.None);
        Assert.Throws<ArgumentException>(() => options.AddEmbeddedFile("factur-x.xml", Cii(InvoiceProfile.Basic), "application/xml", PdfAssociatedFileRelationship.Data));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnrelatedAttachmentPayloadAndMimeDoNotChangeInvoiceProfileAgreement(bool emptyPayload) {
        var options = new PdfOptions().UseFacturX(Cii(InvoiceProfile.En16931), textFallbacks: PdfTextFallbackFeatures.None)
            .AddEmbeddedFile("support.txt", new byte[] { 88 }, "text/plain", PdfAssociatedFileRelationship.Supplement);
        byte[] generated = PdfDocument.Create(options).Paragraph(p => p.Text("Invoice")).ToBytes();
        string original = PdfEncoding.Latin1GetString(generated);
        string altered = emptyPayload
            ? original.Replace("/Length 1 /Params", "/Length 0 /Params")
            : original.Replace("/Subtype /text#2Fplain", "/Subtype /text#20plain");
        Assert.NotEqual(original, altered);
        byte[] pdf = PdfEncoding.Latin1GetBytes(altered);
        var auxiliary = PdfAttachmentExtractor.ExtractAttachments(pdf).Single(file => file.FileName == "support.txt");
        if (emptyPayload) Assert.Empty(auxiliary.Bytes);
        else Assert.Equal("text plain", auxiliary.MimeType);
        var report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.FacturX, pdf);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied,
            report.Requirements.Single(r => r.Id == "readback-einvoice-profile-consistency").Status);
    }

    [Theory]
    [InlineData("<fx:ConformanceLevel><x/>EN 16931</fx:ConformanceLevel>")]
    [InlineData("<fx:ConformanceLevel> EN 16931 </fx:ConformanceLevel>")]
    [InlineData("<fx:ConformanceLevel>EN 16931</fx:ConformanceLevel><fx:ConformanceLevel/>")]
    public void SavedArtifactRejectsNormalizedOrStructuredXmp(string property) {
        const string padding = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX";
        const string canonical = "<fx:ConformanceLevel>EN 16931</fx:ConformanceLevel>";
        byte[] generated = PdfDocument.Create(new PdfOptions().UseFacturX(Cii(InvoiceProfile.En16931), textFallbacks: PdfTextFallbackFeatures.None))
            .Meta(title: padding).Paragraph(p => p.Text("Invoice")).ToBytes();
        string original = PdfEncoding.Latin1GetString(generated);
        string tampered = original.Replace(canonical, property)
            .Replace(">" + padding + "</rdf:li>", ">" + padding.Substring(property.Length - canonical.Length) + "</rdf:li>");
        Assert.NotEqual(original, tampered);
        Assert.Equal(original.Length, tampered.Length);
        var report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.FacturX, PdfEncoding.Latin1GetBytes(tampered));
        Assert.Equal(PdfComplianceRequirementStatus.Missing,
            report.Requirements.Single(r => r.Id == "readback-einvoice-profile-consistency").Status);
    }

    [Fact]
    public void MetadataChangedAfterDocumentCreationIsCheckedBeforeSerialization() {
        var document = PdfDocument.Create(new PdfOptions().UseFacturX(Cii(InvoiceProfile.En16931), textFallbacks: PdfTextFallbackFeatures.None))
            .ElectronicInvoiceMetadata("BASIC").Paragraph(p => p.Text("Invoice"));
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    private static byte[] Cii(InvoiceProfile profile) => Encoding.UTF8.GetBytes(
        "<r:CrossIndustryInvoice xmlns:r='urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100' xmlns:a='urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100'>" +
        "<r:ExchangedDocumentContext><a:GuidelineSpecifiedDocumentContextParameter><a:ID>" + InvoiceProfiles.GetGuidelineId(profile) +
        "</a:ID></a:GuidelineSpecifiedDocumentContextParameter></r:ExchangedDocumentContext></r:CrossIndustryInvoice>");
}
