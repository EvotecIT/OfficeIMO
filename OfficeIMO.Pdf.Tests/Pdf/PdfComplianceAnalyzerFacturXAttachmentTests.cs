using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    [Fact]
    public void FacturXReadinessRejectsNonCanonicalXmlAttachmentName() {
        var options = new PdfOptions {
                IncludeStandardFontToUnicodeMaps = true
            }
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .AddEmbeddedFile("invoice.xml", CreateCiiXml(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, options);

        PdfComplianceRequirement requirement = AssertRequirement(report, "einvoice-xml-attachment", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "einvoice-xml-attachment-params", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("factur-x.xml", requirement.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRejectsSupplementRelationshipForInvoiceXml() {
        var options = new PdfOptions {
                IncludeStandardFontToUnicodeMaps = true
            }
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(), "application/xml", PdfAssociatedFileRelationship.Supplement);

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, options);

        PdfComplianceRequirement requirement = AssertRequirement(report, "einvoice-xml-attachment", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("AFRelationship", requirement.Diagnostic);
        Assert.Contains("Alternative, Data, or Source", requirement.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRejectsMalformedOrWrongRootXmlAttachment() {
        var malformedOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .AddEmbeddedFile("factur-x.xml", Encoding.UTF8.GetBytes("<rsm:CrossIndustryInvoice />"), "application/xml", PdfAssociatedFileRelationship.Data);
        var wrongRootOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .AddEmbeddedFile("factur-x.xml", Encoding.UTF8.GetBytes("<Invoice />"), "text/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceReadinessReport malformedReport = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, malformedOptions);
        PdfComplianceReadinessReport wrongRootReport = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, wrongRootOptions);

        Assert.Contains("parseable XML", AssertRequirement(malformedReport, "einvoice-xml-attachment", PdfComplianceRequirementStatus.Missing).Diagnostic);
        Assert.Contains("CrossIndustryInvoice", AssertRequirement(wrongRootReport, "einvoice-xml-attachment", PdfComplianceRequirementStatus.Missing).Diagnostic);
    }

    [Fact]
    public void FacturXReadbackAcceptsCanonicalInvoiceXmlAttachment() {
        var options = new PdfOptions()
            .ConfigureFacturXGroundwork(CreateCiiXml());
        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Factur-X invoice"))
            .ToBytes();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.FacturX, pdf);

        PdfComplianceRequirement requirement = AssertRequirement(report, "readback-associated-invoice-file", PdfComplianceRequirementStatus.Satisfied);
        Assert.Contains("CrossIndustryInvoice", requirement.Diagnostic);
    }

    [Fact]
    public void FacturXReadbackRejectsNonInvoiceDataAttachment() {
        var options = new PdfOptions {
                IncludeStandardFontToUnicodeMaps = true
            }
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("terms.pdf", Encoding.UTF8.GetBytes("%PDF-1.7\n"), "application/pdf", PdfAssociatedFileRelationship.Data);
        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Factur-X invoice shell"))
            .ToBytes();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.FacturX, pdf);

        PdfComplianceRequirement requirement = AssertRequirement(report, "readback-associated-invoice-file", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("factur-x.xml", requirement.Diagnostic);
    }

}
