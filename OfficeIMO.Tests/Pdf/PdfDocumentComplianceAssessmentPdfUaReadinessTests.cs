using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentComplianceAssessmentTests {

    [Fact]
    public void AssessComplianceUsesDocumentMetadataForPdfUaTitleReadiness() {
        PdfDocument document = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .Meta(title: "Accessible title")
            .Paragraph(paragraph => paragraph.Text("PDF/UA title readiness."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfUa1);

        AssertRequirement(report, "pdfua-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "document-title", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "display-document-title", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "document-language", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "tagged-catalog-markers", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "tagged-parent-tree-next-key", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-document-structure-root", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-document-structure-language", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "tagged-structure", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-text-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-list-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-list-structure-containers", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-cell-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-structure-containers", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-header-scope-attributes", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-span-attributes", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-caption-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-link-annotation-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-link-text-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-form-widget-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-form-field-accessible-names", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-image-alternate-text", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-image-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-drawing-alternate-text", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-drawing-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "alternate-text", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "decorative-drawing-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "decorative-running-page-text-artifacts", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "decorative-flow-rule-artifacts", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "decorative-layout-artifacts", PdfComplianceRequirementStatus.Missing);
    }

    [Fact]
    public void AssessComplianceRecognizesTaggedCatalogMarkersForPdfUaGroundwork() {
        PdfDocument document = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .TaggedPdfCatalogMarkers()
            .Meta(title: "Accessible title")
            .Paragraph(paragraph => paragraph.Text("PDF/UA marker readiness."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfUa1);

        AssertRequirement(report, "tagged-catalog-markers", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "tagged-parent-tree-next-key", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-document-structure-root", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-document-structure-language", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "tagged-structure", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-text-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-list-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-list-structure-containers", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-cell-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-structure-containers", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-header-scope-attributes", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-span-attributes", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-caption-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-link-annotation-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-link-text-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-form-widget-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-form-field-accessible-names", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-image-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-drawing-alternate-text", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-drawing-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "decorative-drawing-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "decorative-running-page-text-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "decorative-flow-rule-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "decorative-layout-artifacts", PdfComplianceRequirementStatus.Satisfied);
        Assert.False(report.IsReady);
    }

    [Fact]
    public void AssessComplianceReportsMissingPdfUaTitleFromDocumentMetadata() {
        PdfDocument document = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .Paragraph(paragraph => paragraph.Text("PDF/UA title readiness without a title."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfUa1);

        PdfComplianceRequirement requirement = AssertRequirement(report, "document-title", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("Meta(title", requirement.Diagnostic, StringComparison.Ordinal);
    }

    [Fact]
    public void AssessComplianceReportsGeneratedImageAlternativeTextReadiness() {
        byte[] png = CreateMinimalRgbPng();
        PdfDocument missingDocument = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .Meta(title: "Accessible title")
            .Image(png, 24, 24);
        PdfDocument satisfiedDocument = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .Meta(title: "Accessible title")
            .Image(png, 24, 24, alternativeText: "Product approval badge");

        PdfComplianceReadinessReport missingReport = missingDocument.AssessCompliance(PdfComplianceProfile.PdfUa1);
        PdfComplianceReadinessReport satisfiedReport = satisfiedDocument.AssessCompliance(PdfComplianceProfile.PdfUa1);

        PdfComplianceRequirement missing = AssertRequirement(missingReport, "generated-image-alternate-text", PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement satisfied = AssertRequirement(satisfiedReport, "generated-image-alternate-text", PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement aggregateMissing = AssertRequirement(missingReport, "alternate-text", PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement aggregateSatisfied = AssertRequirement(satisfiedReport, "alternate-text", PdfComplianceRequirementStatus.Satisfied);
        Assert.Contains("alternativeText", missing.Diagnostic, StringComparison.Ordinal);
        Assert.Contains("Every non-decorative", satisfied.Diagnostic, StringComparison.Ordinal);
        Assert.Contains("generated image", aggregateMissing.Diagnostic, StringComparison.Ordinal);
        Assert.Contains("non-decorative image", aggregateSatisfied.Diagnostic, StringComparison.Ordinal);
        AssertRequirement(satisfiedReport, "generated-text-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "generated-list-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "generated-list-structure-containers", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "generated-table-cell-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "generated-table-structure-containers", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "generated-table-header-scope-attributes", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "generated-table-span-attributes", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "generated-table-caption-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "generated-link-annotation-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "generated-link-text-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "generated-form-widget-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-form-field-accessible-names", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-image-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "generated-drawing-alternate-text", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-drawing-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "decorative-drawing-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "decorative-running-page-text-artifacts", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "decorative-flow-rule-artifacts", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "decorative-layout-artifacts", PdfComplianceRequirementStatus.Missing);
        Assert.False(satisfiedReport.IsReady);
    }

    [Fact]
    public void AssessComplianceUsesDefaultImageStyleAlternativeTextReadiness() {
        byte[] png = CreateMinimalRgbPng();
        PdfDocument document = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .Meta(title: "Accessible title")
            .DefaultImageStyle(new PdfImageStyle {
                AlternativeText = "Default product badge"
            })
            .Image(png, 24, 24);

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfUa1);

        AssertRequirement(report, "generated-image-alternate-text", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "alternate-text", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessComplianceTreatsUncaptionedHeaderFooterImagesAsDecorativeArtifacts() {
        byte[] png = CreateMinimalRgbPng();
        PdfDocument decorativeDocument = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .TaggedPdfCatalogMarkers()
            .Meta(title: "Accessible title")
            .Header(header => header.Image(png, 20, 10))
            .Paragraph(paragraph => paragraph.Text("Header image without alt text."));
        PdfDocument satisfiedDocument = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .TaggedPdfCatalogMarkers()
            .Meta(title: "Accessible title")
            .Header(header => header.Image(png, 20, 10, alternativeText: "Company logo"))
            .Footer(footer => footer.Image(png, 20, 10, alternativeText: "Security footer mark"))
            .Paragraph(paragraph => paragraph.Text("Header and footer image alt text."));

        PdfComplianceReadinessReport decorativeReport = decorativeDocument.AssessCompliance(PdfComplianceProfile.PdfUa1);
        PdfComplianceReadinessReport satisfiedReport = satisfiedDocument.AssessCompliance(PdfComplianceProfile.PdfUa1);

        AssertRequirement(decorativeReport, "generated-image-alternate-text", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(decorativeReport, "decorative-image-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-text-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-list-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-list-structure-containers", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-table-cell-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-table-structure-containers", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-table-header-scope-attributes", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-table-span-attributes", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-table-caption-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-link-annotation-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-link-text-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-form-widget-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-form-field-accessible-names", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-image-alternate-text", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-image-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "decorative-image-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-drawing-alternate-text", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-drawing-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "decorative-drawing-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "decorative-running-page-text-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "decorative-flow-rule-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "decorative-layout-artifacts", PdfComplianceRequirementStatus.Satisfied);
        Assert.False(satisfiedReport.IsReady);
    }

    [Fact]
    public void AssessComplianceTreatsBackgroundAndWatermarkImagesAsDecorativeArtifacts() {
        byte[] png = CreateMinimalRgbPng();
        PdfDocument document = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .TaggedPdfCatalogMarkers()
            .Meta(title: "Accessible title")
            .BackgroundImage(png, OfficeIMO.Drawing.OfficeImageFit.Stretch, opacity: 0.2)
            .ImageWatermark(png, 80, 40, opacity: 0.2)
            .Paragraph(paragraph => paragraph.Text("Decorative image readiness."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfUa1);

        AssertRequirement(report, "generated-text-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-list-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-list-structure-containers", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-cell-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-structure-containers", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-header-scope-attributes", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-span-attributes", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-caption-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-link-annotation-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-link-text-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-form-widget-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-form-field-accessible-names", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-image-alternate-text", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-image-structure-references", PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement artifacts = AssertRequirement(report, "decorative-image-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-drawing-alternate-text", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-drawing-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "decorative-drawing-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "decorative-running-page-text-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "decorative-flow-rule-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "decorative-layout-artifacts", PdfComplianceRequirementStatus.Satisfied);
        Assert.Contains("artifact", artifacts.Diagnostic, StringComparison.OrdinalIgnoreCase);
        AssertRequirement(report, "tagged-structure", PdfComplianceRequirementStatus.Unsupported);
    }

    [Fact]
    public void AssessComplianceReportsGeneratedDrawingAlternativeTextReadiness() {
        OfficeShape shape = CreateComplianceShape();
        PdfDocument missingDocument = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .Meta(title: "Accessible title")
            .Shape(shape);
        PdfDocument satisfiedDocument = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .TaggedPdfCatalogMarkers()
            .Meta(title: "Accessible title")
            .Shape(shape, style: new PdfDrawingStyle {
                AlternativeText = "Risk status badge"
            });

        PdfComplianceReadinessReport missingReport = missingDocument.AssessCompliance(PdfComplianceProfile.PdfUa1);
        PdfComplianceReadinessReport satisfiedReport = satisfiedDocument.AssessCompliance(PdfComplianceProfile.PdfUa1);

        PdfComplianceRequirement missing = AssertRequirement(missingReport, "generated-drawing-alternate-text", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("PdfDrawingStyle.AlternativeText", missing.Diagnostic, StringComparison.Ordinal);
        AssertRequirement(missingReport, "generated-drawing-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "generated-drawing-alternate-text", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "generated-drawing-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(satisfiedReport, "decorative-drawing-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(missingReport, "alternate-text", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(satisfiedReport, "alternate-text", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessComplianceTreatsDecorativeDrawingsAsArtifacts() {
        PdfDocument document = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .TaggedPdfCatalogMarkers()
            .Meta(title: "Accessible title")
            .Shape(CreateComplianceShape(), style: new PdfDrawingStyle {
                Decorative = true
            });

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfUa1);

        AssertRequirement(report, "generated-drawing-alternate-text", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-drawing-structure-references", PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement artifact = AssertRequirement(report, "decorative-drawing-artifacts", PdfComplianceRequirementStatus.Satisfied);
        Assert.Contains("artifact", artifact.Diagnostic, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AssessComplianceReportsGeneratedFormWidgetStructureReadiness() {
        PdfDocument missingDocument = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .Meta(title: "Accessible title")
            .TextField("Contact.Email", value: "info@example.com");
        PdfDocument satisfiedDocument = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .TaggedPdfCatalogMarkers()
            .Meta(title: "Accessible title")
            .TextField("Contact.Email", value: "info@example.com");

        PdfComplianceReadinessReport missingReport = missingDocument.AssessCompliance(PdfComplianceProfile.PdfUa1);
        PdfComplianceReadinessReport satisfiedReport = satisfiedDocument.AssessCompliance(PdfComplianceProfile.PdfUa1);

        PdfComplianceRequirement missing = AssertRequirement(missingReport, "generated-form-widget-structure-references", PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement satisfied = AssertRequirement(satisfiedReport, "generated-form-widget-structure-references", PdfComplianceRequirementStatus.Satisfied);
        Assert.Contains("AcroForm widgets", missing.Diagnostic, StringComparison.Ordinal);
        Assert.Contains("/Form OBJR", satisfied.Diagnostic, StringComparison.Ordinal);
    }

    [Fact]
    public void AssessComplianceReportsGeneratedFormAccessibleNameReadiness() {
        PdfDocument missingDocument = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .TaggedPdfCatalogMarkers()
            .Meta(title: "Accessible title")
            .TextField("Contact.Email", value: "info@example.com");
        PdfDocument satisfiedDocument = PdfDocument.Create(CreatePdfUaGroundworkOptions())
            .TaggedPdfCatalogMarkers()
            .Meta(title: "Accessible title")
            .TextField("Contact.Email", value: "info@example.com", style: new PdfFormFieldStyle {
                AlternateName = "Contact email address"
            });

        PdfComplianceReadinessReport missingReport = missingDocument.AssessCompliance(PdfComplianceProfile.PdfUa1);
        PdfComplianceReadinessReport satisfiedReport = satisfiedDocument.AssessCompliance(PdfComplianceProfile.PdfUa1);

        PdfComplianceRequirement missing = AssertRequirement(missingReport, "generated-form-field-accessible-names", PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement satisfied = AssertRequirement(satisfiedReport, "generated-form-field-accessible-names", PdfComplianceRequirementStatus.Satisfied);
        Assert.Contains("PdfFormFieldStyle.AlternateName", missing.Diagnostic, StringComparison.Ordinal);
        Assert.Contains("alternate field name", satisfied.Diagnostic, StringComparison.Ordinal);
    }
}
