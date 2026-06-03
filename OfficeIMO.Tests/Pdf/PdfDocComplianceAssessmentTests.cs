using System.Collections.Generic;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocComplianceAssessmentTests {
    [Fact]
    public void AssessComplianceUsesGeneratedStandardFontUsageToReportMissingEmbedding() {
        PdfDoc document = CreatePdfA3GroundworkDocument()
            .Paragraph(paragraph => paragraph.Text("Generated body text requires a standard-font resource."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Missing);
    }

    [Fact]
    public void AssessComplianceUsesGeneratedStandardFontUsageToAcceptEmbeddedMapping() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        PdfDoc document = CreatePdfA3GroundworkDocument()
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath), "HelveticaAudit")
            .Paragraph(paragraph => paragraph.Text("Generated body text has an embedded standard-font mapping."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "verapdf-validation", PdfComplianceRequirementStatus.Unsupported);
        Assert.False(report.IsReady);
    }

    [Fact]
    public void AssessComplianceUsesPageScopedEmbeddedFontFamilyCoverage() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        PdfDoc document = CreatePdfA3GroundworkDocument()
            .Page(page => page
                .UseFontFamily("Scoped Audit Font", fontPath)
                .Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Page-scoped embedded font coverage.")))));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessComplianceUsesConfiguredProfileWhenNoProfileIsSupplied() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = CreatePdfA3GroundworkOptions();
        options.ComplianceProfile = PdfComplianceProfile.PdfA3B;

        PdfDoc document = PdfDoc.Create(options)
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath), "HelveticaAudit")
            .Paragraph(paragraph => paragraph.Text("Configured profile readiness."));

        PdfComplianceReadinessReport report = document.AssessCompliance();

        Assert.Equal(PdfComplianceProfile.PdfA3B, report.Profile);
        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessComplianceReportsInvalidEmbeddedMappingBeforeGeneration() {
        PdfDoc document = CreatePdfA3GroundworkDocument()
            .EmbedStandardFont(PdfStandardFont.Helvetica, new byte[] { 1 }, "HelveticaAudit")
            .Paragraph(paragraph => paragraph.Text("Invalid embedded font mapping."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        PdfComplianceRequirement requirement = AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("invalid embedded TrueType", requirement.Diagnostic);
    }

    [Fact]
    public void AssessComplianceUsesDocumentMetadataForPdfUaTitleReadiness() {
        PdfDoc document = PdfDoc.Create(CreatePdfUaGroundworkOptions())
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
        PdfDoc document = PdfDoc.Create(CreatePdfUaGroundworkOptions())
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
        PdfDoc document = PdfDoc.Create(CreatePdfUaGroundworkOptions())
            .Paragraph(paragraph => paragraph.Text("PDF/UA title readiness without a title."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfUa1);

        PdfComplianceRequirement requirement = AssertRequirement(report, "document-title", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("Meta(title", requirement.Diagnostic, StringComparison.Ordinal);
    }

    [Fact]
    public void AssessComplianceReportsGeneratedImageAlternativeTextReadiness() {
        byte[] png = CreateMinimalRgbPng();
        PdfDoc missingDocument = PdfDoc.Create(CreatePdfUaGroundworkOptions())
            .Meta(title: "Accessible title")
            .Image(png, 24, 24);
        PdfDoc satisfiedDocument = PdfDoc.Create(CreatePdfUaGroundworkOptions())
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
        PdfDoc document = PdfDoc.Create(CreatePdfUaGroundworkOptions())
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
    public void AssessComplianceReportsHeaderFooterImageAlternativeTextReadiness() {
        byte[] png = CreateMinimalRgbPng();
        PdfDoc missingDocument = PdfDoc.Create(CreatePdfUaGroundworkOptions())
            .TaggedPdfCatalogMarkers()
            .Meta(title: "Accessible title")
            .Header(header => header.Image(png, 20, 10))
            .Paragraph(paragraph => paragraph.Text("Header image without alt text."));
        PdfDoc satisfiedDocument = PdfDoc.Create(CreatePdfUaGroundworkOptions())
            .TaggedPdfCatalogMarkers()
            .Meta(title: "Accessible title")
            .Header(header => header.Image(png, 20, 10, alternativeText: "Company logo"))
            .Footer(footer => footer.Image(png, 20, 10, alternativeText: "Security footer mark"))
            .Paragraph(paragraph => paragraph.Text("Header and footer image alt text."));

        PdfComplianceReadinessReport missingReport = missingDocument.AssessCompliance(PdfComplianceProfile.PdfUa1);
        PdfComplianceReadinessReport satisfiedReport = satisfiedDocument.AssessCompliance(PdfComplianceProfile.PdfUa1);

        PdfComplianceRequirement missing = AssertRequirement(missingReport, "generated-image-alternate-text", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("header/footer", missing.Diagnostic, StringComparison.Ordinal);
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
        PdfDoc document = PdfDoc.Create(CreatePdfUaGroundworkOptions())
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
        PdfDoc missingDocument = PdfDoc.Create(CreatePdfUaGroundworkOptions())
            .Meta(title: "Accessible title")
            .Shape(shape);
        PdfDoc satisfiedDocument = PdfDoc.Create(CreatePdfUaGroundworkOptions())
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
        PdfDoc document = PdfDoc.Create(CreatePdfUaGroundworkOptions())
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
        PdfDoc missingDocument = PdfDoc.Create(CreatePdfUaGroundworkOptions())
            .Meta(title: "Accessible title")
            .TextField("Contact.Email", value: "info@example.com");
        PdfDoc satisfiedDocument = PdfDoc.Create(CreatePdfUaGroundworkOptions())
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
        PdfDoc missingDocument = PdfDoc.Create(CreatePdfUaGroundworkOptions())
            .TaggedPdfCatalogMarkers()
            .Meta(title: "Accessible title")
            .TextField("Contact.Email", value: "info@example.com");
        PdfDoc satisfiedDocument = PdfDoc.Create(CreatePdfUaGroundworkOptions())
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

    [Fact]
    public void ImageAlternativeTextEmitsMarkedFigureContent() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Image(CreateMinimalRgbPng(), 24, 24, alternativeText: "Company logo")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Figure << /Alt <436F6D70616E79206C6F676F> >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("EMC", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedImageAlternativeTextEmitsFigureStructureReferences() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Image(CreateMinimalRgbPng(), 24, 24, alternativeText: "Company logo")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/MarkInfo << /Marked true >>", content, StringComparison.Ordinal);
        Assert.Contains("/StructTreeRoot", content, StringComparison.Ordinal);
        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <436F6D70616E79206C6F676F> /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTreeNextKey 1", content, StringComparison.Ordinal);
        Assert.Contains("/Nums [0 [", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Document", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Figure", content, StringComparison.Ordinal);
        Assert.Contains("/K << /Type /MCR", content, StringComparison.Ordinal);
        Assert.Contains("/MCID 0", content, StringComparison.Ordinal);
        Assert.Contains("/Alt <436F6D70616E79206C6F676F>", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedDocumentStructureRootEmitsLanguageMetadata() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false,
                Language = "en-US"
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph.Text("Language metadata for generated document structure."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Lang <656E2D5553>", content, StringComparison.Ordinal);
        Assert.Matches(@"/Type /StructElem /S /Document /P \d+ 0 R /K \[[^\]]+\] /Lang <656E2D5553>", content);
    }

    [Fact]
    public void TaggedHeadingParagraphAndImageEmitStructureReferencesWithPageScopedMcids() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .H1("Quarterly summary")
            .Paragraph(paragraph => paragraph.Text("Revenue and risk notes."))
            .Image(CreateMinimalRgbPng(), 24, 24, alternativeText: "Company logo")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Contains("/H1 << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/P << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <436F6D70616E79206C6F676F> /MCID 2 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Document", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /H1", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /P", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Figure", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTreeNextKey 1", content, StringComparison.Ordinal);
        Assert.Contains("/Nums [0 [", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedParagraphLinkEmitsStructureTabOrder() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph
                .Text("Read the ")
                .Link("project site", "https://officeimo.net/", contents: "Project site"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Annots [", content, StringComparison.Ordinal);
        Assert.Contains("/StructParents 0 /Tabs /S", content, StringComparison.Ordinal);
        Assert.Contains("/P << /MCID 0 >> BDC", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedParagraphLinkEmitsAnnotationStructureReferences() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph
                .Text("Read the ")
                .Link("project site", "https://officeimo.net/", contents: "Project site"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Subtype /Link", content, StringComparison.Ordinal);
        Assert.Contains("/StructParent 1", content, StringComparison.Ordinal);
        Assert.Contains("ET\nEMC\n/Link << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Link << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/P << /MCID 2 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Link << /MCID 3 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Document", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Link", content, StringComparison.Ordinal);
        Assert.Contains("/K [<< /Type /MCR /Pg ", content, StringComparison.Ordinal);
        Assert.Contains("/MCID 1 >> << /Type /MCR /Pg ", content, StringComparison.Ordinal);
        Assert.Contains("/MCID 3 >> << /Type /OBJR /Obj ", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTreeNextKey 2", content, StringComparison.Ordinal);
        Assert.Contains("/Nums [0 [", content, StringComparison.Ordinal);
        Assert.Matches(@"/Nums \[0 \[[^\]]+\] 1 \d+ 0 R\]", content);
        Assert.Matches(@"/Nums \[0 \[\d+ 0 R (?<link>\d+) 0 R \d+ 0 R \k<link> 0 R\] 1 \k<link> 0 R\]", content);
    }

    [Fact]
    public void TaggedLinkedHeadingEmitsLinkStructureForVisibleText() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .H1("Premium PDF fonts", linkUri: "https://officeimo.net/", linkContents: "OfficeIMO PDF")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Subtype /Link", content, StringComparison.Ordinal);
        Assert.Contains("/Link << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /H1", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Link", content, StringComparison.Ordinal);
        Assert.Matches(@"/Type /StructElem /S /H1 /P \d+ 0 R /Pg \d+ 0 R /K \[\d+ 0 R\]", content);
        Assert.Matches(@"/Type /StructElem /S /Link /P \d+ 0 R /Pg \d+ 0 R /K \[<< /Type /MCR /Pg \d+ 0 R /MCID 0 >> << /Type /OBJR /Obj \d+ 0 R >>\]", content);
        Assert.Contains("/Type /OBJR /Obj", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedWrappedLinkedHeadingPreservesAllAnnotationStructureReferences() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false,
                PageWidth = 170,
                MarginLeft = 24,
                MarginRight = 24
            })
            .TaggedPdfCatalogMarkers()
            .H1("Premium PDF fonts and compliance evidence", linkUri: "https://officeimo.net/", linkContents: "OfficeIMO PDF")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Subtype /Link") >= 2);
        Assert.Contains("/Type /StructElem /S /H1", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Link", content, StringComparison.Ordinal);
        Assert.Matches(@"/Type /StructElem /S /Link /P \d+ 0 R /Pg \d+ 0 R /K \[<< /Type /MCR /Pg \d+ 0 R /MCID 0 >> << /Type /OBJR /Obj \d+ 0 R >> << /Type /OBJR /Obj \d+ 0 R >>", content);
        Assert.Matches(@"/Nums \[0 \[(?<link>\d+) 0 R\] 1 \k<link> 0 R 2 \k<link> 0 R", content);
    }

    [Fact]
    public void TaggedRichTextDecorationsEmitArtifactMarkedContent() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph
                .Underlined("Underlined")
                .Text(" ")
                .Strikethrough("Strike"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/P << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 2);
    }

    [Fact]
    public void TaggedRichTextBackgroundFillsEmitArtifactMarkedContent() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph
                .BackgroundColor(PdfColor.FromRgb(255, 255, 0))
                .Text("Highlighted text"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/P << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Artifact BMC", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedLinkedBaselineShiftResetsTextRiseBeforeNormalText() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph
                .Text("Value ")
                .Superscript()
                .Link("2", "https://officeimo.net/sup", underline: false)
                .Superscript(false)
                .Text(" normal"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("0 Ts", content, StringComparison.Ordinal);
        Assert.Matches(@"/Link << /MCID \d+ >> BDC[\s\S]+[1-9]\d*(?:\.\d+)? Ts[\s\S]+EMC[\s\S]+0 Ts", content);
    }

    [Fact]
    public void UntaggedParagraphLinksDoNotEmitDanglingMarkedContentReferences() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph
                .Text("Read the ")
                .Link("project site", "https://officeimo.net/", contents: "Project site"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Subtype /Link", content, StringComparison.Ordinal);
        Assert.DoesNotContain("/Link << /MCID", content, StringComparison.Ordinal);
        Assert.DoesNotContain("/StructParent", content, StringComparison.Ordinal);
        Assert.DoesNotContain("/StructTreeRoot", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedFormWidgetsEmitStructureReferences() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .TextField("Contact.Email", value: "info@example.com")
            .CheckBox("Contact.Accepted", isChecked: true)
            .ChoiceField("Contact.Country", new[] { "PL", "DE" }, value: "PL")
            .RadioButtonGroup("Contact.Approval", new[] { "Yes", "No" }, value: "Yes")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Equal(5, CountOccurrences(content, "/Subtype /Widget"));
        Assert.Equal(5, CountOccurrences(content, "/Type /StructElem /S /Form"));
        Assert.Contains("/StructParent 0", content, StringComparison.Ordinal);
        Assert.Contains("/StructParent 1", content, StringComparison.Ordinal);
        Assert.Contains("/StructParent 2", content, StringComparison.Ordinal);
        Assert.Contains("/StructParent 3", content, StringComparison.Ordinal);
        Assert.Contains("/StructParent 4", content, StringComparison.Ordinal);
        Assert.Contains("/Type /OBJR /Obj", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTreeNextKey 5", content, StringComparison.Ordinal);
        Assert.Matches(@"/Nums \[0 \d+ 0 R 1 \d+ 0 R 2 \d+ 0 R 3 \d+ 0 R 4 \d+ 0 R\]", content);
    }

    [Fact]
    public void TaggedListItemsEmitLabelAndBodyStructureReferencesWithPageScopedMcids() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .H1("Checklist")
            .Bullets(new[] { "First item", "Second item" })
            .Image(CreateMinimalRgbPng(), 24, 24, alternativeText: "Company logo")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Contains("/H1 << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Lbl << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/LBody << /MCID 2 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Lbl << /MCID 3 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/LBody << /MCID 4 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <436F6D70616E79206C6F676F> /MCID 5 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /H1", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /L", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /LI", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Lbl", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /LBody", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Figure", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedRowColumnListItemsEmitLabelAndBodyStructureReferences() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Numbered(new[] { "First item", "Second item" }))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Contains("/Lbl << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/LBody << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Lbl << /MCID 2 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/LBody << /MCID 3 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /L", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /LI", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Lbl", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /LBody", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedTableCellsEmitStructureReferencesWithPageScopedMcids() {
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 1;

        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .H1("Table")
            .Table(new[] {
                new[] { "Name", "Status" },
                new[] { "Alpha", "Ready" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Contains("/H1 << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TH << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TH << /MCID 2 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TD << /MCID 3 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TD << /MCID 4 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /H1", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Table", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TR", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TH", content, StringComparison.Ordinal);
        Assert.Contains("/A << /O /Table /Scope /Column >>", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TD", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
        Assert.Contains("/Nums [0 [", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedLinkedTableCellWrapsTextInLinkStructure() {
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 0;

        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Table(new[] {
                new[] { PdfTableCell.TextCell("Resource", linkUri: "https://officeimo.net/table", linkContents: "Linked table resource") }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Subtype /Link", content, StringComparison.Ordinal);
        Assert.Contains("/Link << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Table", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TD", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Link", content, StringComparison.Ordinal);
        Assert.Matches(@"/Type /StructElem /S /TD /P \d+ 0 R /Pg \d+ 0 R /K \[\d+ 0 R\]", content);
        Assert.Matches(@"/Type /StructElem /S /Link /P \d+ 0 R /Pg \d+ 0 R /K \[<< /Type /MCR /Pg \d+ 0 R /MCID \d+ >> << /Type /OBJR /Obj \d+ 0 R >>\]", content);
        Assert.Matches(@"/Nums \[0 \[[^\]]+\] 1 (?<link>\d+) 0 R\]", content);
    }

    [Fact]
    public void TaggedTableDataBarsEmitArtifactMarkedContent() {
        PdfTableStyle style = TableStyles.Minimal();
        style.CellDataBars = new Dictionary<(int Row, int Column), PdfCellDataBar> {
            [(1, 1)] = new PdfCellDataBar {
                Ratio = 0.75,
                Color = PdfColor.FromRgb(34, 197, 94)
            }
        };

        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Table(new[] {
                new[] { "Name", "Progress" },
                new[] { "Alpha", "75%" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Artifact BMC", content, StringComparison.Ordinal);
        Assert.Contains("/TD << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TD", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedMergedTableCellsEmitSpanStructureAttributes() {
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 1;

        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Table(new[] {
                new[] { PdfTableCell.Span("Group", 2) },
                new[] { PdfTableCell.Merge("Alpha", rowSpan: 2), new PdfTableCell("Ready") },
                new[] { new PdfTableCell("Done") }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/TH << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TD << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/A << /O /Table /Scope /Column /ColSpan 2 >>", content, StringComparison.Ordinal);
        Assert.Contains("/A << /O /Table /RowSpan 2 >>", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedTableCaptionEmitsCaptionStructureReferences() {
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.Caption = "Table 1. Status signals";

        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Table(new[] {
                new[] { "Name", "Status" },
                new[] { "Alpha", "Ready" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Caption << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TH << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Table", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Caption", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TR", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedRowColumnTableCaptionEmitsCaptionStructureReferences() {
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.Caption = "Table 1. Column signals";

        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Name", "Status" },
                                    new[] { "Alpha", "Ready" }
                                }, style: style))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Caption << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TH << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Table", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Caption", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TR", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
    }

    [Fact]
    public void HeaderFooterImageAlternativeTextEmitsMarkedFigureContent() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Header(header => header.Image(CreateMinimalRgbPng(), 24, 12, alternativeText: "Header logo"))
            .Footer(footer => footer.Image(CreateMinimalRgbPng(), 24, 12, alternativeText: "Footer logo"))
            .Paragraph(paragraph => paragraph.Text("Header and footer image marked content."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Figure << /Alt <486561646572206C6F676F> >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <466F6F746572206C6F676F> >> BDC", content, StringComparison.Ordinal);
    }

    [Fact]
    public void DecorativeBackgroundAndWatermarkImagesEmitArtifactMarkedContent() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .BackgroundImage(CreateMinimalRgbPng(), OfficeIMO.Drawing.OfficeImageFit.Stretch, opacity: 0.2)
            .ImageWatermark(CreateMinimalRgbPng(), 80, 40, opacity: 0.2)
            .Paragraph(paragraph => paragraph.Text("Decorative image artifact marked content."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 2);
        Assert.DoesNotContain("/Figure << /Alt", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedDecorativePageChromeEmitsArtifactMarkedContent() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Background(PdfColor.White)
            .BackgroundImage(CreateMinimalRgbPng(), OfficeIMO.Drawing.OfficeImageFit.Stretch, opacity: 0.2)
            .ImageWatermark(CreateMinimalRgbPng(), 80, 40, opacity: 0.2)
            .Watermark("DRAFT", fontSize: 48, opacity: 0.18)
            .PageBorder(inset: 30, opacity: 0.4)
            .Paragraph(paragraph => paragraph.Text("Decorative page chrome artifact marked content."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 5);
        Assert.Contains("/Type /StructElem /S /Document", content, StringComparison.Ordinal);
        Assert.DoesNotContain("/Figure << /Alt", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedRunningHeaderFooterTextEmitsArtifactMarkedContent() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Header(header => header.Text("Running header"))
            .Footer(footer => footer.Text("Page {page} of {pages}"))
            .Paragraph(paragraph => paragraph.Text("Body text remains structured."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 2);
        Assert.Contains("/P << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Document", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedHorizontalRulesEmitArtifactMarkedContent() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph.Text("Before top-level rule."))
            .HR(thickness: 1.2, color: PdfColor.FromRgb(148, 163, 184), spacingBefore: 2, spacingAfter: 2)
            .Paragraph(paragraph => paragraph.Text("After top-level rule."))
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row
                .Column(100, column => column
                    .Paragraph(paragraph => paragraph.Text("Before row rule."))
                    .HR(thickness: 0.8, color: PdfColor.FromRgb(203, 213, 225), spacingBefore: 2, spacingAfter: 2)
                    .Paragraph(paragraph => paragraph.Text("After row rule.")))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 2);
        Assert.True(CountOccurrences(content, "/P << /MCID") >= 4);
        Assert.Contains("/Type /StructElem /S /Document", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedDecorativeLayoutChromeEmitsArtifactMarkedContent() {
        var panelStyle = new PanelStyle {
            Background = PdfColor.FromRgb(248, 250, 252),
            BorderColor = PdfColor.FromRgb(37, 99, 235),
            BorderWidth = 0.8,
            PaddingX = 10,
            PaddingY = 8
        };
        PdfTableStyle tableStyle = TableStyles.Minimal();
        tableStyle.HeaderRowCount = 1;
        tableStyle.HeaderFill = PdfColor.FromRgb(229, 231, 235);
        tableStyle.RowStripeFill = PdfColor.FromRgb(248, 250, 252);
        tableStyle.BorderColor = PdfColor.FromRgb(148, 163, 184);
        tableStyle.BorderWidth = 0.7;
        tableStyle.RowSeparatorColor = PdfColor.FromRgb(203, 213, 225);
        tableStyle.RowSeparatorWidth = 0.5;
        tableStyle.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(1, 1)] = PdfColor.FromRgb(220, 252, 231)
        };
        tableStyle.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(1, 1)] = new PdfCellBorder {
                Color = PdfColor.FromRgb(22, 163, 74),
                Width = 0.8
            }
        };

        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Panel(panel => panel.Paragraph(paragraph => paragraph.Text("Decorative panel chrome.")), panelStyle)
            .Table(new[] {
                new[] { "Name", "Status" },
                new[] { "Alpha", "Ready" },
                new[] { "Beta", "Done" }
            }, style: tableStyle)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 8);
        Assert.Contains("/P << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/TH << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/TD << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Table", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TD", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedShapeAndDrawingAlternativeTextEmitFigureStructureReferences() {
        OfficeShape shape = CreateComplianceShape();
        OfficeDrawing drawing = new OfficeDrawing(36, 18)
            .AddShape(CreateComplianceShape(36, 18), 0, 0);

        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Shape(shape, style: new PdfDrawingStyle {
                AlternativeText = "Risk status badge"
            }, linkUri: "https://officeimo.net/shape", linkContents: "Risk shape")
            .Drawing(drawing, style: new PdfDrawingStyle {
                AlternativeText = "Approval workflow diagram"
            }, linkUri: "https://officeimo.net/drawing", linkContents: "Approval drawing")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/MarkInfo << /Marked true >>", content, StringComparison.Ordinal);
        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Equal(2, CountOccurrences(content, "/Subtype /Link"));
        Assert.Contains("/Figure << /Alt <5269736B20737461747573206261646765> /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <417070726F76616C20776F726B666C6F77206469616772616D> /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Figure", content, StringComparison.Ordinal);
        Assert.Matches(@"/Type /StructElem /S /Figure /P \d+ 0 R /Pg \d+ 0 R /K \[<< /Type /MCR /Pg \d+ 0 R /MCID 0 >> << /Type /OBJR /Obj \d+ 0 R >>\] /Alt", content);
        Assert.Matches(@"/Type /StructElem /S /Figure /P \d+ 0 R /Pg \d+ 0 R /K \[<< /Type /MCR /Pg \d+ 0 R /MCID 1 >> << /Type /OBJR /Obj \d+ 0 R >>\] /Alt", content);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
        Assert.Matches(@"/Nums \[0 \[(?<shape>\d+) 0 R (?<drawing>\d+) 0 R\] 1 \k<shape> 0 R 2 \k<drawing> 0 R\]", content);
    }

    [Fact]
    public void TaggedDecorativeShapeAndDrawingEmitArtifactMarkedContent() {
        OfficeDrawing drawing = new OfficeDrawing(36, 18)
            .AddShape(CreateComplianceShape(36, 18), 0, 0);

        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Shape(CreateComplianceShape(), style: new PdfDrawingStyle {
                Decorative = true
            })
            .Drawing(drawing, style: new PdfDrawingStyle {
                Decorative = true
            })
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 2);
        Assert.DoesNotContain("/Figure << /Alt", content, StringComparison.Ordinal);
        Assert.Contains("/StructTreeRoot", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageAlternativeTextRejectsWhitespace() {
        Assert.Throws<ArgumentException>(() => new PdfImageStyle {
            AlternativeText = " "
        });

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Image(CreateMinimalRgbPng(), 24, 24, alternativeText: " "));

        Assert.Throws<ArgumentException>(() => new PdfHeaderFooterImage(
            CreateMinimalRgbPng(),
            24,
            12,
            alternativeText: " "));
    }

    [Fact]
    public void DrawingAlternativeTextRejectsWhitespaceAndDecorativeConflict() {
        Assert.Throws<ArgumentException>(() => new PdfDrawingStyle {
            AlternativeText = " "
        });

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Shape(CreateComplianceShape(), style: new PdfDrawingStyle {
                AlternativeText = "Meaningful badge",
                Decorative = true
            }));
    }

    private static OfficeShape CreateComplianceShape(double width = 24, double height = 24) {
        OfficeShape shape = OfficeShape.Rectangle(width, height);
        shape.FillColor = OfficeColor.FromRgb(219, 234, 254);
        shape.StrokeColor = OfficeColor.FromRgb(37, 99, 235);
        shape.StrokeWidth = 1;
        return shape;
    }

    private static PdfDoc CreatePdfA3GroundworkDocument() {
        return PdfDoc.Create(CreatePdfA3GroundworkOptions());
    }

    private static PdfOptions CreatePdfA3GroundworkOptions() {
        return new PdfOptions {
                IncludeStandardFontToUnicodeMaps = true
            }
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent();
    }

    private static PdfOptions CreatePdfUaGroundworkOptions() {
        return new PdfOptions {
                IncludeStandardFontToUnicodeMaps = true,
                Language = "en-US",
                ViewerPreferences = new PdfViewerPreferencesOptions {
                    DisplayDocTitle = true
                }
            }
            .SetPdfUaIdentification();
    }

    private static PdfComplianceRequirement AssertRequirement(PdfComplianceReadinessReport report, string id, PdfComplianceRequirementStatus status) {
        PdfComplianceRequirement requirement = Assert.Single(report.Requirements, requirement => requirement.Id == id);
        Assert.Equal(status, requirement.Status);
        Assert.False(string.IsNullOrWhiteSpace(requirement.DisplayName));
        Assert.False(string.IsNullOrWhiteSpace(requirement.Diagnostic));
        return requirement;
    }

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += value.Length;
        }

        return count;
    }

}
