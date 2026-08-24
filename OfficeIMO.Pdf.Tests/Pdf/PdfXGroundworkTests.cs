using OfficeIMO.Pdf;
using OfficeIMO.Drawing;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfXGroundworkTests {
    [Fact]
    public void ExternalValidatorEnumPreservesLegacyCustomValue() {
        Assert.Equal(0, (int)PdfExternalValidatorKind.VeraPdf);
        Assert.Equal(1, (int)PdfExternalValidatorKind.PdfUaValidator);
        Assert.Equal(2, (int)PdfExternalValidatorKind.Mustang);
        Assert.Equal(3, (int)PdfExternalValidatorKind.Custom);
        Assert.Equal(4, (int)PdfExternalValidatorKind.PdfXValidator);
    }

    [Fact]
    public void PdfX4GroundworkEmitsTruthfulPrintProductionPrimitivesWithoutClaimingConformance() {
        byte[] cmykProfile = IccMabTestProfiles.CreateCmykLab8Bidirectional();
        var options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX4,
                cmykProfile,
                "FOGRA51",
                PdfTrappingStatus.False);
        options.CompressContentStreams = false;
        options.BackgroundColor = PdfColor.Gray;

        byte[] pdf = PdfDocument.Create(options)
            .Meta(title: "PDF/X-4 groundwork", author: "OfficeIMO")
            .Paragraph(paragraph => paragraph.Text("Neutral black-preservation proof."))
            .Image(PdfPngTestImages.CreateRgbPng(12, 34, 56), 12, 12)
            .ToBytes();
        string raw = Encoding.UTF8.GetString(pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfComplianceReadinessReport readback = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfX4, pdf);
        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();
        PdfPrintProductionStructureEvidence structure = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();
        PdfOptions clone = options.Clone();

        Assert.Equal(PdfComplianceProfile.None, options.ComplianceProfile);
        Assert.Equal(PdfFileVersion.Pdf16, clone.FileVersion);
        Assert.Equal("PDF/X-4", clone.PdfXIdentification!.Version);
        Assert.Null(clone.PdfXIdentification.Conformance);
        Assert.Equal(PdfOutputIntentSubtype.GtsPdfX, clone.OutputIntent!.Subtype);
        Assert.Equal(PdfOutputIntentPolicy.PdfXPrintCondition, clone.OutputIntent.Policy);
        Assert.Equal(4, clone.OutputIntent.ColorComponents);
        Assert.Equal(PdfTrappingStatus.False, clone.TrappingStatus);
        Assert.NotNull(clone.PrintProductionPageBoxes);
        Assert.StartsWith("%PDF-1.6", raw, StringComparison.Ordinal);
        Assert.Contains("xmlns:pdfxid=\"http://www.npes.org/pdfx/ns/id/\"", raw, StringComparison.Ordinal);
        Assert.Contains("<pdfxid:GTS_PDFXVersion>PDF/X-4</pdfxid:GTS_PDFXVersion>", raw, StringComparison.Ordinal);
        Assert.Contains("/S /GTS_PDFX", raw, StringComparison.Ordinal);
        Assert.Contains("/Trapped /False", raw, StringComparison.Ordinal);
        Assert.Contains("/TrimBox [0 0 612 792] /BleedBox [0 0 612 792]", raw, StringComparison.Ordinal);
        Assert.Contains("0 0 0 0.5 k", raw, StringComparison.Ordinal);
        Assert.DoesNotContain(" rg\n", raw, StringComparison.Ordinal);
        Assert.Contains("/ColorSpace /DeviceCMYK", raw, StringComparison.Ordinal);
        Assert.Equal("PDF/X-4", info.XmpMetadata!.PdfXVersion);
        Assert.Equal(PdfTrappingStatus.False, info.Metadata.TrappingStatus);
        Assert.Equal("GTS_PDFX", Assert.Single(info.OutputIntents).Subtype);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readback.FindRequirement("readback-pdfx-color-inspection-complete")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readback.FindRequirement("readback-pdfx-no-device-rgb")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Missing, readback.FindRequirement("readback-pdfx-fonts-embedded")!.Status);
        Assert.True(evidence.IsComplete);
        Assert.False(evidence.HasDeviceRgbUsage);
        Assert.True(evidence.DeviceCmykOperatorCount > 0);
        Assert.Equal(1, evidence.DeviceCmykImageCount);
        Assert.Equal(1, structure.ValidProductionPageBoxCount);
    }

    [Fact]
    public void PdfXFlattenedAnnotationAppearancesAreConvertedToPrintColorSpace() {
        var options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51",
                PdfTrappingStatus.False)
            .SetFlattenVisualAnnotations();
        options.CompressContentStreams = false;

        byte[] pdf = PdfDocument.Create(options)
            .FreeTextAnnotation(
                "CMYK reviewer note",
                width: 140,
                height: 44,
                textColor: new PdfColor(0.15D, 0.35D, 0.75D),
                borderColor: new PdfColor(0.8D, 0.2D, 0.1D),
                fillColor: new PdfColor(0.9D, 0.95D, 1D))
            .HighlightAnnotation(
                "CMYK highlight",
                width: 120,
                height: 14,
                color: new PdfColor(1D, 0.9D, 0.1D))
            .ToBytes();

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.False(evidence.HasDeviceRgbUsage);
        Assert.True(evidence.DeviceCmykOperatorCount > 0);
    }

    [Fact]
    public void PdfXColorNormalizationPreservesPdfLexicalContents() {
        var options = new PdfOptions().ConfigurePdfXGroundwork(
            PdfComplianceProfile.PdfX4,
            IccMabTestProfiles.CreateCmykLab8Bidirectional(),
            "FOGRA51",
            PdfTrappingStatus.False);
        PdfPrintColorTransform transform = Assert.IsType<PdfPrintColorTransform>(PdfPrintColorTransform.Create(options));

        string normalized = transform.NormalizeGeneratedContent(
            "BT (literal\n1 0 0 rg text) Tj ET\n" +
            "<313020302030207267> Tj\n" +
            "<< /Note (dictionary) >> 0.2 0.3 0.4 rg\n" +
            "% 0.4 0.5 0.6 rg remains a comment\n" +
            "0.1 0.2 0.3 rg 0 0 10 10 re f");

        Assert.Contains("(literal\n1 0 0 rg text)", normalized, StringComparison.Ordinal);
        Assert.Contains("<313020302030207267>", normalized, StringComparison.Ordinal);
        Assert.DoesNotContain(">> 0.2 0.3 0.4 rg", normalized, StringComparison.Ordinal);
        Assert.Contains("% 0.4 0.5 0.6 rg remains a comment", normalized, StringComparison.Ordinal);
        Assert.DoesNotContain("0.1 0.2 0.3 rg", normalized, StringComparison.Ordinal);
        Assert.Contains(" k 0 0 10 10 re f", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfXExactArtifactIsInternallyReadyOnlyWithProductionBoxesAndEmbeddedFonts() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) return;
        var options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51",
                PdfTrappingStatus.False)
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath), "PDF/X audit font");

        PdfComplianceArtifact artifact = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Self-contained print artifact."))
            .CreateComplianceArtifact(PdfComplianceProfile.PdfX4);
        PdfPrintProductionStructureEvidence structure = PdfReadDocument.Open(artifact.ToBytes())
            .InspectPrintProductionStructure();

        Assert.True(structure.IsComplete);
        Assert.Equal(1, structure.FontResourceCount);
        Assert.Equal(0, structure.UnembeddedFontResourceCount);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, artifact.Readiness.FindRequirement("readback-pdfx-page-boxes")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, artifact.Readiness.FindRequirement("readback-pdfx-fonts-embedded")!.Status);
        Assert.All(
            artifact.Readiness.Requirements.Where(requirement => requirement.Id != "pdfx-validation"),
            requirement => Assert.Equal(PdfComplianceRequirementStatus.Satisfied, requirement.Status));
        Assert.False(artifact.AssessProof().CanClaimConformance);
    }

    [Fact]
    public void PdfXStructureInspectorAcceptsEmbeddedOpenTypeCffFontPrograms() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        if (fontPath == null) return;
        var options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51")
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath), "PDF/X CFF audit font");

        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Embedded CFF print artifact."))
            .ToBytes();
        PdfPrintProductionStructureEvidence structure = PdfReadDocument.Open(pdf)
            .InspectPrintProductionStructure();

        Assert.Equal(1, structure.FontResourceCount);
        Assert.Equal(0, structure.UnembeddedFontResourceCount);
        Assert.Equal(0, structure.UninspectableFontResourceCount);
    }

    [Fact]
    public void PdfXGeneratedAndExactPoliciesRejectLinks() {
        var options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51");
        PdfDocument document = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Link("External", "https://example.com"));

        PdfComplianceReadinessReport generated = document.AssessCompliance(PdfComplianceProfile.PdfX4);
        PdfComplianceArtifact artifact = document.CreateComplianceArtifact(PdfComplianceProfile.PdfX4);

        Assert.Equal(PdfComplianceRequirementStatus.Missing, generated.FindRequirement("pdfx-no-annotations")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Missing, generated.FindRequirement("pdfx-no-external-references")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Missing, artifact.Readiness.FindRequirement("readback-pdfx-no-annotations")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Missing, artifact.Readiness.FindRequirement("readback-pdfx-no-external-references")!.Status);
    }

    [Fact]
    public void PrintProductionPageBoxesRequireBleedToContainTrim() {
        Assert.Throws<ArgumentException>(() => new PdfPrintProductionPageBoxes(
            PageMargins.Uniform(3D),
            PageMargins.Uniform(6D)));

        var options = new PdfOptions {
            PageWidth = 100D,
            PageHeight = 100D,
            Margins = PageMargins.Uniform(0D),
            PrintProductionPageBoxes = new PdfPrintProductionPageBoxes(
                new PageMargins(60D, 0D, 40D, 0D),
                PageMargins.Uniform(0D))
        };

        Assert.Throws<InvalidOperationException>(() => PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Invalid trim geometry."))
            .ToBytes());
    }

    [Fact]
    public void PdfX1AGroundworkFlattensRasterAlphaAndExactReadbackFindsNoRgbOrTransparency() {
        var raster = new OfficeRasterImage(1, 1);
        raster.SetPixel(0, 0, OfficeColor.FromRgba(20, 40, 60, 96));
        byte[] png = OfficePngWriter.Encode(raster);
        var options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX1A2003,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA39");
        options.CompressContentStreams = false;

        byte[] pdf = PdfDocument.Create(options)
            .Image(png, 12, 12)
            .ToBytes();
        string raw = Encoding.UTF8.GetString(pdf);
        PdfComplianceReadinessReport readback = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfX1A2003, pdf);
        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Contains("/ColorSpace /DeviceCMYK", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/SMask", raw, StringComparison.Ordinal);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readback.FindRequirement("readback-pdfx-color-inspection-complete")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readback.FindRequirement("readback-pdfx-no-device-rgb")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readback.FindRequirement("readback-pdfx1a-no-transparency")!.Status);
        Assert.False(evidence.HasDeviceRgbUsage);
        Assert.False(evidence.HasTransparency);
        Assert.Equal(1, evidence.DeviceCmykImageCount);
    }

    [Fact]
    public void PdfXRasterConversionDoesNotDependOnVectorConversion() {
        var options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51");
        options.CompressContentStreams = false;
        options.ConvertVectorColorsToPdfXPrintCondition = false;
        options.BackgroundColor = PdfColor.Gray;

        byte[] pdf = PdfDocument.Create(options)
            .Image(PdfPngTestImages.CreateRgbPng(12, 34, 56), 12, 12)
            .ToBytes();
        string raw = Encoding.UTF8.GetString(pdf);
        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Contains("0.5 0.5 0.5 rg", raw, StringComparison.Ordinal);
        Assert.Contains("/ColorSpace /DeviceCMYK", raw, StringComparison.Ordinal);
        Assert.True(evidence.HasDeviceRgbUsage);
        Assert.Equal(1, evidence.DeviceCmykImageCount);
    }

    [Fact]
    public void PdfXCmykJpegIsDecodedAndConvertedInsteadOfRelabeled() {
        byte[] cmykJpeg = Convert.FromBase64String(
            "/9j/7gAOQWRvYmUAZAAAAAAA/9sAQwABAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEB/8AAFAgAAQABBEMRAE0RAFkRAEsRAP/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/aAA4EQwBNAFkASwAAPwD+/iv8/wDr/P8A6/v4r//Z");
        var options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51");
        options.CompressContentStreams = false;

        byte[] pdf = PdfDocument.Create(options).Image(cmykJpeg, 12, 12).ToBytes();
        string raw = Encoding.UTF8.GetString(pdf);

        Assert.Contains("/ColorSpace /DeviceCMYK", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/DCTDecode", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfXGroundworkRejectsEncryptionRegardlessOfConfigurationOrder() {
        PdfOptions options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51")
            .SetEncryption("open", "owner");

        Assert.Throws<ArgumentException>(() => PdfDocument.Create(options).ToBytes());
    }

    [Fact]
    public void PdfXReadbackRejectsIccHeaderAndComponentCountMismatch() {
        var options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51");
        byte[] pdf = PdfDocument.Create(options).ToBytes();
        int signatureOffset = FindAscii(pdf, "acsp");
        Assert.True(signatureOffset >= 20);
        pdf[signatureOffset - 20] = (byte)'R';
        pdf[signatureOffset - 19] = (byte)'G';
        pdf[signatureOffset - 18] = (byte)'B';
        pdf[signatureOffset - 17] = (byte)' ';

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfX4, pdf);

        Assert.Equal(PdfComplianceRequirementStatus.Missing, report.FindRequirement("readback-pdfx-output-intent")!.Status);
    }

    [Fact]
    public void PdfXReadbackRejectsHeaderOnlyIccWithoutOutputTransform() {
        var options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51");
        byte[] pdf = PdfDocument.Create(options).ToBytes();
        int signatureOffset = FindAscii(pdf, "acsp");
        Assert.True(signatureOffset >= 36);
        int tagCountOffset = signatureOffset - 36 + 128;
        Assert.True(tagCountOffset + 3 < pdf.Length);
        pdf[tagCountOffset] = 0;
        pdf[tagCountOffset + 1] = 0;
        pdf[tagCountOffset + 2] = 0;
        pdf[tagCountOffset + 3] = 0;

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfX4, pdf);

        Assert.Equal(PdfComplianceRequirementStatus.Missing, report.FindRequirement("readback-pdfx-output-intent")!.Status);
    }

    [Fact]
    public void PrintProductionInspectorFindsNamedAndInlineDeviceRgb() {
        byte[] pdf = BuildInspectionPdf(
            "/CsRgb cs 0.1 0.2 0.3 sc /DeviceRGB CS 0.4 0.5 0.6 SCN " +
            "BI /W 1 /H 1 /BPC 8 /CS /RGB ID abc EI",
            resources: "/ColorSpace << /CsRgb /DeviceRGB >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.True(evidence.HasDeviceRgbUsage);
        Assert.Equal(4, evidence.DeviceRgbOperatorCount);
        Assert.Equal(1, evidence.DeviceRgbImageCount);
    }

    [Fact]
    public void PdfX1AReadbackRejectsDeviceIndependentColorSpaceUsage() {
        byte[] pdf = BuildInspectionPdf(
            "/CsLab cs 0 0 10 10 re f",
            resources: "/ColorSpace << /CsLab [/Lab << /WhitePoint [0.9642 1 0.8249] >>] >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();
        PdfComplianceReadinessReport pdfX1A = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfX1A2003, pdf);
        PdfComplianceReadinessReport pdfX4 = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfX4, pdf);

        Assert.True(evidence.HasDeviceIndependentColorUsage);
        Assert.Equal(1, evidence.DeviceIndependentColorUsageCount);
        Assert.Equal(
            PdfComplianceRequirementStatus.Missing,
            pdfX1A.FindRequirement("readback-pdfx1a-no-device-independent-color")!.Status);
        Assert.Null(pdfX4.FindRequirement("readback-pdfx1a-no-device-independent-color"));
    }

    [Fact]
    public void PrintProductionInspectorRejectsInitialDeviceRgbColorAfterSelection() {
        byte[] pdf = BuildInspectionPdf("/DeviceRGB cs 0 0 10 10 re f");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.HasDeviceRgbUsage);
        Assert.Equal(1, evidence.DeviceRgbOperatorCount);
    }

    [Fact]
    public void PrintProductionInspectorFindsPageGroupsAndUntypedGraphicsStates() {
        byte[] pdf = BuildInspectionPdf(
            "/Alpha gs",
            resources: "/ExtGState << /Alpha 5 0 R >>",
            pageEntries: "/Group << /S /Transparency >>",
            extraObjects: "5 0 obj\n<< /ca 0.5 >>\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.HasTransparency);
        Assert.Equal(1, evidence.NonOpaqueGraphicsStateCount);
        Assert.Equal(1, evidence.TransparencyGroupCount);
    }

    [Fact]
    public void PrintProductionInspectorFindsArrayValuedBlendModes() {
        byte[] pdf = BuildInspectionPdf(
            "/Blend gs",
            resources: "/ExtGState << /Blend 5 0 R >>",
            extraObjects: "5 0 obj\n<< /BM [/Multiply /Normal] >>\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.HasTransparency);
        Assert.Equal(1, evidence.NonOpaqueGraphicsStateCount);
    }

    [Fact]
    public void PdfXFormalGenerationFailsClosedWhenEmbeddedFontCoverageIsMissing() {
        var options = new PdfOptions()
            .ConfigurePdfX(
                PdfComplianceProfile.PdfX1A2003,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA39");

        PdfComplianceReadinessReport readiness = PdfComplianceAnalyzer.Assess(options);
        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(options);

        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readiness.FindRequirement("pdfx-xmp-identification")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readiness.FindRequirement("pdfx-output-intent")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readiness.FindRequirement("pdfx-raster-color-conversion")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readiness.FindRequirement("pdfx-raster-transparency")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Unsupported, readiness.FindRequirement("pdfx-source-color-management")!.Status);
        Assert.Contains(PdfExternalValidatorKind.PdfXValidator, proof.RequiredExternalValidators);
        Assert.False(proof.CanClaimConformance);
        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("Unembedded text must fail closed.")).ToBytes());
        Assert.Contains("font", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PdfXFormalConfigurationRequiresBooleanTrappingStatus() {
        byte[] profile = IccMabTestProfiles.CreateCmykLab8Bidirectional();

        var options = new PdfOptions().ConfigurePdfX(
            PdfComplianceProfile.PdfX4,
            profile,
            "FOGRA51");
        ArgumentException exception = Assert.Throws<ArgumentException>(() => new PdfOptions().ConfigurePdfX(
            PdfComplianceProfile.PdfX4,
            profile,
            "FOGRA51",
            PdfTrappingStatus.Unknown));

        Assert.Equal(PdfTrappingStatus.False, options.TrappingStatus);
        Assert.Equal("trappingStatus", exception.ParamName);
    }

    [Fact]
    public void PdfXReadinessRejectsUnknownTrappingStatus() {
        var options = new PdfOptions().ConfigurePdfXGroundwork(
            PdfComplianceProfile.PdfX4,
            IccMabTestProfiles.CreateCmykLab8Bidirectional(),
            "FOGRA51",
            PdfTrappingStatus.Unknown);
        byte[] pdf = PdfDocument.Create(options).ToBytes();

        PdfComplianceReadinessReport generated = PdfDocument.Create(options).AssessCompliance(PdfComplianceProfile.PdfX4);
        PdfComplianceReadinessReport readback = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfX4, pdf);

        Assert.Equal(PdfComplianceRequirementStatus.Missing, generated.FindRequirement("pdfx-trapping-status")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Missing, readback.FindRequirement("readback-pdfx-trapping-status")!.Status);
    }

    [Theory]
    [InlineData(PdfComplianceProfile.PdfX1A2003)]
    [InlineData(PdfComplianceProfile.PdfX4)]
    public void PdfXFormalGenerationReturnsOnlyAnInternallyReadyExactArtifact(PdfComplianceProfile profile) {
        string? fontPath = profile == PdfComplianceProfile.PdfX1A2003
            ? PdfComplianceTestFonts.FindBundledTrueTypeFont()
            : PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        var options = new PdfOptions()
            .ConfigurePdfX(
                profile,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                profile == PdfComplianceProfile.PdfX1A2003 ? "FOGRA39" : "FOGRA51",
                PdfTrappingStatus.False)
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath!), "PDF/X formal font");

        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Exact CMYK print production artifact."))
            .ToBytes();
        PdfComplianceReadinessReport readiness = PdfComplianceAnalyzer.AssessReadback(profile, pdf);

        Assert.All(
            readiness.Requirements.Where(requirement => requirement.Id != "pdfx-validation"),
            requirement => Assert.Equal(PdfComplianceRequirementStatus.Satisfied, requirement.Status));
        Assert.Equal(PdfComplianceRequirementStatus.Unsupported, readiness.FindRequirement("pdfx-validation")!.Status);
    }

    [Fact]
    public void PdfXProductionMetadataIsReconciledAcrossInfoAndXmp() {
        DateTimeOffset created = new DateTimeOffset(2026, 8, 24, 10, 15, 30, TimeSpan.FromHours(2));
        DateTimeOffset modified = created.AddMinutes(12);
        var production = new PdfXProductionMetadata(
            created,
            modified,
            new Guid("0a955277-5974-43da-aade-163749241dc3"),
            new Guid("afe6603d-0855-4439-a63d-f1674da0a866"),
            "7",
            "proof");
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        var options = new PdfOptions()
            .ConfigurePdfX(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51",
                PdfTrappingStatus.False)
            .SetPdfXProductionMetadata(production)
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath!), "PDF/X metadata font");

        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Reconciled production metadata."))
            .ToBytes();
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        Assert.Equal(created, info.Metadata.CreationDate);
        Assert.Equal(modified, info.Metadata.ModificationDate);
        Assert.Equal("PDF/X-4", info.Metadata.PdfXVersion);
        Assert.Equal(PdfTrappingStatus.False, info.Metadata.TrappingStatus);
        Assert.Equal(created, info.XmpMetadata!.CreationDate);
        Assert.Equal(modified, info.XmpMetadata.ModificationDate);
        Assert.Equal(modified, info.XmpMetadata.MetadataDate);
        Assert.Equal("uuid:0a955277-5974-43da-aade-163749241dc3", info.XmpMetadata.DocumentId);
        Assert.Equal("uuid:afe6603d-0855-4439-a63d-f1674da0a866", info.XmpMetadata.InstanceId);
        Assert.Equal("7", info.XmpMetadata.VersionId);
        Assert.Equal("proof", info.XmpMetadata.RenditionClass);
        Assert.Equal(PdfTrappingStatus.False, info.XmpMetadata.TrappingStatus);
    }

    [Fact]
    public void PdfXAutomaticProductionIdentityIsScopedToEachDocument() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        var options = new PdfOptions()
            .ConfigurePdfX(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51",
                PdfTrappingStatus.False)
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath!), "PDF/X identity font");

        PdfXmpMetadataInfo first = PdfInspector.Inspect(PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("First resource."))
            .ToBytes()).XmpMetadata!;
        PdfXmpMetadataInfo second = PdfInspector.Inspect(PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Second resource."))
            .ToBytes()).XmpMetadata!;

        Assert.NotEqual(first.DocumentId, second.DocumentId);
        Assert.NotEqual(first.InstanceId, second.InstanceId);
    }

    [Fact]
    public void ExplicitProductionMetadataIndependentlyCreatesWellFormedXmp() {
        var metadata = new PdfXProductionMetadata(
            new DateTimeOffset(2026, 8, 24, 12, 0, 0, TimeSpan.Zero),
            new DateTimeOffset(2026, 8, 24, 12, 1, 0, TimeSpan.Zero),
            new Guid("6ad44c01-cdb5-4a3e-a310-2346af838ba5"),
            new Guid("1a814373-f8ca-4b2b-b9e0-0b6553485b03"));

        byte[] pdf = PdfDocument.Create(new PdfOptions().SetPdfXProductionMetadata(metadata)).ToBytes();
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        Assert.True(info.XmpMetadata!.IsWellFormedXml);
        Assert.Equal("uuid:6ad44c01-cdb5-4a3e-a310-2346af838ba5", info.XmpMetadata.DocumentId);
        Assert.Equal(metadata.CreationDate, info.Metadata.CreationDate);
    }

    [Fact]
    public void PdfXMetadataOnlyEditsPreserveProductionMetadata() {
        byte[] source = CreatePdfXMetadataEditFixture();
        var updatedArtifacts = new[] {
            PdfMetadataEditor.UpdateMetadata(source, title: "Full rewrite title"),
            PdfMetadataEditor.SynchronizeMetadata(source, title: "Synchronized title"),
            PdfIncrementalUpdater.UpdateMetadata(source, title: "Append-only title")
        };

        foreach (byte[] updated in updatedArtifacts) {
            PdfDocumentInfo info = PdfInspector.Inspect(updated);
            Assert.Equal(new DateTimeOffset(2026, 8, 24, 13, 0, 0, TimeSpan.Zero), info.Metadata.CreationDate);
            Assert.Equal(new DateTimeOffset(2026, 8, 24, 13, 5, 0, TimeSpan.Zero), info.Metadata.ModificationDate);
            Assert.Equal("PDF/X-4", info.Metadata.PdfXVersion);
            Assert.Equal(PdfTrappingStatus.False, info.Metadata.TrappingStatus);
            Assert.Equal(info.Metadata.CreationDate, info.XmpMetadata!.CreationDate);
            Assert.Equal(info.Metadata.ModificationDate, info.XmpMetadata.ModificationDate);
            Assert.Equal(PdfTrappingStatus.False, info.XmpMetadata.TrappingStatus);
            Assert.Equal("uuid:f16a3ee3-645f-4114-b478-210d114f5265", info.XmpMetadata.DocumentId);
        }
    }

    [Theory]
    [InlineData("D:2026AB")]
    [InlineData("D:20260824130000+02'XX'")]
    [InlineData("D:20260824130000+00'60'")]
    [InlineData("D:20260824130000+14'01'")]
    [InlineData("D:20260824130000Zjunk")]
    public void PdfDateCodecRejectsMalformedOptionalComponents(string value) {
        Assert.Null(PdfDateCodec.TryParse(value));
    }

    [Fact]
    public void PdfXReadbackRejectsCreationAfterModification() {
        byte[] pdf = CreatePdfXMetadataEditFixture();
        ReplaceAsciiAll(
            pdf,
            PdfSyntaxEscaper.TextString("D:20260824130000+00'00'"),
            PdfSyntaxEscaper.TextString("D:20260824131000+00'00'"));
        ReplaceAsciiAll(pdf, "2026-08-24T13:00:00+00:00", "2026-08-24T13:10:00+00:00");

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfX4, pdf);

        Assert.Equal(PdfComplianceRequirementStatus.Missing, report.FindRequirement("readback-pdfx-production-dates")!.Status);
    }

    [Fact]
    public void PdfXFormalStreamFailureDoesNotOverwriteTheDestination() {
        var options = new PdfOptions()
            .ConfigurePdfX(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51");
        byte[] original = { 1, 2, 3, 4 };
        using var destination = new MemoryStream();
        destination.Write(original, 0, original.Length);
        destination.Position = 0;

        Assert.Throws<InvalidOperationException>(() => PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Missing embedded font."))
            .Save(destination));

        Assert.Equal(original, destination.ToArray());
    }

    [Fact]
    public async Task PdfXFormalAsyncCommitDoesNotReportCancellationAfterReplacingSeekableDestination() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        var options = new PdfOptions()
            .ConfigurePdfX(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51",
                PdfTrappingStatus.False)
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath!), "PDF/X cancellation font");
        using var cancellation = new CancellationTokenSource();
        using var destination = new CancelAfterWriteSeekableStream(
            Enumerable.Repeat((byte)0xCC, 16_384).ToArray(),
            cancellation.Cancel);

        PdfSaveResult result = await PdfDocument.Create(options)
            .Meta(title: "Committed PDF/X artifact")
            .Paragraph(paragraph => paragraph.Text("Cancellation begins only after the staged commit starts."))
            .SaveAsync(destination, cancellation.Token);

        Assert.True(result.Succeeded);
        Assert.True(cancellation.IsCancellationRequested);
        Assert.Equal(0, destination.Position);
        Assert.Equal(result.BytesWritten, destination.Length);
        Assert.Equal("Committed PDF/X artifact", PdfInspector.Inspect(destination.ToArray()).Metadata.Title);
    }

    [Fact]
    public void PdfXFormalGenerationAcceptsConfiguredExternalCmykProfile() {
        string? profilePath = Environment.GetEnvironmentVariable("OFFICEIMO_PDFX_ICC_PROFILE");
        if (string.IsNullOrWhiteSpace(profilePath) || !File.Exists(profilePath)) return;
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        byte[] profileBytes = File.ReadAllBytes(profilePath);
        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? parsedProfile));
        foreach (OfficeIccRenderingIntent intent in new[] {
                     OfficeIccRenderingIntent.Perceptual,
                     OfficeIccRenderingIntent.RelativeColorimetric,
                     OfficeIccRenderingIntent.Saturation
                 }) {
            Assert.True(parsedProfile!.TryConvertToDevice(OfficeColor.FromRgb(208, 64, 32), intent, out double[] components));
            Assert.Equal(4, components.Length);
            Assert.All(components, component => Assert.InRange(component, 0D, 1D));
        }
        var options = new PdfOptions()
            .ConfigurePdfX(
                PdfComplianceProfile.PdfX4,
                profileBytes,
                "FOGRA51",
                PdfTrappingStatus.False)
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath!), "PDF/X external profile font");

        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Externally sourced print profile."))
            .ToBytes();
        PdfComplianceReadinessReport readiness = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfX4, pdf);

        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readiness.FindRequirement("readback-pdfx-output-intent")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readiness.FindRequirement("readback-pdfx-no-device-rgb")!.Status);
    }

    [Fact]
    public void PdfXOutputIntentRejectsRgbProfilesAndPreservesSubtypeInClone() {
        Assert.Throws<ArgumentException>(() =>
            PdfOutputIntent.CreatePdfX(CreateMinimalIccProfile("RGB "), "sRGB"));

        PdfOutputIntent clone = new PdfOptions()
            .SetPdfXOutputIntent(IccMabTestProfiles.CreateCmykLab8Bidirectional(), "FOGRA51")
            .Clone()
            .OutputIntent!;

        Assert.Equal(PdfOutputIntentSubtype.GtsPdfX, clone.Subtype);
        Assert.Equal(PdfOutputIntentPolicy.PdfXPrintCondition, clone.Policy);
        Assert.Equal(4, clone.ColorComponents);
        Assert.Equal(OfficeIccProfileClass.OutputDevice, AssertParsedProfile(clone.IccProfile).ProfileClass);
    }

    [Theory]
    [InlineData("scnr")]
    [InlineData("mntr")]
    public void PdfXOutputIntentRejectsNonOutputDeviceIccProfiles(string deviceClass) {
        byte[] profile = IccMabTestProfiles.CreateCmykLab8Bidirectional();
        WriteAscii(profile, 12, deviceClass);

        Assert.Throws<ArgumentException>(() => PdfOutputIntent.CreatePdfX(profile, "FOGRA51"));
    }

    [Fact]
    public void PdfXReadbackRejectsNonOutputDeviceIccProfile() {
        var options = new PdfOptions().ConfigurePdfXGroundwork(
            PdfComplianceProfile.PdfX4,
            IccMabTestProfiles.CreateCmykLab8Bidirectional(),
            "FOGRA51",
            PdfTrappingStatus.False);
        byte[] pdf = PdfDocument.Create(options).ToBytes();
        int signatureOffset = FindAscii(pdf, "acsp");
        Assert.True(signatureOffset >= 36);
        WriteAscii(pdf, signatureOffset - 24, "scnr");

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfX4, pdf);

        Assert.Equal(PdfComplianceRequirementStatus.Missing, report.FindRequirement("readback-pdfx-output-intent")!.Status);
    }

    [Fact]
    public void PdfXGenerationGateRejectsMutatedNonOutputDeviceIntent() {
        byte[] validProfile = IccMabTestProfiles.CreateCmykLab8Bidirectional();
        var options = new PdfOptions().ConfigurePdfX(PdfComplianceProfile.PdfX4, validProfile, "FOGRA51");
        byte[] inputProfile = (byte[])validProfile.Clone();
        WriteAscii(inputProfile, 12, "scnr");
        options.OutputIntent = new PdfOutputIntent(
            inputProfile,
            "FOGRA51",
            PdfOutputIntentPolicy.PdfXPrintCondition,
            PdfOutputIntentSubtype.GtsPdfX);

        PdfComplianceReadinessReport readiness = PdfComplianceAnalyzer.Assess(options);

        Assert.Equal(PdfComplianceRequirementStatus.Missing, readiness.FindRequirement("pdfx-output-intent")!.Status);
        Assert.Throws<InvalidOperationException>(() => PdfDocument.Create(options).ToBytes());
    }

    private static OfficeIccColorProfile AssertParsedProfile(byte[] profileBytes) {
        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        return Assert.IsType<OfficeIccColorProfile>(profile);
    }

    private static byte[] CreateMinimalIccProfile(string colorSpace) {
        byte[] profile = new byte[128];
        profile[3] = 128;
        profile[16] = (byte)colorSpace[0];
        profile[17] = (byte)colorSpace[1];
        profile[18] = (byte)colorSpace[2];
        profile[19] = (byte)colorSpace[3];
        profile[36] = (byte)'a';
        profile[37] = (byte)'c';
        profile[38] = (byte)'s';
        profile[39] = (byte)'p';
        return profile;
    }

    private static byte[] CreatePdfXMetadataEditFixture() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        var metadata = new PdfXProductionMetadata(
            new DateTimeOffset(2026, 8, 24, 13, 0, 0, TimeSpan.Zero),
            new DateTimeOffset(2026, 8, 24, 13, 5, 0, TimeSpan.Zero),
            new Guid("f16a3ee3-645f-4114-b478-210d114f5265"),
            new Guid("1dd6877b-87c6-4949-adde-49795d8e3c22"));
        var options = new PdfOptions()
            .ConfigurePdfX(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51",
                PdfTrappingStatus.False)
            .SetPdfXProductionMetadata(metadata)
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath!), "PDF/X edit font");
        return PdfDocument.Create(options)
            .Meta(title: "Original production metadata")
            .Paragraph(paragraph => paragraph.Text("Metadata editing must retain the PDF/X production contract."))
            .ToBytes();
    }

    private static void ReplaceAsciiAll(byte[] bytes, string oldValue, string newValue) {
        Assert.Equal(oldValue.Length, newValue.Length);
        byte[] needle = Encoding.ASCII.GetBytes(oldValue);
        int replacements = 0;
        for (int index = 0; index <= bytes.Length - needle.Length; index++) {
            bool match = true;
            for (int needleIndex = 0; needleIndex < needle.Length; needleIndex++) {
                if (bytes[index + needleIndex] != needle[needleIndex]) {
                    match = false;
                    break;
                }
            }

            if (!match) continue;
            WriteAscii(bytes, index, newValue);
            replacements++;
            index += needle.Length - 1;
        }

        Assert.True(replacements > 0, "Expected metadata value was not found in the generated PDF.");
    }

    private static int FindAscii(byte[] bytes, string text) {
        byte[] needle = Encoding.ASCII.GetBytes(text);
        for (int index = 0; index <= bytes.Length - needle.Length; index++) {
            bool match = true;
            for (int needleIndex = 0; needleIndex < needle.Length; needleIndex++) {
                if (bytes[index + needleIndex] != needle[needleIndex]) {
                    match = false;
                    break;
                }
            }

            if (match) return index;
        }

        return -1;
    }

    private static byte[] BuildInspectionPdf(
        string content,
        string resources = "",
        string pageEntries = "",
        string extraObjects = "") {
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << " + resources + " >> " + pageEntries + " /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n" + extraObjects + "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteAscii(byte[] bytes, int offset, string value) {
        for (int index = 0; index < value.Length; index++) {
            bytes[offset + index] = (byte)value[index];
        }
    }

    private sealed class CancelAfterWriteSeekableStream : Stream {
        private readonly MemoryStream _inner;
        private readonly Action _cancel;
        private bool _hasCanceled;

        internal CancelAfterWriteSeekableStream(byte[] initialBytes, Action cancel) {
            _inner = new MemoryStream();
            _inner.Write(initialBytes, 0, initialBytes.Length);
            _inner.Position = 0;
            _cancel = cancel;
        }

        internal byte[] ToArray() => _inner.ToArray();
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => true;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        public override void SetLength(long value) => _inner.SetLength(value);
        public override void Write(byte[] buffer, int offset, int count) {
            _inner.Write(buffer, offset, count);
            if (_hasCanceled) return;
            _hasCanceled = true;
            _cancel();
        }
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
