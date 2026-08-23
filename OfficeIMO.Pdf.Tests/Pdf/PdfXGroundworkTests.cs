using OfficeIMO.Pdf;
using OfficeIMO.Drawing;
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
    public void PdfXExactArtifactIsInternallyReadyOnlyWithProductionBoxesAndEmbeddedFonts() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) return;
        var options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51")
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
        Assert.Equal(2, evidence.DeviceRgbOperatorCount);
        Assert.Equal(1, evidence.DeviceRgbImageCount);
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
    public void PdfX1AReadinessRemainsFailClosedUntilColorAndProductionPoliciesAreImplemented() {
        var options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX1A2003,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA39")
            .RequireCompliance(PdfComplianceProfile.PdfX1A2003);

        PdfComplianceReadinessReport readiness = PdfComplianceAnalyzer.Assess(options);
        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(options);

        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readiness.FindRequirement("pdfx-xmp-identification")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readiness.FindRequirement("pdfx-output-intent")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readiness.FindRequirement("pdfx-raster-color-conversion")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readiness.FindRequirement("pdfx-raster-transparency")!.Status);
        Assert.Equal(PdfComplianceRequirementStatus.Unsupported, readiness.FindRequirement("pdfx-source-color-management")!.Status);
        Assert.Contains(PdfExternalValidatorKind.PdfXValidator, proof.RequiredExternalValidators);
        Assert.False(proof.CanClaimConformance);
        Assert.Throws<NotSupportedException>(() => PdfDocument.Create(options).ToBytes());
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
}
