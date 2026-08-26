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
        Assert.True(OfficeIccColorProfile.TryCreate(cmykProfile, out OfficeIccColorProfile? profile));
        Assert.True(profile!.TryDeriveNeutralBlack(OfficeColor.FromRgb(128, 128, 128), options.PdfXRenderingIntent, out double neutralBlack));
        string expectedNeutralOperator = "0 0 0 " + neutralBlack.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture) + " k";

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
        Assert.Contains(expectedNeutralOperator, raw, StringComparison.Ordinal);
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
    public void PdfXNeutralAxisUsesTheOutputProfilesBlackTone() {
        byte[] profileBytes = IccMabTestProfiles.CreateCmykLab8Bidirectional();
        var options = new PdfOptions().ConfigurePdfXGroundwork(
            PdfComplianceProfile.PdfX4,
            profileBytes,
            "FOGRA51",
            PdfTrappingStatus.False);
        PdfPrintColorTransform transform = Assert.IsType<PdfPrintColorTransform>(PdfPrintColorTransform.Create(options));
        OfficeColor neutral = OfficeColor.FromRgb(128, 128, 128);
        var converted = new double[4];

        transform.Convert(neutral, converted);
        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.True(profile!.TryDeriveNeutralBlack(neutral, options.PdfXRenderingIntent, out double expectedBlack));

        Assert.Equal(0D, converted[0]);
        Assert.Equal(0D, converted[1]);
        Assert.Equal(0D, converted[2]);
        Assert.Equal(expectedBlack, converted[3], 6);
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
    public void PdfXStructureInspectorRejectsCffProgramsWithoutCharStrings() {
        byte[] cff = {
            1, 0, 4, 1,
            0, 1, 1, 1, 2, (byte)'A',
            0, 1, 1, 1, 1,
            0, 0,
            0, 0
        };
        byte[] pdf = BuildEmbeddedType1CInspectionPdf(cff);

        PdfPrintProductionStructureEvidence structure = PdfReadDocument.Open(pdf)
            .InspectPrintProductionStructure();

        Assert.Equal(1, structure.FontResourceCount);
        Assert.Equal(1, structure.UnembeddedFontResourceCount);
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
    public void PrintProductionPageBoxesUseTheMediaBoxPrecision() {
        var options = new PdfOptions {
            PageWidth = 595.276D,
            PageHeight = 841.89D
        }.ConfigurePdfXGroundwork(
            PdfComplianceProfile.PdfX4,
            IccMabTestProfiles.CreateCmykLab8Bidirectional(),
            "FOGRA51");

        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Fractional print page."))
            .ToBytes();
        string raw = Encoding.ASCII.GetString(pdf);
        PdfPrintProductionStructureEvidence structure = PdfReadDocument.Open(pdf)
            .InspectPrintProductionStructure();

        Assert.Contains(
            "/MediaBox [0 0 595.276 841.89] /TrimBox [0 0 595.276 841.89] /BleedBox [0 0 595.276 841.89]",
            raw,
            StringComparison.Ordinal);
        Assert.Equal(1, structure.ValidProductionPageBoxCount);
        Assert.Equal(0, structure.InvalidProductionPageBoxCount);
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
    public void PdfXRasterConversionUsesSupportedEmbeddedRgbProfileAndRejectsUnsupportedProfiles() {
        var raster = new OfficeRasterImage(1, 1, OfficeColor.FromRgb(180, 80, 30));
        byte[] untaggedJpeg = OfficeJpegCodec.Encode(raster);
        byte[] taggedRgbJpeg = OfficeJpegCodec.Encode(raster, new OfficeJpegEncodeOptions {
            Metadata = new OfficeJpegMetadata(icc: IccMabTestProfiles.CreateRgbXyz16WithTransformedStages())
        });
        byte[] taggedCmykJpeg = OfficeJpegCodec.Encode(raster, new OfficeJpegEncodeOptions {
            Metadata = new OfficeJpegMetadata(icc: IccMabTestProfiles.CreateCmykLab8Bidirectional())
        });
        byte[] fourComponentJpeg = Convert.FromBase64String(
            "/9j/7gAOQWRvYmUAZAAAAAAA/9sAQwABAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEB/8AAFAgAAQABBEMRAE0RAFkRAEsRAP/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/aAA4EQwBNAFkASwAAPwD+/iv8/wDr/P8A6/v4r//Z");
        byte[] taggedFourComponentJpeg = AddJpegIccProfile(
            fourComponentJpeg,
            IccMabTestProfiles.CreateCmykLab8Bidirectional());
        byte[] taggedGif = AddGifIccApplicationExtension(Convert.FromBase64String(
            "R0lGODlhAQABAJAAAAAAAP///ywAAAAAAQABAAACAkwBADs="));
        var options = new PdfOptions()
            .ConfigurePdfXGroundwork(
                PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(),
                "FOGRA51");
        options.CompressContentStreams = false;

        byte[] pdf = PdfDocument.Create(options).Image(taggedRgbJpeg, 12, 12).ToBytes();
        byte[] untaggedPdf = PdfDocument.Create(options).Image(untaggedJpeg, 12, 12).ToBytes();
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Create(options).Image(taggedCmykJpeg, 12, 12).ToBytes());
        NotSupportedException publicPathException = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Create(options).Image(taggedFourComponentJpeg, 12, 12).ToBytes());
        NotSupportedException taggedGifException = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Create(options).Image(taggedGif, 12, 12).ToBytes());
        byte[] taggedPixels = Assert.Single(PdfImageExtractor.ExtractImages(pdf)).Bytes;
        byte[] untaggedPixels = Assert.Single(PdfImageExtractor.ExtractImages(untaggedPdf)).Bytes;

        Assert.Contains("/ColorSpace /DeviceCMYK", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        Assert.False(taggedPixels.SequenceEqual(untaggedPixels));
        Assert.Contains("embedded ICC profile", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("four-component JPEG", publicPathException.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("embedded ICC profile", taggedGifException.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PdfXRasterConversionRejectsAggregateWorkingSetBeforeOutputAllocation() {
        Assert.True(PdfWriter.IsPdfXImageWorkingSetWithinLimit(
            sourceBytes: 1_024,
            rasterBytes: 4_096,
            profileBytes: 512,
            profileTransformBytes: 1_024,
            cmykBytes: 4_096,
            alphaBytes: 1_024));
        Assert.False(PdfWriter.IsPdfXImageWorkingSetWithinLimit(
            sourceBytes: 1_024,
            rasterBytes: 200_000_000,
            profileBytes: 512,
            profileTransformBytes: 1_024,
            cmykBytes: 200_000_000,
            alphaBytes: 50_000_000));
        Assert.True(PdfWriter.TryGetIccParseAllocationUpperBound(1_024, out long boundedProfileBytes));
        Assert.True(boundedProfileBytes > 1_024);
        Assert.False(PdfWriter.TryGetIccParseAllocationUpperBound(long.MaxValue, out _));
    }

    [Fact]
    public void PdfMetadataResolvesNonzeroGenerationInfoReferences() {
        byte[] matching = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n5 2 obj\n<< /Title (Generation two) /GTS_PDFXVersion (PDF/X-4) /Trapped /False >>\nendobj\n" +
            "trailer\n<< /Note (/Info 9 0 R /Root 7 0 R) % /Info 8 0 R /Root 6 0 R\n /Info 5 2 R /Root 1 3 R >>\n%%EOF\n");
        byte[] mismatched = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n5 2 obj\n<< /Title (Wrong generation) >>\nendobj\n" +
            "trailer\n<< /Info 5 1 R >>\n%%EOF\n");

        PdfMetadata metadata = PdfReadDocument.Open(matching).Metadata;

        Assert.Equal("Generation two", metadata.Title);
        Assert.Equal("PDF/X-4", metadata.PdfXVersion);
        Assert.Equal(PdfTrappingStatus.False, metadata.TrappingStatus);
        Assert.Null(PdfReadDocument.Open(mismatched).Metadata.Title);
        PdfReference infoReference = Assert.IsType<PdfReference>(PdfSyntax.ReadTrailerReference(
            Encoding.ASCII.GetString(matching), "Info"));
        PdfReference rootReference = Assert.IsType<PdfReference>(PdfSyntax.ReadTrailerReference(
            Encoding.ASCII.GetString(matching), "Root"));
        Assert.Equal((5, 2), (infoReference.ObjectNumber, infoReference.Generation));
        Assert.Equal((1, 3), (rootReference.ObjectNumber, rootReference.Generation));
        PdfReference inheritedRoot = Assert.IsType<PdfReference>(PdfSyntax.ReadTrailerReference(
            "trailer\n<< /Prev 42 >>\ntrailer\n<< /Root 4 1 R >>", "Root"));
        Assert.Equal((4, 1), (inheritedRoot.ObjectNumber, inheritedRoot.Generation));
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

    [Theory]
    [InlineData("1e309 0 0 rg 0 0 10 10 re f")]
    [InlineData("BI /W 1e309 /H 1 /BPC 8 /CS /G ID A EI")]
    [InlineData("1e309 /F1 Do")]
    public void PrintProductionInspectorFailsClosedOnMalformedContentOperands(string content) {
        byte[] pdf = BuildInspectionPdf(content);

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void PrintProductionInspectorScopesAliasesToEachPageAndFormResourceDictionary() {
        byte[] pdf = BuildScopedColorSpaceInspectionPdf();

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.True(evidence.HasDeviceRgbUsage);
        Assert.Equal(4, evidence.DeviceRgbOperatorCount);
    }

    [Fact]
    public void PrintProductionInspectorFailsClosedOnUnknownSelectedColorSpace() {
        byte[] pdf = BuildInspectionPdf(
            "/CS1 cs 0.5 scn 0 0 10 10 re f",
            resources: "/ColorSpace << /CS1 /Bogus >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void PrintProductionInspectorIgnoresUnreachableFormStreams() {
        byte[] pdf = BuildInspectionPdf(
            "0 0 0 1 k 0 0 10 10 re f",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length 25 >>\nstream\n" +
                "/DeviceRGB cs 1 0 0 scn f\nendstream\nendobj\n" +
                "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Filter /Unsupported /Length 3 >>\nstream\nabc\nendstream\nendobj\n" +
                "7 0 obj\n<< /Type /Page /MediaBox [0 0 10 10] /Resources << >> /Contents 8 0 R >>\nendobj\n" +
                "8 0 obj\n<< /Length 25 >>\nstream\n/DeviceRGB cs 1 0 0 scn f\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.False(evidence.HasDeviceRgbUsage);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void PrintProductionInspectorInspectsType3CharacterProcedures() {
        byte[] pdf = BuildType3ColorInspectionPdf();

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.True(evidence.HasDeviceRgbUsage);
        Assert.Equal(1, evidence.DeviceRgbOperatorCount);
    }

    [Theory]
    [InlineData("(A) Tj")]
    [InlineData("[(A) 25] TJ")]
    [InlineData("(A) '")]
    [InlineData("0 0 (A) \"")]
    public void PrintProductionInspectorInspectsOnlyPaintedType3CharacterProcedures(string showOperation) {
        byte[] pdf = BuildType3ReachabilityInspectionPdf(showOperation);

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.False(evidence.HasDeviceRgbUsage);
        Assert.Equal(1, evidence.DeviceCmykOperatorCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
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
    public void PrintProductionInspectorBoundsCyclicIndirectObjectGraphs() {
        byte[] pdf = BuildInspectionPdf(
            "/CS1 cs 0.1 0.2 0.3 scn /Blend gs",
            resources: "/ColorSpace << /CS1 6 0 R >> /ExtGState << /Blend 7 0 R >>",
            pageEntries: "/Cycle 9 0 R",
            extraObjects:
                "5 0 obj\n[5 0 R 4 0 R]\nendobj\n" +
                "6 0 obj\n[6 0 R /DeviceRGB]\nendobj\n" +
                "7 0 obj\n<< /BM 8 0 R >>\nendobj\n" +
                "8 0 obj\n[8 0 R /Normal]\nendobj\n" +
                "9 0 obj\n[9 0 R << /S /Transparency >>]\nendobj\n",
            contents: "5 0 R");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Equal(2, evidence.DeviceRgbOperatorCount);
        Assert.Equal(1, evidence.NonOpaqueGraphicsStateCount);
        Assert.Equal(0, evidence.TransparencyGroupCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void PrintProductionInspectorFailsClosedAtConfiguredObjectGraphDepth() {
        byte[] pdf = BuildInspectionPdf(
            string.Empty,
            resources: "/ColorSpace << /CS1 6 0 R >>",
            extraObjects:
                "6 0 obj\n7 0 R\nendobj\n" +
                "7 0 obj\n8 0 R\nendobj\n" +
                "8 0 obj\n9 0 R\nendobj\n" +
                "9 0 obj\n10 0 R\nendobj\n" +
                "10 0 obj\n11 0 R\nendobj\n" +
                "11 0 obj\n12 0 R\nendobj\n" +
                "12 0 obj\n13 0 R\nendobj\n" +
                "13 0 obj\n14 0 R\nendobj\n" +
                "14 0 obj\n15 0 R\nendobj\n" +
                "15 0 obj\n/DeviceRGB\nendobj\n");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxObjectNestingDepth = 8 }
        };
        PdfReadDocument document = PdfReadDocument.Open(pdf, options);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            document.InspectPrintProductionColors());

        Assert.Equal(PdfReadLimitKind.ObjectNestingDepth, exception.Kind);
        Assert.Equal(8, exception.Limit);
        Assert.Equal(9, exception.Actual);
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
    public void XmpMetadataRejectsInvalidUtf8InsteadOfReplacementDecoding() {
        byte[] prefix = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\"?><x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><x:value>");
        byte[] suffix = Encoding.UTF8.GetBytes("</x:value></x:xmpmeta>");
        var xmp = new byte[prefix.Length + 1 + suffix.Length];
        Buffer.BlockCopy(prefix, 0, xmp, 0, prefix.Length);
        xmp[prefix.Length] = 0xFF;
        Buffer.BlockCopy(suffix, 0, xmp, prefix.Length + 1, suffix.Length);

        PdfXmpMetadataInfo? metadata = PdfReadDocument.Open(BuildXmpInspectionPdf(xmp)).XmpMetadata;

        Assert.NotNull(metadata);
        Assert.False(metadata!.IsWellFormedXml);
        Assert.Null(metadata.RawXml);
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

    [Theory]
    [InlineData("D:2026")]
    [InlineData("D:20260824130000")]
    public void PdfDateCodecDoesNotTreatPartialOrTimezoneLessDatesAsProductionPrecise(string value) {
        Assert.NotNull(PdfDateCodec.TryParse(value));
        Assert.False(PdfDateCodec.TryParseProductionDate(value, out _));
    }

    [Fact]
    public void PdfXReadbackRejectsTimezoneLessInfoDatesThatMatchXmpByDefault() {
        byte[] pdf = BuildProductionDateReconciliationPdf(
            "D:20260824130000",
            "D:20260824130500");

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfX4, pdf);

        Assert.Equal(PdfComplianceRequirementStatus.Missing, report.FindRequirement("readback-pdfx-production-dates")!.Status);
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
    public async Task PdfXFormalAsyncPathSaveStopsDuringDeferredLayoutAndPreservesDestination() {
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
        int rowsEnumerated = 0;
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdfx-cancel-" + Guid.NewGuid().ToString("N") + ".pdf");
        byte[] original = { 9, 8, 7, 6 };
        File.WriteAllBytes(path, original);

        try {
            PdfDocument document = PdfDocument.Create(options).TableDeferred(CreateRows, batchSize: 1);

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                document.SaveAsync(path, cancellation.Token));

            Assert.InRange(rowsEnumerated, 2, 3);
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally {
            File.Delete(path);
        }

        IEnumerable<string[]> CreateRows() {
            for (int index = 0; index < 100; index++) {
                rowsEnumerated++;
                if (rowsEnumerated == 2) cancellation.Cancel();
                yield return new[] { "Row " + index.ToString(System.Globalization.CultureInfo.InvariantCulture) };
            }
        }
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

    private static byte[] BuildProductionDateReconciliationPdf(string infoCreationDate, string infoModificationDate) {
        const string xmp =
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\">" +
            "<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
            "<rdf:Description rdf:about=\"\" xmlns:xmp=\"http://ns.adobe.com/xap/1.0/\" " +
            "xmp:CreateDate=\"2026-08-24T13:00:00Z\" " +
            "xmp:ModifyDate=\"2026-08-24T13:05:00Z\" " +
            "xmp:MetadataDate=\"2026-08-24T13:05:00Z\"/>" +
            "</rdf:RDF></x:xmpmeta>";
        byte[] xmpBytes = Encoding.UTF8.GetBytes(xmp);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /Metadata 5 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /Metadata /Subtype /XML /Length " + xmpBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(xmpBytes, 0, xmpBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /CreationDate (" + infoCreationDate + ") /ModDate (" + infoModificationDate + ") >>\nendobj\n");
        WriteAscii(output, "trailer\n<< /Root 1 0 R /Info 6 0 R >>\n%%EOF\n");
        return output.ToArray();
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
        string extraObjects = "",
        string contents = "4 0 R") {
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << " + resources + " >> " + pageEntries + " /Contents " + contents + " >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n" + extraObjects + "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildEmbeddedType1CInspectionPdf(byte[] cff) {
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteInspectionStream(output, 4, string.Empty, "BT /F1 12 Tf (A) Tj ET");
        WriteAscii(output, "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Fake /FontDescriptor 6 0 R >>\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /Type /FontDescriptor /FontName /Fake /FontFile3 7 0 R >>\nendobj\n");
        WriteAscii(output, "7 0 obj\n<< /Subtype /Type1C /Length " + cff.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(cff, 0, cff.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildXmpInspectionPdf(byte[] xmp) {
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /Metadata 4 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << >> >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Type /Metadata /Subtype /XML /Length " + xmp.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(xmp, 0, xmp.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildScopedColorSpaceInspectionPdf() {
        const string rgbContent = "/CS1 cs 0.1 0.2 0.3 scn";
        const string cmykContent = "/CS1 cs 0.1 0.2 0.3 0.4 scn";
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /ColorSpace << /CS1 /DeviceRGB >> /XObject << /Fm1 7 0 R >> >> /Contents 5 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /ColorSpace << /CS1 /DeviceCMYK >> /XObject << /Fm2 8 0 R >> >> /Contents 6 0 R >>\nendobj\n");
        WriteInspectionStream(output, 5, string.Empty, rgbContent + " /Fm1 Do");
        WriteInspectionStream(output, 6, string.Empty, cmykContent + " /Fm2 Do");
        WriteInspectionStream(
            output,
            7,
            "/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /ColorSpace << /CS1 /DeviceRGB >> >>",
            rgbContent);
        WriteInspectionStream(
            output,
            8,
            "/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /ColorSpace << /CS1 /DeviceCMYK >> >>",
            cmykContent);
        WriteAscii(output, "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildType3ColorInspectionPdf() {
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteInspectionStream(output, 4, string.Empty, "BT /F1 12 Tf (A) Tj ET");
        WriteAscii(output, "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj\n");
        WriteInspectionStream(output, 6, string.Empty, "1 0 0 rg 0 0 500 700 re f");
        WriteAscii(output, "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildType3ReachabilityInspectionPdf(string showOperation) {
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteInspectionStream(output, 4, string.Empty, "BT /F1 12 Tf " + showOperation + " ET");
        WriteAscii(output, "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R /B 7 0 R /C 8 0 R >> /Encoding << /Differences [65 /A /B /C] >> /FirstChar 65 /LastChar 67 /Widths [500 500 500] /Resources << >> >>\nendobj\n");
        WriteInspectionStream(output, 6, string.Empty, "0 0 0 1 k 0 0 500 700 re f");
        WriteInspectionStream(output, 7, string.Empty, "1 0 0 rg 0 0 500 700 re f");
        WriteAscii(output, "8 0 obj\n<< /Filter /Unsupported /Length 3 >>\nstream\nabc\nendstream\nendobj\n");
        WriteAscii(output, "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteInspectionStream(MemoryStream output, int objectNumber, string entries, string content) {
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        WriteAscii(
            output,
            objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< " + entries + " /Length " +
            contentBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
    }

    private static byte[] AddJpegIccProfile(byte[] jpeg, byte[] profile) {
        const int maximumPartLength = 65_519;
        int partCount = checked((profile.Length + maximumPartLength - 1) / maximumPartLength);
        if (partCount < 1 || partCount > byte.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(profile));
        }
        using var output = new MemoryStream(checked(jpeg.Length + profile.Length + partCount * 18));
        output.Write(jpeg, 0, 2);
        byte[] prefix = Encoding.ASCII.GetBytes("ICC_PROFILE\0");
        int profileOffset = 0;
        for (int part = 0; part < partCount; part++) {
            int partLength = Math.Min(maximumPartLength, profile.Length - profileOffset);
            int segmentLength = checked(2 + prefix.Length + 2 + partLength);
            output.WriteByte(0xFF);
            output.WriteByte(0xE2);
            output.WriteByte((byte)(segmentLength >> 8));
            output.WriteByte((byte)segmentLength);
            output.Write(prefix, 0, prefix.Length);
            output.WriteByte(checked((byte)(part + 1)));
            output.WriteByte(checked((byte)partCount));
            output.Write(profile, profileOffset, partLength);
            profileOffset += partLength;
        }
        output.Write(jpeg, 2, jpeg.Length - 2);
        return output.ToArray();
    }

    private static byte[] AddGifIccApplicationExtension(byte[] gif) {
        const int globalColorTableEnd = 19;
        byte[] extension = {
            0x21, 0xFF, 0x0B,
            (byte)'I', (byte)'C', (byte)'C', (byte)'R', (byte)'G', (byte)'B',
            (byte)'G', (byte)'1', (byte)'0', (byte)'1', (byte)'2',
            0x03, 0x01, 0x02, 0x03, 0x00
        };
        var result = new byte[gif.Length + extension.Length];
        Buffer.BlockCopy(gif, 0, result, 0, globalColorTableEnd);
        Buffer.BlockCopy(extension, 0, result, globalColorTableEnd, extension.Length);
        Buffer.BlockCopy(gif, globalColorTableEnd, result, globalColorTableEnd + extension.Length, gif.Length - globalColorTableEnd);
        return result;
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
