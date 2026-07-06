using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRedactionVerificationTests {
    [Fact]
    public void AssertVerified_ConfirmsRemovedAndRetainedTextMarkersAfterApply() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();

        Assert.True(proof.Plan.HasMatches);
        Assert.Contains(proof.Plan.Matches, match => match.Text != null && match.Text.Contains("PAY-SECRET-2026", StringComparison.Ordinal));
        Assert.True(proof.Verification.IsVerified);
        Assert.True(proof.Verification.RawPdfBytesChecked);
        Assert.True(proof.Verification.EncodedPdfStringsChecked);
        Assert.True(proof.Verification.DecodedPdfStreamsChecked);
        Assert.Empty(proof.Verification.Issues);
        Assert.DoesNotContain("PAY-SECRET-2026", proof.Verification.ExtractedText, StringComparison.Ordinal);
        Assert.Contains("Visible compliance marker", proof.Verification.ExtractedText, StringComparison.Ordinal);
        Assert.Contains("Public summary marker", proof.Verification.ExtractedText, StringComparison.Ordinal);
    }

    [Fact]
    public void Verify_ReportsRemovedMarkersThatRemainInUnredactedPdf() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(
            proof.Source,
            PdfRedactionProofTestSupport.CreateVerificationOptions());

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, issue => issue.Feature == "RemovedTextMarker" && issue.Marker == "PAY-SECRET-2026");
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RetainedTextMarker");
        Assert.Contains("PDF redaction verification failed", report.Summary, StringComparison.Ordinal);

        var exception = Assert.Throws<InvalidOperationException>(() => report.ThrowIfFailed());
        Assert.Contains("PAY-SECRET-2026", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Verify_ReportsRemovedMarkersThatRemainInPdfHexStringBytes() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithEncodedResidue = AppendPdfHexStringResidue(proof.Redacted, "PAY-SECRET-2026");

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(
            rewrittenWithEncodedResidue,
            PdfRedactionProofTestSupport.CreateVerificationOptions());

        Assert.False(report.IsVerified);
        Assert.True(report.RawPdfBytesChecked);
        Assert.True(report.EncodedPdfStringsChecked);
        Assert.DoesNotContain("PAY-SECRET-2026", report.ExtractedText, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RemovedRawMarker" && issue.Marker == "PAY-SECRET-2026");
        Assert.Contains(report.Issues, issue => issue.Feature == "RemovedEncodedMarker" && issue.Marker == "PAY-SECRET-2026");
    }

    [Fact]
    public void Verify_ReportsRemovedMarkersThatRemainInEscapedPdfLiteralBytes() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithEncodedResidue = AppendPdfLiteralStringResidue(proof.Redacted, "PAY\\055SECRET\\0552026");

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(
            rewrittenWithEncodedResidue,
            PdfRedactionProofTestSupport.CreateVerificationOptions());

        Assert.False(report.IsVerified);
        Assert.DoesNotContain("PAY-SECRET-2026", report.ExtractedText, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RemovedRawMarker" && issue.Marker == "PAY-SECRET-2026");
        Assert.Contains(report.Issues, issue => issue.Feature == "RemovedEncodedMarker" && issue.Marker == "PAY-SECRET-2026");
    }

    [Fact]
    public void Verify_CanSkipEncodedPdfStringResidueChecksWhenRequested() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithEncodedResidue = AppendPdfHexStringResidue(proof.Redacted, "PAY-SECRET-2026");
        PdfRedactionVerificationOptions options = PdfRedactionProofTestSupport.CreateVerificationOptions();
        options.CheckEncodedPdfStrings = false;

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(rewrittenWithEncodedResidue, options);

        Assert.True(report.IsVerified);
        Assert.False(report.EncodedPdfStringsChecked);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RemovedEncodedMarker");
    }

    [Fact]
    public void Verify_ReportsRemovedMarkersThatRemainInDecodedCompressedStreams() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithCompressedResidue = AppendFlateStreamResidue(proof.Redacted, "PAY-SECRET-2026");

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(
            rewrittenWithCompressedResidue,
            PdfRedactionProofTestSupport.CreateVerificationOptions());

        Assert.False(report.IsVerified);
        Assert.True(report.DecodedPdfStreamsChecked);
        Assert.DoesNotContain("PAY-SECRET-2026", report.ExtractedText, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RemovedRawMarker" && issue.Marker == "PAY-SECRET-2026");
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RemovedEncodedMarker" && issue.Marker == "PAY-SECRET-2026");
        Assert.Contains(report.Issues, issue => issue.Feature == "RemovedDecodedStreamMarker" && issue.Marker == "PAY-SECRET-2026");
    }

    [Fact]
    public void Verify_CanSkipDecodedCompressedStreamResidueChecksWhenRequested() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithCompressedResidue = AppendFlateStreamResidue(proof.Redacted, "PAY-SECRET-2026");
        PdfRedactionVerificationOptions options = PdfRedactionProofTestSupport.CreateVerificationOptions();
        options.CheckDecodedPdfStreams = false;

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(rewrittenWithCompressedResidue, options);

        Assert.True(report.IsVerified);
        Assert.False(report.DecodedPdfStreamsChecked);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RemovedDecodedStreamMarker");
    }

    [Fact]
    public void Verify_FailsClosedWhenPdfStreamCannotBeDecoded() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithUndecodableStream = AppendUnsupportedFilteredStream(proof.Redacted);

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(
            rewrittenWithUndecodableStream,
            PdfRedactionProofTestSupport.CreateVerificationOptions());

        Assert.False(report.IsVerified);
        Assert.True(report.DecodedPdfStreamsChecked);
        Assert.DoesNotContain("PAY-SECRET-2026", report.ExtractedText, StringComparison.Ordinal);
        Assert.Contains(report.Issues, issue => issue.Feature == "UndecodablePdfStream" && issue.Marker == "996");
    }

    [Fact]
    public void Verify_CanOptOutOfUndecodableStreamProofFailure() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithUndecodableStream = AppendUnsupportedFilteredStream(proof.Redacted);
        PdfRedactionVerificationOptions options = PdfRedactionProofTestSupport.CreateVerificationOptions();
        options.FailOnUndecodablePdfStreams = false;

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(rewrittenWithUndecodableStream, options);

        Assert.True(report.IsVerified);
        Assert.True(report.DecodedPdfStreamsChecked);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "UndecodablePdfStream");
    }

    [Fact]
    public void Plan_ReportsIntersectingImagePlacementsAsRedactionRisk() {
        byte[] source = BuildImageRedactionPlanningSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement placement = image.PrimaryPlacement!;
        PdfRedactionArea area = CreateImageIntersectionArea(image);

        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area }, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        PdfRedactionMatch match = Assert.Single(plan.Matches, item => item.Kind == PdfRedactionMatchKind.ImagePlacement);
        Assert.Equal(area, match.Area);
        Assert.Equal(placement.PageNumber, match.PageNumber);
        Assert.Equal(placement.X, match.X, 3);
        Assert.Equal(placement.Y, match.Y, 3);
        Assert.Equal(placement.Width, match.Width, 3);
        Assert.Equal(placement.Height, match.Height, 3);
        Assert.Equal(placement.ResourceName, match.ResourceName);
        Assert.Equal(placement.ObjectNumber, match.ObjectNumber);
        Assert.Null(match.Text);
        Assert.Contains(plan.Findings, finding =>
            finding.Code == "RedactionPlanImageIntersection" &&
            finding.Severity == PdfDiagnosticSeverity.Warning &&
            finding.PageNumber == image.PageNumber);
    }

    [Fact]
    public void Apply_FailsClosedWhenRedactionAreaIntersectsImagePlacement() {
        byte[] source = BuildJpegImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfRedactionArea area = CreateImageIntersectionArea(image);

        var exception = Assert.Throws<InvalidOperationException>(() => PdfRedactionApplier.Apply(source, new[] { area }));

        Assert.Contains("intersects image placement", exception.Message, StringComparison.Ordinal);
        Assert.Contains("AllowImagePlacementOverlays", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_RemovesFullyCoveredImagePlacementByDefault() {
        byte[] source = BuildImageRedactionPlanningSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfRedactionArea area = CreateImageCoveringArea(image);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        string text = PdfTextExtractor.ExtractAllText(redacted);
        string raw = PdfEncoding.Latin1GetString(redacted);
        Assert.Contains("Visible image redaction planning marker", text, StringComparison.Ordinal);
        Assert.Contains("Retained text after image", text, StringComparison.Ordinal);
        Assert.Empty(PdfImageExtractor.ExtractImages(redacted));
        Assert.DoesNotContain("/Subtype /Image", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/Im", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_RemovesFullyCoveredNestedFormImagePlacement() {
        byte[] source = BuildNestedFormImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfRedactionArea area = CreateImageCoveringArea(image);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        string raw = PdfEncoding.Latin1GetString(redacted);
        Assert.Empty(PdfImageExtractor.ExtractImages(redacted));
        Assert.DoesNotContain("/Subtype /Image", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/ImNested", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ClonesRepeatedFormInvocationBeforeRemovingNestedImagePlacement() {
        byte[] source = BuildRepeatedFormImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement firstPlacement = image.Placements.OrderBy(placement => placement.X).First();
        PdfRedactionArea area = CreateImageCoveringArea(image, firstPlacement);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        IReadOnlyList<PdfImagePlacement> placements = PdfImageExtractor.ExtractImagePlacements(redacted);
        PdfImagePlacement remaining = Assert.Single(placements);
        Assert.Equal(120D, remaining.X, 3);
        Assert.Single(PdfImageExtractor.ExtractImages(redacted));
    }

    [Fact]
    public void Apply_RewritesPartiallyCoveredSimpleImagePixelsByDefault() {
        byte[] source = BuildSimpleFlateImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement placement = image.PrimaryPlacement!;
        PdfRedactionArea area = CreateImageLeftHalfArea(image, placement);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        IReadOnlyList<PdfImagePlacement> placements = PdfImageExtractor.ExtractImagePlacements(redacted);
        PdfImagePlacement remainingPlacement = Assert.Single(placements);
        Assert.Equal(placement.X, remainingPlacement.X, 3);
        Assert.Equal(placement.Y, remainingPlacement.Y, 3);
        Assert.Single(PdfImageExtractor.ExtractImages(redacted));

        byte[] pixels = DecodeSingleImagePixels(redacted);
        Assert.Equal(24, pixels.Length);
        AssertRedactedLeftHalf(pixels, width: 4, height: 2, components: 3);
    }

    [Fact]
    public void Apply_RewritesPartiallyCoveredSoftMaskedSimpleImagePixelsAndMask() {
        byte[] source = BuildSoftMaskedSimpleFlateImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement placement = image.PrimaryPlacement!;
        PdfRedactionArea area = CreateImageLeftHalfArea(image, placement);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        PdfExtractedImage extracted = Assert.Single(PdfImageExtractor.ExtractImages(redacted));
        Assert.True(extracted.HasTransparencyMask);
        Assert.True(extracted.TransparencyMaskResolved);
        Assert.Equal("soft-mask", extracted.TransparencyMaskKind);
        AssertRedactedLeftHalf(DecodeSingleImagePixels(redacted), width: 4, height: 2, components: 3);
        AssertSoftMaskLeftHalfOpaque(DecodeSoftMaskPixels(redacted, extracted.ObjectNumber), width: 4, height: 2);
    }

    [Fact]
    public void Apply_RewritesPartiallyCoveredDecodeAwareImagePixelsAndMask() {
        byte[] source = BuildDecodeAwareSoftMaskedSimpleFlateImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement placement = image.PrimaryPlacement!;
        PdfRedactionArea area = CreateImageLeftHalfArea(image, placement);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        PdfExtractedImage extracted = Assert.Single(PdfImageExtractor.ExtractImages(redacted));
        Assert.True(extracted.HasTransparencyMask);
        Assert.True(extracted.TransparencyMaskResolved);
        AssertInvertedDecodeRedactedLeftHalf(DecodeSingleImagePixels(redacted), width: 4, height: 2, components: 3);
        AssertInvertedDecodeSoftMaskLeftHalfOpaque(DecodeSoftMaskPixels(redacted, extracted.ObjectNumber), width: 4, height: 2);
    }

    [Fact]
    public void Apply_ClonesRepeatedImageInvocationBeforeRewritingPixels() {
        byte[] source = BuildRepeatedSimpleFlateImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement firstPlacement = image.Placements.OrderBy(placement => placement.X).First();
        PdfRedactionArea area = CreateImageLeftHalfArea(image, firstPlacement);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        string raw = PdfEncoding.Latin1GetString(redacted);
        Assert.Contains("/ImSimpleRedacted1 Do", raw, StringComparison.Ordinal);
        Assert.Contains("/ImSimple Do", raw, StringComparison.Ordinal);
        Assert.Equal(2, PdfImageExtractor.ExtractImagePlacements(redacted).Count);
        byte[][] images = DecodeImagePixelStreams(redacted);
        Assert.Equal(2, images.Length);
        Assert.Contains(images, pixels => PixelRowsMatch(pixels, CreateSimpleFlateImagePixels()));
        Assert.Contains(images, pixels => LeftHalfIsRedacted(pixels, width: 4, height: 2, components: 3));
    }

    [Fact]
    public void Apply_ClonesSharedFormResourceBeforeRemovingNestedImagePlacement() {
        byte[] source = BuildSharedAliasFormImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement firstPlacement = image.Placements.OrderBy(placement => placement.X).First();
        PdfRedactionArea area = CreateImageCoveringArea(image, firstPlacement);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        IReadOnlyList<PdfImagePlacement> placements = PdfImageExtractor.ExtractImagePlacements(redacted);

        PdfImagePlacement remaining = Assert.Single(placements);
        Assert.Equal(120D, remaining.X, 3);
        Assert.Single(PdfImageExtractor.ExtractImages(redacted));
    }

    [Fact]
    public void Apply_PreservesPriorSharedContentArrayImageReplacements() {
        byte[] source = BuildSharedIndirectContentArrayImageRedactionSource();
        IReadOnlyList<PdfImagePlacement> sourcePlacements = PdfImageExtractor.ExtractImagePlacements(source);
        Assert.Equal(4, sourcePlacements.Count);
        var area = new PdfRedactionArea(1, 0, 0, 200, 120, "shared-content-images");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        IReadOnlyList<PdfImagePlacement> placements = PdfImageExtractor.ExtractImagePlacements(redacted);
        Assert.DoesNotContain(placements, placement => placement.PageNumber == 1);
        Assert.Contains(placements, placement => placement.PageNumber == 2);
    }

    [Fact]
    public void Apply_AllowsExplicitImageOverlayWhenWeakerOutcomeIsAccepted() {
        byte[] source = BuildImageRedactionPlanningSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfRedactionArea area = CreateImageIntersectionArea(image);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area }, new PdfRedactionApplyOptions {
            AllowImagePlacementOverlays = true
        });

        Assert.Contains("Retained text after image", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.Single(PdfImageExtractor.ExtractImages(redacted));
    }

    private static byte[] BuildImageRedactionPlanningSource() {
        return PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.Text("Visible image redaction planning marker"))
            .Image(PdfPngTestImages.CreateRgbPng(3, 2), 48, 32, alternativeText: "Sensitive chart pixels")
            .Paragraph(paragraph => paragraph.Text("Retained text after image"))
            .ToBytes();
    }

    private static byte[] BuildNestedFormImageRedactionSource() {
        const string pageContent = "q\n1 0 0 1 100 200 cm\n/Fx Do\nQ\n";
        const string formContent = "q\n10 0 0 10 0 0 cm\n/ImNested Do\nQ\n";
        return BuildNestedImagePdf(pageContent, "<< /Fx 6 0 R >>", formContent, "ImNested");
    }

    private static byte[] BuildRepeatedFormImageRedactionSource() {
        const string pageContent = "q\n1 0 0 1 20 30 cm\n/Fx Do\nQ\nq\n1 0 0 1 120 30 cm\n/Fx Do\nQ\n";
        const string formContent = "q\n10 0 0 10 0 0 cm\n/ImShared Do\nQ\n";
        return BuildNestedImagePdf(pageContent, "<< /Fx 6 0 R >>", formContent, "ImShared");
    }

    private static byte[] BuildSharedAliasFormImageRedactionSource() {
        const string pageContent = "q\n1 0 0 1 20 30 cm\n/FxA Do\nQ\nq\n1 0 0 1 120 30 cm\n/FxB Do\nQ\n";
        const string formContent = "q\n10 0 0 10 0 0 cm\n/ImShared Do\nQ\n";
        return BuildNestedImagePdf(pageContent, "<< /FxA 6 0 R /FxB 6 0 R >>", formContent, "ImShared");
    }

    private static byte[] BuildSharedIndirectContentArrayImageRedactionSource() {
        const string firstContent = "q\n20 0 0 20 20 30 cm\n/ImSharedA Do\nQ\n";
        const string secondContent = "q\n20 0 0 20 80 30 cm\n/ImSharedB Do\nQ\n";
        const string imageBytes = "abc";

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /ImSharedA 9 0 R /ImSharedB 11 0 R >> >> >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 8 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 10 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(firstContent).ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            firstContent,
            "endstream",
            "endobj",
            "6 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(secondContent).ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            secondContent,
            "endstream",
            "endobj",
            "8 0 obj",
            "[5 0 R 6 0 R]",
            "endobj",
            "9 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length " + Encoding.ASCII.GetByteCount(imageBytes).ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            imageBytes,
            "endstream",
            "endobj",
            "10 0 obj",
            "[5 0 R 6 0 R]",
            "endobj",
            "11 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length " + Encoding.ASCII.GetByteCount(imageBytes).ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            imageBytes,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNestedImagePdf(string pageContent, string pageXObjects, string formContent, string imageName) {
        const string imageBytes = "abc";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);
        int imageLength = Encoding.ASCII.GetByteCount(imageBytes);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 4 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /XObject " + pageXObjects + " >>",
            "endobj",
            "5 0 obj",
            "<< /Length " + pageStreamLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Resources << /XObject << /" + imageName + " 7 0 R >> >> /Length " + formStreamLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length " + imageLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            imageBytes,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSimpleFlateImageRedactionSource() {
        const string pageContent = "q\n40 0 0 20 20 30 cm\n/ImSimple Do\nQ\n";
        return BuildSimpleFlateImagePdf(pageContent);
    }

    private static byte[] BuildRepeatedSimpleFlateImageRedactionSource() {
        const string pageContent = "q\n40 0 0 20 20 30 cm\n/ImSimple Do\nQ\nq\n40 0 0 20 80 30 cm\n/ImSimple Do\nQ\n";
        return BuildSimpleFlateImagePdf(pageContent);
    }

    private static byte[] BuildSoftMaskedSimpleFlateImageRedactionSource() {
        const string pageContent = "q\n40 0 0 20 20 30 cm\n/ImSoft Do\nQ\n";
        byte[] pixels = CreateSimpleFlateImagePixels();
        byte[] mask = new byte[] { 64, 64, 128, 128, 192, 192, 224, 224 };
        byte[] compressedPixels = Compress(pixels);
        byte[] compressedMask = Compress(mask);
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent.TrimEnd('\n'));

        using var output = new MemoryStream();
        void WriteAscii(string text) {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
        }

        WriteAscii(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /ImSoft 5 0 R >> >> >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + pageStreamLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 4 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /SMask 6 0 R /Length " + compressedPixels.Length.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream"
        }) + "\n");
        output.Write(compressedPixels, 0, compressedPixels.Length);
        WriteAscii("\nendstream\nendobj\n6 0 obj\n<< /Type /XObject /Subtype /Image /Width 4 /Height 2 /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode /Length " + compressedMask.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(compressedMask, 0, compressedMask.Length);
        WriteAscii("\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildDecodeAwareSoftMaskedSimpleFlateImageRedactionSource() {
        const string pageContent = "q\n40 0 0 20 20 30 cm\n/ImDecoded Do\nQ\n";
        byte[] pixels = CreateSimpleFlateImagePixels();
        byte[] mask = new byte[] { 64, 64, 128, 128, 192, 192, 224, 224 };
        byte[] compressedPixels = Compress(pixels);
        byte[] compressedMask = Compress(mask);
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent.TrimEnd('\n'));

        using var output = new MemoryStream();
        void WriteAscii(string text) {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
        }

        WriteAscii(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /ImDecoded 5 0 R >> >> >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + pageStreamLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 4 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Decode [1 0 1 0 1 0] /Filter /FlateDecode /SMask 6 0 R /Length " + compressedPixels.Length.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream"
        }) + "\n");
        output.Write(compressedPixels, 0, compressedPixels.Length);
        WriteAscii("\nendstream\nendobj\n6 0 obj\n<< /Type /XObject /Subtype /Image /Width 4 /Height 2 /ColorSpace /DeviceGray /BitsPerComponent 8 /Decode [1 0] /Filter /FlateDecode /Length " + compressedMask.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(compressedMask, 0, compressedMask.Length);
        WriteAscii("\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildSimpleFlateImagePdf(string pageContent) {
        byte[] pixels = CreateSimpleFlateImagePixels();
        byte[] compressed = Compress(pixels);
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent.TrimEnd('\n'));

        using var output = new MemoryStream();
        void WriteAscii(string text) {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
        }

        WriteAscii(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /ImSimple 5 0 R >> >> >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + pageStreamLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 4 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length " + compressed.Length.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream"
        }) + "\n");
        output.Write(compressed, 0, compressed.Length);
        WriteAscii("\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildJpegImageRedactionSource() {
        const string pageContent = "q\n20 0 0 20 20 30 cm\n/ImJpeg Do\nQ\n";
        byte[] jpegBytes = new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 };
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent.TrimEnd('\n'));
        using var output = new MemoryStream();
        void WriteAscii(string text) {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
        }

        WriteAscii(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /ImJpeg 5 0 R >> >> >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + pageStreamLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " + jpegBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream"
        }) + "\n");
        output.Write(jpegBytes, 0, jpegBytes.Length);
        WriteAscii("\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] CreateSimpleFlateImagePixels() {
        return new byte[] {
            255, 0, 0, 255, 0, 0, 0, 255, 0, 0, 255, 0,
            0, 0, 255, 0, 0, 255, 255, 255, 255, 255, 255, 255
        };
    }

    private static PdfLogicalImage GetSingleImage(byte[] source) {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(source, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        PdfLogicalImage image = Assert.Single(logical.Images);
        Assert.NotNull(image.PrimaryPlacement);
        return image;
    }

    private static PdfRedactionArea CreateImageIntersectionArea(PdfLogicalImage image) {
        PdfImagePlacement placement = image.PrimaryPlacement!;
        return new PdfRedactionArea(
            image.PageNumber,
            placement.X + 1D,
            placement.Y + 1D,
            Math.Max(1D, placement.Width / 2D),
            Math.Max(1D, placement.Height / 2D),
            "image-risk");
    }

    private static PdfRedactionArea CreateImageLeftHalfArea(PdfLogicalImage image, PdfImagePlacement placement) {
        return new PdfRedactionArea(
            image.PageNumber,
            placement.X,
            placement.Y,
            placement.Width / 2D,
            placement.Height,
            "image-pixel-redact");
    }

    private static PdfRedactionArea CreateImageCoveringArea(PdfLogicalImage image) {
        return CreateImageCoveringArea(image, image.PrimaryPlacement!);
    }

    private static PdfRedactionArea CreateImageCoveringArea(PdfLogicalImage image, PdfImagePlacement placement) {
        return new PdfRedactionArea(
            image.PageNumber,
            placement.X,
            placement.Y,
            placement.Width,
            placement.Height,
            "image-remove");
    }

    private static byte[] AppendPdfHexStringResidue(byte[] pdf, string marker) {
        byte[] suffix = Encoding.ASCII.GetBytes("\n999 0 obj\n<" + ToHex(Encoding.BigEndianUnicode.GetBytes(marker)) + ">\nendobj\n");
        return pdf.Concat(suffix).ToArray();
    }

    private static byte[] AppendPdfLiteralStringResidue(byte[] pdf, string escapedMarker) {
        byte[] suffix = Encoding.ASCII.GetBytes("\n998 0 obj\n(" + escapedMarker + ")\nendobj\n");
        return pdf.Concat(suffix).ToArray();
    }

    private static byte[] AppendFlateStreamResidue(byte[] pdf, string marker) {
        byte[] compressed = Compress(Encoding.UTF8.GetBytes("compressed residue " + marker));
        using var output = new MemoryStream();
        output.Write(pdf, 0, pdf.Length);
        byte[] header = Encoding.ASCII.GetBytes("\n997 0 obj\n<< /Length " + compressed.Length.ToString(CultureInfo.InvariantCulture) + " /Filter /FlateDecode >>\nstream\n");
        output.Write(header, 0, header.Length);
        output.Write(compressed, 0, compressed.Length);
        byte[] footer = Encoding.ASCII.GetBytes("\nendstream\nendobj\n");
        output.Write(footer, 0, footer.Length);
        return output.ToArray();
    }

    private static byte[] AppendUnsupportedFilteredStream(byte[] pdf) {
        byte[] data = new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 };
        using var output = new MemoryStream();
        output.Write(pdf, 0, pdf.Length);
        byte[] header = Encoding.ASCII.GetBytes("\n996 0 obj\n<< /Length " + data.Length.ToString(CultureInfo.InvariantCulture) + " /Filter /DCTDecode >>\nstream\n");
        output.Write(header, 0, header.Length);
        output.Write(data, 0, data.Length);
        byte[] footer = Encoding.ASCII.GetBytes("\nendstream\nendobj\n");
        output.Write(footer, 0, footer.Length);
        return output.ToArray();
    }

    private static byte[] DecodeSingleImagePixels(byte[] pdf) {
        return Assert.Single(DecodeImagePixelStreams(pdf));
    }

    private static byte[][] DecodeImagePixelStreams(byte[] pdf) {
        IReadOnlyList<PdfExtractedImage> images = PdfImageExtractor.ExtractImages(pdf);
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        return images
            .Where(image => image.ObjectNumber > 0)
            .Select(image => objects.TryGetValue(image.ObjectNumber, out PdfIndirectObject? indirect) && indirect.Value is PdfStream stream
                ? StreamDecoder.Decode(stream.Dictionary, stream.Data, objects)
                : Array.Empty<byte>())
            .ToArray();
    }

    private static byte[] DecodeSoftMaskPixels(byte[] pdf, int imageObjectNumber) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        Assert.True(objects.TryGetValue(imageObjectNumber, out PdfIndirectObject? imageIndirect));
        PdfStream imageStream = Assert.IsType<PdfStream>(imageIndirect.Value);
        PdfReference softMaskReference = Assert.IsType<PdfReference>(imageStream.Dictionary.Items["SMask"]);
        Assert.True(objects.TryGetValue(softMaskReference.ObjectNumber, out PdfIndirectObject? softMaskIndirect));
        PdfStream softMaskStream = Assert.IsType<PdfStream>(softMaskIndirect.Value);
        return StreamDecoder.Decode(softMaskStream.Dictionary, softMaskStream.Data, objects);
    }

    private static void AssertRedactedLeftHalf(byte[] pixels, int width, int height, int components) {
        Assert.True(LeftHalfIsRedacted(pixels, width, height, components));
    }

    private static void AssertSoftMaskLeftHalfOpaque(byte[] pixels, int width, int height) {
        Assert.Equal(width * height, pixels.Length);
        for (int row = 0; row < height; row++) {
            for (int column = 0; column < width; column++) {
                byte value = pixels[row * width + column];
                if (column < width / 2) {
                    Assert.Equal((byte)255, value);
                } else {
                    Assert.NotEqual((byte)255, value);
                }
            }
        }
    }

    private static void AssertInvertedDecodeRedactedLeftHalf(byte[] pixels, int width, int height, int components) {
        Assert.Equal(width * height * components, pixels.Length);
        for (int row = 0; row < height; row++) {
            for (int column = 0; column < width; column++) {
                int offset = ((row * width) + column) * components;
                if (column < width / 2) {
                    Assert.Equal((byte)255, pixels[offset]);
                    Assert.Equal((byte)255, pixels[offset + 1]);
                    Assert.Equal((byte)255, pixels[offset + 2]);
                }
            }
        }
    }

    private static void AssertInvertedDecodeSoftMaskLeftHalfOpaque(byte[] pixels, int width, int height) {
        Assert.Equal(width * height, pixels.Length);
        for (int row = 0; row < height; row++) {
            for (int column = 0; column < width; column++) {
                byte value = pixels[row * width + column];
                if (column < width / 2) {
                    Assert.Equal((byte)0, value);
                } else {
                    Assert.NotEqual((byte)0, value);
                }
            }
        }
    }

    private static bool LeftHalfIsRedacted(byte[] pixels, int width, int height, int components) {
        for (int row = 0; row < height; row++) {
            for (int column = 0; column < width; column++) {
                int offset = ((row * width) + column) * components;
                if (column < width / 2) {
                    if (pixels[offset] != 0 || pixels[offset + 1] != 0 || pixels[offset + 2] != 0) {
                        return false;
                    }
                } else {
                    if (pixels[offset] == 0 && pixels[offset + 1] == 0 && pixels[offset + 2] == 0) {
                        return false;
                    }
                }
            }
        }

        return true;
    }

    private static bool PixelRowsMatch(byte[] left, byte[] right) {
        if (left.Length != right.Length) {
            return false;
        }

        for (int i = 0; i < left.Length; i++) {
            if (left[i] != right[i]) {
                return false;
            }
        }

        return true;
    }

    private static byte[] Compress(byte[] bytes) {
        using var output = new MemoryStream();
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(bytes, 0, bytes.Length);
        }

        return output.ToArray();
    }

    private static string ToHex(byte[] bytes) {
        var builder = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            builder.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }
}
