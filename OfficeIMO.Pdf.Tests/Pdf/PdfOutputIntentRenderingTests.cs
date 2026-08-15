using System.Globalization;
using System.Text;
using System.Collections.Concurrent;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOutputIntentRenderingTests {
    [Fact]
    public void RenderPage_AppliesOutputDeviceClassProfileToVectorColor() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16OutputDeviceWithDistinctOutputIntents();
        byte[] pdf = BuildPdf(profileBytes, "0.2 0.4 0.8 rg 10 10 20 20 re f");

        OfficeColor actual = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes).Shape.FillColor!.Value;

        Assert.Equal(
            ExpectedOutputConversion(profileBytes, OfficeColor.FromRgb(51, 102, 204), OfficeIccRenderingIntent.RelativeColorimetric),
            actual);
    }

    [Fact]
    public void RenderPage_ConvertsMatchingDeviceCmykPaintDirectlyThroughOutputProfile() {
        byte[] profileBytes = IccMabTestProfiles.CreateCmykLab8Bidirectional();
        byte[] pdf = BuildPdf(profileBytes, "0.2 0.4 0.6 0.1 k 10 10 20 20 re f", profileEntries: "/N 4");
        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.NotNull(profile);
        Assert.True(profile!.TryConvert(
            new[] { 0.2D, 0.4D, 0.6D, 0.1D },
            OfficeIccRenderingIntent.RelativeColorimetric,
            out OfficeColor expected));

        OfficeColor actual = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes).Shape.FillColor!.Value;

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void OutputIntent_DoesNotTreatIccFallbackAsNativeDeviceCmyk() {
        byte[] profileBytes = IccMabTestProfiles.CreateCmykLab8Bidirectional();
        var profile = new PdfStream(new PdfDictionary(), profileBytes);
        var outputIntent = new PdfDictionary();
        outputIntent.Items["DestOutputProfile"] = new PdfReference(6, 0);
        var outputIntents = new PdfArray();
        outputIntents.Items.Add(outputIntent);
        var catalog = new PdfDictionary();
        catalog.Items["OutputIntents"] = outputIntents;
        var objects = new Dictionary<int, PdfIndirectObject> {
            [6] = new PdfIndirectObject(6, 0, profile)
        };
        PdfOutputIntentColorTransform transform = Assert.IsType<PdfOutputIntentColorTransform>(
            PdfOutputIntentColorTransform.TryCreate(catalog, objects, PdfReadLimits.DefaultMaxDecodedStreamBytes));
        PdfPageColorSpace colorSpace = PdfPageColorSpace.IccFallback(
            PdfPageColorSpaceKind.DeviceCmyk,
            new[] { 1D, 1D, 1D, 1D, 1D, 1D, 1D, 1D });
        double[] components = { 0D, 0D, 0D, 0D };
        Assert.True(colorSpace.TryConvertColor(
            components,
            OfficeIccRenderingIntent.RelativeColorimetric,
            out OfficeColor fallback));

        OfficeColor actual = transform.Apply(
            colorSpace,
            components,
            fallback,
            OfficeIccRenderingIntent.RelativeColorimetric);

        Assert.Equal(transform.Apply(fallback, OfficeIccRenderingIntent.RelativeColorimetric), actual);
    }

    [Fact]
    public void OutputIntent_ConvertsNativeDeviceRgbComponentsDirectly() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        var profile = new PdfStream(new PdfDictionary(), profileBytes);
        var outputIntent = new PdfDictionary();
        outputIntent.Items["DestOutputProfile"] = new PdfReference(6, 0);
        var outputIntents = new PdfArray();
        outputIntents.Items.Add(outputIntent);
        var catalog = new PdfDictionary();
        catalog.Items["OutputIntents"] = outputIntents;
        var objects = new Dictionary<int, PdfIndirectObject> {
            [6] = new PdfIndirectObject(6, 0, profile)
        };
        PdfOutputIntentColorTransform transform = Assert.IsType<PdfOutputIntentColorTransform>(
            PdfOutputIntentColorTransform.TryCreate(catalog, objects, PdfReadLimits.DefaultMaxDecodedStreamBytes));
        double[] components = { 0.2D, 0.4D, 0.6D };
        PdfPageColorSpace colorSpace = PdfPageColorSpaceKind.DeviceRgb;
        Assert.True(colorSpace.TryConvertColor(
            components,
            OfficeIccRenderingIntent.RelativeColorimetric,
            out OfficeColor fallback));
        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? parsedProfile));
        Assert.True(parsedProfile!.TryConvert(
            components,
            OfficeIccRenderingIntent.RelativeColorimetric,
            out OfficeColor expected));

        OfficeColor actual = transform.Apply(
            colorSpace,
            components,
            fallback,
            OfficeIccRenderingIntent.RelativeColorimetric);

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void IccProfile_AcceptsMbaClutWithoutOptionalACurves() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16BidirectionalWithoutOutputCurves();

        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.NotNull(profile);
        Assert.True(profile!.TrySoftProof(OfficeColor.FromRgb(64, 128, 192), out _));
    }

    [Fact]
    public void DrawingEffect_ExplicitRelativeIntentOverridesInheritedPerceptualIntent() {
        PdfPageDrawingEffect inherited = PdfPageDrawingEffect.Default
            .WithRenderingIntent(OfficeIccRenderingIntent.Perceptual);
        PdfPageDrawingEffect local = PdfPageDrawingEffect.Default
            .WithRenderingIntent(OfficeIccRenderingIntent.RelativeColorimetric);

        PdfPageDrawingEffect actual = local.OverlayOn(inherited);

        Assert.Equal(OfficeIccRenderingIntent.RelativeColorimetric, actual.RenderingIntent);
        Assert.True(actual.HasRenderingIntent);
    }

    [Theory]
    [InlineData("Perceptual")]
    [InlineData("RelativeColorimetric")]
    public void RenderPage_AppliesDestinationProfileToDirectVectorColorAfterLateRenderingIntent(string intentName) {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        byte[] pdf = BuildPdf(
            profileBytes,
            "0.2 0.4 0.8 rg /" + intentName + " ri 10 10 20 20 re f",
            profileEntries: string.Empty);

        OfficeColor actual = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes).Shape.FillColor!.Value;

        Assert.Equal(ExpectedOutputConversion(profileBytes, OfficeColor.FromRgb(51, 102, 204), ParseIntent(intentName)), actual);
    }

    [Fact]
    public void RenderPage_AppliesDestinationProfileToImplicitDefaultBlack() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        byte[] pdf = BuildPdf(
            profileBytes,
            "/RelativeColorimetric ri 10 10 20 20 re f");

        OfficeColor actual = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes).Shape.FillColor!.Value;

        Assert.Equal(
            ExpectedSoftProof(profileBytes, OfficeColor.Black, OfficeIccRenderingIntent.RelativeColorimetric),
            actual);
    }

    [Fact]
    public void RenderPage_AppliesDestinationProfileToTextAndInheritedFormPaint() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        const string formContent = "0 0 10 10 re f";
        string resources =
            "/Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> " +
            "/XObject << /Fm 5 0 R >>";
        string extraObjects =
            "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length " +
            Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture) +
            " >>\nstream\n" + formContent + "\nendstream\nendobj\n";
        const string content =
            "0.2 0.4 0.8 rg /Perceptual ri q 1 0 0 1 10 10 cm /Fm Do Q " +
            "BT /F1 12 Tf 10 40 Td (A) Tj ET";
        byte[] pdf = BuildPdf(profileBytes, content, resources, extraObjects);
        OfficeColor expected = ExpectedOutputConversion(
            profileBytes,
            OfficeColor.FromRgb(51, 102, 204),
            OfficeIccRenderingIntent.Perceptual);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Equal(expected, Assert.Single(drawing.Shapes).Shape.FillColor);
        Assert.Equal(expected, Assert.Single(drawing.Elements.OfType<OfficeDrawingText>()).Color);
    }

    [Fact]
    public void RenderPage_DetectsTransparencyInInvokedType3GlyphBeforeSoftProofing() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        const string glyphContent = "500 0 0 0 500 700 d1 /Transparent gs 0.2 0.4 0.8 rg 0 0 500 700 re f";
        string resources = "/Font << /F3 5 0 R >>";
        string extraObjects =
            "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] " +
            "/FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 7 0 R >> " +
            "/Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] " +
            "/Resources << /ExtGState << /Transparent << /ca 0.5 >> >> >> >>\nendobj\n" +
            "7 0 obj\n<< /Length " + Encoding.ASCII.GetByteCount(glyphContent).ToString(CultureInfo.InvariantCulture) +
            " >>\nstream\n" + glyphContent + "\nendstream\nendobj\n";
        byte[] pdf = BuildPdf(
            profileBytes,
            "BT /F3 40 Tf 10 10 Td (A) Tj ET",
            resources,
            extraObjects);
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Equal(OfficeColor.FromRgb(51, 102, 204), FindSingleShapeColor(drawing));
        Assert.Contains(
            page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.OutputIntentTransparencyId);
    }

    [Fact]
    public void ExtractImages_AppliesDestinationProfileToDeviceRgbAndIndexedSamples() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        OfficeColor expected = ExpectedOutputConversion(
            profileBytes,
            OfficeColor.FromRgb(51, 102, 204),
            OfficeIccRenderingIntent.Perceptual);
        byte[] directPdf = BuildImagePdf(
            profileBytes,
            "/DeviceRGB",
            new byte[] { 51, 102, 204 },
            "/Intent /Perceptual");
        byte[] indexedPdf = BuildImagePdf(
            profileBytes,
            "[/Indexed /DeviceRGB 1 <0000003366CC>]",
            new byte[] { 1 },
            "/Intent /Perceptual");

        Assert.Equal(expected, ReadSingleExtractedPixel(directPdf));
        Assert.Equal(expected, ReadSingleExtractedPixel(indexedPdf));
    }

    [Fact]
    public void ExtractImages_NormalizesDctPayloadWhenDestinationProfileRequiresSoftProofing() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        byte[] jpeg = OfficeJpegCodec.Encode(
            OfficeRasterImage.FromRgba32(1, 1, new byte[] { 51, 102, 204, 255 }),
            new OfficeJpegEncodeOptions { Quality = 100, Subsampling = OfficeJpegSubsampling.Y444 });
        Assert.True(OfficeJpegCodec.TryDecode(jpeg, out OfficeRasterImage? decoded));
        OfficeColor expected = ExpectedOutputConversion(
            profileBytes,
            decoded!.GetPixel(0, 0),
            OfficeIccRenderingIntent.RelativeColorimetric);
        byte[] pdf = BuildImagePdf(
            profileBytes,
            "/DeviceRGB",
            jpeg,
            "/Filter /DCTDecode /Intent /RelativeColorimetric");

        Assert.Equal(expected, ReadSingleExtractedPixel(pdf));
    }

    [Fact]
    public void RenderPage_AppliesDestinationProfileToShadingStopsAndPatternTiles() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        OfficeColor expectedBlue = ExpectedOutputConversion(
            profileBytes,
            OfficeColor.FromRgb(51, 102, 204),
            OfficeIccRenderingIntent.RelativeColorimetric);
        OfficeColor expectedRed = ExpectedOutputConversion(
            profileBytes,
            OfficeColor.Red,
            OfficeIccRenderingIntent.RelativeColorimetric);
        const string tileContent = "0.2 0.4 0.8 rg 0 0 10 10 re f";
        string resources = "/Shading << /Sh 5 0 R >> /Pattern << /P 8 0 R >>";
        string extraObjects =
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 100 0] /Function 7 0 R >>\nendobj\n" +
            "7 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0.2 0.4 0.8] /C1 [1 0 0] /N 1 >>\nendobj\n" +
            "8 0 obj\n<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << >> /Length " +
            Encoding.ASCII.GetByteCount(tileContent).ToString(CultureInfo.InvariantCulture) +
            " >>\nstream\n" + tileContent + "\nendstream\nendobj\n";
        byte[] pdf = BuildPdf(
            profileBytes,
            "/RelativeColorimetric ri /Sh sh /Pattern cs /P scn 0 20 20 20 re f",
            resources,
            extraObjects);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeLinearGradient gradient = Assert.Single(drawing.Shapes).Shape.FillGradient!;
        OfficeColor[] stopColors = gradient.Stops.Select(stop => stop.Color).Distinct().ToArray();
        OfficeDrawing patternSurface = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>()).Drawing;
        OfficeDrawing tile = Assert.Single(patternSurface.Elements.OfType<OfficeDrawingTilingPattern>()).Tile;

        Assert.Contains(expectedBlue, stopColors);
        Assert.Contains(expectedRed, stopColors);
        Assert.Equal(expectedBlue, Assert.Single(tile.Shapes).Shape.FillColor);
    }

    [Fact]
    public void ExportImage_ReportsUnsupportedDestinationProfileAndPreservesAuthoredColor() {
        byte[] inputOnlyProfile = IccMabTestProfiles.CreateRgbXyzBOnly();
        byte[] pdf = BuildPdf(inputOnlyProfile, "0.2 0.4 0.8 rg 10 10 20 20 re f");
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeImageExportResult export = page.ExportImage(OfficeImageExportFormat.Png);

        Assert.Equal(OfficeColor.FromRgb(51, 102, 204), Assert.Single(drawing.Shapes).Shape.FillColor);
        Assert.Contains(export.Diagnostics, diagnostic =>
            diagnostic.Code == PdfRenderCapabilities.UnsupportedIccOutputIntentId);
    }

    [Fact]
    public void Open_DoesNotDecodeOversizedDestinationProfileUntilRenderingNeedsIt() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16BidirectionalWithTransformedOutput();
        byte[] compressedProfile = OfficeZlibCodec.Compress(profileBytes);
        byte[] pdf = BuildPdf(
            compressedProfile,
            "0.2 0.4 0.8 rg 10 10 20 20 re f",
            profileEntries: "/Filter /FlateDecode");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = profileBytes.Length - 1 }
        };

        PdfReadDocument document = PdfReadDocument.Open(pdf, options);

        PdfOutputIntentInfo intent = Assert.Single(document.OutputIntents);
        Assert.True(intent.HasDestinationOutputProfile);
        Assert.Equal(6, intent.DestinationOutputProfileObjectNumber);
        Assert.Throws<PdfReadLimitException>(() => _ = intent.DestinationOutputProfileSizeBytes);
        Assert.Throws<PdfReadLimitException>(() => PdfPageImageRenderer.RenderPage(document));
    }

    [Fact]
    public void IccProfileCache_SharesDecodedBytesAcrossConcurrentMetadataAndProfileReads() {
        byte[] profileBytes = PdfIccProfiles.SrgbIec6196621;
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        var stream = new PdfStream(dictionary, OfficeZlibCodec.Compress(profileBytes));
        var objects = new Dictionary<int, PdfIndirectObject>();
        var decoded = new ConcurrentBag<byte[]>();

        Parallel.For(0, 32, index => {
            if ((index & 1) == 0) {
                Assert.True(PdfIccProfileCache.TryReadBytes(stream, objects, profileBytes.Length, out byte[] bytes));
                decoded.Add(bytes);
            } else {
                Assert.True(PdfIccProfileCache.TryRead(stream, objects, profileBytes.Length, out OfficeIccColorProfile? profile));
                Assert.NotNull(profile);
            }
        });

        byte[] first = Assert.Single(decoded.Take(1));
        Assert.All(decoded, bytes => Assert.Same(first, bytes));
    }

    [Fact]
    public void RenderPage_PublishesDestinationProfileSafelyAcrossConcurrentFirstUse() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        byte[] pdf = BuildPdf(profileBytes, "0.2 0.4 0.8 rg 10 10 20 20 re f");
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        OfficeColor expected = ExpectedOutputConversion(
            profileBytes,
            OfficeColor.FromRgb(51, 102, 204),
            OfficeIccRenderingIntent.RelativeColorimetric);
        var colors = new ConcurrentBag<OfficeColor>();

        Parallel.For(0, 32, _ =>
            colors.Add(Assert.Single(PdfPageImageRenderer.RenderPage(document).Shapes).Shape.FillColor!.Value));

        Assert.Equal(32, colors.Count);
        Assert.All(colors, color => Assert.Equal(expected, color));
    }

    [Fact]
    public void RenderPage_TreatsNullAndEmptyOutputIntentDeclarationsAsAbsent() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        byte[] nullPdf = BuildPdf(
            profileBytes,
            "0.2 0.4 0.8 rg 10 10 20 20 re f",
            outputIntents: "null");
        byte[] emptyPdf = BuildPdf(
            profileBytes,
            "0.2 0.4 0.8 rg 10 10 20 20 re f",
            outputIntents: "[]");

        Assert.Equal(OfficeColor.FromRgb(51, 102, 204), Assert.Single(PdfPageImageRenderer.RenderPage(nullPdf).Shapes).Shape.FillColor);
        Assert.Equal(OfficeColor.FromRgb(51, 102, 204), Assert.Single(PdfPageImageRenderer.RenderPage(emptyPdf).Shapes).Shape.FillColor);
        Assert.DoesNotContain(PdfReadDocument.Open(nullPdf).Pages[0].GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code.Contains("output-intent", StringComparison.Ordinal));

        byte[] indirectNullPdf = BuildPdf(
            profileBytes,
            "0.2 0.4 0.8 rg 10 10 20 20 re f",
            extraObjects: "7 0 obj\nnull\nendobj\n",
            outputIntents: "7 0 R");
        Assert.Equal(OfficeColor.FromRgb(51, 102, 204), Assert.Single(PdfPageImageRenderer.RenderPage(indirectNullPdf).Shapes).Shape.FillColor);
    }

    [Theory]
    [InlineData("null", "")]
    [InlineData("7 0 R", "7 0 obj\nnull\nendobj\n")]
    public void RenderPage_DefaultsOptionalNullProfileComponentCount(string nValue, string extraObjects) {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        byte[] pdf = BuildPdf(
            profileBytes,
            "0.2 0.4 0.8 rg 10 10 20 20 re f",
            extraObjects: extraObjects,
            profileEntries: "/N " + nValue);

        OfficeColor actual = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes).Shape.FillColor!.Value;

        Assert.Equal(ExpectedOutputConversion(
            profileBytes,
            OfficeColor.FromRgb(51, 102, 204),
            OfficeIccRenderingIntent.RelativeColorimetric), actual);
    }

    [Fact]
    public void RenderPage_SkipsNullProfileAndUsesLaterValidOutputIntent() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        byte[] pdf = BuildPdf(
            profileBytes,
            "0.2 0.4 0.8 rg 10 10 20 20 re f",
            outputIntents: "[<< /Type /OutputIntent /DestOutputProfile null >> << /Type /OutputIntent /DestOutputProfile 6 0 R >>]");

        OfficeColor actual = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes).Shape.FillColor!.Value;

        Assert.Equal(ExpectedOutputConversion(
            profileBytes,
            OfficeColor.FromRgb(51, 102, 204),
            OfficeIccRenderingIntent.RelativeColorimetric), actual);

        byte[] indirectNullPdf = BuildPdf(
            profileBytes,
            "0.2 0.4 0.8 rg 10 10 20 20 re f",
            extraObjects: "7 0 obj\nnull\nendobj\n",
            outputIntents: "[<< /Type /OutputIntent /DestOutputProfile 7 0 R >> << /Type /OutputIntent /DestOutputProfile 6 0 R >>]");
        Assert.Equal(
            actual,
            Assert.Single(PdfPageImageRenderer.RenderPage(indirectNullPdf).Shapes).Shape.FillColor);
    }

    [Fact]
    public void RenderPage_TreatsNullOnlyProfileCandidatesAsAbsent() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        byte[] pdf = BuildPdf(
            profileBytes,
            "0.2 0.4 0.8 rg 10 10 20 20 re f",
            outputIntents: "[<< /Type /OutputIntent /DestOutputProfile null >>]");
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        Assert.Equal(OfficeColor.FromRgb(51, 102, 204), Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes).Shape.FillColor);
        Assert.DoesNotContain(page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code.Contains("output-intent", StringComparison.Ordinal));
    }

    [Fact]
    public void RenderPage_FailsClosedForCyclicProfileComponentCount() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        byte[] pdf = BuildPdf(
            profileBytes,
            "0.2 0.4 0.8 rg 10 10 20 20 re f",
            extraObjects: "7 0 obj\n8 0 R\nendobj\n8 0 obj\n7 0 R\nendobj\n",
            profileEntries: "/N 7 0 R");
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        Assert.Equal(OfficeColor.FromRgb(51, 102, 204), Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes).Shape.FillColor);
        Assert.Contains(page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedIccOutputIntentId);
    }

    [Fact]
    public void RenderPage_UnmatchedRestorePreservesProofedDefaultAndInheritedFormPaint() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        const string formContent = "Q 0 0 10 10 re f";
        string resources = "/XObject << /Fm 5 0 R >>";
        string extraObjects =
            "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length " +
            Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture) +
            " >>\nstream\n" + formContent + "\nendstream\nendobj\n";
        byte[] pdf = BuildPdf(
            profileBytes,
            "Q 0 0 10 10 re f 0.2 0.4 0.8 rg q 1 0 0 1 20 0 cm /Fm Do Q",
            resources,
            extraObjects);
        OfficeColor expectedBlack = ExpectedSoftProof(
            profileBytes,
            OfficeColor.Black,
            OfficeIccRenderingIntent.RelativeColorimetric);
        OfficeColor expectedBlue = ExpectedOutputConversion(
            profileBytes,
            OfficeColor.FromRgb(51, 102, 204),
            OfficeIccRenderingIntent.RelativeColorimetric);

        OfficeColor[] colors = PdfPageImageRenderer.RenderPage(pdf).Shapes.Select(shape => shape.Shape.FillColor!.Value).ToArray();

        Assert.Contains(expectedBlack, colors);
        Assert.Contains(expectedBlue, colors);
    }

    [Fact]
    public void RenderPage_DiagnosesTransparencyAndPreservesAuthoredColorsBeforeComposition() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        byte[] pdf = BuildPdf(
            profileBytes,
            "0.2 0.4 0.8 rg /GS gs 10 10 20 20 re f",
            resources: "/ExtGState << /GS << /Type /ExtGState /BM [/Multiply /Normal] >> >>");
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Equal(OfficeColor.FromRgb(51, 102, 204), FindSingleShapeColor(drawing));
        Assert.Contains(page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.OutputIntentTransparencyId);
    }

    [Fact]
    public void RenderPage_DiagnosesPageTransparencyGroupBeforeSoftProofing() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        byte[] pdf = BuildPdf(
            profileBytes,
            "0.2 0.4 0.8 rg 10 10 20 20 re f",
            pageEntries: "/Group << /S /Transparency /CS /DeviceRGB >>");
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        OfficeColor color = FindSingleShapeColor(PdfPageImageRenderer.RenderPage(pdf));

        Assert.Equal(OfficeColor.FromRgb(51, 102, 204), color);
        Assert.Contains(page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.OutputIntentTransparencyId);
    }

    [Fact]
    public void RenderPage_ResolvesIndirectTilingPatternTypeBeforeTransparencyInspection() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        const string patternContent = "/GS gs 0 0 10 10 re f";
        string resources = "/Pattern << /P 5 0 R >>";
        string extraObjects =
            "5 0 obj\n<< /Type /Pattern /PatternType 7 0 R /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 " +
            "/Resources << /ExtGState << /GS << /Type /ExtGState /ca 0.5 >> >> >> /Length " +
            Encoding.ASCII.GetByteCount(patternContent).ToString(CultureInfo.InvariantCulture) +
            " >>\nstream\n" + patternContent + "\nendstream\nendobj\n" +
            "7 0 obj\n1\nendobj\n";
        byte[] pdf = BuildPdf(
            profileBytes,
            "/Pattern cs /P scn 10 10 20 20 re f",
            resources,
            extraObjects);
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        _ = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.OutputIntentTransparencyId);
    }

    [Fact]
    public void RenderPage_UnmatchedRestorePreservesProofedDefaultTextColor() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        string resources = "/Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >>";
        byte[] pdf = BuildPdf(profileBytes, "Q BT /F1 12 Tf 10 40 Td (A) Tj ET", resources);
        OfficeColor expected = ExpectedSoftProof(
            profileBytes,
            OfficeColor.Black,
            OfficeIccRenderingIntent.RelativeColorimetric);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Equal(expected, Assert.Single(drawing.Elements.OfType<OfficeDrawingText>()).Color);
    }

    [Fact]
    public void RenderPage_DiagnosesLuminositySoftMaskInsteadOfProofingItsBackdropEarly() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        const string maskContent = "1 g 0 0 20 20 re f";
        string resources = "/ExtGState << /GS << /Type /ExtGState /SMask 7 0 R >> >>";
        string extraObjects =
            "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Group << /S /Transparency /CS /DeviceRGB >> /Length " +
            Encoding.ASCII.GetByteCount(maskContent).ToString(CultureInfo.InvariantCulture) +
            " >>\nstream\n" + maskContent + "\nendstream\nendobj\n" +
            "7 0 obj\n<< /S /Luminosity /G 5 0 R /BC [0.1 0.2 0.3] >>\nendobj\n";
        byte[] pdf = BuildPdf(
            profileBytes,
            "0.2 0.4 0.8 rg /GS gs 10 10 20 20 re f",
            resources,
            extraObjects);
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Equal(OfficeColor.FromRgb(51, 102, 204), FindSingleShapeColor(drawing));
        Assert.Contains(page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.OutputIntentTransparencyId);
    }

    [Fact]
    public void RenderPage_DiagnosesStencilImageMaskBeforeApplyingOutputIntent() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        string resources = "/XObject << /Im 5 0 R >>";
        string extraObjects =
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Length 1 >>\nstream\n@\nendstream\nendobj\n";
        byte[] pdf = BuildPdf(
            profileBytes,
            "0.2 0.4 0.8 rg q 20 0 0 20 10 10 cm /Im Do Q",
            resources,
            extraObjects);
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        _ = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.OutputIntentTransparencyId);
    }

    [Fact]
    public void RenderDiagnostics_RejectsDctThatForcedOutputNormalizationCannotDecode() {
        byte[] profileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        byte[] jpeg = AddAdobeTransform(
            OfficeJpegCodec.Encode(
                OfficeRasterImage.FromRgba32(1, 1, new byte[] { 255, 0, 0, 255 }),
                new OfficeJpegEncodeOptions { Quality = 100, Subsampling = OfficeJpegSubsampling.Y444 }),
            2);
        byte[] pdf = BuildImagePdf(profileBytes, "/DeviceRGB", jpeg, "/Filter /DCTDecode");
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        Assert.False(Assert.Single(PdfImageExtractor.ExtractImages(pdf)).IsImageFile);
        Assert.Contains(page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    private static OfficeColor ExpectedOutputConversion(
        byte[] profileBytes,
        OfficeColor source,
        OfficeIccRenderingIntent intent) {
        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.NotNull(profile);
        Assert.True(profile!.TryConvert(
            new[] { source.R / 255D, source.G / 255D, source.B / 255D },
            intent,
            out OfficeColor expected));
        return expected;
    }

    private static OfficeColor ExpectedSoftProof(
        byte[] profileBytes,
        OfficeColor source,
        OfficeIccRenderingIntent intent) {
        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        Assert.NotNull(profile);
        Assert.True(profile!.TrySoftProof(source, intent, out OfficeColor expected));
        return expected;
    }

    private static OfficeIccRenderingIntent ParseIntent(string name) =>
        name == "Perceptual"
            ? OfficeIccRenderingIntent.Perceptual
            : OfficeIccRenderingIntent.RelativeColorimetric;

    private static OfficeColor ReadSingleExtractedPixel(byte[] pdf) {
        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        Assert.True(image.IsImageFile);
        Assert.Equal("png", image.FileExtension);
        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        return raster!.GetPixel(0, 0);
    }

    private static byte[] BuildImagePdf(
        byte[] profileBytes,
        string imageColorSpace,
        byte[] imageSamples,
        string imageEntries) {
        byte[] contentBytes = Encoding.ASCII.GetBytes("q 20 0 0 20 10 10 cm /Im Do Q");
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OutputIntents [<< /Type /OutputIntent /S /GTS_PDFA1 /DestOutputProfile 6 0 R >>] >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Im 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(
            output,
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /BitsPerComponent 8 /ColorSpace " +
            imageColorSpace + " " + imageEntries + " /Length " +
            imageSamples.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(imageSamples, 0, imageSamples.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /Length " + profileBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(profileBytes, 0, profileBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildPdf(
        byte[] profileBytes,
        string content,
        string resources = "",
        string extraObjects = "",
        string profileEntries = "",
        string? outputIntents = null,
        string pageEntries = "") {
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OutputIntents " +
            (outputIntents ?? "[<< /Type /OutputIntent /S /GTS_PDFA1 /DestOutputProfile 6 0 R >>]") +
            " >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << " + resources + " >> " + pageEntries + " /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, extraObjects);
        WriteAscii(output, "6 0 obj\n<< " + profileEntries + " /Length " + profileBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(profileBytes, 0, profileBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static OfficeColor FindSingleShapeColor(OfficeDrawing drawing) {
        var colors = new List<OfficeColor>();
        Collect(drawing, colors);
        return Assert.Single(colors);

        static void Collect(OfficeDrawing current, List<OfficeColor> colors) {
            colors.AddRange(current.Shapes.Where(shape => shape.Shape.FillColor.HasValue).Select(shape => shape.Shape.FillColor!.Value));
            foreach (OfficeDrawingGroup group in current.Elements.OfType<OfficeDrawingGroup>()) Collect(group.Drawing, colors);
            foreach (OfficeDrawingEffectGroup group in current.Elements.OfType<OfficeDrawingEffectGroup>()) Collect(group.InnerDrawing, colors);
        }
    }

    private static byte[] AddAdobeTransform(byte[] jpeg, byte transform) {
        byte[] marker = {
            0xFF, 0xEE, 0x00, 0x0E,
            (byte)'A', (byte)'d', (byte)'o', (byte)'b', (byte)'e',
            0x00, 0x64,
            0x00, 0x00,
            0x00, 0x00,
            transform
        };
        var result = new byte[jpeg.Length + marker.Length];
        Buffer.BlockCopy(jpeg, 0, result, 0, 2);
        Buffer.BlockCopy(marker, 0, result, 2, marker.Length);
        Buffer.BlockCopy(jpeg, 2, result, 2 + marker.Length, jpeg.Length - 2);
        return result;
    }
}
