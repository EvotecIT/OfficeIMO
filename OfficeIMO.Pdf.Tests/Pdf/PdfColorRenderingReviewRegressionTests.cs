using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfColorRenderingReviewRegressionTests {
    [Fact]
    public void JpegDecoder_ComplementsAdobeYcckColorantsBeforeCmykConversion() {
        byte[] jpeg = Convert.FromBase64String(
            "/9j/7gAOQWRvYmUAZAAAAAAC/9sAQwABAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEB/8AAFAgAAQABBEMRAE0RAFkRAEsRAP/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/aAA4EQwBNAFkASwAAPwD+/iq9V6/v4r//2Q==");

        Assert.True(OfficeJpegCodec.TryDecodeColorComponents(
            jpeg,
            requestedColorTransform: null,
            usePdfColorTransformDefault: true,
            out byte[] components,
            out int width,
            out int height,
            out int componentCount));
        Assert.Equal(1, width);
        Assert.Equal(1, height);
        Assert.Equal(4, componentCount);
        Assert.All(components, component => Assert.InRange(component, 0, 5));

        Assert.True(OfficeJpegCodec.TryDecode(jpeg, out OfficeRasterImage? raster));
        OfficeColor pixel = raster!.GetPixel(0, 0);
        Assert.InRange(pixel.R, 250, 255);
        Assert.InRange(pixel.G, 250, 255);
        Assert.InRange(pixel.B, 250, 255);
    }

    [Fact]
    public void GraphicsEffectTimeline_FramesNamedIccInlineImageBeforeReadingIntentTransitions() {
        string payload = new string('x', 10) + " /AbsoluteColorimetric ri ";
        payload = payload.PadRight(30, 'x');
        string content =
            "BI /W 10 /H 1 /BPC 8 /CS /CsIcc ID " + payload +
            " EI /Perceptual ri";

        PdfPageDrawingEffectTransition transition = Assert.Single(
            PdfPageGraphicsEffectTimelineParser.Parse(
                content,
                graphicsStates: null,
                initialEffect: PdfPageDrawingEffect.Default,
                initialTransform: Matrix2D.Identity,
                inlineImageComponentCount: static name => name == "CsIcc" ? 3 : 1));

        Assert.Equal(OfficeIccRenderingIntent.Perceptual, transition.Effect.RenderingIntent);
    }

    [Fact]
    public void RenderPage_FramesNamedIccInlineImageBeforeOutputIntentCompositionScan() {
        byte[] outputProfile = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        string payload = (new string('x', 10) + " /Alpha gs ").PadRight(30, 'x');
        byte[] pdf = BuildNamedInlineImageOutputIntentPdf(outputProfile, payload);
        Assert.True(OfficeIccColorProfile.TryCreate(outputProfile, out OfficeIccColorProfile? profile));
        Assert.True(profile!.TryConvert(
            new[] { 0.2D, 0.4D, 0.8D },
            OfficeIccRenderingIntent.RelativeColorimetric,
            out OfficeColor expected));

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Equal(expected, Assert.Single(drawing.Shapes).Shape.FillColor);
    }

    [Fact]
    public void SampledShading_RejectsIncompleteExactKnotSets() {
        const int sampleCount = 129;
        var dictionary = new PdfDictionary();
        dictionary.Items["FunctionType"] = new PdfNumber(0);
        dictionary.Items["Domain"] = NumberArray(0, 1);
        dictionary.Items["Range"] = NumberArray(0, 1);
        dictionary.Items["Size"] = NumberArray(sampleCount);
        dictionary.Items["BitsPerSample"] = new PdfNumber(8);
        var function = new PdfStream(dictionary, new byte[sampleCount]);
        var objects = new Dictionary<int, PdfIndirectObject>();

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            function,
            1,
            1,
            objects,
            4096,
            out _));
        Assert.False(PdfColorSpaceFunctionResolver.TryCreateShadingFunction(
            function,
            1,
            objects,
            4096,
            out _));
    }

    [Fact]
    public void ImageIccProfiles_ConsumeOneAggregateCallerOwnedRetentionBudget() {
        byte[] profileBytes = PdfIccProfiles.SrgbIec6196621;
        PdfStream first = CreateIccProfileStream(profileBytes);
        PdfStream second = CreateIccProfileStream(profileBytes);
        PdfArray firstColorSpace = IccColorSpace(first);
        PdfArray secondColorSpace = IccColorSpace(second);
        var objects = new Dictionary<int, PdfIndirectObject>();
        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? parsedProfile));
        long retainedProfileBytes = Math.Max(profileBytes.LongLength, parsedProfile!.RetainedByteCount);
        int retainedLimit = checked((int)(retainedProfileBytes * 2L - 1L));
        var context = new PdfColorFunctionResolutionContext(retainedLimit);

        Assert.True(TryResolve(firstColorSpace, objects, context, out PdfImageColorSpaceNormalization firstNormalization));
        Assert.True(TryResolve(firstColorSpace, objects, context, out PdfImageColorSpaceNormalization aliasNormalization));
        Assert.Equal(firstNormalization.Kind, aliasNormalization.Kind);
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            TryResolve(secondColorSpace, objects, context, out _));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(retainedLimit, exception.Limit);
        Assert.Equal(retainedProfileBytes * 2L, exception.Actual);
    }

    [Fact]
    public void ImageIccProfiles_ChargeExpandedSharedSampledCurvesToRetentionBudget() {
        byte[] profileBytes = CreateRgbProfileWithSharedSampledCurves(sampleCount: 4096);
        PdfStream profileStream = CreateIccProfileStream(profileBytes);
        PdfArray colorSpace = IccColorSpace(profileStream);
        int retainedLimit = checked(profileBytes.Length * 2);
        var context = new PdfColorFunctionResolutionContext(retainedLimit);
        Assert.True(PdfIccProfileCache.TryReadBytes(
            profileStream,
            new Dictionary<int, PdfIndirectObject>(),
            retainedLimit,
            context.IccProfileRetentionBudget,
            out byte[] cachedBytes));
        Assert.Equal(profileBytes.Length, cachedBytes.Length);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            TryResolve(colorSpace, new Dictionary<int, PdfIndirectObject>(), context, out _));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(retainedLimit, exception.Limit);
        Assert.True(exception.Actual > retainedLimit);
    }

    [Fact]
    public void IccProfileCache_ChargesParsedProfileAndDecodedBytesCumulatively() {
        byte[] profileBytes = CreateRgbProfileWithSharedSampledCurves(sampleCount: 4096);
        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile));
        PdfStream profileStream = CreateIccProfileStream(profileBytes);
        long retainedProfileBytes = Math.Max(profileBytes.LongLength, profile!.RetainedByteCount);
        int retainedLimit = checked((int)(retainedProfileBytes + profileBytes.LongLength - 1L));
        var budget = new PdfIccProfileRetentionBudget(retainedLimit);
        var objects = new Dictionary<int, PdfIndirectObject>();

        Assert.True(PdfIccProfileCache.TryRead(
            profileStream,
            objects,
            retainedLimit,
            budget,
            out OfficeIccColorProfile? cachedProfile));
        Assert.NotNull(cachedProfile);
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfIccProfileCache.TryReadBytes(profileStream, objects, retainedLimit, budget, out _));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(retainedLimit, exception.Limit);
        Assert.Equal(retainedProfileBytes + profileBytes.LongLength, exception.Actual);
    }

    [Fact]
    public void SeparationImageDirectOutputConversion_PreservesAuthoredRenderingIntent() {
        byte[] outputProfileBytes = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        var outputProfileStream = new PdfStream(new PdfDictionary(), outputProfileBytes);
        var outputIntent = new PdfDictionary();
        outputIntent.Items["DestOutputProfile"] = outputProfileStream;
        var outputIntents = new PdfArray();
        outputIntents.Items.Add(outputIntent);
        var catalog = new PdfDictionary();
        catalog.Items["OutputIntents"] = outputIntents;
        PdfOutputIntentColorTransform transform = Assert.IsType<PdfOutputIntentColorTransform>(
            PdfOutputIntentColorTransform.TryCreate(
                catalog,
                new Dictionary<int, PdfIndirectObject>(),
                PdfReadLimits.DefaultMaxDecodedStreamBytes));

        var tintTransform = new PdfDictionary();
        tintTransform.Items["FunctionType"] = new PdfNumber(2);
        tintTransform.Items["Domain"] = NumberArray(0, 1);
        tintTransform.Items["C0"] = NumberArray(0.1, 0.2, 0.3);
        tintTransform.Items["C1"] = NumberArray(0.6, 0.7, 0.8);
        tintTransform.Items["N"] = new PdfNumber(1);
        var colorSpace = new PdfArray();
        colorSpace.Items.Add(new PdfName("Separation"));
        colorSpace.Items.Add(new PdfName("Spot"));
        colorSpace.Items.Add(new PdfName("DeviceRGB"));
        colorSpace.Items.Add(tintTransform);
        Assert.True(PdfImageColorSpaceNormalization.TryResolve(
            colorSpace,
            string.Empty,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            OfficeIccRenderingIntent.Saturation,
            transform,
            out PdfImageColorSpaceNormalization normalization));
        Assert.True(OfficeIccColorProfile.TryCreate(outputProfileBytes, out OfficeIccColorProfile? outputProfile));
        Assert.True(outputProfile!.TryConvert(
            new[] { 0.35D, 0.45D, 0.55D },
            OfficeIccRenderingIntent.Saturation,
            out OfficeColor expected));

        Assert.True(normalization.TryConvertComponents(new[] { 0.5D }, out OfficeColor actual));

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void RenderPage_DoesNotMaterializeOverwrittenUnusedColorSpace() {
        byte[] outputProfile = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        byte[] unusedProfile = CreateRgbProfileWithSharedSampledCurves(sampleCount: 4096);
        Assert.True(unusedProfile.Length > outputProfile.Length);
        byte[] pdf = BuildOverwrittenColorSpaceOutputIntentPdf(outputProfile, unusedProfile);
        var options = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = unusedProfile.Length - 1 }
        };

        PdfReadDocument document = PdfReadDocument.Open(pdf, options);
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(document);

        Assert.Single(drawing.Shapes);
    }

    [Fact]
    public void TextParser_PreservesNoneSeparationAsLogicalNonMarkingText() {
        var colorSpaces = new Dictionary<string, PdfPageColorSpace>(StringComparer.Ordinal) {
            ["NoneSpot"] = PdfPageColorSpace.SeparationNone()
        };

        PdfTextSpan span = Assert.Single(TextContentParser.Parse(
            "BT /F1 12 Tf /NoneSpot cs 1 scn (Hidden paint) Tj ET",
            static (_, bytes) => Encoding.ASCII.GetString(bytes),
            static (_, bytes) => bytes.Length * 500D,
            colorSpaces: colorSpaces));

        Assert.Equal("Hidden paint", span.Text);
        Assert.False(span.IsVisible);
        Assert.False(span.CanRestamp);
        Assert.Equal(0, span.Color!.Value.A);
    }

    private static bool TryResolve(
        PdfArray colorSpace,
        Dictionary<int, PdfIndirectObject> objects,
        PdfColorFunctionResolutionContext context,
        out PdfImageColorSpaceNormalization normalization) =>
        PdfImageColorSpaceNormalization.TryResolve(
            colorSpace,
            string.Empty,
            objects,
            context.MaximumRetainedBytes,
            OfficeIccRenderingIntent.RelativeColorimetric,
            outputIntentColorTransform: null,
            colorFunctionEvaluationBudget: null,
            functionResolutionContext: context,
            out normalization);

    private static PdfStream CreateIccProfileStream(byte[] profileBytes) {
        var dictionary = new PdfDictionary();
        dictionary.Items["N"] = new PdfNumber(3);
        return new PdfStream(dictionary, (byte[])profileBytes.Clone());
    }

    private static PdfArray IccColorSpace(PdfStream profile) {
        var colorSpace = new PdfArray();
        colorSpace.Items.Add(new PdfName("ICCBased"));
        colorSpace.Items.Add(profile);
        return colorSpace;
    }

    private static PdfArray NumberArray(params double[] values) {
        var array = new PdfArray();
        foreach (double value in values) array.Items.Add(new PdfNumber(value));
        return array;
    }

    private static byte[] CreateRgbProfileWithSharedSampledCurves(int sampleCount) {
        byte[] source = PdfIccProfiles.SrgbIec6196621;
        int curveLength = checked(12 + sampleCount * 2);
        int paddedCurveLength = checked((curveLength + 3) & ~3);
        int curveOffset = checked((source.Length + 3) & ~3);
        var profile = new byte[checked(curveOffset + paddedCurveLength)];
        Buffer.BlockCopy(source, 0, profile, 0, source.Length);
        WriteUInt32(profile, 0, (uint)profile.Length);
        WriteUInt32(profile, curveOffset, 0x63757276U); // curv
        WriteUInt32(profile, curveOffset + 8, (uint)sampleCount);
        for (int index = 0; index < sampleCount; index++) {
            ushort sample = (ushort)Math.Round(index * 65535D / (sampleCount - 1));
            int offset = curveOffset + 12 + index * 2;
            profile[offset] = (byte)(sample >> 8);
            profile[offset + 1] = (byte)sample;
        }
        RedirectTag(profile, "rTRC", curveOffset, curveLength);
        RedirectTag(profile, "gTRC", curveOffset, curveLength);
        RedirectTag(profile, "bTRC", curveOffset, curveLength);
        return profile;
    }

    private static void RedirectTag(byte[] profile, string signature, int offset, int length) {
        uint target = ((uint)signature[0] << 24) | ((uint)signature[1] << 16) | ((uint)signature[2] << 8) | signature[3];
        int count = checked((int)ReadUInt32(profile, 128));
        for (int index = 0; index < count; index++) {
            int entry = 132 + index * 12;
            if (ReadUInt32(profile, entry) != target) continue;
            WriteUInt32(profile, entry + 4, (uint)offset);
            WriteUInt32(profile, entry + 8, (uint)length);
            return;
        }
        throw new InvalidOperationException("ICC tag was not found: " + signature + ".");
    }

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        unchecked(((uint)bytes[offset] << 24) |
                  ((uint)bytes[offset + 1] << 16) |
                  ((uint)bytes[offset + 2] << 8) |
                  bytes[offset + 3]);

    private static void WriteUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }

    private static byte[] BuildNamedInlineImageOutputIntentPdf(byte[] outputProfile, string payload) {
        byte[] sourceProfile = PdfIccProfiles.SrgbIec6196621;
        string content =
            "BI /W 10 /H 1 /BPC 8 /CS /CsIcc ID " + payload +
            " EI 0.2 0.4 0.8 rg 10 10 20 20 re f";
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OutputIntents [<< /Type /OutputIntent /S /GTS_PDFA1 /DestOutputProfile 6 0 R >>] >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /ColorSpace << /CsIcc [/ICCBased 7 0 R] >> /ExtGState << /Alpha << /ca 0.5 >> >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteProfile(output, 6, outputProfile, includeComponentCount: false);
        WriteProfile(output, 7, sourceProfile, includeComponentCount: true);
        WriteAscii(output, "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildOverwrittenColorSpaceOutputIntentPdf(
        byte[] outputProfile,
        byte[] unusedProfile) {
        string content = "/Unused cs 0 scn /DeviceRGB cs 0.2 0.4 0.8 scn 10 10 20 20 re f";
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OutputIntents [<< /Type /OutputIntent /S /GTS_PDFA1 /DestOutputProfile 6 0 R >>] >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /ColorSpace << /Unused [/ICCBased 7 0 R] >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteProfile(output, 6, outputProfile, includeComponentCount: false);
        WriteProfile(output, 7, unusedProfile, includeComponentCount: true);
        WriteAscii(output, "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteProfile(Stream output, int objectNumber, byte[] profile, bool includeComponentCount) {
        WriteAscii(
            output,
            objectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj\n<< " +
            (includeComponentCount ? "/N 3 " : string.Empty) +
            "/Length " + profile.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(profile, 0, profile.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
    }

    private static void WriteAscii(Stream output, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        output.Write(bytes, 0, bytes.Length);
    }
}
