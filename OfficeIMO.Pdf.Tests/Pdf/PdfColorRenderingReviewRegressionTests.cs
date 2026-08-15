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
        var context = new PdfColorFunctionResolutionContext(profileBytes.Length * 2 - 1);

        Assert.True(TryResolve(firstColorSpace, objects, context, out PdfImageColorSpaceNormalization firstNormalization));
        Assert.True(TryResolve(firstColorSpace, objects, context, out PdfImageColorSpaceNormalization aliasNormalization));
        Assert.Equal(firstNormalization.Kind, aliasNormalization.Kind);
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            TryResolve(secondColorSpace, objects, context, out _));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(profileBytes.Length * 2 - 1, exception.Limit);
        Assert.Equal(profileBytes.Length * 2, exception.Actual);
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
