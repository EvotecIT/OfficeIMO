using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfColorFunctionTests {
    [Fact]
    public void Type0_Order3UsesLinearInterpolationOnlyInSmallSampleDimensions() {
        PdfStream functionObject = SampledFunction(
            inputCount: 2,
            outputCount: 1,
            sizes: new[] { 4, 2 },
            bitsPerSample: 8,
            samples: new byte[] { 0, 64, 192, 255, 0, 64, 192, 255 },
            order: 3);

        Assert.True(TryCreateTint(functionObject, 2, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform));
        IReadOnlyList<double>? result = transform(new[] { 1D / 6D, 0.5D });

        Assert.NotNull(result);
        Assert.Equal(28D / 255D, result![0], 8);
    }

    [Fact]
    public void Type0_Order3UsesNaturalCubicSplineForThreeSampleDimension() {
        PdfStream functionObject = SampledFunction(
            inputCount: 1,
            outputCount: 1,
            sizes: new[] { 3 },
            bitsPerSample: 8,
            samples: new byte[] { 0, 64, 255 },
            order: 3);

        Assert.True(TryCreateTint(functionObject, 1, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform));
        IReadOnlyList<double>? result = transform(new[] { 0.25D });

        Assert.NotNull(result);
        Assert.Equal(643D / 8160D, result![0], 8);
    }

    [Fact]
    public void Type0_Order3UsesNaturalCubicSplineForOneInputAndEveryOutput() {
        PdfStream functionObject = SampledFunction(
            inputCount: 1,
            outputCount: 2,
            sizes: new[] { 4 },
            bitsPerSample: 8,
            samples: new byte[] { 0, 255, 255, 0, 0, 255, 255, 0 },
            order: 3);

        Assert.True(TryCreateTint(functionObject, 1, 2, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform));
        IReadOnlyList<double>? result = transform(new[] { 1D / 6D });

        Assert.NotNull(result);
        Assert.Equal(0.75D, result![0], 8);
        Assert.Equal(0.25D, result[1], 8);
    }

    [Fact]
    public void Type0_Order3PricesNaturalSplineByItsTwoReadsPerOutput() {
        PdfStream functionObject = SampledFunction(
            inputCount: 1,
            outputCount: 3,
            sizes: new[] { 4 },
            bitsPerSample: 8,
            samples: new byte[] { 0, 0, 0, 64, 64, 64, 192, 192, 192, 255, 255, 255 },
            order: 3);

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            functionObject,
            1,
            3,
            new Dictionary<int, PdfIndirectObject>(),
            1024,
            out PdfColorFunction function));

        Assert.Equal(6, function.CubicEvaluationCost);
    }

    [Fact]
    public void Type0_Order3HonorsReversedAndFractionalEncodeIntervals() {
        PdfStream reversed = SampledFunction(
            1,
            1,
            new[] { 4 },
            8,
            new byte[] { 0, 255, 0, 255 },
            encode: new[] { 3D, 0D },
            order: 3);
        PdfStream fractional = SampledFunction(
            1,
            1,
            new[] { 4 },
            8,
            new byte[] { 0, 255, 0, 255 },
            encode: new[] { 0.5D, 2.5D },
            order: 3);

        Assert.True(TryCreateTint(reversed, 1, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> reversedTransform));
        Assert.True(TryCreateTint(fractional, 1, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> fractionalTransform));

        Assert.Equal(0.75D, Assert.Single(reversedTransform(new[] { 5D / 6D })!), 8);
        Assert.Equal(1D, Assert.Single(fractionalTransform(new[] { 0.25D })!), 8);
    }

    [Fact]
    public void Type0_Order3RetainsCubicInterpolationAcrossNarrowAndConstantEncodeIntervals() {
        PdfStream functionObject = SampledFunction(
            1,
            1,
            new[] { 4 },
            8,
            new byte[] { 0, 0, 255, 0 },
            encode: new[] { 1.25D, 1.75D },
            order: 3);
        PdfStream constant = SampledFunction(
            1,
            1,
            new[] { 4 },
            8,
            new byte[] { 0, 0, 255, 0 },
            encode: new[] { 1.5D, 1.5D },
            order: 3);

        Assert.True(TryCreateTint(functionObject, 1, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform));
        Assert.True(TryCreateTint(constant, 1, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> constantTransform));

        Assert.Equal(0.575D, Assert.Single(transform(new[] { 0.5D })!), 8);
        Assert.Equal(0.575D, Assert.Single(constantTransform(new[] { 0D })!), 8);
        Assert.Equal(0.575D, Assert.Single(constantTransform(new[] { 1D })!), 8);
    }

    [Fact]
    public void Type0_OptionalNullOrderEncodeAndDecodeUseSpecifiedDefaults() {
        PdfStream direct = SampledFunction(1, 1, new[] { 2 }, 8, new byte[] { 0, 255 });
        direct.Dictionary.Items["Order"] = PdfNull.Instance;
        direct.Dictionary.Items["Encode"] = PdfNull.Instance;
        direct.Dictionary.Items["Decode"] = PdfNull.Instance;

        var objects = new Dictionary<int, PdfIndirectObject> {
            [20] = new PdfIndirectObject(20, 0, PdfNull.Instance),
            [21] = new PdfIndirectObject(21, 0, new PdfReference(20, 0))
        };
        PdfStream indirect = SampledFunction(1, 1, new[] { 2 }, 8, new byte[] { 0, 255 });
        indirect.Dictionary.Items["Order"] = new PdfReference(21, 0);
        indirect.Dictionary.Items["Encode"] = new PdfReference(21, 0);
        indirect.Dictionary.Items["Decode"] = new PdfReference(21, 0);

        Assert.True(TryCreateTint(direct, 1, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> directTransform));
        Assert.True(PdfColorSpaceFunctionResolver.TryCreateTintTransform(
            indirect,
            1,
            1,
            objects,
            1024,
            out PdfColorSpaceTintTransform indirectTransform));

        Assert.Equal(0.5D, Assert.Single(directTransform(new[] { 0.5D })!), 8);
        var indirectOutput = new double[1];
        Assert.True(indirectTransform(new[] { 0.5D }, indirectOutput));
        Assert.Equal(0.5D, indirectOutput[0], 8);
    }

    [Fact]
    public void Type0_Order3UsesFirstInputAsFastestTensorDimension() {
        byte[] samples = {
            0, 1, 4, 9,
            10, 11, 14, 19,
            20, 21, 24, 29,
            30, 31, 34, 39
        };
        PdfStream functionObject = SampledFunction(2, 1, new[] { 4, 4 }, 8, samples, order: 3);

        Assert.True(TryCreateTint(functionObject, 2, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform));
        IReadOnlyList<double>? result = transform(new[] { 0.5D, 0.5D });

        Assert.NotNull(result);
        Assert.Equal(17.25D / 255D, result![0], 8);
    }

    [Fact]
    public void Type0_Order3PreservesSampleKnotsAndClipsSplineOutputToRange() {
        PdfStream functionObject = SampledFunction(1, 1, new[] { 4 }, 8, new byte[] { 0, 255, 0, 255 }, order: 3);
        functionObject.Dictionary.Items["Range"] = Numbers(0D, 0.7D);
        functionObject.Dictionary.Items["Decode"] = Numbers(0D, 1D);

        Assert.True(TryCreateTint(functionObject, 1, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform));

        Assert.Equal(0D, Assert.Single(transform(new[] { 0D })!), 8);
        Assert.Equal(0.7D, Assert.Single(transform(new[] { 1D / 3D })!), 8);
        Assert.Equal(0D, Assert.Single(transform(new[] { 2D / 3D })!), 8);
    }

    [Fact]
    public void Type0_Order3FailsClosedWhenSplineArithmeticCannotRemainFinite() {
        PdfStream functionObject = SampledFunction(1, 1, new[] { 4 }, 8, new byte[] { 0, 255, 0, 255 }, order: 3);
        functionObject.Dictionary.Items["Range"] = Numbers(-1E308D, 1E308D);
        functionObject.Dictionary.Items["Decode"] = Numbers(-1E308D, 1E308D);

        Assert.False(PdfColorSpaceFunctionResolver.TryCreateFunction(
            functionObject,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024,
            out _));
    }

    [Fact]
    public void Type0_Order3AccountsSplineWorkingMemoryAgainstTheFunctionBudget() {
        PdfStream functionObject = SampledFunction(1, 1, new[] { 4 }, 8, new byte[] { 0, 255, 0, 255 }, order: 3);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfColorSpaceFunctionResolver.TryCreateFunction(
                functionObject,
                1,
                1,
                new Dictionary<int, PdfIndirectObject>(),
                64,
                out _));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(64, exception.Limit);
        Assert.Equal(68, exception.Actual);
    }

    [Fact]
    public void ImageNormalization_BoundsFourDimensionalCubicTintWorkBeforePixelConversion() {
        PdfStream function = SampledFunction(
            4,
            3,
            new[] { 4, 4, 4, 4 },
            8,
            new byte[4 * 4 * 4 * 4 * 3],
            order: 3);
        PdfArray colorSpace = Array(
            new PdfName("DeviceN"),
            Array(new PdfName("C1"), new PdfName("C2"), new PdfName("C3"), new PdfName("C4")),
            new PdfName("DeviceRGB"),
            function);

        Assert.True(PdfImageColorSpaceNormalization.TryResolve(
            colorSpace,
            string.Empty,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            out PdfImageColorSpaceNormalization normalization));

        Assert.True(normalization.CanConvertPixelCount(43_690));
        Assert.False(normalization.CanConvertPixelCount(43_691));
        Assert.False(normalization.CanConvertPixelCount(1_000_000));
    }

    [Fact]
    public void ImageNormalization_PreservesCubicWorkThroughUnsupportedIccAlternate() {
        PdfStream function = SampledFunction(
            1,
            3,
            new[] { 4 },
            8,
            new byte[] { 0, 0, 0, 64, 64, 64, 192, 192, 192, 255, 255, 255 },
            order: 3);
        PdfArray separation = Array(
            new PdfName("Separation"),
            new PdfName("Spot"),
            new PdfName("DeviceRGB"),
            function);
        PdfStream unsupportedProfile = new(
            Dictionary(
                ("N", Number(1)),
                ("Alternate", separation)),
            new byte[] { 0 });
        PdfArray colorSpace = Array(new PdfName("ICCBased"), unsupportedProfile);

        Assert.True(PdfImageColorSpaceNormalization.TryResolve(
            colorSpace,
            string.Empty,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            out PdfImageColorSpaceNormalization normalization));

        long maximumPixels = (PdfReadLimits.DefaultMaxDecodedStreamBytes / sizeof(double)) / 6;
        Assert.True(normalization.CanConvertPixelCount(maximumPixels));
        Assert.False(normalization.CanConvertPixelCount(maximumPixels + 1));
    }

    [Fact]
    public void ImageCapabilityDiagnosticsRejectTheSameOverBudgetCubicTintWorkAsRendering() {
        PdfStream function = SampledFunction(
            4,
            3,
            new[] { 4, 4, 4, 4 },
            8,
            new byte[4 * 4 * 4 * 4 * 3],
            order: 3);
        PdfArray colorSpace = Array(
            new PdfName("DeviceN"),
            Array(new PdfName("C1"), new PdfName("C2"), new PdfName("C3"), new PdfName("C4")),
            new PdfName("DeviceRGB"),
            function);
        PdfDictionary image = new();
        image.Items["Width"] = Number(43_690);
        image.Items["Height"] = Number(1);
        image.Items["BitsPerComponent"] = Number(8);
        image.Items["ColorSpace"] = colorSpace;

        Assert.True(ResourceResolver.CanProjectImageColorSpace(
            image,
            null,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes));

        image.Items["Width"] = Number(43_691);

        Assert.False(ResourceResolver.CanProjectImageColorSpace(
            image,
            null,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes));
    }

    [Fact]
    public void Type0_Order3BoundsCurvatureBreakpointsForLargeSampleTables() {
        PdfStream sampled = SampledFunction(1, 1, new[] { 1000 }, 8, new byte[1000], order: 3);

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            sampled,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            32 * 1024,
            out PdfColorFunction function));

        Assert.Equal(128, function.Breakpoints.Count);
        Assert.Equal(0D, function.Breakpoints[0], 8);
        Assert.Equal(1D, function.Breakpoints[127], 8);
    }

    [Fact]
    public void RenderPage_AppliesCubicSampledSeparationTintToContentPaint() {
        const string sampleHex = "000000FF0000000000FF0000>";
        byte[] pdf = BuildSinglePagePdf(
            "/Spot cs 0.1666666667 scn 20 20 100 100 re f",
            "<< /ColorSpace << /Spot [/Separation /Brand /DeviceRGB 5 0 R] >> >>",
            SampledStreamObject(5, 1, 3, "[4]", 8, sampleHex, order: 3));

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeColor fill = Assert.Single(drawing.Shapes).Shape.FillColor!.Value;

        Assert.InRange(fill.R, 190, 192);
        Assert.Equal(0, fill.G);
        Assert.Equal(0, fill.B);
    }

    [Fact]
    public void RenderPage_NormalizesCubicSampledSeparationTintForImagePixels() {
        const string sampleHex = "000000FF0000000000FF0000>";
        byte[] pdf = BuildSinglePagePdf(
            "q 20 0 0 20 40 80 cm /Im1 Do Q",
            "<< /XObject << /Im1 5 0 R >> >>",
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace [/Separation /Brand /DeviceRGB 6 0 R] /BitsPerComponent 8 /Filter /ASCIIHexDecode /Length 3 >>\nstream\n2A>\nendstream\nendobj",
            SampledStreamObject(6, 1, 3, "[4]", 8, sampleHex, order: 3));

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] pixelBytes = PdfPngTestImages.DecodeStoredPngIdat(Assert.Single(drawing.Images).Bytes);

        Assert.Equal(0, pixelBytes[0]);
        Assert.InRange(pixelBytes[1], 188, 191);
        Assert.Equal(0, pixelBytes[2]);
        Assert.Equal(0, pixelBytes[3]);
    }

    [Fact]
    public void RenderPage_ProjectsCubicSampledShadingCurvatureAsBoundedGradientStops() {
        const string sampleHex = "000000FF0000000000FF0000>";
        byte[] pdf = BuildSinglePagePdf(
            "/Sh1 sh",
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] /Function 6 0 R /Extend [true true] >>\nendobj",
            SampledStreamObject(6, 1, 3, "[4]", 8, sampleHex, order: 3));

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeLinearGradient gradient = Assert.Single(drawing.Shapes).Shape.FillGradient!;
        OfficeGradientStop cubicQuarter = Assert.Single(
            gradient.Stops,
            stop => Math.Abs(stop.Offset - 1D / 6D) < 0.0000001D);

        Assert.InRange(cubicQuarter.Color.R, 190, 192);
        Assert.Equal(0, cubicQuarter.Color.G);
        Assert.Equal(0, cubicQuarter.Color.B);
        Assert.InRange(gradient.Stops.Count, 12, 15);
    }
}
