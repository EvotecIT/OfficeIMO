using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfColorFunctionTests {
    [Theory]
    [InlineData("{ abs }", -0.25D, 0.25D)]
    [InlineData("{ 0.25 add }", 0.5D, 0.75D)]
    [InlineData("{ 0.25 sub }", 0.5D, 0.25D)]
    [InlineData("{ 4 mul }", 0.5D, 2D)]
    [InlineData("{ 4 div }", 0.5D, 0.125D)]
    [InlineData("{ 2 exp }", 3D, 9D)]
    [InlineData("{ sqrt }", 9D, 3D)]
    [InlineData("{ ln }", 1D, 0D)]
    [InlineData("{ log }", 10D, 1D)]
    [InlineData("{ 2 atan }", 2D, 45D)]
    [InlineData("{ sin }", 90D, 1D)]
    [InlineData("{ cos }", 180D, -1D)]
    [InlineData("{ ceiling }", 1.25D, 2D)]
    [InlineData("{ floor }", 1.75D, 1D)]
    [InlineData("{ round }", -1.5D, -1D)]
    [InlineData("{ truncate }", -1.75D, -1D)]
    [InlineData("{ cvi cvr }", 1.75D, 1D)]
    public void Type4_ExecutesArithmeticAndConversionOperators(string program, double input, double expected) {
        Assert.Equal(expected, EvaluateCalculator(program, new[] { input }, 1)[0], 8);
    }

    [Theory]
    [InlineData("{ cvi 3 idiv }", 7.9D, 2D)]
    [InlineData("{ cvi 3 mod }", 7.9D, 1D)]
    [InlineData("{ cvi 1 bitshift }", 3D, 6D)]
    [InlineData("{ cvi -1 bitshift }", -1D, 2147483647D)]
    [InlineData("{ cvi 6 and }", 3D, 2D)]
    [InlineData("{ cvi 6 or }", 3D, 7D)]
    [InlineData("{ cvi 6 xor }", 3D, 5D)]
    [InlineData("{ cvi not }", 0D, -1D)]
    [InlineData("{ cvi 2147483647 add }", 1D, 2147483648D)]
    public void Type4_ExecutesIntegerAndBitwiseOperators(string program, double input, double expected) {
        Assert.Equal(expected, EvaluateCalculator(program, new[] { input }, 1)[0], 8);
    }

    [Theory]
    [InlineData("{ dup add }", 0.25D, 0.5D)]
    [InlineData("{ 1 index add add }", 1D, 2D, 4D)]
    [InlineData("{ 2 copy add add add }", 1D, 2D, 6D)]
    [InlineData("{ exch sub }", 1D, 2D, 1D)]
    [InlineData("{ 3 1 roll sub sub }", 1D, 2D, 3D, 4D)]
    public void Type4_ExecutesBoundedStackOperators(string program, params double[] valuesAndExpected) {
        double expected = valuesAndExpected[valuesAndExpected.Length - 1];
        double[] inputs = valuesAndExpected.Take(valuesAndExpected.Length - 1).ToArray();

        Assert.Equal(expected, EvaluateCalculator(program, inputs, 1)[0], 8);
    }

    [Theory]
    [InlineData("{ 1 eq { 1 } { 0 } ifelse }", 1D, 1D)]
    [InlineData("{ 2 ne { 1 } { 0 } ifelse }", 1D, 1D)]
    [InlineData("{ 1 le { 1 } { 0 } ifelse }", 1D, 1D)]
    [InlineData("{ 1 ge { 1 } { 0 } ifelse }", 1D, 1D)]
    [InlineData("{ pop true false and { 1 } { 0 } ifelse }", 0D, 0D)]
    [InlineData("{ pop true false xor { 1 } { 0 } ifelse }", 0D, 1D)]
    [InlineData("{ pop true not { 1 } { 0 } ifelse }", 0D, 0D)]
    [InlineData("{ pop true 1 eq { 1 } { 0 } ifelse }", 0D, 0D)]
    [InlineData("{ pop true 1 ne { 1 } { 0 } ifelse }", 0D, 1D)]
    [InlineData("{ pop 1 }", 0D, 1D)]
    public void Type4_ExecutesRelationalBooleanAndPopOperators(string program, double input, double expected) {
        Assert.Equal(expected, EvaluateCalculator(program, new[] { input }, 1)[0], 8);
    }

    [Fact]
    public void Type4_ExecutesNestedConditionalAndBooleanOperatorsWithComments() {
        const string program = "{% preserve the source value\n dup .5 lt { dup mul } { dup .75 gt { 1 exch sub } { pop true false or { .5 } if } ifelse } ifelse }";

        Assert.Equal(0.0625D, EvaluateCalculator(program, new[] { 0.25D }, 1)[0], 8);
        Assert.Equal(0.5D, EvaluateCalculator(program, new[] { 0.5D }, 1)[0], 8);
        Assert.Equal(0.1D, EvaluateCalculator(program, new[] { 0.9D }, 1)[0], 8);
    }

    [Fact]
    public void Type4_AllowsNonAsciiBytesInsideCommentsOnly() {
        byte[] source = { (byte)'{', (byte)'%', 0xC2, 0xA9, (byte)'\n', (byte)'d', (byte)'u', (byte)'p', (byte)'}' };

        Assert.True(PdfCalculatorProgram.TryParse(source, out PdfCalculatorProgram program));
        Assert.NotNull(program.Evaluate(new[] { 0.25D }, 2));

        source[1] = 0xC2;
        Assert.False(PdfCalculatorProgram.TryParse(source, out _));
    }

    [Fact]
    public void Type4_PreservesAuthoredConditionalBreakpointWhenUniformSamplesFillTheLimit() {
        PdfStream calculator = CalculatorFunction(1, 1, "{ dup .5 lt { pop 0 } { pop 1 } ifelse }");

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            calculator,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out PdfColorFunction function));

        Assert.Contains(0.5D, function.Breakpoints);
        Assert.Contains(0.5D, function.Discontinuities);
        Assert.True(function.Breakpoints.Count <= 128);
    }

    [Fact]
    public void Type4_DerivesConditionalBoundaryFromAffineOperands() {
        PdfStream calculator = CalculatorFunction(1, 1, "{ dup 2 mul 1 lt { pop 0 } { pop 1 } ifelse }");

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            calculator,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out PdfColorFunction function));

        Assert.Contains(0.5D, function.Breakpoints);
        Assert.Contains(0.5D, function.Discontinuities);
    }

    [Theory]
    [InlineData("{ dup dup }", false)]
    [InlineData("{ 360 mul sin dup dup }", true)]
    public void Type4_ClassifiesAdaptiveShadingSamplingWithoutLosingTheAffineFastPath(
        string program,
        bool expected) {
        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            CalculatorFunction(1, 3, program),
            1,
            3,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out PdfColorFunction function));

        Assert.Equal(expected, function.RequiresAdaptiveShadingSampling);
    }

    [Fact]
    public void Type4_RejectsConditionalThresholdThatCannotBeBounded() {
        PdfStream calculator = CalculatorFunction(1, 1, "{ dup 360 mul sin 0 lt { pop 0 } { pop 1 } ifelse }");

        Assert.False(PdfColorSpaceFunctionResolver.TryCreateFunction(
            calculator,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out _));
    }

    [Fact]
    public void Type4_ShadingRejectsPeriodicOutputThatCanAliasAdaptiveProbes() {
        Assert.False(PdfColorSpaceFunctionResolver.TryCreateShadingFunction(
            CalculatorFunction(1, 3, "{ 7200 mul sin dup dup }"),
            3,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out _));
    }

    [Fact]
    public void Type4_ShadingRejectsSharpAbsoluteValueFeaturesThatCanAliasAdaptiveProbes() {
        Assert.False(PdfColorSpaceFunctionResolver.TryCreateShadingFunction(
            CalculatorFunction(1, 3, "{ 3 mul 1 sub abs 10000 mul 1 exch sub dup dup }"),
            3,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out _));
    }

    [Fact]
    public void Type4_ShadingRejectsNarrowRationalPeakThatCanAliasAdaptiveProbes() {
        Assert.False(PdfColorSpaceFunctionResolver.TryCreateShadingFunction(
            CalculatorFunction(1, 3, "{ 3 mul 1 sub dup mul 1000000000 mul 1 add 1 exch div dup dup }"),
            3,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out _));
    }

    [Fact]
    public void Type4_ShadingAcceptsCertifiedMonotonicQuadraticOutput() {
        Assert.True(PdfColorSpaceFunctionResolver.TryCreateShadingFunction(
            CalculatorFunction(1, 3, "{ dup mul dup dup }"),
            3,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out _));
    }

    [Fact]
    public void Type4_BoundsValidationWorkAcrossAuthoredConstants() {
        string body = string.Join(" ", Enumerable.Range(1, 2000).Select(index =>
            (index / 2001D).ToString("0.000000", System.Globalization.CultureInfo.InvariantCulture) + " pop"));
        PdfStream calculator = CalculatorFunction(1, 1, "{ " + body + " }");

        Assert.False(PdfColorSpaceFunctionResolver.TryCreateFunction(
            calculator,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            out _));
    }

    [Fact]
    public void Type4_ExecutesTheIsoDoubleDotCalculatorExample() {
        const string program = "{ 360 mul sin 2 div exch 360 mul sin 2 div add }";

        Assert.Equal(1D, EvaluateCalculator(program, new[] { 0.25D, 0.25D }, 1)[0], 8);
    }

    [Fact]
    public void Type4_ClipsInputsAndOutputsThroughTheSharedFunctionContract() {
        PdfStream calculator = CalculatorFunction(1, 1, "{ 2 mul }");

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            calculator,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out PdfColorFunction function));

        Assert.Equal(1D, Assert.Single(function.Evaluate(new[] { 2D })!), 8);
        Assert.Equal(0D, Assert.Single(function.Evaluate(new[] { -1D })!), 8);
    }

    [Theory]
    [InlineData("dup")]
    [InlineData("{ unknown }")]
    [InlineData("{ 1e2 }")]
    [InlineData("{ { 1 } }")]
    [InlineData("{ true { 1 } { 0 } if }")]
    [InlineData("{ pop pop }")]
    [InlineData("{ 0 div }")]
    [InlineData("{ -1 sqrt }")]
    [InlineData("{ pop 3000000000 cvi }")]
    public void Type4_RejectsMalformedOrUndefinedPrograms(string program) {
        Assert.False(PdfColorSpaceFunctionResolver.TryCreateFunction(
            CalculatorFunction(1, 1, program),
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out _));
    }

    [Fact]
    public void Type4_RejectsProgramsThatOverflowTheBoundedOperandStack() {
        string program = "{ " + string.Join(" ", Enumerable.Repeat("dup", 256)) + " }";

        Assert.False(PdfColorSpaceFunctionResolver.TryCreateFunction(
            CalculatorFunction(1, 1, program),
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out _));
    }

    [Fact]
    public void Type4_AccountsProgramEvaluationWorkBeforeImagePixelConversion() {
        const int operationPairs = 2000;
        string program = "{ " + string.Concat(Enumerable.Repeat("dup pop ", operationPairs)) + "dup dup }";
        PdfStream function = CalculatorFunction(1, 3, program);
        PdfArray colorSpace = Array(
            new PdfName("Separation"),
            new PdfName("Brand"),
            new PdfName("DeviceRGB"),
            function);
        int maximumPixels = (int)((PdfReadLimits.DefaultMaxDecodedStreamBytes / sizeof(double)) / (operationPairs * 2 + 2));
        PdfDictionary image = Dictionary(
            ("Width", Number(maximumPixels)),
            ("Height", Number(1)),
            ("BitsPerComponent", Number(8)),
            ("ColorSpace", colorSpace));

        Assert.True(ResourceResolver.CanProjectImageColorSpace(
            image,
            null,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes));

        image.Items["Width"] = Number(maximumPixels + 1);

        Assert.False(ResourceResolver.CanProjectImageColorSpace(
            image,
            null,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes));
    }

    [Theory]
    [InlineData("{ 1 copy pop dup dup }", 256)]
    [InlineData("{ dup dup 3 1 roll }", 768)]
    public void Type4_WeightsVariableCostStackOperatorsForImageWorkBudgets(string program, int minimumVariableWork) {
        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            CalculatorFunction(1, 3, program),
            1,
            3,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            out PdfColorFunction function));

        Assert.True(function.EvaluationCost >= minimumVariableWork);
        int maximumPixels = (int)((PdfReadLimits.DefaultMaxDecodedStreamBytes / sizeof(double)) / function.EvaluationCost);
        PdfDictionary image = Dictionary(
            ("Width", Number(maximumPixels + 1)),
            ("Height", Number(1)),
            ("BitsPerComponent", Number(8)),
            ("ColorSpace", Array(
                new PdfName("Separation"),
                new PdfName("Brand"),
                new PdfName("DeviceRGB"),
                CalculatorFunction(1, 3, program))));

        Assert.False(ResourceResolver.CanProjectImageColorSpace(
            image,
            null,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes));
    }

    [Fact]
    public void Type4_PreservesNestedIccAlternateEvaluationWorkForImageBudgets() {
        const int operationPairs = 2000;
        string program = "{ " + string.Concat(Enumerable.Repeat("dup pop ", operationPairs)) + "dup dup }";
        PdfArray alternate = Array(
            new PdfName("Separation"),
            new PdfName("Brand"),
            new PdfName("DeviceRGB"),
            CalculatorFunction(1, 3, program));
        PdfStream unsupportedProfile = new PdfStream(
            Dictionary(("N", Number(1)), ("Alternate", alternate)),
            new byte[] { 0, 1, 2, 3 });
        PdfDictionary image = Dictionary(
            ("Width", Number(1)),
            ("Height", Number(1)),
            ("BitsPerComponent", Number(8)),
            ("ColorSpace", Array(new PdfName("ICCBased"), unsupportedProfile)));
        int maximumPixels = (int)((PdfReadLimits.DefaultMaxDecodedStreamBytes / sizeof(double)) / (operationPairs * 2 + 2));

        image.Items["Width"] = Number(maximumPixels);
        Assert.True(ResourceResolver.CanProjectImageColorSpace(
            image,
            null,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes));

        image.Items["Width"] = Number(maximumPixels + 1);
        Assert.False(ResourceResolver.CanProjectImageColorSpace(
            image,
            null,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes));
    }

    [Fact]
    public void Type4_RejectsIndexedPalettesThatExceedNestedTransformWorkBudgets() {
        string program = "{ dup dup " + string.Concat(Enumerable.Repeat("3 1 roll ", 171)) + "}";
        PdfArray indexed = Array(
            new PdfName("Indexed"),
            Array(
                new PdfName("Separation"),
                new PdfName("Brand"),
                new PdfName("DeviceRGB"),
                CalculatorFunction(1, 3, program)),
            Number(255),
            new PdfStringObj(Enumerable.Range(0, 256).Select(static value => (byte)value).ToArray()));
        var objects = new Dictionary<int, PdfIndirectObject>();

        Assert.False(PdfIndexedImageNormalizer.CanNormalizeColorSpace(
            indexed,
            8,
            objects,
            PdfReadLimits.DefaultMaxDecodedStreamBytes));
        Assert.False(PdfImageColorSpaceNormalization.TryResolve(
            indexed,
            string.Empty,
            objects,
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            out _));

        PdfDictionary image = Dictionary(
            ("Width", Number(1)),
            ("Height", Number(1)),
            ("BitsPerComponent", Number(8)),
            ("ColorSpace", indexed));
        Assert.False(ResourceResolver.CanProjectImageColorSpace(
            image,
            null,
            objects,
            PdfReadLimits.DefaultMaxDecodedStreamBytes));
    }

    [Fact]
    public void IndexedCapabilityProbeDoesNotEvaluateTintPaletteEntries() {
        PdfArray indexed = Array(
            new PdfName("Indexed"),
            Array(
                new PdfName("Separation"),
                new PdfName("Brand"),
                new PdfName("DeviceRGB"),
                CalculatorFunction(1, 3, "{ 3 mul 1 sub abs 0.01 sub sqrt dup dup }")),
            Number(0),
            new PdfStringObj(new byte[] { 85 }));
        var objects = new Dictionary<int, PdfIndirectObject>();

        Assert.True(PdfIndexedImageNormalizer.CanNormalizeColorSpace(
            indexed,
            8,
            objects,
            PdfReadLimits.DefaultMaxDecodedStreamBytes));
        Assert.False(PdfImageColorSpaceNormalization.TryResolve(
            indexed,
            string.Empty,
            objects,
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            out _));
    }

    [Fact]
    public void IccIndexedAlternateChargesPaletteMaterializationToSharedBudget() {
        PdfArray indexedAlternate = Array(
            new PdfName("Indexed"),
            Array(
                new PdfName("Separation"),
                new PdfName("Brand"),
                new PdfName("DeviceRGB"),
                CalculatorFunction(1, 3, "{ dup dup }")),
            Number(1),
            new PdfStringObj(new byte[] { 0, 255 }));
        var profileDictionary = Dictionary(
            ("N", Number(1)),
            ("Alternate", indexedAlternate));
        PdfArray icc = Array(
            new PdfName("ICCBased"),
            new PdfStream(profileDictionary, new byte[] { 0 }));
        long chargedWork = 0L;

        Assert.False(PdfImageColorSpaceNormalization.TryResolve(
            icc,
            string.Empty,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            OfficeIccRenderingIntent.RelativeColorimetric,
            outputIntentColorTransform: null,
            (evaluationCost, evaluationCount) => {
                chargedWork += (long)evaluationCost * evaluationCount;
                return false;
            },
            out _));
        Assert.True(chargedWork > 0L);
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void Type4_ReusesItsBoundedOperandStackAcrossPixelScaleEvaluations() {
        PdfStream calculator = CalculatorFunction(1, 3, "{ dup dup }");
        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            calculator,
            1,
            3,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out PdfColorFunction function));
        double[] input = { 0.5D };
        Assert.NotNull(function.Evaluate(input));

        long before = GC.GetAllocatedBytesForCurrentThread();
        for (int index = 0; index < 1000; index++) Assert.NotNull(function.Evaluate(input));
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.InRange(allocated, 1L, 200_000L);
    }
#endif

    [Fact]
    public void RenderPage_AppliesCalculatorSeparationTintToContentPaint() {
        byte[] pdf = BuildSinglePagePdf(
            "/Spot cs .5 scn 20 20 100 100 re f",
            "<< /ColorSpace << /Spot [/Separation /Brand /DeviceRGB 5 0 R] >> >>",
            CalculatorStreamObject(5, 1, 3, "{ dup dup mul exch dup }"));

        OfficeColor fill = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes).Shape.FillColor!.Value;

        Assert.InRange(fill.R, 63, 64);
        Assert.InRange(fill.G, 127, 128);
        Assert.InRange(fill.B, 127, 128);
    }

    [Fact]
    public void RenderPage_NormalizesCalculatorSeparationTintForImagePixels() {
        byte[] pdf = BuildSinglePagePdf(
            "q 20 0 0 20 40 80 cm /Im1 Do Q",
            "<< /XObject << /Im1 5 0 R >> >>",
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace [/Separation /Brand /DeviceRGB 6 0 R] /BitsPerComponent 8 /Filter /ASCIIHexDecode /Length 3 >>\nstream\n80>\nendstream\nendobj",
            CalculatorStreamObject(6, 1, 3, "{ dup dup mul exch dup }"));

        byte[] pixelBytes = PdfPngTestImages.DecodeStoredPngIdat(Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Images).Bytes);

        Assert.Equal(0, pixelBytes[0]);
        Assert.InRange(pixelBytes[1], 63, 65);
        Assert.InRange(pixelBytes[2], 127, 129);
        Assert.InRange(pixelBytes[3], 127, 129);
    }

    [Fact]
    public void RenderPage_ProjectsCalculatorShadingAsBoundedGradientStops() {
        byte[] pdf = BuildSinglePagePdf(
            "/Sh1 sh",
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] /Function 6 0 R /Extend [true true] >>\nendobj",
            CalculatorStreamObject(6, 1, 3, "{ dup dup mul exch dup }"));

        OfficeLinearGradient gradient = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes).Shape.FillGradient!;
        OfficeGradientStop midpoint = gradient.Stops.OrderBy(stop => Math.Abs(stop.Offset - 0.5D)).First();

        Assert.InRange(midpoint.Offset, 0.49D, 0.51D);
        Assert.InRange(midpoint.Color.R, 61, 67);
        Assert.InRange(midpoint.Color.G, 125, 130);
        Assert.InRange(midpoint.Color.B, 125, 130);
        Assert.Equal(128, gradient.Stops.Count);
    }

    private static double[] EvaluateCalculator(string program, double[] inputs, int outputCount) {
        PdfStream calculator = new PdfStream(
            Dictionary(
                ("FunctionType", Number(4)),
                ("Domain", Numbers(inputs.SelectMany(static value => new[] { value, value }).ToArray())),
                ("Range", Numbers(Enumerable.Range(0, outputCount).SelectMany(static _ => new[] { -3000000000D, 3000000000D }).ToArray()))),
            System.Text.Encoding.ASCII.GetBytes(program));
        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            calculator,
            inputs.Length,
            outputCount,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out PdfColorFunction function));
        return Assert.IsType<double[]>(function.Evaluate(inputs));
    }

    private static PdfStream CalculatorFunction(
        int inputCount,
        int outputCount,
        string program,
        double minimum = 0D,
        double maximum = 1D) =>
        new PdfStream(
            Dictionary(
                ("FunctionType", Number(4)),
                ("Domain", Numbers(Enumerable.Range(0, inputCount).SelectMany(_ => new[] { minimum, maximum }).ToArray())),
                ("Range", Numbers(Enumerable.Range(0, outputCount).SelectMany(_ => new[] { minimum, maximum }).ToArray()))),
            System.Text.Encoding.ASCII.GetBytes(program));

    private static string CalculatorStreamObject(int objectNumber, int inputCount, int outputCount, string program) {
        string domain = string.Join(" ", Enumerable.Range(0, inputCount).Select(static _ => "0 1"));
        string range = string.Join(" ", Enumerable.Range(0, outputCount).Select(static _ => "0 1"));
        return objectNumber + " 0 obj\n<< /FunctionType 4 /Domain [" + domain + "] /Range [" + range + "] /Length " +
               System.Text.Encoding.ASCII.GetByteCount(program) + " >>\nstream\n" + program + "\nendstream\nendobj";
    }
}
