using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfColorFunctionTests {
    public static IEnumerable<object[]> SampleWidths() {
        yield return new object[] { 1, new byte[] { 0x40 } };
        yield return new object[] { 2, new byte[] { 0x30 } };
        yield return new object[] { 4, new byte[] { 0x0F } };
        yield return new object[] { 8, new byte[] { 0x00, 0xFF } };
        yield return new object[] { 12, new byte[] { 0x00, 0x0F, 0xFF } };
        yield return new object[] { 16, new byte[] { 0x00, 0x00, 0xFF, 0xFF } };
        yield return new object[] { 24, new byte[] { 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF } };
        yield return new object[] { 32, new byte[] { 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF } };
    }

    [Theory]
    [MemberData(nameof(SampleWidths))]
    public void Type0_InterpolatesEverySpecifiedSampleWidth(int bitsPerSample, byte[] samples) {
        PdfStream functionObject = SampledFunction(1, 1, new[] { 2 }, bitsPerSample, samples);

        Assert.True(TryCreateTint(functionObject, 1, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform));
        IReadOnlyList<double>? result = transform(new[] { 0.5D });

        Assert.NotNull(result);
        Assert.Equal(0.5D, result![0], 8);
    }

    [Fact]
    public void Type0_InterpolatesPackedFourBitOutputs() {
        PdfStream functionObject = SampledFunction(
            inputCount: 1,
            outputCount: 2,
            sizes: new[] { 2 },
            bitsPerSample: 4,
            samples: new byte[] { 0x0F, 0xF0 });

        Assert.True(TryCreateTint(functionObject, 1, 2, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform));
        IReadOnlyList<double>? result = transform(new[] { 0.25D });
        Assert.NotNull(result);

        Assert.Equal(0.25D, result![0], 8);
        Assert.Equal(0.75D, result[1], 8);
    }

    [Fact]
    public void Type0_UsesFirstInputAsFastestSampleDimensionAndHonorsReversedEncode() {
        PdfStream defaultFunctionObject = SampledFunction(
            inputCount: 2,
            outputCount: 1,
            sizes: new[] { 2, 2 },
            bitsPerSample: 8,
            samples: new byte[] { 0, 64, 128, 255 });
        PdfStream functionObject = SampledFunction(
            inputCount: 2,
            outputCount: 1,
            sizes: new[] { 2, 2 },
            bitsPerSample: 8,
            samples: new byte[] { 0, 64, 128, 255 },
            encode: new[] { 1D, 0D, 0D, 1D });

        Assert.True(TryCreateTint(defaultFunctionObject, 2, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> defaultTransform));
        Assert.True(TryCreateTint(functionObject, 2, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform));
        IReadOnlyList<double>? defaultResult = defaultTransform(new[] { 0.25D, 0.5D });
        IReadOnlyList<double>? result = transform(new[] { 0.25D, 0.5D });
        Assert.NotNull(defaultResult);
        Assert.NotNull(result);

        Assert.Equal(87.875D / 255D, defaultResult![0], 8);
        Assert.Equal(135.625D / 255D, result![0], 8);
    }

    [Fact]
    public void Type0_ReadsTwelveBitSamplesAcrossByteBoundaries() {
        PdfStream functionObject = SampledFunction(
            inputCount: 1,
            outputCount: 1,
            sizes: new[] { 2 },
            bitsPerSample: 12,
            samples: new byte[] { 0xAB, 0xC1, 0x23 });

        Assert.True(TryCreateTint(functionObject, 1, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform));

        IReadOnlyList<double>? first = transform(new[] { 0D });
        IReadOnlyList<double>? second = transform(new[] { 1D });
        Assert.NotNull(first);
        Assert.NotNull(second);
        Assert.Equal(0xABC / 4095D, first![0], 8);
        Assert.Equal(0x123 / 4095D, second![0], 8);
    }

    [Fact]
    public void Type0_IgnoresDecodedBytesAfterTheRequiredSampleTable() {
        PdfStream functionObject = SampledFunction(
            inputCount: 1,
            outputCount: 1,
            sizes: new[] { 2 },
            bitsPerSample: 8,
            samples: new byte[] { 0, 255, 0xAA, 0xBB });

        Assert.True(TryCreateTint(functionObject, 1, 1, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform));

        Assert.Equal(0D, Assert.Single(transform(new[] { 0D })!), 8);
        Assert.Equal(1D, Assert.Single(transform(new[] { 1D })!), 8);
    }

    [Fact]
    public void Type0_AllowsAlternateColorOutputsOutsideTheUnitInterval() {
        PdfStream functionObject = SampledFunction(
            inputCount: 1,
            outputCount: 3,
            sizes: new[] { 2 },
            bitsPerSample: 8,
            samples: new byte[] { 0, 128, 255, 255, 128, 0 });
        functionObject.Dictionary.Items["Range"] = Numbers(0D, 100D, -128D, 127D, -128D, 127D);
        functionObject.Dictionary.Items["Decode"] = Numbers(0D, 100D, -128D, 127D, -128D, 127D);

        Assert.True(TryCreateTint(functionObject, 1, 3, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform));
        IReadOnlyList<double>? result = transform(new[] { 0D });

        Assert.NotNull(result);
        Assert.Equal(0D, result![0], 8);
        Assert.InRange(result[1], -0.001D, 0.001D);
        Assert.Equal(127D, result[2], 8);
    }

    [Fact]
    public void Type3_ResolvesReferenceChainsSelectsHalfOpenIntervalsAndClipsOuterRange() {
        var objects = new Dictionary<int, PdfIndirectObject>();
        PdfDictionary lower = Type2(new[] { 0D }, new[] { 0.4D });
        PdfDictionary upper = Type2(new[] { 0.8D }, new[] { 1D });
        objects[10] = new PdfIndirectObject(10, 0, new PdfReference(11, 0));
        objects[11] = new PdfIndirectObject(11, 0, lower);
        objects[12] = new PdfIndirectObject(12, 0, upper);

        PdfDictionary stitching = Dictionary(
            ("FunctionType", Number(3)),
            ("Domain", Numbers(0D, 1D)),
            ("Range", Numbers(0D, 0.9D)),
            ("Functions", Array(new PdfReference(10, 0), new PdfReference(12, 0))),
            ("Bounds", Numbers(0.5D)),
            ("Encode", Numbers(0D, 1D, 0D, 1D)));

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(stitching, 1, 1, objects, 1024, out PdfColorFunction function));

        double[]? lowerResult = function.Evaluate(new[] { 0.499D });
        double[]? boundaryResult = function.Evaluate(new[] { 0.5D });
        double[]? upperResult = function.Evaluate(new[] { 1D });
        Assert.NotNull(lowerResult);
        Assert.NotNull(boundaryResult);
        Assert.NotNull(upperResult);
        Assert.Equal(0.3992D, lowerResult![0], 5);
        Assert.Equal(0.8D, boundaryResult![0], 8);
        Assert.Equal(0.9D, upperResult![0], 8);
    }

    [Fact]
    public void Type3_AllowsTheSpecifiedTerminalZeroWidthSubdomain() {
        PdfDictionary stitching = Dictionary(
            ("FunctionType", Number(3)),
            ("Domain", Numbers(0D, 1D)),
            ("Functions", Array(
                Type2(new[] { 0D }, new[] { 0.4D }),
                Type2(new[] { 0.8D }, new[] { 1D }))),
            ("Bounds", Numbers(1D)),
            ("Encode", Numbers(0D, 1D, 0D, 1D)));

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            stitching,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024,
            out PdfColorFunction function));
        double[]? before = function.Evaluate(new[] { 0.999D });
        double[]? atEnd = function.Evaluate(new[] { 1D });

        Assert.NotNull(before);
        Assert.NotNull(atEnd);
        Assert.Equal(0.3996D, before![0], 5);
        Assert.Equal(0.8D, atEnd![0], 8);
        Assert.Contains(1D, function.Discontinuities);
    }

    [Fact]
    public void Type0_ImageConversionReusesTintOutputBuffers() {
        PdfStream sampled = SampledFunction(
            inputCount: 1,
            outputCount: 3,
            sizes: new[] { 2 },
            bitsPerSample: 8,
            samples: new byte[] { 0, 0, 0, 0, 255, 0 });

        AssertImageTintConversionReusesBuffers(sampled);
    }

    [Fact]
    public void Type3_ImageConversionReusesTintOutputBuffers() {
        PdfDictionary stitching = Dictionary(
            ("FunctionType", Number(3)),
            ("Domain", Numbers(0D, 1D)),
            ("Functions", Array(
                Type2(new[] { 0D, 0D, 0D }, new[] { 0D, 0D, 0D }),
                Type2(new[] { 0D, 1D, 0D }, new[] { 0D, 1D, 0D }))),
            ("Bounds", Numbers(0.5D)),
            ("Encode", Numbers(0D, 1D, 0D, 1D)));

        AssertImageTintConversionReusesBuffers(stitching);
    }

    [Fact]
    public void FunctionResolver_RejectsReferenceCyclesAndUnsupportedSampleOrder() {
        var objects = new Dictionary<int, PdfIndirectObject> {
            [1] = new PdfIndirectObject(1, 0, new PdfReference(2, 0)),
            [2] = new PdfIndirectObject(2, 0, new PdfReference(1, 0))
        };
        PdfStream unsupported = SampledFunction(1, 1, new[] { 2 }, 8, new byte[] { 0, 255 }, order: 2);

        Assert.False(PdfColorSpaceFunctionResolver.TryCreateFunction(new PdfReference(1, 0), 1, 1, objects, 1024, out _));
        Assert.False(PdfColorSpaceFunctionResolver.TryCreateFunction(unsupported, 1, 1, new Dictionary<int, PdfIndirectObject>(), 1024, out _));
    }

    [Fact]
    public void Type3_RejectsRecursiveFunctionCyclesAndExcessiveNestingDepth() {
        PdfDictionary self = Dictionary(
            ("FunctionType", Number(3)),
            ("Domain", Numbers(0D, 1D)),
            ("Functions", Array(new PdfReference(30, 0))),
            ("Bounds", Numbers()),
            ("Encode", Numbers(0D, 1D)));
        var objects = new Dictionary<int, PdfIndirectObject> {
            [30] = new PdfIndirectObject(30, 0, self)
        };

        Assert.False(PdfColorSpaceFunctionResolver.TryCreateFunction(
            new PdfReference(30, 0), 1, 1, objects, 1024, out _));

        PdfObject nested = Type2(new[] { 0D }, new[] { 1D });
        for (int depth = 0; depth < 17; depth++) {
            nested = Dictionary(
                ("FunctionType", Number(3)),
                ("Domain", Numbers(0D, 1D)),
                ("Functions", Array(nested)),
                ("Bounds", Numbers()),
                ("Encode", Numbers(0D, 1D)));
        }

        Assert.False(PdfColorSpaceFunctionResolver.TryCreateFunction(
            nested, 1, 1, new Dictionary<int, PdfIndirectObject>(), 1024, out _));
    }

    [Fact]
    public void Type0_PreservesDecodedStreamBudgetFailure() {
        PdfStream functionObject = SampledFunction(1, 1, new[] { 100 }, 8, new byte[100]);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfColorSpaceFunctionResolver.TryCreateFunction(
                functionObject,
                1,
                1,
                new Dictionary<int, PdfIndirectObject>(),
                16,
                out _));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(16, exception.Limit);
        Assert.Equal(100, exception.Actual);
    }

    [Fact]
    public void Type3_AccountsNestedSampleTablesAgainstOneFunctionBudget() {
        PdfStream first = SampledFunction(1, 1, new[] { 10 }, 8, new byte[10]);
        PdfStream second = SampledFunction(1, 1, new[] { 10 }, 8, new byte[10]);
        PdfDictionary stitching = Dictionary(
            ("FunctionType", Number(3)),
            ("Domain", Numbers(0D, 1D)),
            ("Functions", Array(first, second)),
            ("Bounds", Numbers(0.5D)),
            ("Encode", Numbers(0D, 1D, 0D, 1D)));

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfColorSpaceFunctionResolver.TryCreateFunction(
                stitching,
                1,
                1,
                new Dictionary<int, PdfIndirectObject>(),
                16,
                out _));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(20, exception.Actual);
    }

    [Fact]
    public void Type3_AccountsDiscardedSamplePaddingAgainstOneFunctionBudget() {
        PdfStream padded = SampledFunction(1, 1, new[] { 2 }, 8, new byte[10]);
        PdfStream second = SampledFunction(1, 1, new[] { 2 }, 8, new byte[2]);
        PdfDictionary stitching = Dictionary(
            ("FunctionType", Number(3)),
            ("Domain", Numbers(0D, 1D)),
            ("Functions", Array(padded, second)),
            ("Bounds", Numbers(0.5D)),
            ("Encode", Numbers(0D, 1D, 0D, 1D)));

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfColorSpaceFunctionResolver.TryCreateFunction(
                stitching,
                1,
                1,
                new Dictionary<int, PdfIndirectObject>(),
                11,
                out _));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(12, exception.Actual);
    }

    [Fact]
    public void Type3_MemoizesCompactRepeatedReferenceGraphs() {
        var objects = new Dictionary<int, PdfIndirectObject>();
        const int leafObjectNumber = 100;
        objects[leafObjectNumber] = new PdfIndirectObject(leafObjectNumber, 0, Type2(new[] { 0D }, new[] { 1D }));
        int childObjectNumber = leafObjectNumber;
        for (int level = 5; level >= 1; level--) {
            int objectNumber = level;
            PdfObject[] functions = Enumerable.Repeat<PdfObject>(new PdfReference(childObjectNumber, 0), 32).ToArray();
            double[] bounds = Enumerable.Range(1, 31).Select(static index => index / 32D).ToArray();
            double[] encode = Enumerable.Range(0, 32).SelectMany(static _ => new[] { 0D, 1D }).ToArray();
            PdfDictionary stitching = Dictionary(
                ("FunctionType", Number(3)),
                ("Domain", Numbers(0D, 1D)),
                ("Functions", Array(functions)),
                ("Bounds", Numbers(bounds)),
                ("Encode", Numbers(encode)));
            objects[objectNumber] = new PdfIndirectObject(objectNumber, 0, stitching);
            childObjectNumber = objectNumber;
        }

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            new PdfReference(1, 0),
            1,
            1,
            objects,
            1024,
            out PdfColorFunction function));
        Assert.Equal(0D, Assert.Single(function.Evaluate(new[] { 0.5D })!), 8);
    }

    [Fact]
    public void Type3_RejectsMoreThanTheAggregateParsedFunctionNodeBudget() {
        double[] bounds = Enumerable.Range(1, 31).Select(static index => index / 32D).ToArray();
        double[] encode = Enumerable.Range(0, 32).SelectMany(static _ => new[] { 0D, 1D }).ToArray();
        var branches = new PdfObject[32];
        for (int branch = 0; branch < branches.Length; branch++) {
            PdfObject[] leaves = Enumerable.Range(0, 32)
                .Select(_ => (PdfObject)Type2(new[] { 0D }, new[] { 1D }))
                .ToArray();
            branches[branch] = Dictionary(
                ("FunctionType", Number(3)),
                ("Domain", Numbers(0D, 1D)),
                ("Functions", Array(leaves)),
                ("Bounds", Numbers(bounds)),
                ("Encode", Numbers(encode)));
        }
        PdfDictionary root = Dictionary(
            ("FunctionType", Number(3)),
            ("Domain", Numbers(0D, 1D)),
            ("Functions", Array(branches)),
            ("Bounds", Numbers(bounds)),
            ("Encode", Numbers(encode)));

        Assert.False(PdfColorSpaceFunctionResolver.TryCreateFunction(
            root,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024,
            out _));
    }

    [Fact]
    public void FunctionResolver_AllowsDegenerateDomainsForEveryFunctionType() {
        PdfStream sampled = SampledFunction(1, 1, new[] { 1 }, 8, new byte[] { 128 });
        sampled.Dictionary.Items["Domain"] = Numbers(0.5D, 0.5D);
        PdfDictionary exponential = Type2(new[] { 0D }, new[] { 1D });
        exponential.Items["Domain"] = Numbers(0.5D, 0.5D);
        PdfDictionary stitching = Dictionary(
            ("FunctionType", Number(3)),
            ("Domain", Numbers(0.5D, 0.5D)),
            ("Functions", Array(exponential)),
            ("Bounds", Numbers()),
            ("Encode", Numbers(0D, 1D)));
        PdfStream calculator = new PdfStream(
            Dictionary(
                ("FunctionType", Number(4)),
                ("Domain", Numbers(0.5D, 0.5D)),
                ("Range", Numbers(0D, 1D))),
            Encoding.ASCII.GetBytes("{}"));

        foreach (PdfObject candidate in new PdfObject[] { sampled, exponential, stitching, calculator }) {
            Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
                candidate,
                1,
                1,
                new Dictionary<int, PdfIndirectObject>(),
                1024,
                out PdfColorFunction function));
            Assert.NotNull(function.Evaluate(new[] { 0D }));
        }
    }

    [Fact]
    public void Type2_ValidatesNonpositiveExponentAgainstTheAuthoredDomain() {
        PdfDictionary valid = Type2(new[] { 0D }, new[] { 1D });
        valid.Items["Domain"] = Numbers(1D, 2D);
        valid.Items["N"] = Number(-1D);
        PdfDictionary singular = Type2(new[] { 0D }, new[] { 1D });
        singular.Items["Domain"] = Numbers(0D, 1D);
        singular.Items["N"] = Number(-1D);
        PdfDictionary nonReal = Type2(new[] { 0D }, new[] { 1D });
        nonReal.Items["Domain"] = Numbers(-2D, -1D);
        nonReal.Items["N"] = Number(0.5D);

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            valid,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024,
            out PdfColorFunction function));
        Assert.Equal(0.5D, Assert.Single(function.Evaluate(new[] { 2D })!), 8);
        Assert.False(PdfColorSpaceFunctionResolver.TryCreateFunction(
            singular,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024,
            out _));
        Assert.False(PdfColorSpaceFunctionResolver.TryCreateFunction(
            nonReal,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024,
            out _));
    }

    [Fact]
    public void Type3_SelectsTheZeroWidthLeftEndpointFunctionOnlyAtTheEndpoint() {
        PdfDictionary stitching = Dictionary(
            ("FunctionType", Number(3)),
            ("Domain", Numbers(0D, 1D)),
            ("Functions", Array(
                Type2(new[] { 0D }, new[] { 0D }),
                Type2(new[] { 1D }, new[] { 1D }))),
            ("Bounds", Numbers(0D)),
            ("Encode", Numbers(0D, 1D, 0D, 1D)));

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            stitching,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024,
            out PdfColorFunction function));

        Assert.Equal(0D, Assert.Single(function.Evaluate(new[] { 0D })!), 8);
        Assert.Equal(1D, Assert.Single(function.Evaluate(new[] { double.Epsilon })!), 8);
        Assert.Contains(0D, function.Discontinuities);
    }

    [Fact]
    public void Type3_PreservesNestedLeftEndpointDiscontinuity() {
        PdfDictionary nested = Dictionary(
            ("FunctionType", Number(3)),
            ("Domain", Numbers(0D, 1D)),
            ("Functions", Array(
                Type2(new[] { 0D }, new[] { 0D }),
                Type2(new[] { 1D }, new[] { 1D }))),
            ("Bounds", Numbers(0D)),
            ("Encode", Numbers(0D, 1D, 0D, 1D)));
        PdfDictionary outer = Dictionary(
            ("FunctionType", Number(3)),
            ("Domain", Numbers(0D, 1D)),
            ("Functions", Array(nested)),
            ("Bounds", Numbers()),
            ("Encode", Numbers(0D, 1D)));

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            outer,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024,
            out PdfColorFunction function));

        Assert.Contains(0D, function.Discontinuities);
    }

    [Fact]
    public void Type0_LimitsBreakpointsAfterRestrictingToTheReachableEncodeInterval() {
        PdfStream sampled = SampledFunction(1, 1, new[] { 1000 }, 8, new byte[1000], encode: new[] { 100D, 200D });

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            sampled,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            2048,
            out PdfColorFunction function));

        Assert.Equal(101, function.Breakpoints.Count);
        Assert.Equal(0D, function.Breakpoints[0], 8);
        Assert.Equal(1D, function.Breakpoints[100], 8);
    }

    [Fact]
    public void ShadingFunctionArray_AllowsIndependentComponentDomains() {
        PdfDictionary green = Type2(new[] { 0D }, new[] { 1D });
        green.Items["Domain"] = Numbers(0.2D, 0.8D);
        PdfArray functions = Array(
            Type2(new[] { 0D }, new[] { 1D }),
            green,
            Type2(new[] { 0D }, new[] { 1D }));

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateShadingFunction(
            functions,
            3,
            new Dictionary<int, PdfIndirectObject>(),
            1024,
            out PdfColorFunction function));
        double[]? result = function.Evaluate(new[] { 0.1D });

        Assert.NotNull(result);
        Assert.Equal(0.1D, result![0], 8);
        Assert.Equal(0.2D, result[1], 8);
        Assert.Equal(0.1D, result[2], 8);
        Assert.Contains(0.2D, function.Breakpoints);
        Assert.Contains(0.8D, function.Breakpoints);
    }

    [Fact]
    public void ShadingFunctionArray_RetainsAuthoredDiscontinuitiesBeforeSampleKnots() {
        PdfStream sampled = SampledFunction(1, 1, new[] { 1000 }, 8, new byte[1000]);
        PdfObject[] stitchedChildren = Enumerable.Range(0, 8)
            .Select(index => (PdfObject)Type2(new[] { index / 8D }, new[] { index / 8D }))
            .ToArray();
        double[] bounds = Enumerable.Range(1, 7).Select(static index => index / 8D).ToArray();
        double[] encode = Enumerable.Range(0, 8).SelectMany(static _ => new[] { 0D, 1D }).ToArray();
        PdfDictionary stitching = Dictionary(
            ("FunctionType", Number(3)),
            ("Domain", Numbers(0D, 1D)),
            ("Functions", Array(stitchedChildren)),
            ("Bounds", Numbers(bounds)),
            ("Encode", Numbers(encode)));

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateShadingFunction(
            Array(sampled, stitching),
            2,
            new Dictionary<int, PdfIndirectObject>(),
            2048,
            out PdfColorFunction function));

        Assert.True(function.Breakpoints.Count <= 128);
        Assert.All(function.Discontinuities, discontinuity => Assert.Contains(discontinuity, function.Breakpoints));
        Assert.All(bounds, bound => Assert.Contains(bound, function.Breakpoints));
    }

    [Fact]
    public void ShadingFunctionArray_RejectsAuthoredDiscontinuityUnionAboveTheExactLimit() {
        PdfObject[] components = Enumerable.Range(0, 5)
            .Select(component => (PdfObject)Dictionary(
                ("FunctionType", Number(3)),
                ("Domain", Numbers(0D, 1D)),
                ("Functions", Array(Enumerable.Range(0, 32)
                    .Select(index => (PdfObject)Type2(new[] { (component * 32 + index) / 160D }, new[] { (component * 32 + index) / 160D }))
                    .ToArray())),
                ("Bounds", Numbers(Enumerable.Range(1, 31).Select(index => (component * 32 + index) / 160D).ToArray())),
                ("Encode", Numbers(Enumerable.Range(0, 32).SelectMany(static _ => new[] { 0D, 1D }).ToArray()))))
            .ToArray();

        Assert.False(PdfColorSpaceFunctionResolver.TryCreateShadingFunction(
            Array(components),
            5,
            new Dictionary<int, PdfIndirectObject>(),
            1024 * 1024,
            out _));
    }

    [Fact]
    public void FunctionResolutionContext_CachesAliasesAndAggregatesRetainedBytesAcrossResources() {
        PdfStream first = SampledFunction(1, 1, new[] { 64 }, 8, new byte[64]);
        PdfStream second = SampledFunction(1, 1, new[] { 64 }, 8, new byte[64]);
        var context = new PdfColorFunctionResolutionContext(100);
        var objects = new Dictionary<int, PdfIndirectObject>();

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(first, 1, 1, objects, 1024, context, out PdfColorFunction cached));
        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(first, 1, 1, objects, 1024, context, out PdfColorFunction alias));
        Assert.Same(cached, alias);
        Assert.Throws<PdfReadLimitException>(() =>
            PdfColorSpaceFunctionResolver.TryCreateFunction(second, 1, 1, objects, 1024, context, out _));
    }

    [Fact]
    public void Type0_HandlesExtremeFiniteEncodeValuesWithoutIndexingOutsideSamples() {
        PdfStream sampled = SampledFunction(
            1,
            1,
            new[] { 2 },
            8,
            new byte[] { 0, 255 },
            encode: new[] { -1E308D, 1E308D });

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(
            sampled,
            1,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024,
            out PdfColorFunction function));

        Assert.Equal(0D, Assert.Single(function.Evaluate(new[] { 0D })!), 8);
        Assert.Equal(0D, Assert.Single(function.Evaluate(new[] { 0.5D })!), 8);
        Assert.Equal(1D, Assert.Single(function.Evaluate(new[] { 1D })!), 8);
    }

    [Fact]
    public void Type0_RejectsSampleTablesAboveTheBoundedFourInputContract() {
        PdfStream functionObject = SampledFunction(5, 1, new[] { 1, 1, 1, 1, 1 }, 8, new byte[] { 0 });

        Assert.False(PdfColorSpaceFunctionResolver.TryCreateFunction(
            functionObject,
            5,
            1,
            new Dictionary<int, PdfIndirectObject>(),
            1024,
            out _));
    }

    [Fact]
    public void ShadingFunctionArray_ComposesOneOutputFunctions() {
        PdfArray functions = Array(
            Type2(new[] { 0D }, new[] { 1D }),
            Type2(new[] { 1D }, new[] { 0D }),
            Type2(new[] { 0.5D }, new[] { 0.5D }));

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateShadingFunction(
            functions,
            3,
            new Dictionary<int, PdfIndirectObject>(),
            1024,
            out PdfColorFunction function));
        double[]? result = function.Evaluate(new[] { 0.25D });
        Assert.NotNull(result);

        Assert.Equal(new[] { 0.25D, 0.75D, 0.5D }, result!);
    }

    [Fact]
    public void RenderPage_AppliesSampledSeparationTintToContentPaint() {
        const string sampleHex = "FF00000000FF>";
        byte[] pdf = BuildSinglePagePdf(
            "/Spot cs 0.5 scn 20 20 100 100 re f",
            "<< /ColorSpace << /Spot [/Separation /Brand /DeviceRGB 5 0 R] >> >>",
            SampledStreamObject(5, 1, 3, "[2]", 8, sampleHex));

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeColor fill = Assert.Single(drawing.Shapes).Shape.FillColor!.Value;

        Assert.InRange(fill.R, 126, 129);
        Assert.InRange(fill.G, 0, 1);
        Assert.InRange(fill.B, 126, 129);
    }

    [Fact]
    public void RenderPage_NormalizesSampledSeparationTintForImagePixels() {
        const string sampleHex = "FF00000000FF>";
        byte[] pdf = BuildSinglePagePdf(
            "q 20 0 0 20 40 80 cm /Im1 Do Q",
            "<< /XObject << /Im1 5 0 R >> >>",
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace [/Separation /Brand /DeviceRGB 6 0 R] /BitsPerComponent 8 /Filter /ASCIIHexDecode /Length 3 >>\nstream\n80>\nendstream\nendobj",
            SampledStreamObject(6, 1, 3, "[2]", 8, sampleHex));

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(
            pdf,
            options: new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg, ContinueOnError = true }));

        Assert.Equal("image/png", Assert.Single(drawing.Images).ContentType);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void RenderPage_ProjectsSampledShadingKnotsAsGradientStops() {
        const string sampleHex = "FF000000FF000000FF>";
        byte[] pdf = BuildSinglePagePdf(
            "/Sh1 sh",
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] /Function 6 0 R /Extend [true true] >>\nendobj",
            SampledStreamObject(6, 1, 3, "[3]", 8, sampleHex));

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeLinearGradient gradient = Assert.Single(drawing.Shapes).Shape.FillGradient!;

        Assert.Equal(3, gradient.Stops.Count);
        Assert.Equal(OfficeColor.Red, gradient.Stops[0].Color);
        Assert.Equal(0.5D, gradient.Stops[1].Offset, 8);
        Assert.Equal(OfficeColor.Lime, gradient.Stops[1].Color);
        Assert.Equal(OfficeColor.Blue, gradient.Stops[2].Color);
    }

    [Fact]
    public void RenderPage_PreservesNonlinearExponentialShadingThroughBoundedStops() {
        byte[] pdf = BuildSinglePagePdf(
            "/Sh1 sh",
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] /Function << /FunctionType 2 /Domain [0 1] /C0 [0 0 0] /C1 [1 0 0] /N 2 >> /Extend [true true] >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeLinearGradient gradient = Assert.Single(drawing.Shapes).Shape.FillGradient!;
        OfficeGradientStop midpoint = Assert.Single(gradient.Stops, stop => Math.Abs(stop.Offset - 0.5D) < 0.0000001D);

        Assert.InRange(midpoint.Color.R, 63, 65);
        Assert.Equal(0, midpoint.Color.G);
        Assert.Equal(0, midpoint.Color.B);
    }

    [Fact]
    public void RenderPage_AdaptivelyRefinesNonlinearDeviceRgbCalculatorShading() {
        const string program = "{ dup mul dup dup }";
        byte[] pdf = BuildSinglePagePdf(
            "/Sh1 sh",
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] /Function 6 0 R /Extend [true true] >>\nendobj",
            CalculatorStreamObject(6, program));

        OfficeLinearGradient gradient = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes).Shape.FillGradient!;

        Assert.True(gradient.Stops.Count > 2);
        Assert.Contains(gradient.Stops, static stop => stop.Color.R > 0);
    }

    [Fact]
    public void RenderPage_RefinesAffineShadingWhenAuthoredRangeClipsOutput() {
        byte[] pdf = BuildSinglePagePdf(
            "/Sh1 sh",
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] " +
            "/Function << /FunctionType 2 /Domain [0 1] /Range [0 1 0 1 0 1] " +
            "/C0 [-1 0 0] /C1 [1 0 0] /N 1 >> /Extend [true true] >>\nendobj");

        OfficeLinearGradient gradient = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes).Shape.FillGradient!;

        Assert.Contains(gradient.Stops, static stop => Math.Abs(stop.Offset - 0.5D) < 0.01D && stop.Color.R == 0);
    }

    [Fact]
    public void RenderPage_RejectsCalculatorShadingWithUnboundedRoundingDiscontinuities() {
        const string program = "{ 4 mul floor 4 div dup dup }";
        byte[] pdf = BuildSinglePagePdf(
            "/Sh1 sh",
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] /Function 6 0 R /Extend [true true] >>\nendobj",
            CalculatorStreamObject(6, program));

        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        Assert.DoesNotContain(page.ToDrawing().Shapes, static item => item.Shape.FillGradient != null);
        Assert.Contains(
            page.GetRenderCapabilityDiagnostics(),
            static diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
    }

    [Fact]
    public void RenderPage_BoundsCalculatorTintEvaluationAcrossContentPaint() {
        string program = "{ " + string.Concat(Enumerable.Repeat("dup pop ", 100)) + "dup dup }";
        byte[] pdf = BuildSinglePagePdf(
            "/Spot cs 0.5 scn 20 20 100 100 re f",
            "<< /ColorSpace << /Spot [/Separation /Brand /DeviceRGB 5 0 R] >> >>",
            CalculatorStreamObject(5, program));
        PdfReadPage page = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxContentOperations = 100 }
        }).Pages[0];

        OfficeShape shape = Assert.Single(page.ToDrawing().Shapes).Shape;

        Assert.Equal(OfficeColor.Black, shape.FillColor);
    }

    [Fact]
    public void RenderPage_PreservesClippedFunctionDomainPlateausInShading() {
        byte[] pdf = BuildSinglePagePdf(
            "/Sh1 sh",
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] /Domain [0 1] /Function << /FunctionType 2 /Domain [0.25 0.75] /C0 [0 0 0] /C1 [1 0 0] /N 1 >> /Extend [true true] >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeLinearGradient gradient = Assert.Single(drawing.Shapes).Shape.FillGradient!;

        Assert.Equal(4, gradient.Stops.Count);
        Assert.Equal(0.25D, gradient.Stops[1].Offset, 8);
        Assert.Equal(0.75D, gradient.Stops[2].Offset, 8);
        Assert.Equal(gradient.Stops[0].Color, gradient.Stops[1].Color);
        Assert.Equal(gradient.Stops[2].Color, gradient.Stops[3].Color);
    }

    [Fact]
    public void RenderPage_PreservesType3DiscontinuityAcrossDescendingShadingDomain() {
        byte[] pdf = BuildSinglePagePdf(
            "/Sh1 sh",
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] /Domain [1 0] " +
            "/Function << /FunctionType 3 /Domain [0 1] " +
            "/Functions [<< /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [1 0 0] /N 1 >> " +
            "<< /FunctionType 2 /Domain [0 1] /C0 [0 0 1] /C1 [0 0 1] /N 1 >>] " +
            "/Bounds [0.5] /Encode [0 1 0 1] >> /Extend [true true] >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeLinearGradient gradient = Assert.Single(drawing.Shapes).Shape.FillGradient!;
        OfficeGradientStop[] boundary = gradient.Stops.Where(stop => Math.Abs(stop.Offset - 0.5D) < 0.0000001D).ToArray();

        Assert.Equal(2, boundary.Length);
        Assert.Equal(OfficeColor.Blue, boundary[0].Color);
        Assert.Equal(OfficeColor.Red, boundary[1].Color);
    }

    [Fact]
    public void RenderPage_PreservesZeroWidthLeftEndpointFunctionInShading() {
        byte[] pdf = BuildSinglePagePdf(
            "/Sh1 sh",
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] " +
            "/Function << /FunctionType 3 /Domain [0 1] " +
            "/Functions [<< /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [1 0 0] /N 1 >> " +
            "<< /FunctionType 2 /Domain [0 1] /C0 [0 0 1] /C1 [0 0 1] /N 1 >>] " +
            "/Bounds [0] /Encode [0 1 0 1] >> /Extend [true true] >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeLinearGradient gradient = Assert.Single(drawing.Shapes).Shape.FillGradient!;
        OfficeGradientStop[] endpoint = gradient.Stops.Where(stop => stop.Offset == 0D).ToArray();

        Assert.Equal(2, endpoint.Length);
        Assert.Equal(OfficeColor.Red, endpoint[0].Color);
        Assert.Equal(OfficeColor.Blue, endpoint[1].Color);
    }

    private static bool TryCreateTint(
        PdfObject function,
        int inputCount,
        int outputCount,
        out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform) {
        transform = null!;
        if (!PdfColorSpaceFunctionResolver.TryCreateTintTransform(
                function,
                inputCount,
                outputCount,
                new Dictionary<int, PdfIndirectObject>(),
                1024 * 1024,
                out PdfColorSpaceTintTransform boundedTransform)) return false;

        transform = components => {
            var output = new double[outputCount];
            return boundedTransform(components, output) ? output : null;
        };
        return true;
    }

    private static PdfStream SampledFunction(
        int inputCount,
        int outputCount,
        int[] sizes,
        int bitsPerSample,
        byte[] samples,
        double[]? encode = null,
        int order = 1) {
        var entries = new List<(string Key, PdfObject Value)> {
            ("FunctionType", Number(0)),
            ("Domain", Numbers(Enumerable.Range(0, inputCount).SelectMany(static _ => new[] { 0D, 1D }).ToArray())),
            ("Range", Numbers(Enumerable.Range(0, outputCount).SelectMany(static _ => new[] { 0D, 1D }).ToArray())),
            ("Size", Numbers(sizes.Select(static value => (double)value).ToArray())),
            ("BitsPerSample", Number(bitsPerSample))
        };
        if (encode != null) entries.Add(("Encode", Numbers(encode)));
        if (order != 1) entries.Add(("Order", Number(order)));
        return new PdfStream(Dictionary(entries.ToArray()), samples);
    }

    private static PdfDictionary Type2(double[] c0, double[] c1) => Dictionary(
        ("FunctionType", Number(2)),
        ("Domain", Numbers(0D, 1D)),
        ("C0", Numbers(c0)),
        ("C1", Numbers(c1)),
        ("N", Number(1D)));

    private static void AssertImageTintConversionReusesBuffers(PdfObject function) {
        PdfArray colorSpace = Array(
            new PdfName("Separation"),
            new PdfName("Spot"),
            new PdfName("DeviceRGB"),
            function);
        Assert.True(PdfImageColorSpaceNormalization.TryResolve(
            colorSpace,
            string.Empty,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            out PdfImageColorSpaceNormalization normalization));
        PdfImageColorConversionBuffer conversionBuffer = normalization.CreateConversionBuffer();
        byte[] sample = { 255 };
        for (int index = 0; index < 32; index++) {
            Assert.True(normalization.TryConvertPixel(sample, 0, null, conversionBuffer, out _));
        }

        bool converted = true;
        OfficeColor color = OfficeColor.Black;
#if NET8_0_OR_GREATER
        long before = GC.GetAllocatedBytesForCurrentThread();
#endif
        for (int index = 0; index < 4096; index++) {
            converted &= normalization.TryConvertPixel(sample, 0, null, conversionBuffer, out color);
        }
#if NET8_0_OR_GREATER
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
#endif

        Assert.True(converted);
        Assert.Equal(OfficeColor.FromRgb(0, 255, 0), color);
#if NET8_0_OR_GREATER
        Assert.InRange(allocated, 0, 1024);
#endif
    }

    private static PdfDictionary Dictionary(params (string Key, PdfObject Value)[] entries) {
        var dictionary = new PdfDictionary();
        foreach ((string key, PdfObject value) in entries) dictionary.Items[key] = value;
        return dictionary;
    }

    private static PdfArray Array(params PdfObject[] values) {
        var array = new PdfArray();
        array.Items.AddRange(values);
        return array;
    }

    private static PdfArray Numbers(params double[] values) => Array(values.Select(Number).ToArray());

    private static PdfNumber Number(double value) => new(value);

    private static string SampledStreamObject(
        int objectNumber,
        int inputCount,
        int outputCount,
        string size,
        int bitsPerSample,
        string asciiHexSamples,
        int order = 1) {
        string intervals = string.Join(" ", Enumerable.Range(0, inputCount).Select(static _ => "0 1"));
        string ranges = string.Join(" ", Enumerable.Range(0, outputCount).Select(static _ => "0 1"));
        string orderEntry = order == 1 ? string.Empty : " /Order " + order;
        return objectNumber + " 0 obj\n<< /FunctionType 0 /Domain [" + intervals + "] /Range [" + ranges + "] /Size " + size +
               " /BitsPerSample " + bitsPerSample + orderEntry + " /Filter /ASCIIHexDecode /Length " + asciiHexSamples.Length + " >>\nstream\n" +
               asciiHexSamples + "\nendstream\nendobj";
    }

    private static string CalculatorStreamObject(int objectNumber, string program) =>
        objectNumber + " 0 obj\n<< /FunctionType 4 /Domain [0 1] /Range [0 1 0 1 0 1] /Length " +
        Encoding.ASCII.GetByteCount(program) + " >>\nstream\n" + program + "\nendstream\nendobj";

    private static byte[] BuildSinglePagePdf(string content, string resources, params string[] extraObjects) {
        content = content.TrimEnd('\r', '\n');
        string[] objects = {
            "%PDF-1.4",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 160] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources " + resources + " /Contents 4 0 R >>\nendobj",
            "4 0 obj\n<< /Length " + Encoding.ASCII.GetByteCount(content) + " >>\nstream\n" + content + "\nendstream\nendobj"
        };
        string pdf = string.Join("\n", objects.Concat(extraObjects).Concat(new[] {
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        })) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }
}
