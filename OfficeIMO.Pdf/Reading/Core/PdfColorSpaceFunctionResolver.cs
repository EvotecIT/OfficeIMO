using System.IO;

namespace OfficeIMO.Pdf;

internal delegate bool PdfColorSpaceTintTransform(IReadOnlyList<double> components, double[] output);

/// <summary>Resolves bounded PDF functions shared by content, image, and shading color projection.</summary>
internal static partial class PdfColorSpaceFunctionResolver {
    private const int MaxFunctionDepth = 16;
    private const int MaxParsedFunctionNodes = 1024;
    private const int MaxSampledInputs = 4;
    private const int MaxStitchingFunctions = 32;
    private const int MaxSuggestedSampleBreakpoints = 128;
    private static readonly double[] DefaultC0 = { 0D };
    private static readonly double[] DefaultC1 = { 1D };
    [ThreadStatic]
    private static double[]? _scalarInput;

    internal static bool TryCreateTintTransform(
        PdfObject? value,
        int inputCount,
        int outputCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        out PdfColorSpaceTintTransform transform) =>
        TryCreateTintTransform(
            value,
            inputCount,
            outputCount,
            objects,
            maxDecodedStreamBytes,
            out transform,
            out _);

    internal static bool TryCreateTintTransform(
        PdfObject? value,
        int inputCount,
        int outputCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        out PdfColorSpaceTintTransform transform,
        out int evaluationCost) {
        transform = null!;
        evaluationCost = 0;
        if (!TryCreateFunction(value, inputCount, outputCount, objects, maxDecodedStreamBytes, out PdfColorFunction function) ||
            !HasUnitIntervals(function.Domain, inputCount)) return false;

        transform = function.TryEvaluate;
        evaluationCost = function.EvaluationCost;
        return true;
    }

    internal static bool TryCreateFunction(
        PdfObject? value,
        int inputCount,
        int outputCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        out PdfColorFunction function) {
        function = null!;
        if (inputCount < 1 || outputCount < 1 || maxDecodedStreamBytes <= 0) return false;
        long retainedFunctionBytes = 0L;
        long remainingCalculatorValidationWork = PdfCalculatorProgram.MaxValidationWork;
        int parsedFunctionNodes = 0;
        return TryCreateFunction(
            value,
            inputCount,
            outputCount,
            objects,
            maxDecodedStreamBytes,
            depth: 0,
            new HashSet<PdfObject>(),
            new Dictionary<PdfObject, Dictionary<long, PdfColorFunction>>(PdfObjectReferenceComparer.Instance),
            ref parsedFunctionNodes,
            ref retainedFunctionBytes,
            ref remainingCalculatorValidationWork,
            out function);
    }

    internal static bool TryCreateShadingFunction(
        PdfObject? value,
        int outputCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        out PdfColorFunction function) {
        function = null!;
        if (!TryResolveObject(value, objects, out PdfObject? resolved)) return false;
        if (resolved is not PdfArray functions) {
            return TryCreateFunction(resolved, 1, outputCount, objects, maxDecodedStreamBytes, out function);
        }
        if (functions.Items.Count != outputCount || outputCount < 1) return false;

        var components = new PdfColorFunction[outputCount];
        var activeFunctions = new HashSet<PdfObject>();
        var functionCache = new Dictionary<PdfObject, Dictionary<long, PdfColorFunction>>(PdfObjectReferenceComparer.Instance);
        long retainedFunctionBytes = 0L;
        long remainingCalculatorValidationWork = PdfCalculatorProgram.MaxValidationWork;
        int parsedFunctionNodes = 0;
        for (int index = 0; index < components.Length; index++) {
            if (!TryCreateFunction(
                    functions.Items[index],
                    1,
                    1,
                    objects,
                    maxDecodedStreamBytes,
                    depth: 0,
                    activeFunctions,
                    functionCache,
                    ref parsedFunctionNodes,
                    ref retainedFunctionBytes,
                    ref remainingCalculatorValidationWork,
                    out components[index])) return false;
        }

        double[] domain = {
            components.Min(static component => component.Domain[0]),
            components.Max(static component => component.Domain[1])
        };
        double[] componentDomainBoundaries = components
            .SelectMany(static component => component.Domain)
            .Distinct()
            .OrderBy(static value => value)
            .ToArray();
        if (componentDomainBoundaries.Length > MaxSuggestedSampleBreakpoints) return false;
        double[] authoredDiscontinuities = components
            .SelectMany(static component => component.Discontinuities)
            .Distinct()
            .OrderBy(static value => value)
            .ToArray();
        int reservedDomainBoundaries = componentDomainBoundaries.Count(
            boundary => !authoredDiscontinuities.Contains(boundary));
        double[] discontinuities = LimitSuggestedPoints(
            authoredDiscontinuities,
            required: null,
            MaxSuggestedSampleBreakpoints - reservedDomainBoundaries);
        double[] requiredBreakpoints = componentDomainBoundaries
            .Concat(discontinuities)
            .Distinct()
            .ToArray();
        double[] breakpoints = LimitSuggestedPoints(
            components.SelectMany(static component => component.Domain.Concat(component.Breakpoints)),
            requiredBreakpoints);
        function = new PdfColorFunction(
            1,
            outputCount,
            domain,
            range: null,
            (values, output, outputOffset) => EvaluateFunctionArray(values, output, outputOffset, components),
            breakpoints,
            discontinuities,
            evaluationCost: SumEvaluationCost(components));
        return true;
    }

    private static bool TryCreateFunction(
        PdfObject? value,
        int inputCount,
        int outputCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        int depth,
        HashSet<PdfObject> activeFunctions,
        Dictionary<PdfObject, Dictionary<long, PdfColorFunction>> functionCache,
        ref int parsedFunctionNodes,
        ref long retainedFunctionBytes,
        ref long remainingCalculatorValidationWork,
        out PdfColorFunction function) {
        function = null!;
        if (depth > MaxFunctionDepth ||
            !TryResolveObject(value, objects, out PdfObject? resolved) ||
            resolved == null) return false;

        long cacheKey = ((long)inputCount << 32) | (uint)outputCount;
        if (functionCache.TryGetValue(resolved, out Dictionary<long, PdfColorFunction>? variants) &&
            variants.TryGetValue(cacheKey, out PdfColorFunction? cachedFunction)) {
            function = cachedFunction;
            return true;
        }
        if (++parsedFunctionNodes > MaxParsedFunctionNodes || !activeFunctions.Add(resolved)) return false;

        try {
            PdfDictionary? dictionary = resolved switch {
                PdfStream stream => stream.Dictionary,
                PdfDictionary direct => direct,
                _ => null
            };
            if (dictionary == null) return false;

            int? functionType = TryReadInteger(dictionary.Items.TryGetValue("FunctionType", out PdfObject? type) ? type : null, objects);
            bool created = functionType switch {
                0 => TryCreateSampledFunction(resolved as PdfStream, dictionary, inputCount, outputCount, objects, maxDecodedStreamBytes, ref retainedFunctionBytes, out function),
                2 => TryCreateExponentialFunction(dictionary, inputCount, outputCount, objects, out function),
                3 => TryCreateStitchingFunction(dictionary, inputCount, outputCount, objects, maxDecodedStreamBytes, depth, activeFunctions, functionCache, ref parsedFunctionNodes, ref retainedFunctionBytes, ref remainingCalculatorValidationWork, out function),
                4 => TryCreateCalculatorFunction(resolved as PdfStream, dictionary, inputCount, outputCount, objects, maxDecodedStreamBytes, ref retainedFunctionBytes, ref remainingCalculatorValidationWork, out function),
                _ => false
            };
            if (!created || function == null) return false;
            if (variants == null) {
                variants = new Dictionary<long, PdfColorFunction>();
                functionCache[resolved] = variants;
            }
            variants[cacheKey] = function;
            return true;
        } finally {
            activeFunctions.Remove(resolved);
        }
    }

    private static bool TryCreateSampledFunction(
        PdfStream? stream,
        PdfDictionary dictionary,
        int inputCount,
        int outputCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        ref long retainedFunctionBytes,
        out PdfColorFunction function) {
        function = null!;
        if (stream == null || inputCount > MaxSampledInputs ||
            !TryReadIntervals(dictionary, "Domain", inputCount, objects, allowEqual: true, required: true, out double[] domain) ||
            !TryReadIntervals(dictionary, "Range", outputCount, objects, allowEqual: true, required: true, out double[] range) ||
            !TryReadIntegerArray(dictionary, "Size", inputCount, objects, out int[] sizes) ||
            sizes.Any(static size => size < 1) ||
            !dictionary.Items.TryGetValue("BitsPerSample", out PdfObject? bitsObject) ||
            TryReadInteger(bitsObject, objects) is not int bitsPerSample ||
            !IsSupportedSampleWidth(bitsPerSample)) return false;

        if (!TryResolveOptionalEntry(dictionary, "Order", objects, out PdfObject? orderObject, out bool hasOrder)) return false;
        int order = hasOrder ? TryReadInteger(orderObject, objects) ?? 0 : 1;
        if (order is not (1 or 3)) return false;

        double[] encode;
        if (!TryResolveOptionalEntry(dictionary, "Encode", objects, out PdfObject? encodeObject, out bool hasEncode)) return false;
        if (hasEncode) {
            encode = ReadNumberArray(encodeObject, objects);
            if (!HasFinitePairs(encode, inputCount)) return false;
        } else {
            encode = new double[inputCount * 2];
            for (int index = 0; index < inputCount; index++) encode[index * 2 + 1] = sizes[index] - 1D;
        }

        if (!TryResolveOptionalEntry(dictionary, "Decode", objects, out PdfObject? decodeObject, out bool hasDecode)) return false;
        double[] decode = hasDecode ? ReadNumberArray(decodeObject, objects) : (double[])range.Clone();
        if (!HasFinitePairs(decode, outputCount)) return false;

        long samplePointCount = 1;
        try {
            for (int index = 0; index < sizes.Length; index++) samplePointCount = checked(samplePointCount * sizes[index]);
            long sampleValueCount = checked(samplePointCount * outputCount);
            long expectedBits = checked(sampleValueCount * bitsPerSample);
            long expectedBytes = checked((expectedBits + 7L) / 8L);
            int naturalSplinePointCount = GetNaturalSplinePointCount(order, inputCount, encode, sizes);
            bool usesNaturalCubicSpline = naturalSplinePointCount > 0;
            long derivativeBytes = usesNaturalCubicSpline
                ? checked((long)naturalSplinePointCount * outputCount * sizeof(double))
                : 0L;
            long splineWorkspaceBytes = usesNaturalCubicSpline
                ? checked((long)naturalSplinePointCount * sizeof(double))
                : 0L;
            long minimumRetainedBytes = checked(retainedFunctionBytes + expectedBytes + derivativeBytes);
            long peakFunctionBytes = checked(minimumRetainedBytes + splineWorkspaceBytes);
            if (peakFunctionBytes > maxDecodedStreamBytes || expectedBytes > int.MaxValue) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maxDecodedStreamBytes, peakFunctionBytes);
            }

            byte[] decoded;
            try {
                decoded = Filters.StreamDecoder.DecodeRequired(
                    dictionary,
                    stream.Data,
                    objects,
                    checked((int)(maxDecodedStreamBytes - retainedFunctionBytes)));
            } catch (InvalidDataException) {
                return false;
            }
            if (decoded.LongLength < expectedBytes) return false;
            long copiedSampleBytes = decoded.LongLength == expectedBytes ? 0L : expectedBytes;
            peakFunctionBytes = checked(
                retainedFunctionBytes + decoded.LongLength + copiedSampleBytes + derivativeBytes + splineWorkspaceBytes);
            if (peakFunctionBytes > maxDecodedStreamBytes) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.DecodedStreamBytes,
                    maxDecodedStreamBytes,
                    peakFunctionBytes);
            }
            byte[] samples;
            if (decoded.LongLength == expectedBytes) {
                samples = decoded;
            } else {
                samples = new byte[(int)expectedBytes];
                Buffer.BlockCopy(decoded, 0, samples, 0, samples.Length);
            }
            long totalRetainedBytes = checked(retainedFunctionBytes + decoded.LongLength + derivativeBytes);
            ulong maximumSample = bitsPerSample == 32 ? uint.MaxValue : (1UL << bitsPerSample) - 1UL;
            if (!TryCreateSampledEvaluator(
                    order,
                    domain,
                    encode,
                    sizes,
                    bitsPerSample,
                    maximumSample,
                    decode,
                    samples,
                    outputCount,
                    out PdfColorFunctionEvaluator evaluator,
                    out int cubicEvaluationCost)) return false;
            retainedFunctionBytes = totalRetainedBytes;
            double[] breakpoints = CreateSampleBreakpoints(domain, encode, sizes, order);

            function = new PdfColorFunction(
                inputCount,
                outputCount,
                domain,
                range,
                evaluator,
                breakpoints,
                evaluationCost: cubicEvaluationCost);
            return true;
        } catch (OverflowException) {
            return false;
        }
    }

    private static bool TryCreateExponentialFunction(
        PdfDictionary dictionary,
        int inputCount,
        int outputCount,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfColorFunction function) {
        function = null!;
        if (inputCount != 1 ||
            !TryReadIntervals(dictionary, "Domain", 1, objects, allowEqual: true, required: true, out double[] domain) ||
            !TryReadOptionalRange(dictionary, outputCount, objects, out double[]? range)) return false;

        if (!TryReadOptionalNumberArray(dictionary, "C0", DefaultC0, objects, out double[] c0) ||
            !TryReadOptionalNumberArray(dictionary, "C1", DefaultC1, objects, out double[] c1) ||
            !dictionary.Items.TryGetValue("N", out PdfObject? exponentValue) ||
            TryReadNumber(exponentValue, objects) is not double exponent ||
            c0.Length != outputCount || c1.Length != outputCount ||
            c0.Any(static value => !IsFinite(value)) || c1.Any(static value => !IsFinite(value)) ||
            !IsSupportedType2Exponent(exponent, domain)) return false;

        function = new PdfColorFunction(
            inputCount,
            outputCount,
            domain,
            range,
            (values, output, outputOffset) => EvaluateType2(values, output, outputOffset, c0, c1, exponent),
            exponent == 1D ? domain : CreateUniformBreakpoints(domain, 65));
        return true;
    }

    private static bool IsSupportedType2Exponent(double exponent, double[] domain) {
        if (!IsFinite(exponent) || domain.Length != 2) return false;
        double minimum = domain[0];
        double maximum = domain[1];
        if (minimum < 0D && exponent != Math.Truncate(exponent)) return false;
        if (exponent < 0D && minimum <= 0D && maximum >= 0D) return false;
        return IsFinite(Math.Pow(minimum, exponent)) && IsFinite(Math.Pow(maximum, exponent));
    }

    private static bool TryCreateStitchingFunction(
        PdfDictionary dictionary,
        int inputCount,
        int outputCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        int depth,
        HashSet<PdfObject> activeFunctions,
        Dictionary<PdfObject, Dictionary<long, PdfColorFunction>> functionCache,
        ref int parsedFunctionNodes,
        ref long retainedFunctionBytes,
        ref long remainingCalculatorValidationWork,
        out PdfColorFunction function) {
        function = null!;
        if (inputCount != 1 ||
            !TryReadIntervals(dictionary, "Domain", 1, objects, allowEqual: true, required: true, out double[] domain) ||
            !TryReadOptionalRange(dictionary, outputCount, objects, out double[]? range) ||
            !dictionary.Items.TryGetValue("Functions", out PdfObject? functionsObject) ||
            !TryResolveObject(functionsObject, objects, out PdfObject? resolvedFunctions) ||
            resolvedFunctions is not PdfArray functions ||
            functions.Items.Count < 1 || functions.Items.Count > MaxStitchingFunctions) return false;

        double[] bounds = dictionary.Items.TryGetValue("Bounds", out PdfObject? boundsObject)
            ? ReadNumberArray(boundsObject, objects)
            : Array.Empty<double>();
        double[] encode = dictionary.Items.TryGetValue("Encode", out PdfObject? encodeObject)
            ? ReadNumberArray(encodeObject, objects)
            : Array.Empty<double>();
        if (bounds.Length != functions.Items.Count - 1 || !HasFinitePairs(encode, functions.Items.Count)) return false;
        double domainStart = domain[0];
        double domainEnd = domain[1];
        double previous = domainStart;
        for (int index = 0; index < bounds.Length; index++) {
            bool isLast = index == bounds.Length - 1;
            if (!IsFinite(bounds[index]) || bounds[index] < previous ||
                (index > 0 && bounds[index] == previous) ||
                bounds[index] > domainEnd || (!isLast && bounds[index] == domainEnd)) return false;
            previous = bounds[index];
        }

        var children = new PdfColorFunction[functions.Items.Count];
        for (int index = 0; index < children.Length; index++) {
            if (!TryCreateFunction(
                    functions.Items[index],
                    1,
                    outputCount,
                    objects,
                    maxDecodedStreamBytes,
                    depth + 1,
                    activeFunctions,
                    functionCache,
                    ref parsedFunctionNodes,
                    ref retainedFunctionBytes,
                    ref remainingCalculatorValidationWork,
                    out children[index])) return false;
        }

        double[] discontinuities = CreateStitchingDiscontinuities(domain, bounds, encode, children);
        double[] requiredBreakpoints = domain.Concat(discontinuities).Distinct().ToArray();
        double[] breakpoints = CreateStitchingBreakpoints(domain, bounds, encode, children, requiredBreakpoints);
        function = new PdfColorFunction(
            inputCount,
            outputCount,
            domain,
            range,
            (values, output, outputOffset) => EvaluateStitching(
                values[0], output, outputOffset, domain, bounds, encode, children),
            breakpoints,
            discontinuities,
            evaluationCost: children.Max(static child => child.EvaluationCost));
        return true;
    }
    private static bool EvaluateType2(
        double[] values,
        double[] output,
        int outputOffset,
        double[] c0,
        double[] c1,
        double exponent) {
        double factor = Math.Pow(values[0], exponent);
        if (!IsFinite(factor)) return false;
        for (int index = 0; index < c0.Length; index++) {
            output[outputOffset + index] = c0[index] + factor * (c1[index] - c0[index]);
        }
        return true;
    }

    private static bool EvaluateStitching(
        double value,
        double[] output,
        int outputOffset,
        double[] domain,
        double[] bounds,
        double[] encode,
        PdfColorFunction[] children) {
        int childIndex = 0;
        bool selectsLeftEndpointFunction = bounds.Length > 0 && value == domain[0] && bounds[0] == domain[0];
        if (!selectsLeftEndpointFunction) {
            while (childIndex < bounds.Length && value >= bounds[childIndex]) childIndex++;
        }
        double sourceStart = childIndex == 0 ? domain[0] : bounds[childIndex - 1];
        double sourceEnd = childIndex == bounds.Length ? domain[1] : bounds[childIndex];
        double encoded = PdfColorFunction.Interpolate(
            value,
            sourceStart,
            sourceEnd,
            encode[childIndex * 2],
            encode[childIndex * 2 + 1]);
        if (!IsFinite(encoded)) return false;
        double[] input = _scalarInput ??= new double[1];
        input[0] = encoded;
        return children[childIndex].TryEvaluate(input, output, outputOffset);
    }

    private static bool EvaluateFunctionArray(
        double[] values,
        double[] output,
        int outputOffset,
        PdfColorFunction[] components) {
        for (int index = 0; index < components.Length; index++) {
            if (!components[index].TryEvaluate(values, output, outputOffset + index)) return false;
        }
        return true;
    }

    private static int SumEvaluationCost(IEnumerable<PdfColorFunction> functions) {
        int total = 0;
        foreach (PdfColorFunction function in functions) {
            total = checked(total + function.EvaluationCost);
        }
        return total;
    }

    private static double[] CreateUniformBreakpoints(double[] domain, int pointCount) {
        var result = new double[pointCount];
        for (int index = 0; index < pointCount; index++) {
            result[index] = PdfColorFunction.Interpolate(index, 0D, pointCount - 1D, domain[0], domain[1]);
        }
        return result;
    }

    private static double[] CreateStitchingBreakpoints(
        double[] domain,
        double[] bounds,
        double[] encode,
        PdfColorFunction[] children,
        IReadOnlyCollection<double> requiredBreakpoints) {
        var result = new List<double>(bounds.Length + 8);
        result.AddRange(domain);
        result.AddRange(bounds);
        for (int index = 0; index < children.Length; index++) {
            double encodedStart = encode[index * 2];
            double encodedEnd = encode[index * 2 + 1];
            if (encodedStart == encodedEnd) continue;
            double sourceStart = index == 0 ? domain[0] : bounds[index - 1];
            double sourceEnd = index == bounds.Length ? domain[1] : bounds[index];
            foreach (double childBreakpoint in children[index].Breakpoints) {
                double input = PdfColorFunction.Interpolate(childBreakpoint, encodedStart, encodedEnd, sourceStart, sourceEnd);
                if (IsFinite(input) && input > sourceStart && input < sourceEnd) result.Add(input);
            }
        }
        return LimitSuggestedPoints(result, requiredBreakpoints);
    }

    private static double[] CreateStitchingDiscontinuities(
        double[] domain,
        double[] bounds,
        double[] encode,
        PdfColorFunction[] children) {
        var result = new List<double>(bounds.Length + 8);
        result.AddRange(bounds);
        for (int index = 0; index < children.Length; index++) {
            double encodedStart = encode[index * 2];
            double encodedEnd = encode[index * 2 + 1];
            if (encodedStart == encodedEnd) continue;
            double sourceStart = index == 0 ? domain[0] : bounds[index - 1];
            double sourceEnd = index == bounds.Length ? domain[1] : bounds[index];
            foreach (double childDiscontinuity in children[index].Discontinuities) {
                double input = PdfColorFunction.Interpolate(childDiscontinuity, encodedStart, encodedEnd, sourceStart, sourceEnd);
                if (IsFinite(input) && input >= sourceStart && input <= sourceEnd) result.Add(input);
            }
        }
        int reservedDomainBoundaries = domain.Count(boundary => !result.Contains(boundary));
        return LimitSuggestedPoints(
            result,
            bounds,
            MaxSuggestedSampleBreakpoints - reservedDomainBoundaries);
    }

    private static double[] LimitSuggestedPoints(
        IEnumerable<double> values,
        IReadOnlyCollection<double>? required,
        int maxPoints = MaxSuggestedSampleBreakpoints) {
        double[] ordered = values.Distinct().OrderBy(static value => value).ToArray();
        if (ordered.Length <= maxPoints) return ordered;

        var selected = required == null
            ? new SortedSet<double>()
            : new SortedSet<double>(required);
        int remaining = maxPoints - selected.Count;
        if (remaining <= 0) return selected.Take(maxPoints).ToArray();

        double[] optional = ordered.Where(value => !selected.Contains(value)).ToArray();
        if (optional.Length <= remaining) {
            selected.UnionWith(optional);
            return selected.ToArray();
        }

        if (remaining == 1) {
            selected.Add(optional[optional.Length / 2]);
        } else {
            for (int index = 0; index < remaining; index++) {
                int sourceIndex = (int)Math.Round(index * (optional.Length - 1D) / (remaining - 1D));
                selected.Add(optional[sourceIndex]);
            }
        }
        return selected.ToArray();
    }

    private static bool TryReadIntervals(
        PdfDictionary dictionary,
        string key,
        int count,
        Dictionary<int, PdfIndirectObject> objects,
        bool allowEqual,
        bool required,
        out double[] values) {
        values = Array.Empty<double>();
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value)) return !required;
        values = ReadNumberArray(value, objects);
        if (values.Length != count * 2) return false;
        for (int index = 0; index < count; index++) {
            double minimum = values[index * 2];
            double maximum = values[index * 2 + 1];
            if (!IsFinite(minimum) || !IsFinite(maximum) || (allowEqual ? maximum < minimum : maximum <= minimum)) return false;
        }
        return true;
    }

    private static bool TryReadOptionalRange(
        PdfDictionary dictionary,
        int outputCount,
        Dictionary<int, PdfIndirectObject> objects,
        out double[]? range) {
        range = null;
        if (!TryResolveOptionalEntry(dictionary, "Range", objects, out PdfObject? rangeObject, out bool hasRange)) return false;
        if (!hasRange) return true;
        double[] values = ReadNumberArray(rangeObject, objects);
        if (values.Length != outputCount * 2) return false;
        for (int index = 0; index < outputCount; index++) {
            double minimum = values[index * 2];
            double maximum = values[index * 2 + 1];
            if (!IsFinite(minimum) || !IsFinite(maximum) || maximum < minimum) return false;
        }
        range = values;
        return true;
    }

    private static bool TryReadIntegerArray(
        PdfDictionary dictionary,
        string key,
        int count,
        Dictionary<int, PdfIndirectObject> objects,
        out int[] values) {
        values = Array.Empty<int>();
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value) ||
            !TryResolveObject(value, objects, out PdfObject? resolved) ||
            resolved is not PdfArray array || array.Items.Count != count) return false;
        values = new int[count];
        for (int index = 0; index < count; index++) {
            if (TryReadInteger(array.Items[index], objects) is not int number) return false;
            values[index] = number;
        }
        return true;
    }

    private static bool HasUnitIntervals(IReadOnlyList<double> values, int count) {
        if (values.Count != count * 2) return false;
        for (int index = 0; index < count; index++) {
            if (values[index * 2] != 0D || values[index * 2 + 1] != 1D) return false;
        }
        return true;
    }

    private static bool HasFinitePairs(double[] values, int count) =>
        values.Length == count * 2 && values.All(static value => IsFinite(value));

    private static bool TryReadOptionalObject(
        PdfDictionary dictionary,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfObject? value) {
        value = null;
        if (!dictionary.Items.TryGetValue(key, out PdfObject? candidate)) return true;
        if (!TryResolveObject(candidate, objects, out PdfObject? resolved)) return false;
        if (resolved is null or PdfNull) return true;
        value = resolved;
        return true;
    }

    private static bool IsSupportedSampleWidth(int bitsPerSample) =>
        bitsPerSample is 1 or 2 or 4 or 8 or 12 or 16 or 24 or 32;

    private static int? TryReadInteger(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) {
        double? number = TryReadNumber(value, objects);
        if (number is not double resolved || Math.Truncate(resolved) != resolved ||
            resolved < int.MinValue || resolved > int.MaxValue) return null;
        return (int)resolved;
    }

    private static double? TryReadNumber(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) {
        return TryResolveObject(value, objects, out PdfObject? resolved) && resolved is PdfNumber number && IsFinite(number.Value)
            ? number.Value
            : null;
    }

    private static double[] ReadNumberArray(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) {
        if (!TryResolveObject(value, objects, out PdfObject? resolved) || resolved is not PdfArray array) return Array.Empty<double>();
        var result = new double[array.Items.Count];
        for (int index = 0; index < result.Length; index++) {
            if (TryReadNumber(array.Items[index], objects) is not double number) return Array.Empty<double>();
            result[index] = number;
        }
        return result;
    }

    private static bool TryResolveObject(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfObject? resolved) {
        resolved = value;
        var visited = new HashSet<long>();
        for (int depth = 0; resolved is PdfReference reference && depth <= MaxFunctionDepth; depth++) {
            long key = ((long)reference.ObjectNumber << 32) ^ (uint)reference.Generation;
            if (!visited.Add(key) || !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
                resolved = null;
                return false;
            }
            resolved = indirect.Value;
        }
        return resolved is not PdfReference;
    }

    private static bool TryReadOptionalNumberArray(
        PdfDictionary dictionary,
        string key,
        double[] defaultValue,
        Dictionary<int, PdfIndirectObject> objects,
        out double[] values) {
        values = defaultValue;
        if (!TryResolveOptionalEntry(dictionary, key, objects, out PdfObject? resolved, out bool hasValue)) return false;
        if (!hasValue) return true;
        values = ReadNumberArray(resolved, objects);
        return true;
    }

    private static bool TryResolveOptionalEntry(
        PdfDictionary dictionary,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfObject? resolved,
        out bool hasValue) {
        resolved = null;
        hasValue = false;
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value)) return true;
        if (!TryResolveObject(value, objects, out resolved) || resolved == null) return false;
        hasValue = resolved is not PdfNull;
        return true;
    }
    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private sealed class PdfObjectReferenceComparer : IEqualityComparer<PdfObject> {
        internal static readonly PdfObjectReferenceComparer Instance = new PdfObjectReferenceComparer();
        public bool Equals(PdfObject? left, PdfObject? right) => ReferenceEquals(left, right);
        public int GetHashCode(PdfObject value) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value);
    }
}
