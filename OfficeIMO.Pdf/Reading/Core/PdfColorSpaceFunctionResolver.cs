using System.Text;

namespace OfficeIMO.Pdf;

internal delegate bool PdfColorSpaceTintTransform(IReadOnlyList<double> components, double[] output);

/// <summary>Resolves the bounded tint functions shared by content and image color-space projection.</summary>
internal static class PdfColorSpaceFunctionResolver {
    private static readonly double[] DefaultC0 = { 0D };
    private static readonly double[] DefaultC1 = { 1D };

    internal static bool TryCreateTintTransform(
        PdfObject? value,
        int inputCount,
        int outputCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        out PdfColorSpaceTintTransform transform) {
        transform = null!;
        PdfObject? resolved = ResolveObject(value, objects);
        PdfDictionary? dictionary = resolved switch {
            PdfStream stream => stream.Dictionary,
            PdfDictionary direct => direct,
            _ => null
        };
        if (dictionary == null) return false;

        int? functionType = TryReadInteger(dictionary.Items.TryGetValue("FunctionType", out PdfObject? type) ? type : null, objects);
        if (functionType == 2 && inputCount == 1) {
            if (!TryReadOptionalNumberArray(dictionary, "C0", DefaultC0, objects, out double[] c0) ||
                !TryReadOptionalNumberArray(dictionary, "C1", DefaultC1, objects, out double[] c1) ||
                !dictionary.Items.TryGetValue("N", out PdfObject? exponentValue) ||
                ResolveObject(exponentValue, objects) is not PdfNumber exponentNumber) return false;
            double exponent = exponentNumber.Value;
            if (c0.Length != outputCount || c1.Length != outputCount ||
                c0.Any(value => !IsFinite(value)) || c1.Any(value => !IsFinite(value)) ||
                !IsFinite(exponent) || exponent <= 0D ||
                !HasUnitFunctionBounds(dictionary, inputCount, outputCount, requireRange: false, objects)) return false;
            if (!TryResolveOptionalEntry(dictionary, "Range", objects, out _, out bool clipOutputsToUnitRange)) return false;
            transform = (components, output) => EvaluateType2(components, output, c0, c1, exponent, clipOutputsToUnitRange);
            return true;
        }

        if (functionType == 4 && resolved is PdfStream calculator &&
            HasUnitFunctionBounds(dictionary, inputCount, outputCount, requireRange: true, objects) &&
            TryReadBoundedCalculatorProgram(calculator, inputCount, outputCount, objects, maxDecodedStreamBytes, out int duplicateCount)) {
            transform = (components, output) => EvaluateIdentity(components, output, inputCount, outputCount, duplicateCount);
            return true;
        }

        return false;
    }

    private static bool HasUnitFunctionBounds(
        PdfDictionary dictionary,
        int inputCount,
        int outputCount,
        bool requireRange,
        Dictionary<int, PdfIndirectObject> objects) {
        if (!dictionary.Items.TryGetValue("Domain", out PdfObject? domainObject) ||
            !HasUnitIntervals(ReadNumberArray(domainObject, objects), inputCount)) return false;
        if (!TryResolveOptionalEntry(dictionary, "Range", objects, out PdfObject? rangeObject, out bool hasRange)) return false;
        return hasRange ? HasUnitIntervals(ReadNumberArray(rangeObject, objects), outputCount) : !requireRange;
    }

    private static bool HasUnitIntervals(double[] values, int count) {
        if (values.Length != count * 2) return false;
        for (int index = 0; index < count; index++) {
            if (values[index * 2] != 0D || values[index * 2 + 1] != 1D) return false;
        }
        return true;
    }

    private static bool TryReadBoundedCalculatorProgram(
        PdfStream stream,
        int inputCount,
        int outputCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        out int duplicateCount) {
        duplicateCount = 0;
        if (Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, objects).Count != 0) return false;
        int decodeLimit = Math.Min(257, maxDecodedStreamBytes);
        byte[] bytes;
        try {
            bytes = Filters.StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, decodeLimit);
        } catch (PdfReadLimitException) {
            throw;
        } catch (InvalidDataException) {
            return false;
        }
        if (bytes.Length > 256) return false;
        string program = Encoding.ASCII.GetString(bytes).Trim();
        if (program.Length < 2 || program[0] != '{' || program[program.Length - 1] != '}') return false;
        string body = program.Substring(1, program.Length - 2).Trim();
        if (body.Length == 0) return inputCount == outputCount;
        string[] tokens = body.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (inputCount != 1 || tokens.Length > 32 || tokens.Any(token => !string.Equals(token, "dup", StringComparison.Ordinal))) return false;
        duplicateCount = tokens.Length;
        return outputCount == duplicateCount + 1;
    }

    private static bool EvaluateType2(
        IReadOnlyList<double> components,
        double[] output,
        double[] c0,
        double[] c1,
        double exponent,
        bool clipOutputsToUnitRange) {
        if (components.Count < 1 || output.Length < c0.Length || !IsFinite(components[0])) return false;
        double factor = Math.Pow(Clamp01(components[0]), exponent);
        for (int index = 0; index < c0.Length; index++) {
            double value = c0[index] + factor * (c1[index] - c0[index]);
            output[index] = clipOutputsToUnitRange ? Clamp01(value) : value;
        }
        return true;
    }

    private static bool EvaluateIdentity(
        IReadOnlyList<double> components,
        double[] output,
        int inputCount,
        int outputCount,
        int duplicateCount) {
        if (components.Count < inputCount || output.Length < outputCount) return false;
        for (int index = 0; index < inputCount; index++) if (!IsFinite(components[index])) return false;
        if (duplicateCount == 0) {
            for (int index = 0; index < outputCount; index++) output[index] = Clamp01(components[index]);
        } else {
            double value = Clamp01(components[0]);
            for (int index = 0; index < outputCount; index++) output[index] = value;
        }
        return true;
    }

    private static int? TryReadInteger(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) {
        if (ResolveObject(value, objects) is not PdfNumber number ||
            !IsFinite(number.Value) || Math.Truncate(number.Value) != number.Value ||
            number.Value < int.MinValue || number.Value > int.MaxValue) return null;
        return (int)number.Value;
    }

    private static double[] ReadNumberArray(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) {
        if (ResolveObject(value, objects) is not PdfArray array) return Array.Empty<double>();
        var result = new double[array.Items.Count];
        for (int index = 0; index < result.Length; index++) {
            if (ResolveObject(array.Items[index], objects) is not PdfNumber number) return Array.Empty<double>();
            result[index] = number.Value;
        }
        return result;
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
        resolved = ResolveObject(value, objects);
        if (resolved == null) return false;
        hasValue = resolved is not PdfNull;
        return true;
    }

    private static PdfObject? ResolveObject(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        PdfObject? resolved = value;
        while (resolved is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject indirect)) return null;
            resolved = indirect.Value;
        }
        return resolved;
    }

    private static double Clamp01(double value) => value < 0D ? 0D : value > 1D ? 1D : value;
    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
