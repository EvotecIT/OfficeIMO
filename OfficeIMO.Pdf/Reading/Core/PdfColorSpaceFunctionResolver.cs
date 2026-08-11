using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>Resolves the bounded tint functions shared by content and image color-space projection.</summary>
internal static class PdfColorSpaceFunctionResolver {
    internal static bool TryCreateTintTransform(
        PdfObject? value,
        int inputCount,
        int outputCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform) {
        transform = null!;
        PdfObject? resolved = PdfObjectLookup.Resolve(objects, value);
        PdfDictionary? dictionary = resolved switch {
            PdfStream stream => stream.Dictionary,
            PdfDictionary direct => direct,
            _ => null
        };
        if (dictionary == null) return false;

        int? functionType = TryReadInteger(dictionary.Items.TryGetValue("FunctionType", out PdfObject? type) ? type : null, objects);
        if (functionType == 2 && inputCount == 1) {
            double[] c0 = dictionary.Items.TryGetValue("C0", out PdfObject? c0Value)
                ? ReadNumberArray(c0Value, objects)
                : new[] { 0D };
            double[] c1 = dictionary.Items.TryGetValue("C1", out PdfObject? c1Value)
                ? ReadNumberArray(c1Value, objects)
                : new[] { 1D };
            if (!dictionary.Items.TryGetValue("N", out PdfObject? exponentValue) ||
                PdfObjectLookup.Resolve(objects, exponentValue) is not PdfNumber exponentNumber) return false;
            double exponent = exponentNumber.Value;
            if (c0.Length != outputCount || c1.Length != outputCount ||
                c0.Any(value => !IsFinite(value)) || c1.Any(value => !IsFinite(value)) ||
                !IsFinite(exponent) || exponent <= 0D ||
                !HasUnitFunctionBounds(dictionary, inputCount, outputCount, requireRange: false, objects)) return false;
            transform = components => EvaluateType2(components, c0, c1, exponent);
            return true;
        }

        if (functionType == 4 && resolved is PdfStream calculator &&
            HasUnitFunctionBounds(dictionary, inputCount, outputCount, requireRange: true, objects) &&
            TryReadBoundedCalculatorProgram(calculator, inputCount, outputCount, objects, maxDecodedStreamBytes, out int duplicateCount)) {
            transform = components => EvaluateIdentity(components, inputCount, outputCount, duplicateCount);
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
        if (!dictionary.Items.TryGetValue("Range", out PdfObject? rangeObject)) return !requireRange;
        return HasUnitIntervals(ReadNumberArray(rangeObject, objects), outputCount);
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
        byte[] bytes = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects, decodeLimit);
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

    private static double[]? EvaluateType2(
        IReadOnlyList<double> components,
        double[] c0,
        double[] c1,
        double exponent) {
        if (components.Count < 1 || !IsFinite(components[0])) return null;
        double factor = Math.Pow(Clamp01(components[0]), exponent);
        var result = new double[c0.Length];
        for (int index = 0; index < result.Length; index++) result[index] = c0[index] + factor * (c1[index] - c0[index]);
        return result;
    }

    private static double[]? EvaluateIdentity(
        IReadOnlyList<double> components,
        int inputCount,
        int outputCount,
        int duplicateCount) {
        if (components.Count < inputCount || components.Take(inputCount).Any(value => !IsFinite(value))) return null;
        if (duplicateCount == 0) return components.Take(outputCount).Select(Clamp01).ToArray();
        return Enumerable.Repeat(Clamp01(components[0]), outputCount).ToArray();
    }

    private static int? TryReadInteger(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) {
        if (PdfObjectLookup.Resolve(objects, value) is not PdfNumber number ||
            !IsFinite(number.Value) || Math.Truncate(number.Value) != number.Value ||
            number.Value < int.MinValue || number.Value > int.MaxValue) return null;
        return (int)number.Value;
    }

    private static double[] ReadNumberArray(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) {
        if (PdfObjectLookup.Resolve(objects, value) is not PdfArray array) return Array.Empty<double>();
        var result = new double[array.Items.Count];
        for (int index = 0; index < result.Length; index++) {
            if (PdfObjectLookup.Resolve(objects, array.Items[index]) is not PdfNumber number) return Array.Empty<double>();
            result[index] = number.Value;
        }
        return result;
    }

    private static double Clamp01(double value) => value < 0D ? 0D : value > 1D ? 1D : value;
    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
