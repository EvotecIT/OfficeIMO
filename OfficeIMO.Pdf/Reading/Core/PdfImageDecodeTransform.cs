namespace OfficeIMO.Pdf;

internal sealed class PdfImageDecodeTransform {
    private readonly double[] _minimums;
    private readonly double[] _maximums;

    private PdfImageDecodeTransform(double[] minimums, double[] maximums) {
        _minimums = minimums;
        _maximums = maximums;
    }

    internal static bool TryCreateColor(
        PdfDictionary dictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageDecodeTransform? transform) {
        // An explicit identity Decode array is still authored behavior. Calibrated and ICCBased
        // spaces have non-unit default ranges, so callers must be able to distinguish it from an
        // omitted Decode entry.
        return TryCreate(dictionary, componentCount, objects, out transform);
    }

    internal static bool TryCreateColorDeclaration(
        PdfDictionary dictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageDecodeTransform? transform) =>
        TryCreate(dictionary, componentCount, objects, out transform);

    internal static bool TryCreateIndexed(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageDecodeTransform? transform) =>
        TryCreate(dictionary, 1, objects, out transform);

    internal static bool IsIdentityColorDecodeOrAbsent(
        PdfDictionary dictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects) {
        if (!TryCreate(dictionary, componentCount, objects, out PdfImageDecodeTransform? transform)) return false;
        if (transform is null) return true;
        for (int component = 0; component < componentCount; component++) {
            if (transform._minimums[component] != 0D || transform._maximums[component] != 1D) return false;
        }
        return true;
    }

    private static bool TryResolveReferenceChain(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfObject? resolved) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        resolved = value;
        while (resolved is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject indirect)) {
                resolved = null;
                return false;
            }
            resolved = indirect.Value;
        }
        return true;
    }

    internal static bool TryCreateIndexedDeclaration(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageDecodeTransform? transform) =>
        TryCreate(dictionary, 1, objects, out transform);

    internal byte TransformColorComponent(byte sample, int componentIndex) {
        double decoded = TransformColorComponentValue(sample, componentIndex);
        return ClampToByte(decoded * 255D);
    }

    internal double TransformColorComponentValue(byte sample, int componentIndex) =>
        Decode(sample / 255D, componentIndex);

    internal int TransformIndexedSample(int sample, int bitsPerComponent, int highValue) {
        int maxSample = (1 << bitsPerComponent) - 1;
        if (maxSample <= 0) {
            return 0;
        }

        double decoded = Decode(sample / (double)maxSample, 0);
        int value = (int)System.Math.Round(decoded, System.MidpointRounding.AwayFromZero);
        if (value < 0) {
            return 0;
        }

        return value > highValue ? highValue : value;
    }

    private static bool TryCreate(
        PdfDictionary dictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageDecodeTransform? transform) {
        transform = null;
        if (componentCount <= 0) {
            return false;
        }

        if (!dictionary.Items.TryGetValue("Decode", out PdfObject? decodeObj)) return true;
        if (!TryResolveReferenceChain(decodeObj, objects, out PdfObject? resolvedDecode)) return false;
        if (resolvedDecode is PdfNull) return true;
        if (resolvedDecode is not PdfArray decodeArray || decodeArray.Items.Count != componentCount * 2) return false;

        var minimums = new double[componentCount];
        var maximums = new double[componentCount];
        for (int component = 0; component < componentCount; component++) {
            if (!TryResolveReferenceChain(decodeArray.Items[component * 2], objects, out PdfObject? resolvedMinimum) ||
                resolvedMinimum is not PdfNumber minimum ||
                !TryResolveReferenceChain(decodeArray.Items[component * 2 + 1], objects, out PdfObject? resolvedMaximum) ||
                resolvedMaximum is not PdfNumber maximum) {
                return false;
            }

            if (double.IsNaN(minimum.Value) || double.IsInfinity(minimum.Value) ||
                double.IsNaN(maximum.Value) || double.IsInfinity(maximum.Value)) return false;
            minimums[component] = minimum.Value;
            maximums[component] = maximum.Value;
        }

        transform = new PdfImageDecodeTransform(minimums, maximums);
        return true;
    }

    private double Decode(double normalizedSample, int componentIndex) {
        return _minimums[componentIndex] + normalizedSample * (_maximums[componentIndex] - _minimums[componentIndex]);
    }

    private static byte ClampToByte(double value) {
        if (value <= 0) {
            return 0;
        }

        if (value >= 255) {
            return 255;
        }

        return (byte)System.Math.Round(value, System.MidpointRounding.AwayFromZero);
    }
}
