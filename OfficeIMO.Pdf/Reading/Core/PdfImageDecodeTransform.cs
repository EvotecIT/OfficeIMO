namespace OfficeIMO.Pdf;

internal sealed class PdfImageDecodeTransform {
    private readonly double[] _minimums;
    private readonly double[] _maximums;

    private PdfImageDecodeTransform(double[] minimums, double[] maximums) {
        _minimums = minimums;
        _maximums = maximums;
    }

    internal static PdfImageDecodeTransform? CreateColor(
        PdfDictionary dictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects) {
        // An explicit identity Decode array is still authored behavior. Calibrated and ICCBased
        // spaces have non-unit default ranges, so callers must be able to distinguish it from an
        // omitted Decode entry.
        return TryCreate(dictionary, componentCount, objects, out var transform) ? transform : null;
    }

    internal static PdfImageDecodeTransform? CreateIndexed(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects) {
        return TryCreate(dictionary, 1, objects, out var transform) ? transform : null;
    }

    internal static bool IsIdentityColorDecodeOrAbsent(
        PdfDictionary dictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects) {
        if (componentCount <= 0 || !dictionary.Items.TryGetValue("Decode", out PdfObject? decodeObject)) {
            return componentCount > 0;
        }

        PdfObject? resolved = ResolveReferenceChain(decodeObject, objects);
        if (resolved is PdfNull) return true;
        if (resolved is not PdfArray decodeArray || decodeArray.Items.Count != componentCount * 2) return false;

        for (int component = 0; component < componentCount; component++) {
            if (ResolveReferenceChain(decodeArray.Items[component * 2], objects) is not PdfNumber { Value: 0D } ||
                ResolveReferenceChain(decodeArray.Items[component * 2 + 1], objects) is not PdfNumber { Value: 1D }) {
                return false;
            }
        }

        return true;
    }

    private static PdfObject? ResolveReferenceChain(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        PdfObject? resolved = value;
        while (resolved is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject indirect)) {
                return null;
            }
            resolved = indirect.Value;
        }
        return resolved;
    }

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
        out PdfImageDecodeTransform transform) {
        transform = null!;
        if (componentCount <= 0 ||
            !dictionary.Items.TryGetValue("Decode", out var decodeObj) ||
            PdfObjectLookup.Resolve(objects, decodeObj) is not PdfArray decodeArray ||
            decodeArray.Items.Count < componentCount * 2) {
            return false;
        }

        var minimums = new double[componentCount];
        var maximums = new double[componentCount];
        for (int component = 0; component < componentCount; component++) {
            if (PdfObjectLookup.Resolve(objects, decodeArray.Items[component * 2]) is not PdfNumber minimum ||
                PdfObjectLookup.Resolve(objects, decodeArray.Items[component * 2 + 1]) is not PdfNumber maximum) {
                return false;
            }

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
