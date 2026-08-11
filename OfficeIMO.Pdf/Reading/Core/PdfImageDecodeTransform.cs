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

    internal static bool TryCreateColorDeclaration(
        PdfDictionary dictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageDecodeTransform? transform) =>
        TryCreateDeclaration(dictionary, componentCount, objects, out transform);

    internal static PdfImageDecodeTransform? CreateIndexed(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects) {
        return TryCreate(dictionary, 1, objects, out var transform) ? transform : null;
    }

    internal static bool TryCreateIndexedDeclaration(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageDecodeTransform? transform) =>
        TryCreateDeclaration(dictionary, 1, objects, out transform);

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

    private static bool TryCreateDeclaration(
        PdfDictionary dictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageDecodeTransform? transform) {
        transform = null;
        if (!dictionary.Items.TryGetValue("Decode", out PdfObject? decodeObject)) return true;
        PdfObject? resolvedDecode = PdfObjectLookup.ResolveChain(objects, decodeObject);
        if (resolvedDecode is null or PdfNull) return true;
        if (componentCount <= 0 ||
            resolvedDecode is not PdfArray decodeArray ||
            decodeArray.Items.Count != componentCount * 2) {
            return false;
        }

        var minimums = new double[componentCount];
        var maximums = new double[componentCount];
        for (int component = 0; component < componentCount; component++) {
            if (PdfObjectLookup.ResolveChain(objects, decodeArray.Items[component * 2]) is not PdfNumber minimum ||
                PdfObjectLookup.ResolveChain(objects, decodeArray.Items[component * 2 + 1]) is not PdfNumber maximum ||
                double.IsNaN(minimum.Value) ||
                double.IsInfinity(minimum.Value) ||
                double.IsNaN(maximum.Value) ||
                double.IsInfinity(maximum.Value)) {
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
