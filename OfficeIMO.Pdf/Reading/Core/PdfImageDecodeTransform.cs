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
        return TryCreate(dictionary, componentCount, objects, out var transform) &&
            !transform.IsIdentity(0, 1)
            ? transform
            : null;
    }

    internal static PdfImageDecodeTransform? CreateIndexed(
        PdfDictionary dictionary,
        int highValue,
        Dictionary<int, PdfIndirectObject> objects) {
        return TryCreate(dictionary, 1, objects, out var transform) &&
            !transform.IsIdentity(0, highValue)
            ? transform
            : null;
    }

    internal byte TransformColorComponent(byte sample, int componentIndex) {
        double decoded = Decode(sample / 255D, componentIndex);
        return ClampToByte(decoded * 255D);
    }

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

    private bool IsIdentity(double expectedMinimum, double expectedMaximum) {
        for (int i = 0; i < _minimums.Length; i++) {
            if (System.Math.Abs(_minimums[i] - expectedMinimum) > double.Epsilon ||
                System.Math.Abs(_maximums[i] - expectedMaximum) > double.Epsilon) {
                return false;
            }
        }

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
