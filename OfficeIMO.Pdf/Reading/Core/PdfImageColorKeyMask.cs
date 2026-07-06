namespace OfficeIMO.Pdf;

internal sealed class PdfImageColorKeyMask {
    private readonly int[] _minimums;
    private readonly int[] _maximums;

    private PdfImageColorKeyMask(int[] minimums, int[] maximums) {
        _minimums = minimums;
        _maximums = maximums;
    }

    internal static PdfImageColorKeyMask? Create(
        PdfDictionary dictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects) {
        if (componentCount <= 0 ||
            !dictionary.Items.TryGetValue("Mask", out var maskObj) ||
            PdfObjectLookup.Resolve(objects, maskObj) is not PdfArray maskArray ||
            maskArray.Items.Count < componentCount * 2) {
            return null;
        }

        var minimums = new int[componentCount];
        var maximums = new int[componentCount];
        for (int component = 0; component < componentCount; component++) {
            if (PdfObjectLookup.Resolve(objects, maskArray.Items[component * 2]) is not PdfNumber minimum ||
                PdfObjectLookup.Resolve(objects, maskArray.Items[component * 2 + 1]) is not PdfNumber maximum) {
                return null;
            }

            minimums[component] = ClampSample((int)minimum.Value);
            maximums[component] = ClampSample((int)maximum.Value);
        }

        return new PdfImageColorKeyMask(minimums, maximums);
    }

    internal bool IsTransparent(byte[] samples, int sampleOffset) {
        for (int component = 0; component < _minimums.Length; component++) {
            int sample = samples[sampleOffset + component];
            if (sample < _minimums[component] || sample > _maximums[component]) {
                return false;
            }
        }

        return true;
    }

    internal bool IsTransparentSample(int sample) {
        return _minimums.Length == 1 &&
            sample >= _minimums[0] &&
            sample <= _maximums[0];
    }

    private static int ClampSample(int value) {
        if (value < 0) {
            return 0;
        }

        return value > 255 ? 255 : value;
    }
}
