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
            PdfObjectLookup.ResolveChain(objects, maskObj) is not PdfArray maskArray ||
            maskArray.Items.Count < componentCount * 2) {
            return null;
        }

        var minimums = new int[componentCount];
        var maximums = new int[componentCount];
        for (int component = 0; component < componentCount; component++) {
            if (PdfObjectLookup.ResolveChain(objects, maskArray.Items[component * 2]) is not PdfNumber minimum ||
                PdfObjectLookup.ResolveChain(objects, maskArray.Items[component * 2 + 1]) is not PdfNumber maximum) {
                return null;
            }

            minimums[component] = ClampSample((int)minimum.Value);
            maximums[component] = ClampSample((int)maximum.Value);
        }

        return new PdfImageColorKeyMask(minimums, maximums);
    }

    internal static bool TryCreateDeclaration(
        PdfDictionary dictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageColorKeyMask? mask) {
        mask = null;
        if (!dictionary.Items.TryGetValue("Mask", out PdfObject? maskObject)) return true;
        PdfObject? resolvedMask = PdfObjectLookup.ResolveChain(objects, maskObject);
        if (resolvedMask is null or PdfNull ||
            resolvedMask is PdfName { Name: "None" } ||
            resolvedMask is PdfStream) {
            return true;
        }
        if (componentCount <= 0 ||
            resolvedMask is not PdfArray maskArray ||
            maskArray.Items.Count != componentCount * 2) {
            return false;
        }

        var minimums = new int[componentCount];
        var maximums = new int[componentCount];
        for (int component = 0; component < componentCount; component++) {
            if (PdfObjectLookup.ResolveChain(objects, maskArray.Items[component * 2]) is not PdfNumber minimum ||
                PdfObjectLookup.ResolveChain(objects, maskArray.Items[component * 2 + 1]) is not PdfNumber maximum ||
                double.IsNaN(minimum.Value) ||
                double.IsInfinity(minimum.Value) ||
                double.IsNaN(maximum.Value) ||
                double.IsInfinity(maximum.Value) ||
                minimum.Value != Math.Truncate(minimum.Value) ||
                maximum.Value != Math.Truncate(maximum.Value) ||
                minimum.Value < int.MinValue ||
                minimum.Value > int.MaxValue ||
                maximum.Value < int.MinValue ||
                maximum.Value > int.MaxValue) {
                return false;
            }
            minimums[component] = ClampSample((int)minimum.Value);
            maximums[component] = ClampSample((int)maximum.Value);
        }
        mask = new PdfImageColorKeyMask(minimums, maximums);
        return true;
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
