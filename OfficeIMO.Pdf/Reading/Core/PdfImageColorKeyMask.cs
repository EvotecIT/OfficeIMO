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
        int bitsPerComponent,
        Dictionary<int, PdfIndirectObject> objects) {
        if (componentCount <= 0 || bitsPerComponent <= 0 || bitsPerComponent > 16 ||
            !dictionary.Items.TryGetValue("Mask", out var maskObj) ||
            !PdfObjectLookup.TryResolveReferenceChain(objects, maskObj, out PdfObject? resolvedMask) ||
            resolvedMask is not PdfArray maskArray ||
            maskArray.Items.Count != componentCount * 2) {
            return null;
        }

        int maximumSample = (1 << bitsPerComponent) - 1;
        var minimums = new int[componentCount];
        var maximums = new int[componentCount];
        for (int component = 0; component < componentCount; component++) {
            if (!PdfObjectLookup.TryResolveReferenceChain(objects, maskArray.Items[component * 2], out PdfObject? resolvedMinimum) ||
                resolvedMinimum is not PdfNumber minimum ||
                !PdfObjectLookup.TryResolveReferenceChain(objects, maskArray.Items[component * 2 + 1], out PdfObject? resolvedMaximum) ||
                resolvedMaximum is not PdfNumber maximum ||
                !TryReadSample(minimum.Value, maximumSample, out int minimumSample) ||
                !TryReadSample(maximum.Value, maximumSample, out int maximumSampleValue) ||
                minimumSample > maximumSampleValue) {
                return null;
            }

            minimums[component] = minimumSample;
            maximums[component] = maximumSampleValue;
        }

        return new PdfImageColorKeyMask(minimums, maximums);
    }

    internal static bool TryCreateDeclaration(
        PdfDictionary dictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageColorKeyMask? mask) =>
        TryCreateDeclaration(dictionary, componentCount, bitsPerComponent: 8, objects, out mask);

    internal static bool TryCreateDeclaration(
        PdfDictionary dictionary,
        int componentCount,
        int bitsPerComponent,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageColorKeyMask? mask) {
        mask = null;
        if (!dictionary.Items.TryGetValue("Mask", out PdfObject? maskObject)) return true;
        if (!PdfObjectLookup.TryResolveReferenceChain(objects, maskObject, out PdfObject? resolvedMask)) return false;
        if (resolvedMask is PdfNull ||
            resolvedMask is PdfName { Name: "None" } ||
            resolvedMask is PdfStream) {
            return true;
        }
        if (componentCount <= 0 ||
            resolvedMask is not PdfArray maskArray ||
            maskArray.Items.Count != componentCount * 2) {
            return false;
        }

        mask = Create(dictionary, componentCount, bitsPerComponent, objects);
        return mask is not null;
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

    private static bool TryReadSample(double value, int maximumSample, out int sample) {
        sample = 0;
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value > maximumSample || value != Math.Truncate(value)) {
            return false;
        }
        sample = (int)value;
        return true;
    }
}
