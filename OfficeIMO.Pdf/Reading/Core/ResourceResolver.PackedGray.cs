namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    private static bool IsPackedGray(string colorSpace, int bitsPerComponent) =>
        colorSpace is "DeviceGray" or "G" && bitsPerComponent is 1 or 2 or 4;

    // Expand packed samples before the existing decode, color management, and mask pipeline.
    // Each supported bit depth maps exactly to an 8-bit sample, including color-key endpoints.
    private static bool TryExpandPackedGray(PdfStream stream, int width, int height, int bitsPerComponent,
        Dictionary<int, PdfIndirectObject> objects, int maximumBytes, out PdfStream expanded) {
        expanded = stream;
        if (!PdfImageBufferLimits.TryGetScanlineBufferSize(width, height, 1, maximumBytes, out _, out _) ||
            !PdfImageStreamDecoder.TryDecode(stream, objects, out byte[] packed, maximumBytes)) return false;
        int stride = checked((int)(((long)width * bitsPerComponent + 7) / 8));
        if ((long)stride * height != packed.LongLength) return false;
        int maximumSample = (1 << bitsPerComponent) - 1;
        var samples = new byte[checked(width * height)];
        for (int row = 0; row < height; row++) {
            for (int column = 0; column < width; column++) {
                long bit = (long)column * bitsPerComponent;
                int sample = (packed[row * stride + (int)(bit / 8)] >> (8 - bitsPerComponent - (int)(bit & 7))) & maximumSample;
                samples[row * width + column] = (byte)(sample * 255 / maximumSample);
            }
        }
        var dictionary = new PdfDictionary();
        foreach (var pair in stream.Dictionary.Items) dictionary.Items.Add(pair.Key, pair.Value);
        dictionary.Items.Remove("Filter");
        dictionary.Items.Remove("DecodeParms");
        dictionary.Items.Remove("F");
        dictionary.Items.Remove("DP");
        dictionary.Items["BitsPerComponent"] = new PdfNumber(8);
        if (dictionary.Items.TryGetValue("Mask", out PdfObject? maskObject) &&
            ResolveObject(maskObject, objects) is PdfArray mask) {
            if (mask.Items.Count != 2) return false;
            var scaledMask = new PdfArray();
            foreach (PdfObject item in mask.Items) {
                if (ResolveObject(item, objects) is not PdfNumber number || number.Value < 0 ||
                    number.Value > maximumSample || number.Value != Math.Floor(number.Value)) return false;
                scaledMask.Items.Add(new PdfNumber(number.Value * 255 / maximumSample));
            }
            dictionary.Items["Mask"] = scaledMask;
        }
        expanded = new PdfStream(dictionary, samples);
        return true;
    }
}
