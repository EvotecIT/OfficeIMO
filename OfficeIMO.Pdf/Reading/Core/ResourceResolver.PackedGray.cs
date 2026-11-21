using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    private static bool IsPackedGray(string colorSpace, int bitsPerComponent) =>
        colorSpace is "DeviceGray" or "G" && bitsPerComponent is 1 or 2 or 4;

    /// <summary>
    /// Writes unmasked device-gray 1/2/4-bit samples as a gray PNG of the same depth. The expanded 8-bit
    /// route gives the same colors, but a page-sized bilevel scan would otherwise be copied three times at
    /// eight times its size. <paramref name="handled"/> is false when the image needs that general route.
    /// </summary>
    private static bool TryBuildPackedGrayPngFile(PdfStream stream, int width, int height, int bitsPerComponent,
        PdfObject? colorSpaceObj, string colorSpace, Dictionary<int, PdfIndirectObject> objects, int maximumBytes,
        OfficeIccRenderingIntent renderingIntent, PdfOutputIntentColorTransform? outputIntentColorTransform,
        Func<int, long, bool>? colorFunctionEvaluationBudget, PdfColorFunctionResolutionContext? functionResolutionContext,
        out bool handled, out byte[] pngBytes, CancellationToken cancellationToken) {
        handled = false;
        pngBytes = Array.Empty<byte>();
        if (stream.Dictionary.Items.ContainsKey("Mask") || stream.Dictionary.Items.ContainsKey("SMask") ||
            !PdfImageColorSpaceNormalization.TryResolve(colorSpaceObj, colorSpace, objects, maximumBytes, renderingIntent,
                outputIntentColorTransform, colorFunctionEvaluationBudget, functionResolutionContext, out var normalization) ||
            normalization.SourceColorCount != 1 || normalization.RequiresColorConversion ||
            normalization.ProducesTransparency || normalization.PngColorType != 0 ||
            !PdfImageDecodeTransform.TryCreateColorDeclaration(stream.Dictionary, 1, objects, out PdfImageDecodeTransform? decode)) return false;

        // Map each packed sample through the same 8-bit expansion and Decode array as the general route.
        // PNG scales an n-bit gray sample to 8 bits as sample * 255 / (2^n - 1), so the output must land on that grid.
        int maximumSample = (1 << bitsPerComponent) - 1;
        int step = 255 / maximumSample;
        var map = new byte[maximumSample + 1];
        bool identity = true;
        bool inverted = true;
        for (int sample = 0; sample <= maximumSample; sample++) {
            byte expanded = (byte)(sample * 255 / maximumSample);
            byte gray = decode is null ? expanded : decode.TransformColorComponent(expanded, 0);
            if (gray % step != 0) return false;
            map[sample] = (byte)(gray / step);
            identity &= map[sample] == sample;
            inverted &= map[sample] == maximumSample - sample;
        }

        handled = true;
        cancellationToken.ThrowIfCancellationRequested();
        if (!PdfImageBufferLimits.TryGetScanlineBufferSize(width, height, 1, maximumBytes, out _, out _) ||
            !PdfImageStreamDecoder.TryDecode(stream, objects, out byte[] packed, maximumBytes, cancellationToken)) return false;
        int stride = checked((int)(((long)width * bitsPerComponent + 7) / 8));
        if ((long)stride * height != packed.LongLength || packed.Length == 0 ||
            !normalization.TryConsumeEvaluationWork((long)width * height, colorFunctionEvaluationBudget)) return false;

        // Rows go straight into the PNG's stored blocks; the output starts zeroed like a fresh scanline buffer.
        pngBytes = OfficePngWriter.EncodeStoredScanlines(width, height, bitsPerComponent, 0, null, (scanlines, start) => {
            for (int row = 0; row < height; row++) {
                cancellationToken.ThrowIfCancellationRequested();
                int source = row * stride;
                int target = start + row * (stride + 1) + 1;
                if (identity) {
                    Buffer.BlockCopy(packed, source, scanlines, target, stride);
                } else if (inverted) {
                    for (int index = 0; index < stride; index++) scanlines[target + index] = (byte)~packed[source + index];
                } else {
                    for (int column = 0; column < width; column++) {
                        int bit = column * bitsPerComponent;
                        int shift = 8 - bitsPerComponent - (bit & 7);
                        int sample = (packed[source + (bit >> 3)] >> shift) & maximumSample;
                        scanlines[target + (bit >> 3)] |= (byte)(map[sample] << shift);
                    }
                }
            }
        });
        return true;
    }

    // Expand packed samples before the existing decode, color management, and mask pipeline.
    // Each supported bit depth maps exactly to an 8-bit sample, including color-key endpoints.
    private static bool TryExpandPackedGray(PdfStream stream, int width, int height, int bitsPerComponent,
        Dictionary<int, PdfIndirectObject> objects, int maximumBytes, out PdfStream expanded, CancellationToken cancellationToken = default) {
        expanded = stream;
        cancellationToken.ThrowIfCancellationRequested();
        if (!PdfImageBufferLimits.TryGetScanlineBufferSize(width, height, 1, maximumBytes, out _, out _) ||
            !PdfImageStreamDecoder.TryDecode(stream, objects, out byte[] packed, maximumBytes, cancellationToken)) return false;
        int stride = checked((int)(((long)width * bitsPerComponent + 7) / 8));
        if ((long)stride * height != packed.LongLength) return false;
        int maximumSample = (1 << bitsPerComponent) - 1;
        var samples = new byte[checked(width * height)];
        for (int row = 0; row < height; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int column = 0; column < width; column++) {
                if ((column & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
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
