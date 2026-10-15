using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    /// <summary>
    /// True when a non-indexed image needs its samples rewritten before the 8-bit pipeline: 1, 2, 4 or 16-bit
    /// components, or a soft mask whose size, depth or /Matte the pipeline cannot pair directly.
    /// </summary>
    private static bool RequiresSampleNormalization(PdfStream stream, int width, int height, int bitsPerComponent,
        Dictionary<int, PdfIndirectObject> objects) {
        if (bitsPerComponent is 1 or 2 or 4 or 16) return true;
        if (bitsPerComponent != 8 || ResolveStream(stream.Dictionary.Items.TryGetValue("SMask", out PdfObject? maskObject) ? maskObject : null, objects)
                is not PdfStream softMask) return false;
        int maskWidth = (int)(softMask.Dictionary.Get<PdfNumber>("Width")?.Value ?? 0);
        int maskHeight = (int)(softMask.Dictionary.Get<PdfNumber>("Height")?.Value ?? 0);
        int maskBits = (int)(softMask.Dictionary.Get<PdfNumber>("BitsPerComponent")?.Value ?? 0);
        return maskWidth != width || maskHeight != height || maskBits != 8 ||
            softMask.Dictionary.Items.TryGetValue("Matte", out PdfObject? matte) && ResolveObject(matte, objects) is not PdfNull;
    }

    /// <summary>
    /// Rewrites the samples as 8 bits per component and pairs them with an 8-bit soft mask of the same size,
    /// un-premultiplying /Matte colours (ISO 32000-1, 11.6.5.3). Colour-key ranges scale with the samples.
    /// </summary>
    private static bool TryNormalizeImageSamples(PdfStream stream, int width, int height, int bitsPerComponent, int colorCount,
        Dictionary<int, PdfIndirectObject> objects, int maximumBytes, out PdfStream normalized, CancellationToken cancellationToken) {
        normalized = stream;
        cancellationToken.ThrowIfCancellationRequested();
        if (colorCount < 1 || colorCount > 4 || bitsPerComponent is not (1 or 2 or 4 or 8 or 16) ||
            !PdfImageBufferLimits.TryGetScanlineBufferSize(width, height, colorCount + 1, maximumBytes, out _, out _)) return false;
        byte[] samples;
        if (IsDctFilterChain(stream.Dictionary, objects)) {
            if (bitsPerComponent != 8 || !TryDecodeDctImage(stream, width, height, colorCount, objects, maximumBytes, out samples)) return false;
        } else {
            if (!PdfImageStreamDecoder.TryDecode(stream, objects, out byte[] raw, maximumBytes, cancellationToken) ||
                !TryUnpackSamples(raw, width, height, colorCount, bitsPerComponent, out samples, cancellationToken)) return false;
        }

        var dictionary = new PdfDictionary();
        foreach (KeyValuePair<string, PdfObject> pair in stream.Dictionary.Items) dictionary.Items.Add(pair.Key, pair.Value);
        foreach (string key in new[] { "Filter", "DecodeParms", "F", "DP", "SMask" }) dictionary.Items.Remove(key);
        dictionary.Items["BitsPerComponent"] = new PdfNumber(8);
        if (dictionary.Items.TryGetValue("Mask", out PdfObject? maskObject) && ResolveObject(maskObject, objects) is PdfArray colorKey) {
            if (colorKey.Items.Count != 2 * colorCount) return false;
            int maximumSample = bitsPerComponent == 16 ? 65535 : (1 << bitsPerComponent) - 1;
            var scaled = new PdfArray();
            for (int index = 0; index < colorKey.Items.Count; index++) {
                if (ResolveObject(colorKey.Items[index], objects) is not PdfNumber number || number.Value < 0 || number.Value > maximumSample) return false;
                double value = number.Value * 255D / maximumSample;
                scaled.Items.Add(new PdfNumber(index % 2 == 0 ? Math.Floor(value) : Math.Ceiling(value)));
            }
            dictionary.Items["Mask"] = scaled;
        }

        if (ResolveStream(stream.Dictionary.Items.TryGetValue("SMask", out PdfObject? softMaskObject) ? softMaskObject : null, objects) is PdfStream softMask) {
            if (!TryNormalizeSoftMask(softMask, width, height, objects, maximumBytes, out byte[] alpha, cancellationToken)) return false;
            if (TryReadMatte(softMask, colorCount, objects, out double[]? matte) && matte != null &&
                !stream.Dictionary.Items.ContainsKey("Decode")) {
                UnpremultiplyMatte(samples, alpha, colorCount, matte);
            }
            var maskDictionary = new PdfDictionary();
            maskDictionary.Items["Type"] = new PdfName("XObject");
            maskDictionary.Items["Subtype"] = new PdfName("Image");
            maskDictionary.Items["Width"] = new PdfNumber(width);
            maskDictionary.Items["Height"] = new PdfNumber(height);
            maskDictionary.Items["ColorSpace"] = new PdfName("DeviceGray");
            maskDictionary.Items["BitsPerComponent"] = new PdfNumber(8);
            dictionary.Items["SMask"] = new PdfStream(maskDictionary, alpha);
        }
        normalized = new PdfStream(dictionary, samples);
        return true;
    }

    private static bool TryUnpackSamples(byte[] raw, int width, int height, int colorCount, int bitsPerComponent,
        out byte[] samples, CancellationToken cancellationToken) {
        samples = Array.Empty<byte>();
        long rowBits = (long)width * colorCount * bitsPerComponent;
        long stride = (rowBits + 7) / 8;
        if (stride * height > raw.LongLength || (long)width * colorCount * height > int.MaxValue) return false;
        int perRow = width * colorCount;
        samples = new byte[checked(perRow * height)];
        int maximumSample = (1 << Math.Min(bitsPerComponent, 8)) - 1;
        for (int row = 0; row < height; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            long rowStart = row * stride;
            for (int index = 0; index < perRow; index++) {
                byte value;
                if (bitsPerComponent == 8) {
                    value = raw[rowStart + index];
                } else if (bitsPerComponent == 16) {
                    int sample = (raw[rowStart + index * 2] << 8) | raw[rowStart + index * 2 + 1];
                    value = (byte)((sample * 255 + 32767) / 65535);
                } else {
                    long bit = (long)index * bitsPerComponent;
                    int sample = (raw[rowStart + bit / 8] >> (8 - bitsPerComponent - (int)(bit & 7))) & maximumSample;
                    value = (byte)(sample * 255 / maximumSample);
                }
                samples[row * perRow + index] = value;
            }
        }
        return true;
    }

    // Soft masks are DeviceGray images of any supported depth and size; they are resampled to the base image grid.
    private static bool TryNormalizeSoftMask(PdfStream softMask, int width, int height, Dictionary<int, PdfIndirectObject> objects,
        int maximumBytes, out byte[] alpha, CancellationToken cancellationToken) {
        alpha = Array.Empty<byte>();
        int maskWidth = (int)(softMask.Dictionary.Get<PdfNumber>("Width")?.Value ?? 0);
        int maskHeight = (int)(softMask.Dictionary.Get<PdfNumber>("Height")?.Value ?? 0);
        int maskBits = (int)(softMask.Dictionary.Get<PdfNumber>("BitsPerComponent")?.Value ?? 0);
        string colorSpace = GetNameOrEmpty(softMask.Dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject) ? colorSpaceObject : null, objects);
        if (maskWidth <= 0 || maskHeight <= 0 || !string.Equals(colorSpace, "DeviceGray", StringComparison.Ordinal) ||
            Filters.StreamDecoder.GetUnsupportedFilters(softMask.Dictionary, objects).Count != 0 ||
            !PdfImageBufferLimits.TryGetScanlineBufferSize(maskWidth, maskHeight, 1, maximumBytes, out _, out _)) return false;
        byte[] maskSamples;
        if (IsDctFilterChain(softMask.Dictionary, objects)) {
            if (maskBits != 8 || !TryDecodeDctImage(softMask, maskWidth, maskHeight, 1, objects, maximumBytes, out maskSamples)) return false;
        } else if (!PdfImageStreamDecoder.TryDecode(softMask, objects, out byte[] raw, maximumBytes, cancellationToken) ||
                   !TryUnpackSamples(raw, maskWidth, maskHeight, 1, maskBits, out maskSamples, cancellationToken)) {
            return false;
        }
        if (ResolveArray(softMask.Dictionary.Items.TryGetValue("Decode", out PdfObject? decodeObject) ? decodeObject : null, objects) is PdfArray decode &&
            decode.Items.Count == 2 && ResolveObject(decode.Items[0], objects) is PdfNumber low && ResolveObject(decode.Items[1], objects) is PdfNumber high) {
            for (int index = 0; index < maskSamples.Length; index++) {
                double value = low.Value + maskSamples[index] / 255D * (high.Value - low.Value);
                maskSamples[index] = (byte)Math.Round(Math.Max(0D, Math.Min(1D, value)) * 255D);
            }
        }
        if (maskWidth == width && maskHeight == height) {
            alpha = maskSamples;
            return true;
        }
        if ((long)width * height > int.MaxValue) return false;
        alpha = new byte[width * height];
        for (int row = 0; row < height; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            int sourceRow = Math.Min(maskHeight - 1, (int)((row + 0.5D) * maskHeight / height));
            for (int column = 0; column < width; column++) {
                int sourceColumn = Math.Min(maskWidth - 1, (int)((column + 0.5D) * maskWidth / width));
                alpha[row * width + column] = maskSamples[sourceRow * maskWidth + sourceColumn];
            }
        }
        return true;
    }

    private static bool TryReadMatte(PdfStream softMask, int colorCount, Dictionary<int, PdfIndirectObject> objects, out double[]? matte) {
        matte = null;
        if (!softMask.Dictionary.Items.TryGetValue("Matte", out PdfObject? matteObject) || ResolveArray(matteObject, objects) is not PdfArray array ||
            array.Items.Count != colorCount) return false;
        matte = new double[colorCount];
        for (int index = 0; index < colorCount; index++) {
            if (ResolveObject(array.Items[index], objects) is not PdfNumber number || double.IsNaN(number.Value) || double.IsInfinity(number.Value)) return false;
            matte[index] = Math.Max(0D, Math.Min(1D, number.Value));
        }
        return true;
    }

    // c = m + (c' - m) / alpha, the inverse of the pre-blend with the matte colour.
    private static void UnpremultiplyMatte(byte[] samples, byte[] alpha, int colorCount, double[] matte) {
        for (int pixel = 0; pixel < alpha.Length; pixel++) {
            double a = alpha[pixel] / 255D;
            if (a <= 0D) continue;
            for (int component = 0; component < colorCount; component++) {
                int index = pixel * colorCount + component;
                double value = matte[component] + (samples[index] / 255D - matte[component]) / a;
                samples[index] = (byte)Math.Round(Math.Max(0D, Math.Min(1D, value)) * 255D);
            }
        }
    }
}
