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
    private static bool TryNormalizeImageSamples(PdfStream stream, int width, int height, int bitsPerComponent, PdfImageColorSpaceNormalization colorSpace,
        Dictionary<int, PdfIndirectObject> objects, int maximumBytes, out PdfStream normalized, CancellationToken cancellationToken) {
        normalized = stream;
        int colorCount = colorSpace.SourceColorCount;
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
                if (number.Value != Math.Floor(number.Value)) return false;
                int quantized = (int)Math.Floor(number.Value * 255D / maximumSample + 0.5D);
                if (bitsPerComponent == 16 && !IsLosslessColorKeyEndpoint(number.Value, index)) return false;
                scaled.Items.Add(new PdfNumber(quantized));
            }
            dictionary.Items["Mask"] = scaled;
        }

        if (ResolveStream(stream.Dictionary.Items.TryGetValue("SMask", out PdfObject? softMaskObject) ? softMaskObject : null, objects) is PdfStream softMask) {
            if (!TryNormalizeSoftMask(softMask, width, height, objects, maximumBytes, out byte[] alpha, cancellationToken)) return false;
            if (softMask.Dictionary.Items.TryGetValue("Matte", out PdfObject? matteObject) && ResolveObject(matteObject, objects) is not PdfNull) {
                if (!TryReadMatte(softMask, colorSpace, objects, out double[]? matte) || matte == null) return false;
                if (!PdfImageDecodeTransform.TryCreateColor(stream.Dictionary, colorCount, objects, out PdfImageDecodeTransform? decode)) return false;
                // Decode before un-premultiplying, then remove Decode to avoid applying it twice.
                if (decode != null) {
                    for (int index = 0; index < samples.Length; index++) {
                        if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                        int component = index % colorCount;
                        double value = colorSpace.MapComponentToUnitRange(component, decode.TransformColorComponentValue(samples[index], component));
                        samples[index] = (byte)Math.Round(value * 255D);
                    }
                    dictionary.Items.Remove("Decode");
                }
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

    internal static bool CanRepresentSixteenBitColorKey(PdfDictionary dictionary, Dictionary<int, PdfIndirectObject> objects) {
        if (!dictionary.Items.TryGetValue("Mask", out PdfObject? mask) || ResolveObject(mask, objects) is not PdfArray key) return true;
        for (int index = 0; index < key.Items.Count; index++) {
            if (ResolveObject(key.Items[index], objects) is not PdfNumber number || !IsLosslessColorKeyEndpoint(number.Value, index)) return false;
        }
        return true;
    }

    private static bool IsLosslessColorKeyEndpoint(double value, int index) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value > 65535D || value != Math.Floor(value)) return false;
        int quantized = (int)Math.Floor(value * 255D / 65535D + 0.5D);
        double boundary = index % 2 == 0
            ? Math.Max(0D, Math.Ceiling((quantized * 65535D - 32767D) / 255D))
            : Math.Min(65535D, Math.Ceiling(((quantized + 1) * 65535D - 32767D) / 255D) - 1D);
        return value == boundary;
    }

    private static bool TryUnpackSamples(byte[] raw, int width, int height, int colorCount, int bitsPerComponent,
        out byte[] samples, CancellationToken cancellationToken) {
        samples = Array.Empty<byte>();
        if (bitsPerComponent is not (1 or 2 or 4 or 8 or 16) || width <= 0 || height <= 0 || colorCount <= 0) return false;
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
        if (maskBits is not (1 or 2 or 4 or 8 or 16) || maskWidth <= 0 || maskHeight <= 0 || !string.Equals(colorSpace, "DeviceGray", StringComparison.Ordinal) ||
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
        bool interpolate = ResolveObject(softMask.Dictionary.Items.TryGetValue("Interpolate", out PdfObject? interpolation) ? interpolation : null, objects)
            is PdfBoolean { Value: true };
        for (int row = 0; row < height; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            int sourceRow = Math.Min(maskHeight - 1, (int)((row + 0.5D) * maskHeight / height));
            for (int column = 0; column < width; column++) {
                int sourceColumn = Math.Min(maskWidth - 1, (int)((column + 0.5D) * maskWidth / width));
                if (!interpolate) {
                    alpha[row * width + column] = maskSamples[sourceRow * maskWidth + sourceColumn];
                    continue;
                }
                double x = Math.Max(0D, Math.Min(maskWidth - 1D, (column + 0.5D) * maskWidth / width - 0.5D));
                double y = Math.Max(0D, Math.Min(maskHeight - 1D, (row + 0.5D) * maskHeight / height - 0.5D));
                int x0 = (int)x, y0 = (int)y, x1 = Math.Min(x0 + 1, maskWidth - 1), y1 = Math.Min(y0 + 1, maskHeight - 1);
                double fx = x - x0, fy = y - y0;
                double top = maskSamples[y0 * maskWidth + x0] * (1D - fx) + maskSamples[y0 * maskWidth + x1] * fx;
                double bottom = maskSamples[y1 * maskWidth + x0] * (1D - fx) + maskSamples[y1 * maskWidth + x1] * fx;
                alpha[row * width + column] = (byte)Math.Round(top * (1D - fy) + bottom * fy);
            }
        }
        return true;
    }

    private static bool TryReadMatte(PdfStream softMask, PdfImageColorSpaceNormalization colorSpace, Dictionary<int, PdfIndirectObject> objects, out double[]? matte) {
        matte = null;
        int colorCount = colorSpace.SourceColorCount;
        if (!softMask.Dictionary.Items.TryGetValue("Matte", out PdfObject? matteObject) || ResolveArray(matteObject, objects) is not PdfArray array ||
            array.Items.Count != colorCount) return false;
        matte = new double[colorCount];
        for (int index = 0; index < colorCount; index++) {
            if (ResolveObject(array.Items[index], objects) is not PdfNumber number || double.IsNaN(number.Value) || double.IsInfinity(number.Value)) return false;
            matte[index] = colorSpace.MapComponentToUnitRange(index, number.Value);
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
