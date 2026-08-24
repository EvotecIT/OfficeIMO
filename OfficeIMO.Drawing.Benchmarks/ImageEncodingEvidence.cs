namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Validates output size and fidelity separately from timed benchmark measurements.</summary>
internal static class ImageEncodingEvidence {
    private static readonly int[] Sizes = { 256, 512, 1024 };

    internal static void Validate(TextWriter writer) {
        writer.WriteLine();
        writer.WriteLine("Encoding size and fidelity matrix (encoded bytes versus source RGBA bytes):");
        writer.WriteLine("Size       Format/options                  Bytes    Ratio       MAE     PSNR");

        foreach (int size in Sizes) {
            OfficeRasterImage source = ImageBenchmarkCorpus.CreatePattern(size, size);
            WriteLossless(writer, size, source, "PNG optimal", OfficePngWriter.Encode(source, new OfficePngEncodeOptions {
                Compression = OfficePngCompression.Optimal
            }));
            WriteLossless(writer, size, source, "PNG stored", OfficePngWriter.Encode(source, new OfficePngEncodeOptions {
                Compression = OfficePngCompression.Stored
            }));
            WriteJpeg(writer, size, source, quality: 60, OfficeJpegSubsampling.Y420, progressive: false, optimizeHuffman: false);
            WriteJpeg(writer, size, source, quality: 85, OfficeJpegSubsampling.Y420, progressive: false, optimizeHuffman: false);
            WriteJpeg(writer, size, source, quality: 95, OfficeJpegSubsampling.Y420, progressive: false, optimizeHuffman: false);
            WriteJpeg(writer, size, source, quality: 85, OfficeJpegSubsampling.Y444, progressive: false, optimizeHuffman: false);
            WriteJpeg(writer, size, source, quality: 85, OfficeJpegSubsampling.Y420, progressive: true, optimizeHuffman: true);
            WriteLossless(writer, size, source, "TIFF none", OfficeTiffCodec.Encode(source, new OfficeTiffEncodeOptions {
                Compression = OfficeTiffCompression.None
            }));
            WriteLossless(writer, size, source, "TIFF PackBits", OfficeTiffCodec.Encode(source, new OfficeTiffEncodeOptions {
                Compression = OfficeTiffCompression.PackBits
            }));
            WriteLossless(writer, size, source, "TIFF LZW", OfficeTiffCodec.Encode(source, new OfficeTiffEncodeOptions {
                Compression = OfficeTiffCompression.Lzw
            }));
            WriteLossless(writer, size, source, "TIFF Deflate", OfficeTiffCodec.Encode(source, new OfficeTiffEncodeOptions {
                Compression = OfficeTiffCompression.Deflate
            }));
            WriteLossless(writer, size, source, "WebP lossless", OfficeWebpCodec.Encode(source));
        }
    }

    private static void WriteLossless(TextWriter writer, int size, OfficeRasterImage source, string label, byte[] encoded) {
        OfficeRasterImage decoded = ImageBenchmarkCorpus.Decode(encoded, label);
        byte[] expected = source.GetPixels();
        byte[] actual = decoded.GetPixels();
        if (!expected.AsSpan().SequenceEqual(actual)) {
            throw new InvalidOperationException(label + " did not preserve the source RGBA pixels.");
        }
        WriteRow(writer, size, label, encoded.Length, expected.Length, meanAbsoluteError: 0D, psnr: double.PositiveInfinity);
    }

    private static void WriteJpeg(
        TextWriter writer,
        int size,
        OfficeRasterImage source,
        int quality,
        OfficeJpegSubsampling subsampling,
        bool progressive,
        bool optimizeHuffman) {
        byte[] encoded = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
            Quality = quality,
            Subsampling = subsampling,
            Progressive = progressive,
            OptimizeHuffman = optimizeHuffman,
            Background = OfficeColor.White
        });
        OfficeRasterImage decoded = ImageBenchmarkCorpus.Decode(encoded, "JPEG evidence");
        byte[] expected = FlattenAgainstWhite(source.GetPixels());
        (double mae, double psnr) = MeasureRgbFidelity(expected, decoded.GetPixels());
        if (mae > 40D || psnr < 15D) {
            throw new InvalidOperationException(
                $"JPEG Q{quality} {subsampling} fidelity was outside the validation envelope: MAE {mae:F3}, PSNR {psnr:F2} dB.");
        }

        string mode = progressive ? " progressive+huffman" : string.Empty;
        WriteRow(writer, size, $"JPEG Q{quality} {subsampling}{mode}", encoded.Length, expected.Length, mae, psnr);
    }

    internal static byte[] FlattenAgainstWhite(byte[] rgba) {
        var flattened = new byte[rgba.Length];
        for (int offset = 0; offset < rgba.Length; offset += 4) {
            int alpha = rgba[offset + 3];
            int inverse = 255 - alpha;
            flattened[offset] = (byte)((rgba[offset] * alpha + 255 * inverse + 127) / 255);
            flattened[offset + 1] = (byte)((rgba[offset + 1] * alpha + 255 * inverse + 127) / 255);
            flattened[offset + 2] = (byte)((rgba[offset + 2] * alpha + 255 * inverse + 127) / 255);
            flattened[offset + 3] = 255;
        }
        return flattened;
    }

    internal static (double MeanAbsoluteError, double Psnr) MeasureRgbFidelity(byte[] expected, byte[] actual) {
        if (expected.Length != actual.Length || expected.Length % 4 != 0) {
            throw new InvalidOperationException("JPEG evidence did not preserve the expected pixel dimensions.");
        }

        long absoluteError = 0L;
        double squaredError = 0D;
        long channelCount = expected.LongLength / 4L * 3L;
        for (int offset = 0; offset < expected.Length; offset += 4) {
            for (int channel = 0; channel < 3; channel++) {
                int difference = expected[offset + channel] - actual[offset + channel];
                absoluteError += Math.Abs(difference);
                squaredError += difference * difference;
            }
        }

        double meanAbsoluteError = absoluteError / (double)channelCount;
        double meanSquaredError = squaredError / channelCount;
        double psnr = meanSquaredError == 0D
            ? double.PositiveInfinity
            : 10D * Math.Log10(255D * 255D / meanSquaredError);
        return (meanAbsoluteError, psnr);
    }

    private static void WriteRow(
        TextWriter writer,
        int size,
        string label,
        int encodedLength,
        int rgbaLength,
        double meanAbsoluteError,
        double psnr) {
        string psnrText = double.IsPositiveInfinity(psnr) ? "lossless" : psnr.ToString("F2");
        writer.WriteLine(
            $"{size,4}x{size,-4} {label,-31} {encodedLength,10:N0} {encodedLength / (double)rgbaLength,8:P1} " +
            $"{meanAbsoluteError,9:F3} {psnrText,8}");
    }
}
