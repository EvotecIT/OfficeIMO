using System.Diagnostics;

namespace OfficeIMO.Drawing.Benchmarks;

internal static class ImageCancellationEvidence {
    private const int Width = 4096;
    private const int Height = 1025;
    private static readonly TimeSpan CancellationDelay = TimeSpan.FromMilliseconds(1);
    private static readonly TimeSpan MaximumObservedLatency = TimeSpan.FromSeconds(2);

    internal static void Validate(TextWriter writer) {
        OfficeRasterImage source = ImageBenchmarkCorpus.CreatePattern(Width, Height);
        var tiffOptions = new OfficeRasterEncodingOptions {
            Tiff = new OfficeTiffEncodeOptions {
                Compression = OfficeTiffCompression.None,
                Predictor = OfficeTiffPredictor.None
            }
        };
        byte[] tiff = OfficeRasterImageEncoder.Encode(source, OfficeImageExportFormat.Tiff, tiffOptions);
        byte[] webp = OfficeRasterImageEncoder.Encode(source, OfficeImageExportFormat.Webp);

        writer.WriteLine();
        writer.WriteLine("Cancellation latency (cancel requested 1 ms after bounded decode starts):");
        WriteResult(writer, "TIFF", Measure(tiff));
        WriteResult(writer, "WebP", Measure(webp));
    }

    private static TimeSpan Measure(byte[] encoded) {
        using var cancellation = new CancellationTokenSource();
        cancellation.CancelAfter(CancellationDelay);
        var options = new OfficeRasterDecodeOptions {
            CancellationToken = cancellation.Token,
            MaximumEncodedBytes = 32 * 1024 * 1024,
            MaximumDecodedPixels = 8_000_000
        };
        var stopwatch = Stopwatch.StartNew();
        try {
            OfficeRasterImageDecoder.TryDecode(encoded, options, out _, out _);
        } catch (OperationCanceledException) {
            stopwatch.Stop();
            if (stopwatch.Elapsed > MaximumObservedLatency) {
                throw new InvalidOperationException(
                    $"Cancellation took {stopwatch.Elapsed.TotalMilliseconds:N1} ms, above the evidence ceiling of {MaximumObservedLatency.TotalMilliseconds:N0} ms.");
            }
            return stopwatch.Elapsed;
        }
        throw new InvalidOperationException("The bounded decoder completed without observing scheduled cancellation.");
    }

    private static void WriteResult(TextWriter writer, string format, TimeSpan elapsed) =>
        writer.WriteLine($"{format,-5} {elapsed.TotalMilliseconds,8:N1} ms");
}
