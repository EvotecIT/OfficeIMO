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
        writer.WriteLine("Cancellation latency (cancel requested 1 ms after synchronized decode start):");
        WriteResult(writer, "TIFF", Measure(tiff));
        WriteResult(writer, "WebP", Measure(webp));
    }

    internal static TimeSpan Measure(OfficeImageExportFormat format) {
        if (format is not (OfficeImageExportFormat.Tiff or OfficeImageExportFormat.Webp)) {
            throw new ArgumentOutOfRangeException(nameof(format));
        }
        OfficeRasterImage source = ImageBenchmarkCorpus.CreatePattern(Width, Height);
        var options = new OfficeRasterEncodingOptions {
            Tiff = new OfficeTiffEncodeOptions {
                Compression = OfficeTiffCompression.None,
                Predictor = OfficeTiffPredictor.None
            }
        };
        return Measure(OfficeRasterImageEncoder.Encode(source, format, options));
    }

    private static TimeSpan Measure(byte[] encoded) => TimeSpan.FromMilliseconds(
        MeasureRepresentativeLatency(() => MeasureSingle(encoded).TotalMilliseconds));

    // A single observation can include an unrelated OS scheduling pause. The median
    // still exposes a persistent cancellation slowdown to the regression gate.
    internal static double MeasureRepresentativeLatency(Func<double> measure) {
        double first = measure();
        double second = measure();
        double third = measure();
        if (first > second) (first, second) = (second, first);
        if (second > third) (second, third) = (third, second);
        return Math.Max(first, second);
    }

    private static TimeSpan MeasureSingle(byte[] encoded) {
        using var cancellation = new CancellationTokenSource();
        var options = new OfficeRasterDecodeOptions {
            CancellationToken = cancellation.Token,
            MaximumEncodedBytes = 32 * 1024 * 1024,
            MaximumDecodedPixels = 8_000_000
        };
        using var beginDecode = new ManualResetEventSlim();
        long cancellationTimestamp = 0L;
        var cancellationThread = new Thread(() => {
            beginDecode.Wait();
            Thread.Sleep(CancellationDelay);
            Volatile.Write(ref cancellationTimestamp, Stopwatch.GetTimestamp());
            cancellation.Cancel();
        }) {
            IsBackground = true,
            Name = "OfficeIMO image cancellation evidence"
        };
        cancellationThread.Start();
        beginDecode.Set();
        bool cancelled = false;
        long observedAt = 0L;
        try {
            OfficeRasterImageDecoder.TryDecode(encoded, options, out _, out _);
        } catch (OperationCanceledException) {
            observedAt = Stopwatch.GetTimestamp();
            cancelled = true;
        } finally {
            cancellationThread.Join();
        }
        if (!cancelled) throw new InvalidOperationException("The bounded decoder completed without observing scheduled cancellation.");
        long requestedAt = Volatile.Read(ref cancellationTimestamp);
        if (requestedAt == 0L) {
            throw new InvalidOperationException("Cancellation was observed before the synchronized request.");
        }
        TimeSpan elapsed = Stopwatch.GetElapsedTime(requestedAt, observedAt);
        if (elapsed > MaximumObservedLatency) {
            throw new InvalidOperationException(
                $"Cancellation took {elapsed.TotalMilliseconds:N1} ms, above the evidence ceiling of {MaximumObservedLatency.TotalMilliseconds:N0} ms.");
        }
        return elapsed;
    }

    private static void WriteResult(TextWriter writer, string format, TimeSpan elapsed) =>
        writer.WriteLine($"{format,-5} {elapsed.TotalMilliseconds,8:N1} ms");
}
