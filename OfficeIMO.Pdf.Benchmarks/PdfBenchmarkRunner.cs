using System.Diagnostics;
using OfficeIMO.Pdf;

internal static class PdfBenchmarkRunner {
    internal const string AnalysisCold = "analysis-cold";
    internal const string AnalysisCached = "analysis-cached";
    internal const string RenderSvg = "render-svg-12";
    internal const string RenderPng = "render-png-4";
    internal const string SerializeBuffered = "serialize-buffered-60";
    internal const string SerializeForward = "serialize-forward-60";
    internal const string SerializeHarfBuzz = "serialize-harfbuzz-60";

    internal static IReadOnlyList<PdfPerformanceMeasurement> Measure(byte[] corpus) {
        (PdfPerformanceMeasurement cold, PdfPerformanceMeasurement cached) = MeasureAnalysis(corpus);
        PdfPerformanceMeasurement svg = MeasureWorkflow(RenderSvg, corpus, RunSvgRender, sampleCount: 5);
        PdfPerformanceMeasurement png = MeasureWorkflow(RenderPng, corpus, RunPngRender, sampleCount: 3);
        PdfPerformanceMeasurement buffered = MeasureSerialization(
            SerializeBuffered,
            PdfObjectSerializationMode.Buffered,
            sampleCount: 3);
        PdfPerformanceMeasurement forward = MeasureSerialization(
            SerializeForward,
            PdfObjectSerializationMode.ForwardOnly,
            sampleCount: 3);
        PdfPerformanceMeasurement harfBuzz = MeasureSerialization(
            SerializeHarfBuzz,
            PdfObjectSerializationMode.ForwardOnly,
            sampleCount: 3,
            static _ => PdfBenchmarkCorpus.CreateHarfBuzzDocument());
        return new[] { cold, cached, svg, png, buffered, forward, harfBuzz };
    }

    private static (PdfPerformanceMeasurement Cold, PdfPerformanceMeasurement Cached) MeasureAnalysis(byte[] corpus) {
        const int sampleCount = 7;
        RunColdAnalysis(corpus);
        RunCachedAnalysis(corpus);

        var coldSamples = new List<PdfPerformanceSample>(sampleCount);
        var cachedSamples = new List<PdfPerformanceSample>(sampleCount);
        for (int sample = 0; sample < sampleCount; sample++) {
            if (sample % 2 == 0) {
                coldSamples.Add(MeasureOnce(corpus, RunColdAnalysis));
                cachedSamples.Add(MeasureOnce(corpus, RunCachedAnalysis));
            } else {
                cachedSamples.Add(MeasureOnce(corpus, RunCachedAnalysis));
                coldSamples.Add(MeasureOnce(corpus, RunColdAnalysis));
            }
        }

        return (
            Summarize(AnalysisCold, coldSamples),
            Summarize(AnalysisCached, cachedSamples));
    }

    private static PdfPerformanceMeasurement MeasureWorkflow(
        string name,
        byte[] corpus,
        Func<byte[], long> operation,
        int sampleCount) {
        operation(corpus);
        var samples = new List<PdfPerformanceSample>(sampleCount);
        for (int sample = 0; sample < sampleCount; sample++) {
            samples.Add(MeasureOnce(corpus, operation));
        }

        return Summarize(name, samples);
    }

    private static long RunColdAnalysis(byte[] corpus) {
        string text = PdfDocument.Open(corpus).Read.Text();
        PdfDocumentInfo info = PdfDocument.Open(corpus).Inspect();
        PdfDocumentPreflight preflight = PdfDocument.Open(corpus).Preflight();
        PdfAnalysisReport analysis = PdfDocument.Open(corpus).Analyze();
        return text.Length + info.PageCount + preflight.Diagnostics.Count + analysis.Diagnostics.ObjectCount;
    }

    private static long RunCachedAnalysis(byte[] corpus) {
        PdfDocument document = PdfDocument.Open(corpus);
        string text = document.Read.Text();
        PdfDocumentInfo info = document.Inspect();
        PdfDocumentPreflight preflight = document.Preflight();
        PdfAnalysisReport analysis = document.Analyze();
        return text.Length + info.PageCount + preflight.Diagnostics.Count + analysis.Diagnostics.ObjectCount;
    }

    private static long RunSvgRender(byte[] corpus) =>
        Render(
            corpus,
            "1-12",
            12,
            new PdfPageRenderOptions {
                Format = PdfPageRenderFormat.Svg,
                MaxPages = 12,
                ContinueOnError = false
            });

    private static long RunPngRender(byte[] corpus) =>
        Render(
            corpus,
            "1-4",
            4,
            new PdfPageRenderOptions {
                Format = PdfPageRenderFormat.Png,
                ThumbnailMaxDimension = 512,
                MaxPages = 4,
                MaxPixelsPerPage = 512L * 512L,
                ContinueOnError = false
            });

    private static long Render(
        byte[] corpus,
        string pageRanges,
        int expectedPages,
        PdfPageRenderOptions options) {
        IReadOnlyList<PdfPageRenderResult> results = PdfDocument
            .Open(corpus)
            .Read
            .RenderPages(pageRanges, options);
        if (results.Count != expectedPages ||
            results.Any(result => !result.Succeeded || result.Bytes is null || result.Bytes.Length == 0)) {
            throw new InvalidOperationException(
                $"PDF {options.Format} render workload did not produce {expectedPages} complete pages.");
        }

        return results.Sum(result => (long)result.Bytes!.Length);
    }

    private static PdfPerformanceMeasurement MeasureSerialization(
        string name,
        PdfObjectSerializationMode mode,
        int sampleCount,
        Func<PdfObjectSerializationMode, PdfDocument>? createDocument = null) {
        RunSerialization(mode, createDocument);
        var samples = new List<PdfPerformanceSample>(sampleCount);
        for (int sample = 0; sample < sampleCount; sample++) {
            GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
            long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
            using var heap = new ManagedHeapSampler();
            var stopwatch = Stopwatch.StartNew();
            PdfPerformanceSample result = RunSerialization(mode, createDocument);
            stopwatch.Stop();
            long peakManagedHeapBytes = heap.Stop();
            samples.Add(result with {
                ElapsedMilliseconds = stopwatch.Elapsed.TotalMilliseconds,
                AllocatedBytes = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore,
                PeakManagedHeapBytes = peakManagedHeapBytes
            });
        }

        return Summarize(name, samples);
    }

    private static PdfPerformanceSample RunSerialization(
        PdfObjectSerializationMode mode,
        Func<PdfObjectSerializationMode, PdfDocument>? createDocument) {
        PdfDocument document = createDocument?.Invoke(mode) ?? PdfBenchmarkCorpus.CreateDocument(mode);
        using var output = new MemoryStream();
        PdfSaveResult save = document.Save(output);
        PdfSerializationReport serialization = save.Serialization
            ?? throw new InvalidOperationException("PDF serialization benchmark did not return runtime evidence.");
        byte[] bytes = output.ToArray();
        if (PdfDocument.Open(bytes).Inspect().PageCount != PdfBenchmarkCorpus.PageCount) {
            throw new InvalidOperationException("PDF serialization benchmark produced an invalid page count.");
        }
        if (serialization.IsForwardOnlyObjectSerialization != (mode == PdfObjectSerializationMode.ForwardOnly)) {
            throw new InvalidOperationException("PDF serialization benchmark observed the wrong object writer mode.");
        }
        if (mode == PdfObjectSerializationMode.ForwardOnly && serialization.PeakRetainedObjectBytes != 0L) {
            throw new InvalidOperationException("Forward-only object serialization retained completed object bodies.");
        }

        return new PdfPerformanceSample(
            0D,
            0L,
            bytes.LongLength,
            serialization.PeakRetainedPageContentBytes,
            serialization.PeakRetainedObjectBytes,
            serialization.LargestSerializedObjectBytes,
            serialization.IsForwardOnlyObjectSerialization);
    }

    private static PdfPerformanceSample MeasureOnce(byte[] corpus, Func<byte[], long> operation) {
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
        using var heap = new ManagedHeapSampler();
        var stopwatch = Stopwatch.StartNew();
        long output = operation(corpus);
        stopwatch.Stop();
        long peakManagedHeapBytes = heap.Stop();
        long allocated = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
        if (output <= 0) {
            throw new InvalidOperationException("PDF performance workflow produced no observable output.");
        }

        return new PdfPerformanceSample(
            stopwatch.Elapsed.TotalMilliseconds,
            allocated,
            output,
            PeakManagedHeapBytes: peakManagedHeapBytes);
    }

    private static PdfPerformanceMeasurement Summarize(string name, IReadOnlyList<PdfPerformanceSample> samples) {
        long output = samples[0].Output;
        if (samples.Any(sample => sample.Output != output)) {
            throw new InvalidOperationException(name + " produced inconsistent output between samples.");
        }

        double elapsed = samples
            .Select(sample => sample.ElapsedMilliseconds)
            .OrderBy(value => value)
            .ElementAt(samples.Count / 2);
        long allocated = samples
            .Select(sample => sample.AllocatedBytes)
            .OrderBy(value => value)
            .ElementAt(samples.Count / 2);
        PdfPerformanceSample representative = samples
            .OrderBy(sample => sample.ElapsedMilliseconds)
            .ElementAt(samples.Count / 2);
        return new PdfPerformanceMeasurement(
            name,
            elapsed,
            allocated,
            output,
            representative.PeakRetainedPageContentBytes,
            representative.PeakRetainedObjectBytes,
            representative.LargestSerializedObjectBytes,
            representative.IsForwardOnlyObjectSerialization,
            representative.PeakManagedHeapBytes);
    }
}
