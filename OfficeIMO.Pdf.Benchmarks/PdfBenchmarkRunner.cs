using System.Diagnostics;
using System.Text.Json;
using OfficeIMO.Pdf;

internal static class PdfBenchmarkRunner {
    internal const string AnalysisCold = "analysis-cold";
    internal const string AnalysisCached = "analysis-cached";
    internal const string RenderSvg = "render-svg-12";
    internal const string RenderPng = "render-png-4";
    internal const string ComposeSerializeBuffered = "compose-serialize-buffered-60";
    internal const string ComposeSerializeForward = "compose-serialize-forward-60";
    internal const string SerializePrecomposedBuffered = "layout-serialize-precomposed-buffered-60";
    internal const string SerializePrecomposedForward = "layout-serialize-precomposed-forward-60";
    internal const string ComposeSerializeHarfBuzz = "compose-serialize-harfbuzz-60";

    internal static IReadOnlyList<PdfPerformanceMeasurement> Measure(
        byte[] corpus,
        out double cachedSpeedup,
        out double cachedAllocationReduction) {
        (PdfPerformanceMeasurement cold, PdfPerformanceMeasurement cached, cachedSpeedup, cachedAllocationReduction) = MeasureAnalysis(corpus);
        PdfPerformanceMeasurement svg = MeasureWorkflow(RenderSvg, corpus, RunSvgRender, sampleCount: 5);
        PdfPerformanceMeasurement png = MeasureWorkflow(RenderPng, corpus, RunPngRender, sampleCount: 3);
        PdfPerformanceMeasurement buffered = MeasureSerialization(
            ComposeSerializeBuffered,
            PdfObjectSerializationMode.Buffered,
            sampleCount: 3,
            includeComposition: true);
        PdfPerformanceMeasurement forward = MeasureSerialization(
            ComposeSerializeForward,
            PdfObjectSerializationMode.ForwardOnly,
            sampleCount: 3,
            includeComposition: true);
        PdfPerformanceMeasurement precomposedBuffered = MeasureSerialization(
            SerializePrecomposedBuffered,
            PdfObjectSerializationMode.Buffered,
            sampleCount: 3,
            includeComposition: false);
        PdfPerformanceMeasurement precomposedForward = MeasureSerialization(
            SerializePrecomposedForward,
            PdfObjectSerializationMode.ForwardOnly,
            sampleCount: 3,
            includeComposition: false);
        PdfPerformanceMeasurement harfBuzz = MeasureSerialization(
            ComposeSerializeHarfBuzz,
            PdfObjectSerializationMode.ForwardOnly,
            sampleCount: 3,
            includeComposition: true,
            static _ => PdfBenchmarkCorpus.CreateHarfBuzzDocument());
        return new[] {
            cold, cached, svg, png,
            buffered, forward,
            precomposedBuffered, precomposedForward,
            harfBuzz
        };
    }

    private static (
        PdfPerformanceMeasurement Cold,
        PdfPerformanceMeasurement Cached,
        double CachedSpeedup,
        double CachedAllocationReduction) MeasureAnalysis(byte[] corpus) {
        const int sampleCount = 11;
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

        PdfPerformanceMeasurement cold = Summarize(AnalysisCold, coldSamples) with {
            PeakManagedHeapBytes = MeasurePeakManagedHeap(corpus, RunColdAnalysis)
        };
        PdfPerformanceMeasurement cached = Summarize(AnalysisCached, cachedSamples) with {
            PeakManagedHeapBytes = MeasurePeakManagedHeap(corpus, RunCachedAnalysis)
        };
        return (
            cold,
            cached,
            RatioOfMedians(coldSamples, cachedSamples, sample => sample.ElapsedMilliseconds),
            RatioOfMedians(coldSamples, cachedSamples, sample => sample.AllocatedBytes));
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

        return Summarize(name, samples) with {
            PeakManagedHeapBytes = MeasurePeakManagedHeap(corpus, operation)
        };
    }

    private static long RunColdAnalysis(byte[] corpus) {
        // Keep this comparison focused on the canonical parse cache. Mixing unrelated analysis
        // stages into the ratio hides the reuse signal behind their independent work.
        long output = 0L;
        PdfDocumentInfo? last = null;
        for (int operation = 0; operation < 4; operation++) {
            last = PdfDocument.Load(corpus).Inspect();
            output += last.PageCount;
        }
        return CombineInspectionOutput(output, last!);
    }

    private static long RunCachedAnalysis(byte[] corpus) {
        PdfDocument document = PdfDocument.Load(corpus);
        long output = 0L;
        PdfDocumentInfo? last = null;
        for (int operation = 0; operation < 4; operation++) {
            last = document.Inspect();
            output += last.PageCount;
        }
        return CombineInspectionOutput(output, last!);
    }

    private static long CombineInspectionOutput(long pageCount, PdfDocumentInfo info) {
        byte[] contract = JsonSerializer.SerializeToUtf8Bytes(info);
        ulong checksum = 14695981039346656037UL;
        foreach (byte value in contract) {
            checksum ^= value;
            checksum *= 1099511628211UL;
        }
        return (long)((checksum ^ (ulong)pageCount) & 0x3FFF_FFFF_FFFF_FFFFUL) + 1L;
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
            .Load(corpus)
            .Render
            .Pages(pageRanges, options);
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
        bool includeComposition,
        Func<PdfObjectSerializationMode, PdfDocument>? createDocument = null) {
        Func<PdfObjectSerializationMode, PdfDocument> factory =
            createDocument ?? PdfBenchmarkCorpus.CreateDocument;
        PdfSerializationArtifact warmup = RunSerialization(mode, factory, includeComposition ? null : factory(mode));
        ValidateSerialization(warmup, mode);
        var samples = new List<PdfPerformanceSample>(sampleCount);
        for (int sample = 0; sample < sampleCount; sample++) {
            PdfDocument? precomposed = includeComposition ? null : factory(mode);
            GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
            long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
            var stopwatch = Stopwatch.StartNew();
            PdfSerializationArtifact artifact = RunSerialization(mode, factory, precomposed);
            stopwatch.Stop();
            long allocatedBytes = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
            ValidateSerialization(artifact, mode);
            samples.Add(artifact.Sample with {
                ElapsedMilliseconds = stopwatch.Elapsed.TotalMilliseconds,
                AllocatedBytes = allocatedBytes
            });
        }

        return Summarize(name, samples) with {
            PeakManagedHeapBytes = MeasureSerializationPeakManagedHeap(mode, factory, includeComposition)
        };
    }

    private static PdfSerializationArtifact RunSerialization(
        PdfObjectSerializationMode mode,
        Func<PdfObjectSerializationMode, PdfDocument> createDocument,
        PdfDocument? precomposed) {
        PdfDocument document = precomposed ?? createDocument(mode);
        using var output = new MemoryStream();
        PdfSaveResult save = document.Save(output);
        PdfSerializationReport serialization = save.Serialization
            ?? throw new InvalidOperationException("PDF serialization benchmark did not return runtime evidence.");
        byte[] bytes = output.ToArray();
        return new PdfSerializationArtifact(
            bytes,
            new PdfPerformanceSample(
                0D,
                0L,
                bytes.LongLength,
                serialization.PeakRetainedPageContentBytes,
                serialization.PeakRetainedObjectBytes,
                serialization.LargestSerializedObjectBytes,
                serialization.IsForwardOnlyObjectSerialization));
    }

    private static void ValidateSerialization(PdfSerializationArtifact artifact, PdfObjectSerializationMode mode) {
        if (PdfDocument.Load(artifact.Bytes).Inspect().PageCount != PdfBenchmarkCorpus.PageCount) {
            throw new InvalidOperationException("PDF serialization benchmark produced an invalid page count.");
        }
        if (artifact.Sample.IsForwardOnlyObjectSerialization != (mode == PdfObjectSerializationMode.ForwardOnly)) {
            throw new InvalidOperationException("PDF serialization benchmark observed the wrong object writer mode.");
        }
        if (mode == PdfObjectSerializationMode.ForwardOnly && artifact.Sample.PeakRetainedObjectBytes != 0L) {
            throw new InvalidOperationException("Forward-only object serialization retained completed object bodies.");
        }
    }

    private sealed record PdfSerializationArtifact(byte[] Bytes, PdfPerformanceSample Sample);

    private static PdfPerformanceSample MeasureOnce(
        byte[] corpus,
        Func<byte[], long> operation) {
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
        var stopwatch = Stopwatch.StartNew();
        long output = operation(corpus);
        stopwatch.Stop();
        long allocated = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
        if (output <= 0) {
            throw new InvalidOperationException("PDF performance workflow produced no observable output.");
        }

        return new PdfPerformanceSample(
            stopwatch.Elapsed.TotalMilliseconds,
            allocated,
            output);
    }

    private static long MeasurePeakManagedHeap(byte[] corpus, Func<byte[], long> operation) {
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        using var heap = new ManagedHeapSampler();
        long output = operation(corpus);
        long peakManagedHeapBytes = heap.Stop();
        if (output <= 0) {
            throw new InvalidOperationException("PDF peak-heap workflow produced no observable output.");
        }
        return peakManagedHeapBytes;
    }

    private static long MeasureSerializationPeakManagedHeap(
        PdfObjectSerializationMode mode,
        Func<PdfObjectSerializationMode, PdfDocument> factory,
        bool includeComposition) {
        PdfDocument? precomposed = includeComposition ? null : factory(mode);
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        using var heap = new ManagedHeapSampler();
        PdfSerializationArtifact artifact = RunSerialization(mode, factory, precomposed);
        long peakManagedHeapBytes = heap.Stop();
        ValidateSerialization(artifact, mode);
        return peakManagedHeapBytes;
    }

    private static double RatioOfMedians(
        IReadOnlyList<PdfPerformanceSample> numerators,
        IReadOnlyList<PdfPerformanceSample> denominators,
        Func<PdfPerformanceSample, double> selector) {
        if (numerators.Count != denominators.Count || numerators.Count == 0) {
            throw new InvalidOperationException("PDF comparison workloads require matching samples.");
        }

        double numeratorMedian = numerators
            .Select(selector)
            .OrderBy(value => value)
            .ElementAt(numerators.Count / 2);
        double denominatorMedian = denominators
            .Select(selector)
            .OrderBy(value => value)
            .ElementAt(denominators.Count / 2);
        return numeratorMedian / Math.Max(denominatorMedian, 0.001D);
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
