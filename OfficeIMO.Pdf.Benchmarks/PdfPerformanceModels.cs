internal sealed record PdfPerformanceBudget(
    Dictionary<string, PdfWorkloadBudget> Workloads,
    double MinimumCachedSpeedup,
    double MinimumCachedAllocationReduction);

internal sealed record PdfWorkloadBudget(
    double MaxElapsedMilliseconds,
    long MaxAllocatedBytes,
    long MaxPeakManagedHeapBytes);

internal sealed record PdfPerformanceMeasurement(
    string Name,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long Output,
    long PeakRetainedPageContentBytes = 0L,
    long PeakRetainedObjectBytes = 0L,
    long LargestSerializedObjectBytes = 0L,
    bool IsForwardOnlyObjectSerialization = false,
    long PeakManagedHeapBytes = 0L) {
    internal long LargestTransientBufferBytes =>
        Math.Max(
            PeakRetainedPageContentBytes,
            Math.Max(PeakRetainedObjectBytes, LargestSerializedObjectBytes));
}

internal sealed record PdfPerformanceSample(
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long Output,
    long PeakRetainedPageContentBytes = 0L,
    long PeakRetainedObjectBytes = 0L,
    long LargestSerializedObjectBytes = 0L,
    bool IsForwardOnlyObjectSerialization = false,
    long PeakManagedHeapBytes = 0L);
