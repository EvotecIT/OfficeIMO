internal sealed record PdfPerformanceBudget(
    Dictionary<string, PdfWorkloadBudget> Workloads,
    double MinimumCachedSpeedup,
    double MinimumCachedAllocationReduction);

internal sealed record PdfWorkloadBudget(
    double MaxElapsedMilliseconds,
    long MaxAllocatedBytes);

internal sealed record PdfPerformanceMeasurement(
    string Name,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long Output);

internal sealed record PdfPerformanceSample(
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long Output);
