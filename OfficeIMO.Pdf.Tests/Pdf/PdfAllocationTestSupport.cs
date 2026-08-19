namespace OfficeIMO.Tests.Pdf;

internal static class PdfAllocationTestSupport {
#if NET8_0_OR_GREATER
    private const int MeasurementAttempts = 3;

    internal static long MeasureMinimumThreadAllocation(Action operation) {
        ArgumentNullException.ThrowIfNull(operation);

        // A full test run can cross a runtime/JIT bookkeeping boundary during any one
        // sample. Persistent per-operation allocation remains present in every sample.
        long minimum = long.MaxValue;
        for (int attempt = 0; attempt < MeasurementAttempts; attempt++) {
            long before = GC.GetAllocatedBytesForCurrentThread();
            operation();
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
            minimum = Math.Min(minimum, allocated);
        }
        return minimum;
    }
#endif
}
