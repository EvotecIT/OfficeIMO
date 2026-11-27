namespace OfficeIMO.Archives;

/// <summary>Local and operation-wide decompression accounting for formats with whole-block decoding.</summary>
internal sealed class ArchiveBlockBudget {
    private readonly ArchiveLimits _limits;
    private long _work;

    internal ArchiveBlockBudget(ArchiveLimits limits) => _limits = limits;

    internal void Charge(long expanded, long work) {
        long next = _work;
        ArchiveLimits.Add(ref next, work, _limits.DecompressionWork.MaximumBytes, nameof(ArchiveExtractionOptions.MaxDecompressionWorkBytes));
        _limits.Operation.Decode(expanded, work);
        _work = next;
    }
}
