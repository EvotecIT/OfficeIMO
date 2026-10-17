namespace OfficeIMO.Archives;

/// <summary>Detached logical resource charges. These counters do not measure live allocations or peak process memory.</summary>
public sealed class ArchiveResourceUsage {
    internal ArchiveResourceUsage(long input, long entries, long expanded, long selected, long work, int depth) {
        InputBytes = input; EntryCount = entries; ExpandedBytes = expanded; SelectedBytes = selected;
        DecompressionWorkBytes = work; MaximumDepth = depth;
    }
    /// <summary>Compressed input admitted across the operation.</summary>
    public long InputBytes { get; }
    /// <summary>Physical records admitted across the operation.</summary>
    public long EntryCount { get; }
    /// <summary>Attempted decoded expansion, including repeated and failed reads.</summary>
    public long ExpandedBytes { get; }
    /// <summary>Cumulative selected payload admission, including attempts that fail after reservation.</summary>
    public long SelectedBytes { get; }
    /// <summary>Charged processing units, including failed attempts after reservation.</summary>
    public long DecompressionWorkBytes { get; }
    /// <summary>Greatest admitted container depth.</summary>
    public int MaximumDepth { get; }
}
