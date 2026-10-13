using System;
using System.IO;
using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.Archives;

internal sealed class ArchiveLimits {
    internal ArchiveLimits(ArchiveExtractionOptions options) {
        Input = Positive(options.MaxInputBytes, nameof(options.MaxInputBytes));
        Entries = checked((int)Positive(options.MaxEntryCount, nameof(options.MaxEntryCount)));
        Cabinets = checked((int)Positive(options.MaxCabinetCount, nameof(options.MaxCabinetCount)));
        Entry = Positive(options.MaxEntryBytes, nameof(options.MaxEntryBytes));
        Expanded = Positive(options.MaxExpandedBytes, nameof(options.MaxExpandedBytes));
        DecompressionWork = new OfficeCabinetWorkBudget(Positive(options.MaxDecompressionWorkBytes, nameof(options.MaxDecompressionWorkBytes)));
        Metadata = Positive(options.MaxMetadataBytes, nameof(options.MaxMetadataBytes));
        Rows = checked((int)Positive(options.MaxTableRows, nameof(options.MaxTableRows)));
        Strings = checked((int)Positive(options.MaxStringPoolEntries, nameof(options.MaxStringPoolEntries)));
        Depth = checked((int)Positive(options.MaxPathDepth, nameof(options.MaxPathDepth)));
        PathLength = checked((int)Positive(options.MaxPathLength, nameof(options.MaxPathLength)));
        Include = options.IncludeFile;
    }
    internal long Input { get; }
    internal int Entries { get; }
    internal int Cabinets { get; }
    internal long Entry { get; }
    internal long Expanded { get; }
    internal OfficeCabinetWorkBudget DecompressionWork { get; }
    internal long Metadata { get; }
    internal int Rows { get; }
    internal int Strings { get; }
    internal int Depth { get; }
    internal int PathLength { get; }
    internal Func<string, bool>? Include { get; }
    internal bool Select(string path) => Include?.Invoke(path) ?? true;
    internal void CheckEntry(long bytes) {
        if (bytes < 0 || bytes > Entry) throw new InvalidDataException("Archive member exceeds the configured byte limit.");
    }
    internal static void Add(ref long total, long bytes, long maximum, string description) {
        if (bytes < 0 || total > maximum - bytes) throw new InvalidDataException(description + " exceeds the configured byte limit.");
        total += bytes;
    }
    private static long Positive(long value, string name) => value > 0 ? value : throw new ArgumentOutOfRangeException(name);
}
