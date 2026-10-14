using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.Archives;

internal sealed partial class ArchiveLimits {
    internal ArchiveLimits(ArchiveExtractionOptions options) {
        if (options.AllowedFormats != null) {
            AllowedFormats = new HashSet<ArchiveFormat>(options.AllowedFormats);
            foreach (ArchiveFormat format in AllowedFormats)
                if (format != ArchiveFormat.Cab && format != ArchiveFormat.Msi && format != ArchiveFormat.Zip)
                    throw new ArgumentOutOfRangeException(nameof(options.AllowedFormats));
        }
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
    internal HashSet<ArchiveFormat>? AllowedFormats { get; }
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
        Check(bytes, Entry, nameof(ArchiveExtractionOptions.MaxEntryBytes));
    }
    internal static void Check(long value, long maximum, string limitName) {
        if (value < 0 || maximum < 0) throw new InvalidDataException("Archive sizes or counts cannot be negative.");
        if (value > maximum) throw Exceeded(limitName, maximum);
    }
    internal static OfficeInspectionResourceLimitException Exceeded(string limitName, long maximum, Exception? inner = null) =>
        new("archives", limitName, maximum, inner);
    internal static void Add(ref long total, long bytes, long maximum, string limitName) {
        if (bytes < 0 || total < 0 || maximum < 0) throw new InvalidDataException("Archive aggregate sizes cannot be negative.");
        if (total > maximum || bytes > maximum - total) throw Exceeded(limitName, maximum);
        total += bytes;
    }
    private static long Positive(long value, string name) => value > 0 ? value : throw new ArgumentOutOfRangeException(name);
}
