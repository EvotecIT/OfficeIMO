namespace OfficeIMO.AsciiDoc;

/// <summary>Immutable origin ranges for one explicitly processed source document.</summary>
/// <remarks>Offsets refer to the original <see cref="AsciiDocProcessingResult.ProcessedSource"/> snapshot. Editing the returned document does not update this map.</remarks>
public sealed class AsciiDocProcessingSourceMap {
    private readonly IReadOnlyList<AsciiDocSourceMapping> _entries;
    internal AsciiDocProcessingSourceMap(IEnumerable<AsciiDocSourceMapping> entries, int length) {
        _entries = Array.AsReadOnly(entries.ToArray());
        Length = length;
    }
    /// <summary>Number of UTF-16 characters in the processed snapshot.</summary>
    public int Length { get; }
    /// <summary>Nonoverlapping origin ranges in processed order.</summary>
    public IReadOnlyList<AsciiDocSourceMapping> Entries => _entries;
    /// <summary>Finds the origin range at a processed offset, including the end-of-source boundary.</summary>
    /// <returns>Null for an empty processed document.</returns>
    public AsciiDocSourceMapping? Find(int processedOffset) {
        if (processedOffset < 0 || processedOffset > Length) throw new ArgumentOutOfRangeException(nameof(processedOffset));
        if (_entries.Count == 0) return null;
        if (processedOffset == Length) return _entries[_entries.Count - 1];
        int low = 0, high = _entries.Count - 1;
        while (low <= high) {
            int middle = low + (high - low) / 2;
            AsciiDocSourceMapping entry = _entries[middle];
            if (processedOffset < entry.ProcessedStart) high = middle - 1;
            else if (processedOffset >= entry.ProcessedEnd) low = middle + 1;
            else return entry;
        }
        return null;
    }
}

/// <summary>A processed text range and the original line or directive that produced it.</summary>
public sealed class AsciiDocSourceMapping {
    internal AsciiDocSourceMapping(int start, int length, string? sourceName, AsciiDocSourceSpan originalSpan, bool exact) {
        ProcessedStart = start; ProcessedEnd = checked(start + length); SourceName = sourceName; OriginalSpan = originalSpan; IsExact = exact;
    }
    /// <summary>Inclusive zero-based UTF-16 offset in processed source.</summary>
    public int ProcessedStart { get; }
    /// <summary>Exclusive zero-based UTF-16 offset in processed source.</summary>
    public int ProcessedEnd { get; }
    /// <summary>Origin identifier supplied by the caller or include resolver.</summary>
    public string? SourceName { get; }
    /// <summary>Original source range. Generated text refers to its producing directive.</summary>
    public AsciiDocSourceSpan OriginalSpan { get; }
    /// <summary>True when this range is a character-for-character slice of one original line.</summary>
    public bool IsExact { get; }
    /// <summary>Maps an offset to an exact original position, or the producing range's start for transformed text.</summary>
    public AsciiDocSourcePosition GetSourcePosition(int processedOffset) {
        if (processedOffset < ProcessedStart || processedOffset > ProcessedEnd) throw new ArgumentOutOfRangeException(nameof(processedOffset));
        if (!IsExact) return OriginalSpan.Start;
        int relative = processedOffset - ProcessedStart;
        if (relative == OriginalSpan.Length) return OriginalSpan.End;
        return new AsciiDocSourcePosition(OriginalSpan.Start.Offset + relative, OriginalSpan.Start.Line, OriginalSpan.Start.Column + relative);
    }
    internal AsciiDocSourceMapping Shift(int offset) => new AsciiDocSourceMapping(checked(ProcessedStart + offset), ProcessedEnd - ProcessedStart, SourceName, OriginalSpan, IsExact);
}
