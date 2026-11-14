namespace OfficeIMO.Diff;

/// <summary>Selects the sequence-alignment strategy used by <see cref="OfficeDiff"/>.</summary>
public enum OfficeDiffAlgorithm {
    /// <summary>Myers' O(ND) algorithm with linear-space refinement. Produces a minimal edit script unless a work limit is reached.</summary>
    Myers,
    /// <summary>Patience diff: aligns on items that occur exactly once on both sides, then uses Myers between anchors. Often reads better for moved or repeated blocks; not guaranteed minimal.</summary>
    Patience,
    /// <summary>Histogram diff: aligns on the rarest common items first, then uses Myers where no rare anchor exists. Not guaranteed minimal.</summary>
    Histogram
}

/// <summary>Operation of one <see cref="OfficeDiffEdit"/>.</summary>
public enum OfficeDiffOperation {
    /// <summary>Items present unchanged on both sides.</summary>
    Equal,
    /// <summary>Items present only in the old sequence.</summary>
    Delete,
    /// <summary>Items present only in the new sequence.</summary>
    Insert
}

/// <summary>One contiguous run of an edit script.</summary>
/// <remarks>
/// <see cref="OldStart"/> and <see cref="NewStart"/> are always the positions in each sequence where the run applies.
/// For <see cref="OfficeDiffOperation.Insert"/> the old length is zero; for <see cref="OfficeDiffOperation.Delete"/> the new length is zero.
/// </remarks>
public readonly struct OfficeDiffEdit : IEquatable<OfficeDiffEdit> {
    /// <summary>Creates an edit run.</summary>
    public OfficeDiffEdit(OfficeDiffOperation operation, int oldStart, int oldLength, int newStart, int newLength) {
        if (oldStart < 0) throw new ArgumentOutOfRangeException(nameof(oldStart));
        if (oldLength < 0) throw new ArgumentOutOfRangeException(nameof(oldLength));
        if (newStart < 0) throw new ArgumentOutOfRangeException(nameof(newStart));
        if (newLength < 0) throw new ArgumentOutOfRangeException(nameof(newLength));
        if (operation == OfficeDiffOperation.Equal && oldLength != newLength) throw new ArgumentException("An equal run covers the same number of items on both sides.", nameof(newLength));
        if (operation == OfficeDiffOperation.Delete && newLength != 0) throw new ArgumentException("A delete run has no new items.", nameof(newLength));
        if (operation == OfficeDiffOperation.Insert && oldLength != 0) throw new ArgumentException("An insert run has no old items.", nameof(oldLength));
        Operation = operation;
        OldStart = oldStart;
        OldLength = oldLength;
        NewStart = newStart;
        NewLength = newLength;
    }

    /// <summary>Run operation.</summary>
    public OfficeDiffOperation Operation { get; }

    /// <summary>Zero-based start in the old sequence.</summary>
    public int OldStart { get; }

    /// <summary>Number of old items covered.</summary>
    public int OldLength { get; }

    /// <summary>Zero-based start in the new sequence.</summary>
    public int NewStart { get; }

    /// <summary>Number of new items covered.</summary>
    public int NewLength { get; }

    /// <inheritdoc />
    public bool Equals(OfficeDiffEdit other) =>
        Operation == other.Operation && OldStart == other.OldStart && OldLength == other.OldLength &&
        NewStart == other.NewStart && NewLength == other.NewLength;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeDiffEdit other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = (int)Operation;
            hash = (hash * 397) ^ OldStart;
            hash = (hash * 397) ^ OldLength;
            hash = (hash * 397) ^ NewStart;
            return (hash * 397) ^ NewLength;
        }
    }

    /// <inheritdoc />
    public override string ToString() => Operation + " old[" + OldStart + "+" + OldLength + "] new[" + NewStart + "+" + NewLength + "]";
}

/// <summary>Settings for <see cref="OfficeDiff"/>.</summary>
public sealed class OfficeDiffOptions {
    /// <summary>Alignment strategy. The default is <see cref="OfficeDiffAlgorithm.Myers"/>.</summary>
    public OfficeDiffAlgorithm Algorithm { get; set; } = OfficeDiffAlgorithm.Myers;

    /// <summary>
    /// Merges short equal runs that sit between changes into the surrounding change, so that a change reads as one replacement
    /// instead of alternating fragments. The weight of a run is its item count, or its character count for text tokens.
    /// The default is <c>false</c>.
    /// </summary>
    public bool SemanticCleanup { get; set; }

    /// <summary>
    /// Maximum number of comparison steps before the remaining regions are reported as whole deletions and insertions.
    /// Zero or less means no limit. The default is 50,000,000.
    /// </summary>
    public long MaxWorkUnits { get; set; } = 50_000_000;

    /// <summary>Maximum time spent aligning before the remaining regions are reported as whole deletions and insertions. Null means no limit.</summary>
    public TimeSpan? Timeout { get; set; }

    /// <summary>Text comparison ignores case when <c>true</c>. Applies to the text helpers only.</summary>
    public bool IgnoreCase { get; set; }

    /// <summary>
    /// Text comparison treats every whitespace token as equal to any other whitespace token when <c>true</c>, so changes in spacing or line breaks
    /// between words are not reported. Applies to the word and line helpers only.
    /// </summary>
    public bool IgnoreWhitespace { get; set; }

    internal OfficeDiffOptions Clone() => (OfficeDiffOptions)MemberwiseClone();
}
