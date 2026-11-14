namespace OfficeIMO.Diff;

/// <summary>Edit script between two sequences produced by <see cref="OfficeDiff"/>.</summary>
/// <typeparam name="T">Item type.</typeparam>
public sealed class OfficeDiffResult<T> {
    internal OfficeDiffResult(IReadOnlyList<T> oldItems, IReadOnlyList<T> newItems, IReadOnlyList<OfficeDiffEdit> edits,
        IEqualityComparer<T> comparer, OfficeDiffAlgorithm algorithm, bool limitReached, long workUnits) {
        Old = oldItems;
        New = newItems;
        Edits = edits;
        Comparer = comparer;
        Algorithm = algorithm;
        LimitReached = limitReached;
        WorkUnits = workUnits;
        foreach (OfficeDiffEdit edit in edits) {
            if (edit.Operation == OfficeDiffOperation.Delete) DeletedCount += edit.OldLength;
            else if (edit.Operation == OfficeDiffOperation.Insert) InsertedCount += edit.NewLength;
            else EqualCount += edit.OldLength;
        }
    }

    /// <summary>Old sequence.</summary>
    public IReadOnlyList<T> Old { get; }

    /// <summary>New sequence.</summary>
    public IReadOnlyList<T> New { get; }

    /// <summary>
    /// Edit runs in order. Together they cover every old item and every new item exactly once; within each changed region
    /// the deletion comes before the insertion.
    /// </summary>
    public IReadOnlyList<OfficeDiffEdit> Edits { get; }

    /// <summary>Comparer used to match items.</summary>
    public IEqualityComparer<T> Comparer { get; }

    /// <summary>Algorithm that produced the alignment.</summary>
    public OfficeDiffAlgorithm Algorithm { get; }

    /// <summary>
    /// True when a work, time or depth limit stopped the alignment early. The edit script is still complete and correct
    /// (applying it to <see cref="Old"/> yields <see cref="New"/>), but some unchanged items may be reported as deleted and reinserted.
    /// </summary>
    public bool LimitReached { get; }

    /// <summary>Comparison steps spent on the alignment.</summary>
    public long WorkUnits { get; }

    /// <summary>Number of old items that are unchanged.</summary>
    public int EqualCount { get; }

    /// <summary>Number of old items deleted.</summary>
    public int DeletedCount { get; }

    /// <summary>Number of new items inserted.</summary>
    public int InsertedCount { get; }

    /// <summary>Edit distance: deletions plus insertions.</summary>
    public int EditDistance => DeletedCount + InsertedCount;

    /// <summary>True when the sequences are equal under <see cref="Comparer"/>.</summary>
    public bool AreEqual => DeletedCount == 0 && InsertedCount == 0;

    /// <summary>
    /// Applies the edit script to <paramref name="source"/> and returns the new sequence.
    /// Every kept and deleted item is checked against the item recorded in the script.
    /// </summary>
    /// <exception cref="InvalidOperationException">The source does not match the old sequence this script was computed from.</exception>
    public IReadOnlyList<T> Apply(IReadOnlyList<T> source) {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (source.Count != Old.Count) throw new InvalidOperationException("The source length does not match the old sequence of this diff.");
        var output = new List<T>(New.Count);
        foreach (OfficeDiffEdit edit in Edits) {
            switch (edit.Operation) {
                case OfficeDiffOperation.Equal:
                    for (int index = 0; index < edit.OldLength; index++) {
                        T item = source[edit.OldStart + index];
                        if (!Comparer.Equals(item, Old[edit.OldStart + index])) throw new InvalidOperationException("The source differs from the old sequence at item " + (edit.OldStart + index) + ".");
                        output.Add(item);
                    }

                    break;
                case OfficeDiffOperation.Delete:
                    for (int index = 0; index < edit.OldLength; index++) {
                        if (!Comparer.Equals(source[edit.OldStart + index], Old[edit.OldStart + index])) throw new InvalidOperationException("The source differs from the old sequence at item " + (edit.OldStart + index) + ".");
                    }

                    break;
                default:
                    for (int index = 0; index < edit.NewLength; index++) output.Add(New[edit.NewStart + index]);
                    break;
            }
        }

        return output;
    }

    /// <summary>Reverts the edit script on <paramref name="target"/> and returns the old sequence.</summary>
    /// <exception cref="InvalidOperationException">The target does not match the new sequence this script was computed from.</exception>
    public IReadOnlyList<T> Revert(IReadOnlyList<T> target) {
        if (target is null) throw new ArgumentNullException(nameof(target));
        if (target.Count != New.Count) throw new InvalidOperationException("The target length does not match the new sequence of this diff.");
        var output = new List<T>(Old.Count);
        foreach (OfficeDiffEdit edit in Edits) {
            switch (edit.Operation) {
                case OfficeDiffOperation.Equal:
                    for (int index = 0; index < edit.NewLength; index++) {
                        if (!Comparer.Equals(target[edit.NewStart + index], New[edit.NewStart + index])) throw new InvalidOperationException("The target differs from the new sequence at item " + (edit.NewStart + index) + ".");
                        output.Add(Old[edit.OldStart + index]);
                    }

                    break;
                case OfficeDiffOperation.Insert:
                    for (int index = 0; index < edit.NewLength; index++) {
                        if (!Comparer.Equals(target[edit.NewStart + index], New[edit.NewStart + index])) throw new InvalidOperationException("The target differs from the new sequence at item " + (edit.NewStart + index) + ".");
                    }

                    break;
                default:
                    for (int index = 0; index < edit.OldLength; index++) output.Add(Old[edit.OldStart + index]);
                    break;
            }
        }

        return output;
    }

    /// <summary>Groups the edit script into changed regions: each region is a maximal run of deletions and insertions between unchanged items.</summary>
    public IReadOnlyList<OfficeDiffHunk> GetChangedRegions() {
        var regions = new List<OfficeDiffHunk>();
        int index = 0;
        while (index < Edits.Count) {
            if (Edits[index].Operation == OfficeDiffOperation.Equal) {
                index++;
                continue;
            }

            int oldStart = Edits[index].OldStart, newStart = Edits[index].NewStart, oldLength = 0, newLength = 0;
            while (index < Edits.Count && Edits[index].Operation != OfficeDiffOperation.Equal) {
                oldLength += Edits[index].OldLength;
                newLength += Edits[index].NewLength;
                index++;
            }

            regions.Add(new OfficeDiffHunk(oldStart, oldLength, newStart, newLength));
        }

        return regions;
    }
}

/// <summary>A changed region: old items replaced by new items. Either side can be empty.</summary>
public readonly struct OfficeDiffHunk {
    /// <summary>Creates a changed region.</summary>
    public OfficeDiffHunk(int oldStart, int oldLength, int newStart, int newLength) {
        OldStart = oldStart;
        OldLength = oldLength;
        NewStart = newStart;
        NewLength = newLength;
    }

    /// <summary>Zero-based start in the old sequence.</summary>
    public int OldStart { get; }

    /// <summary>Number of old items removed.</summary>
    public int OldLength { get; }

    /// <summary>Zero-based start in the new sequence.</summary>
    public int NewStart { get; }

    /// <summary>Number of new items added.</summary>
    public int NewLength { get; }
}
