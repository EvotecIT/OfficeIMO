namespace OfficeIMO.Pdf;

/// <summary>
/// A font registration map that records which change produced its current entries. Text measurement
/// compares its snapshots with the live options many times per document; equal stamps prove the same
/// entries without walking both maps.
/// </summary>
/// <remarks>
/// Writes must go through this type so every change takes a new stamp. Copies made with
/// <see cref="CopyWithSameEntries"/> keep the stamp because they hold the same entry references.
/// </remarks>
internal sealed class PdfFontMembershipMap<TKey, TValue> : Dictionary<TKey, TValue> where TKey : notnull {
    private static long _lastStamp;

    internal PdfFontMembershipMap(IEqualityComparer<TKey>? comparer = null) : base(comparer) {
        Stamp = NextStamp();
    }

    private PdfFontMembershipMap(PdfFontMembershipMap<TKey, TValue> source) : base(source, source.Comparer) {
        Stamp = source.Stamp;
    }

    /// <summary>Unique to the change that produced the current entries.</summary>
    internal long Stamp { get; private set; }

    internal PdfFontMembershipMap<TKey, TValue> CopyWithSameEntries() => new(this);

    internal new TValue this[TKey key] {
        get => base[key];
        set {
            base[key] = value;
            Stamp = NextStamp();
        }
    }

    internal new void Add(TKey key, TValue value) {
        base.Add(key, value);
        Stamp = NextStamp();
    }

    internal new bool Remove(TKey key) {
        if (!base.Remove(key)) return false;
        Stamp = NextStamp();
        return true;
    }

    internal new void Clear() {
        if (Count == 0) return;
        base.Clear();
        Stamp = NextStamp();
    }

    internal static bool HaveSameEntries(PdfFontMembershipMap<TKey, TValue>? first, PdfFontMembershipMap<TKey, TValue>? second) =>
        first != null && second != null && first.Stamp == second.Stamp;

    private static long NextStamp() => Interlocked.Increment(ref _lastStamp);
}
