namespace OfficeIMO.Pdf;

/// <summary>
/// The owners of each key, as a per-key set that stays a single reference until a key gets a second owner.
/// Owners enumerate in the order they were first added, as a <see cref="HashSet{T}"/> without removals does.
/// </summary>
internal class PdfOwnerIndex<TKey, TOwner> where TKey : notnull where TOwner : class {
    private readonly Dictionary<TKey, object> _owners;

    internal PdfOwnerIndex(IEqualityComparer<TKey>? comparer = null) => _owners = new Dictionary<TKey, object>(comparer);

    internal void Add(TKey run, TOwner owner) {
        if (!_owners.TryGetValue(run, out object? existing)) {
            _owners.Add(run, owner);
            return;
        }
        if (existing is HashSet<TOwner> owners) {
            owners.Add(owner);
            return;
        }
        var single = (TOwner)existing;
        if (EqualityComparer<TOwner>.Default.Equals(single, owner)) return;
        _owners[run] = new HashSet<TOwner> { single, owner };
    }

    /// <summary>The run's owners: one in <paramref name="single"/>, or several in <paramref name="several"/>.</summary>
    internal bool TryGet(TKey run, out TOwner? single, out HashSet<TOwner>? several) {
        single = null;
        several = null;
        if (!_owners.TryGetValue(run, out object? existing)) return false;
        if (existing is HashSet<TOwner> owners) several = owners;
        else single = (TOwner)existing;
        return true;
    }
}

/// <summary>The owners of each source run; see <see cref="PdfOwnerIndex{TKey, TOwner}"/>.</summary>
internal sealed class PdfSourceRunIndex<TOwner> : PdfOwnerIndex<PdfTextSpan, TOwner> where TOwner : class {
}
