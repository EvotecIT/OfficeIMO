namespace OfficeIMO.Pdf;

/// <summary>
/// The owners of each source run, as a per-run set that stays a single reference until a run gets a second owner.
/// Owners enumerate in the order they were first added, as a <see cref="HashSet{T}"/> without removals does.
/// </summary>
internal sealed class PdfSourceRunIndex<TOwner> where TOwner : class {
    private readonly Dictionary<PdfTextSpan, object> _owners = new();

    internal void Add(PdfTextSpan run, TOwner owner) {
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
    internal bool TryGet(PdfTextSpan run, out TOwner? single, out HashSet<TOwner>? several) {
        single = null;
        several = null;
        if (!_owners.TryGetValue(run, out object? existing)) return false;
        if (existing is HashSet<TOwner> owners) several = owners;
        else single = (TOwner)existing;
        return true;
    }
}
