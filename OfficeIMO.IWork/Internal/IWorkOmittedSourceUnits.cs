namespace OfficeIMO.IWork.Internal;

/// <summary>Keeps the first omission for each physical source record in encounter order.</summary>
internal sealed class IWorkOmittedSourceUnits {
    private readonly List<IWorkObjectIdentity> _items = new();
    private readonly HashSet<ulong> _identifiers = new();

    internal IReadOnlyList<IWorkObjectIdentity> Items => _items;

    internal void Add(IWorkObjectIdentity identity) {
        if (_identifiers.Add(identity.RecordIdentifier)) _items.Add(identity);
    }
}
