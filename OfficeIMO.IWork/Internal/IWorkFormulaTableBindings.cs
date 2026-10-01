using System.Threading;

namespace OfficeIMO.IWork.Internal;

internal static class IWorkFormulaTableBindings {
    // Duplicate native identities retire every candidate, including one whose label cannot be rendered.
    internal static IReadOnlyDictionary<Guid, IWorkFormulaTableBinding> Create(IEnumerable<IWorkTable> tables,
        Func<IWorkTable, string?> qualifier, CancellationToken cancellationToken, bool boundBodyRanges = false) {
        var result = new Dictionary<Guid, IWorkFormulaTableBinding>();
        var seen = new HashSet<Guid>();
        foreach (IWorkTable table in tables) {
            cancellationToken.ThrowIfCancellationRequested();
            if (table.FormulaIdentifier is not Guid identifier) continue;
            if (!seen.Add(identifier)) { result.Remove(identifier); continue; }
            string? name = qualifier(table);
            if (name != null) result.Add(identifier, new IWorkFormulaTableBinding(table, name, boundBodyRanges));
        }
        return result;
    }
}

/// <summary>Connects a source identity to a context label and its validated table body.</summary>
internal sealed class IWorkFormulaTableBinding {
    internal IWorkFormulaTableBinding(IWorkTable table, string qualifier, bool boundBodyRanges) {
        Table = table; Qualifier = qualifier; BoundBodyRanges = boundBodyRanges;
    }
    internal IWorkTable Table { get; }
    internal string Qualifier { get; }
    internal bool BoundBodyRanges { get; }
}
