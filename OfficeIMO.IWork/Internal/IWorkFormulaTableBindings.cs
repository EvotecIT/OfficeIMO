using System.Threading;

namespace OfficeIMO.IWork.Internal;

internal static class IWorkFormulaTableBindings {
    // Duplicate native identities retire every candidate, including one whose label cannot be rendered.
    internal static IReadOnlyDictionary<Guid, string> Create(IEnumerable<IWorkTable> tables,
        Func<IWorkTable, string?> qualifier, CancellationToken cancellationToken) {
        var result = new Dictionary<Guid, string>();
        var seen = new HashSet<Guid>();
        foreach (IWorkTable table in tables) {
            cancellationToken.ThrowIfCancellationRequested();
            if (table.FormulaIdentifier is not Guid identifier) continue;
            if (!seen.Add(identifier)) { result.Remove(identifier); continue; }
            string? name = qualifier(table);
            if (name != null) result.Add(identifier, name);
        }
        return result;
    }
}
