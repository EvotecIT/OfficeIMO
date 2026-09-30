namespace OfficeIMO.IWork.Internal;

internal static class IWorkFormulaInventory {
    internal static IReadOnlyList<IWorkFormulaCellStatus> Create(IWorkSourceDocument source, IEnumerable<IWorkTable>? tables) {
        var result = new List<IWorkFormulaCellStatus>();
        if (tables == null) return result;
        foreach (IWorkTable table in tables) {
            source.CancellationToken.ThrowIfCancellationRequested();
            foreach (IWorkTableCell cell in table.Cells) {
                source.CancellationToken.ThrowIfCancellationRequested();
                if (cell.Kind == IWorkCellKind.Formula) result.Add(new IWorkFormulaCellStatus(table, cell));
            }
        }
        return result;
    }
}
