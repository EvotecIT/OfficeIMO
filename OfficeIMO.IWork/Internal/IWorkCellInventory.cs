namespace OfficeIMO.IWork.Internal;

/// <summary>Assesses only cells already materialized under the source-wide cell budget.</summary>
internal sealed class IWorkCellInventory {
    internal IWorkCellInventory(IWorkSourceDocument source, IEnumerable<IWorkTable>? tables) {
        var formulas = new List<IWorkFormulaCellStatus>();
        var issues = new List<IWorkSourceCellIssue>();
        if (tables != null) {
            foreach (IWorkTable table in tables) {
                source.CancellationToken.ThrowIfCancellationRequested();
                foreach (IWorkTableCell cell in table.Cells) {
                    source.CancellationToken.ThrowIfCancellationRequested();
                    if (cell.SourceFormulaIsDeclared) formulas.Add(new IWorkFormulaCellStatus(table, cell));
                    if (cell.HasDecodeError) issues.Add(new IWorkSourceCellIssue(table, cell));
                }
            }
        }
        FormulaCells = formulas;
        SourceCellIssues = issues;
    }

    internal IReadOnlyList<IWorkFormulaCellStatus> FormulaCells { get; }
    internal IReadOnlyList<IWorkSourceCellIssue> SourceCellIssues { get; }
}
