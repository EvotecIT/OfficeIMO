using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

internal static partial class IWorkNumbersReader {
    private static void BindTableFormulas(IWorkSourceDocument source, List<IWorkNumbersSheet> sheets,
        IWorkProjectionBudget budget, List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        var labels = new Dictionary<IWorkTable, string?>();
        foreach (IWorkNumbersSheet sheet in sheets) foreach (IWorkTable table in sheet.Tables) {
            source.CancellationToken.ThrowIfCancellationRequested();
            string? sheetName = IWorkFormulaReader.QuoteTableName(sheet.Name, source.Options.MaximumFormulaCharacters);
            string? tableName = IWorkFormulaReader.QuoteTableName(table.Name, source.Options.MaximumFormulaCharacters);
            labels.Add(table, sheetName != null && tableName != null ? sheetName + "::" + tableName + "::" : null);
        }
        IReadOnlyDictionary<Guid, string> qualifiers = IWorkFormulaTableBindings.Create(labels.Keys,
            table => labels[table], source.CancellationToken);
        var resolved = new Dictionary<IWorkTable, IWorkTable>();
        foreach (IWorkTable table in labels.Keys) {
            var cells = new List<IWorkTableCell>(table.Cells.Count);
            foreach (IWorkTableCell cell in table.Cells) {
                source.CancellationToken.ThrowIfCancellationRequested();
                if (cell.FormulaDefinition == null) cells.Add(cell);
                else {
                    IWorkFormulaResult result = cell.FormulaDefinition.Render(qualifiers, budget);
                    budget.AddTextCharacters(result.Text.Length);
                    cells.Add(cell.WithFormula(result));
                }
            }
            if (table.ModelRecord != null)
                IWorkTableReader.AssessFormulas(cells, table.Name, table.ModelRecord, diagnostics, ref supportsEditableReconstruction);
            resolved.Add(table, table.Cells.Any(cell => cell.FormulaDefinition != null) ? table.WithCells(cells) : table);
        }
        for (int index = 0; index < sheets.Count; index++) {
            IWorkNumbersSheet sheet = sheets[index];
            sheets[index] = new IWorkNumbersSheet(sheet.Name, sheet.Tables.Select(table => resolved[table]).ToArray(),
                sheet.TextBoxes, sheet.Drawables.Select(drawable => drawable.Table != null
                    ? new IWorkNumbersDrawable(resolved[drawable.Table]) : drawable).ToArray(), sheet.SourceIdentity);
        }
    }
}
