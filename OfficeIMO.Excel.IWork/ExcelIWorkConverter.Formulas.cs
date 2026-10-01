using System.Threading;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.Excel.IWork;

public static partial class ExcelIWorkConverter {
    private static Dictionary<IWorkTable, ExcelSheet> CreateTableWorksheets(ExcelDocument document,
        IWorkNumbersProjection projection, List<NumbersWorksheetMapping> mappings,
        ExcelSheetNameValidationMode nameMode, CancellationToken cancellationToken) {
        var tables = new Dictionary<IWorkTable, ExcelSheet>();
        for (int sheetIndex = 0; sheetIndex < projection.Sheets.Count; sheetIndex++) {
            IWorkNumbersSheet sourceSheet = projection.Sheets[sheetIndex];
            cancellationToken.ThrowIfCancellationRequested();
            if (sourceSheet.TextBoxes.Count > 0 || sourceSheet.Tables.Count == 0) {
                ExcelSheet text = document.AddWorksheet(sourceSheet.Name, nameMode);
                mappings.Add(new NumbersWorksheetMapping(sheetIndex + 1, sourceSheet.Name, null, null, sourceSheet.Name, text.Name));
                for (int index = 0; index < sourceSheet.TextBoxes.Count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    text.CellAt(index + 1, 1).SetValue(sourceSheet.TextBoxes[index]);
                }
            }
            for (int tableIndex = 0; tableIndex < sourceSheet.Tables.Count; tableIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                IWorkTable table = sourceSheet.Tables[tableIndex];
                string requested = sourceSheet.Tables.Count == 1 && sourceSheet.TextBoxes.Count == 0
                    ? sourceSheet.Name : sourceSheet.Name + " - " + (table.Name.Length > 0 ? table.Name : $"Table {tableIndex + 1}");
                ExcelSheet sheet = document.AddWorksheet(requested, nameMode);
                mappings.Add(new NumbersWorksheetMapping(sheetIndex + 1, sourceSheet.Name,
                    tableIndex + 1, table.Name, requested, sheet.Name));
                tables.Add(table, sheet);
            }
        }
        return tables;
    }

    private static Dictionary<IWorkTableCell, string> BindExcelFormulas(IWorkNumbersProjection projection,
        Dictionary<IWorkTable, ExcelSheet> tables, CancellationToken cancellationToken, out string? limitation) {
        IReadOnlyDictionary<Guid, string> qualifiers = IWorkFormulaTableBindings.Create(tables.Keys,
            table => {
                string? name = IWorkFormulaReader.QuoteTableName(tables[table].Name, 8192);
                return name != null ? name + "!" : null;
            }, cancellationToken);
        var formulas = new Dictionary<IWorkTableCell, string>();
        limitation = null;
        foreach (IWorkTable table in tables.Keys) foreach (IWorkTableCell cell in table.Cells) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!cell.FormulaIsComplete || cell.Formula == null) continue;
            IWorkFormulaResult? result = cell.FormulaDefinition?.Render(qualifiers, projection.FormulaBudget!);
            string text = result?.Text ?? cell.Formula;
            if (result?.IsComplete == false || text.Length > 8192) {
                limitation = $"Numbers table '{table.Name}' contains a formula that cannot be bound within the XLSX formula limits.";
                return formulas;
            }
            formulas.Add(cell, text);
        }
        return formulas;
    }
}
