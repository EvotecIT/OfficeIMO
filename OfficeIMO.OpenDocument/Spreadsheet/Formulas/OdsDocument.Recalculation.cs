namespace OfficeIMO.OpenDocument;

public sealed partial class OdsDocument {
    /// <summary>Evaluates the supported OpenFormula subset and updates cached values within configured bounds.</summary>
    public OdsRecalculationReport Recalculate(OdsFormulaEvaluationOptions? options = null) {
        OdsFormulaEvaluationOptions effective = (options ?? new OdsFormulaEvaluationOptions()).Normalize();
        var report = new OdsRecalculationReport();
        var coordinates = new List<(string Sheet, long Row, long Column)>();
        foreach (OdsSheet sheet in Sheets) {
            foreach (OdsRowRun rowRun in sheet.RowRuns) {
                foreach (OdsCellRun cellRun in rowRun.CellRuns.Where(cell => cell.Formula != null && !cell.IsCovered)) {
                    for (long rowOffset = 0; rowOffset < rowRun.RepeatCount; rowOffset++) {
                        for (long columnOffset = 0; columnOffset < cellRun.RepeatCount; columnOffset++) {
                            if (coordinates.Count >= effective.MaximumFormulaCells) { report.Truncated = true; goto Collected; }
                            coordinates.Add((sheet.Name, checked(rowRun.StartRow + rowOffset), checked(cellRun.StartColumn + columnOffset)));
                        }
                    }
                }
            }
        }

    Collected:
        report.FormulaCells = coordinates.Count;
        var context = new OdsFormulaEvaluationContext(this, effective);
        foreach (var coordinate in coordinates) {
            OdsFormulaValue value = OdsFormulaEvaluator.EvaluateCell(context, coordinate.Sheet, coordinate.Row, coordinate.Column, 0);
            if (value.Kind == OdsFormulaValueKind.Error) {
                report.AddFailure(coordinate.Sheet, coordinate.Row, coordinate.Column, value.AsText());
                continue;
            }
            OdsCell cell = GetSheet(coordinate.Sheet)!.Cell(coordinate.Row, coordinate.Column);
            switch (value.Kind) {
                case OdsFormulaValueKind.Empty: cell.ClearValue(); break;
                case OdsFormulaValueKind.Number: cell.SetNumber(value.AsNumber()); break;
                case OdsFormulaValueKind.Boolean: cell.SetBoolean(value.AsBoolean()); break;
                default: cell.SetString(value.AsText()); break;
            }
            report.UpdatedCells++;
        }
        return report;
    }
}
