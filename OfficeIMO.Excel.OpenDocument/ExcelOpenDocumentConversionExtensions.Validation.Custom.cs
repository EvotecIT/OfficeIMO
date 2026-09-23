using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.Spreadsheet;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    private static bool TryCreateOdsCustomValidationCondition(
        ExcelDataValidationSnapshot validation, out OdsValidationConditionSyntax? condition) {
        condition = null;
        if (validation.A1Ranges.Count != 1 || !string.IsNullOrWhiteSpace(validation.Formula2)
            || !string.IsNullOrWhiteSpace(validation.Operator)
            || string.IsNullOrWhiteSpace(validation.Formula1)) return false;

        string formula = validation.Formula1!.Trim();
        if (!formula.StartsWith("=", StringComparison.Ordinal)) formula = "=" + formula;
        SpreadsheetFormulaSyntaxTree syntax = SpreadsheetFormulaSyntaxTree.Parse(
            formula, SpreadsheetFormulaDialect.ExcelA1);
        if (!OdsPortableValidationFormula.IsSupported(syntax)) return false;
        SpreadsheetFormulaTranslationResult translated = syntax.TranslateTo(SpreadsheetFormulaDialect.OpenFormula);
        if (!translated.IsSuccessful || !translated.Formula.StartsWith("of:=", StringComparison.Ordinal)) return false;
        condition = OdsValidationConditionSyntax.CreateFormula(translated.Formula.Substring(4));
        return true;
    }

    private static bool TryFormatExcelValidationBaseCell(string rangeText, string sheetName,
        out string? baseCellAddress) {
        baseCellAddress = null;
        if (!SpreadsheetRangeReference.TryParse(rangeText, SpreadsheetAddressDialect.ExcelA1,
                out SpreadsheetRangeReference? range)
            || !range!.Start.IsCell || range.Start.SheetName != null
            || range.End != null && (!range.End.IsCell || range.End.SheetName != null
                || range.End.Row < range.Start.Row || range.End.Column < range.Start.Column)) return false;

        string a1 = SpreadsheetAddressConverter.ToA1(checked((int)range.Start.Row!.Value),
            range.Start.Column!.Value);
        int rowOffset = 0;
        while (rowOffset < a1.Length && !char.IsDigit(a1[rowOffset])) rowOffset++;
        string escapedSheet = sheetName.Replace("'", "''");
        baseCellAddress = "$'" + escapedSheet + "'.$" + a1.Substring(0, rowOffset) + "$"
            + a1.Substring(rowOffset);
        return true;
    }

    private static bool TryApplyOdsCustomValidation(ExcelSheet sheet, string sourceSheetName,
        string references, OdsValidation validation, OdsValidationConditionSyntax condition) {
        if (string.IsNullOrWhiteSpace(condition.FirstOperand)
            || !SpreadsheetRangeReference.TryParse(validation.BaseCellAddress,
                SpreadsheetAddressDialect.OpenDocument, out SpreadsheetRangeReference? baseCell)
            || baseCell!.End != null || !baseCell.Start.IsCell
            || !string.Equals(baseCell.Start.SheetName, sourceSheetName, StringComparison.Ordinal)
            || !TryGetRectangularValidationRange(references, baseCell.Start,
                out string? range)) return false;

        SpreadsheetFormulaSyntaxTree syntax = SpreadsheetFormulaSyntaxTree.Parse(
            "of:=" + condition.FirstOperand, SpreadsheetFormulaDialect.OpenFormula);
        if (!OdsPortableValidationFormula.IsSupported(syntax)) return false;
        SpreadsheetFormulaTranslationResult translated = syntax.TranslateTo(SpreadsheetFormulaDialect.ExcelA1);
        if (!translated.IsSuccessful || !translated.Formula.StartsWith("=", StringComparison.Ordinal)) return false;
        string formula = translated.Formula.Substring(1);
        if (formula.Length > ExcelSheet.MaximumDataValidationFormulaLength) return false;
        sheet.ValidationCustomFormula(range!, formula, validation.AllowEmptyCell);
        return true;
    }

    private static bool TryGetRectangularValidationRange(string references, SpreadsheetCellReference baseCell,
        out string? range) {
        range = null;
        string[] addresses = references.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (addresses.Length == 0) return false;
        var cells = new HashSet<(long Row, int Column)>();
        long minRow = long.MaxValue, maxRow = 0;
        int minColumn = int.MaxValue, maxColumn = 0;
        foreach (string address in addresses) {
            if (!SpreadsheetRangeReference.TryParse(address, SpreadsheetAddressDialect.ExcelA1,
                    out SpreadsheetRangeReference? parsed)
                || parsed!.End != null || !parsed.Start.IsCell || parsed.Start.SheetName != null) return false;
            long row = parsed.Start.Row!.Value;
            int column = parsed.Start.Column!.Value;
            if (!cells.Add((row, column))) return false;
            minRow = Math.Min(minRow, row);
            maxRow = Math.Max(maxRow, row);
            minColumn = Math.Min(minColumn, column);
            maxColumn = Math.Max(maxColumn, column);
        }
        if (baseCell.Row != minRow || baseCell.Column != minColumn
            || (maxRow - minRow + 1L) * (maxColumn - minColumn + 1L) != cells.Count) return false;
        string first = SpreadsheetAddressConverter.ToA1(checked((int)minRow), minColumn);
        string last = SpreadsheetAddressConverter.ToA1(checked((int)maxRow), maxColumn);
        range = first == last ? first : first + ":" + last;
        return true;
    }

}
