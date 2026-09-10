using AngleSharp.Dom;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    /// <summary>Imports one bounded value and tells formatting whether it must preserve that value.</summary>
    private static bool SetCellValue(
        ExcelSheet sheet,
        int row,
        int column,
        IElement cell,
        string fallbackText,
        HtmlToExcelResult result,
        HtmlToExcelOptions options,
        HtmlImportBudget budget,
        HashSet<long>? importedFormulaCells,
        bool useSemanticValues,
        out bool preserveValue) {
        preserveValue = false;
        string? kind = cell.GetAttribute("data-officeimo-value-kind");
        string? rawValue = cell.GetAttribute("data-officeimo-value");
        if (!useSemanticValues && options.ImportTypedCellValues && (kind != null || rawValue != null)) {
            if (!IsScalarCellValueKind(kind) || rawValue == null) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
                    "Cell " + BuildCellReference(row, column) + " did not declare a supported scalar value and was imported from visible text. Generic imports accept text, number, boolean, and date-time values; formulas require the semantic workbook path.",
                    lossKind: OfficeConversionLossKind.Approximation);
                return TrySetCellTextValue(sheet, row, column, fallbackText, result, budget);
            }
        } else if (!useSemanticValues || string.IsNullOrWhiteSpace(kind) || rawValue == null) {
            return TrySetCellTextValue(sheet, row, column, fallbackText, result, budget);
        }

        bool isFormula = kind!.Equals("formula", StringComparison.OrdinalIgnoreCase);
        if (isFormula && options.ImportFormulas) {
            // Remember even a budget-rejected formula so the compatibility inventory cannot
            // make a second, conflicting decision for the same cell.
            importedFormulaCells?.Add(GetImportCellKey(row, column));
        }

        if (!budget.IsMetadataWithinLimit(rawValue!, out string metadataLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                "Cell " + BuildCellReference(row, column) + " semantic value exceeded the shared metadata limit and was imported from visible text.",
                lossKind: OfficeConversionLossKind.Approximation, detail: metadataLimit);
            return TrySetCellTextValue(sheet, row, column, fallbackText, result, budget);
        }

        if (kind.Equals("number", StringComparison.OrdinalIgnoreCase)) {
            if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                && !double.IsNaN(number) && !double.IsInfinity(number)) {
                sheet.CellValue(row, column, number);
                preserveValue = true;
                return true;
            }
            AddInvalidCellValueDiagnostic(row, column, "number", result);
        } else if (kind.Equals("boolean", StringComparison.OrdinalIgnoreCase)) {
            if (rawValue!.Equals("1", StringComparison.OrdinalIgnoreCase) || rawValue.Equals("true", StringComparison.OrdinalIgnoreCase)) {
                sheet.CellValue(row, column, true);
                preserveValue = true;
                return true;
            }
            if (rawValue.Equals("0", StringComparison.OrdinalIgnoreCase) || rawValue.Equals("false", StringComparison.OrdinalIgnoreCase)) {
                sheet.CellValue(row, column, false);
                preserveValue = true;
                return true;
            }
            AddInvalidCellValueDiagnostic(row, column, "boolean", result);
        } else if (kind.Equals("text", StringComparison.OrdinalIgnoreCase)) {
            bool stored = TrySetCellTextValue(sheet, row, column, rawValue!, result, budget);
            preserveValue = stored;
            return stored;
        } else if (kind.Equals("date-time", StringComparison.OrdinalIgnoreCase)) {
            // Excel stores no time zone. Normalize explicit offsets to UTC, while
            // retaining the authored wall clock when metadata has no zone.
            if (DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal, out DateTime dateTime)
                && dateTime.Year >= 100) {
                sheet.CellValue(row, column, dateTime);
                preserveValue = true;
                return true;
            }
            AddInvalidCellValueDiagnostic(row, column, "date/time", result);
        } else if (isFormula) {
            if (!options.ImportFormulas) {
                return TrySetCellTextValue(sheet, row, column, fallbackText, result, budget);
            }

            string formulaLimit = string.Empty;
            string annotationLimit = string.Empty;
            if (!IsWithinExcelFieldLimit(rawValue!, budget, ExcelFormulaCharacterLimit, "ExcelFormulaCharacterLimit", out formulaLimit)
                || !budget.TryReserveAnnotation(out annotationLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                    "Cell " + BuildCellReference(row, column) + " formula was omitted because a semantic or native formula limit was reached.",
                    lossKind: OfficeConversionLossKind.Omission, detail: formulaLimit.Length > 0 ? formulaLimit : annotationLimit);
                return TrySetCellTextValue(sheet, row, column, fallbackText, result, budget);
            }

            sheet.CellFormula(row, column, rawValue!);
            result.Formulas++;
            preserveValue = true;
            return true;
        } else if (kind.Equals("error", StringComparison.OrdinalIgnoreCase)) {
            sheet.CellError(row, column, rawValue!);
            preserveValue = true;
            return true;
        }

        return TrySetCellTextValue(sheet, row, column, fallbackText, result, budget);
    }

    private static bool IsScalarCellValueKind(string? kind) =>
        string.Equals(kind, "text", StringComparison.OrdinalIgnoreCase)
        || string.Equals(kind, "number", StringComparison.OrdinalIgnoreCase)
        || string.Equals(kind, "boolean", StringComparison.OrdinalIgnoreCase)
        || string.Equals(kind, "date-time", StringComparison.OrdinalIgnoreCase);

    private static void AddInvalidCellValueDiagnostic(int row, int column, string kind, HtmlToExcelResult result) =>
        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
            "Cell " + BuildCellReference(row, column) + " contained a semantic " + kind + " value that could not be parsed and was imported as text.",
            lossKind: OfficeConversionLossKind.Approximation);

    /// <summary>Applies the first visible run's cell-wide style without replacing a scalar with rich text.</summary>
    private static void ApplyScalarCellFormatting(ExcelCell cell, IReadOnlyList<HtmlSemanticRun> runs) {
        HtmlSemanticRun? source = runs.FirstOrDefault(run => !string.IsNullOrWhiteSpace(run.Text));
        if (source == null) return;
        ExcelRichTextRun run = ToExcelRun(source);
        if (run.BoldSpecified) cell.SetBold(run.Bold);
        if (run.ItalicSpecified) cell.SetItalic(run.Italic);
        if (run.UnderlineSpecified) {
            if (run.UnderlineStyle.HasValue) cell.SetUnderline(run.UnderlineStyle.Value);
            else cell.SetUnderline(run.Underline);
        }
        if (run.StrikethroughSpecified) cell.SetStrikethrough(run.Strikethrough);
        if (run.VerticalTextAlignment.HasValue) cell.SetVerticalTextAlignment(run.VerticalTextAlignment.Value);
        if (!string.IsNullOrWhiteSpace(run.FontName)) cell.SetFontName(run.FontName!);
        if (run.FontSize.HasValue) cell.SetFontSize(run.FontSize.Value);
        if (!string.IsNullOrWhiteSpace(run.FontColor)) cell.SetFontColor(run.FontColor!);
    }
}
