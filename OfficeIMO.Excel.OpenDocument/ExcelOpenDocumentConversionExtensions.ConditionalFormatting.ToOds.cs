using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    private const int MaximumExcelConditionalRulesPerSheet = 16;
    private const int MaximumExcelConditionalCellsPerSheet = 4096;

    private static int ApplyExcelConditionalFormatting(ExcelSheet sourceSheet,
        ExcelWorksheetSnapshot worksheet, OdsDocument document, OdsSheet targetSheet,
        ExcelOpenDocumentConversionOptions options, HashSet<(int Row, int Column)> materializedCoordinates,
        ref long materializedCells, ref bool truncated) {
        int ruleCount = worksheet.ConditionalFormattingRuleCount;
        if (ruleCount == 0 || ruleCount > MaximumExcelConditionalRulesPerSheet) return 0;
        IReadOnlyList<ExcelConditionalFormattingInfo> rules = sourceSheet.GetConditionalFormattingRules();
        if (rules.Count != ruleCount || !TryGetExcelConditionalRange(rules, worksheet, options,
                out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) return 0;

        long cellCount = (long)(lastRow - firstRow + 1) * (lastColumn - firstColumn + 1);
        if (cellCount > MaximumExcelConditionalCellsPerSheet) return 0;
        var ordered = rules.OrderBy(rule => rule.Priority).ToArray();
        var conditions = new List<(string Condition, OdfColor Fill)>(ordered.Length);
        for (int index = 0; index < ordered.Length; index++) {
            ExcelConditionalFormattingInfo rule = ordered[index];
            if (rule.Source != ExcelConditionalFormattingSource.Standard
                || !string.Equals(rule.Type, "CellIs", StringComparison.OrdinalIgnoreCase)
                || rule.Priority <= 0 || index > 0 && ordered[index - 1].Priority == rule.Priority
                || index < ordered.Length - 1 && !rule.StopIfTrue
                || rule.HasPreservedUnknownMarkup || !rule.IsSolidFillOnlyDifferentialStyle
                || !TryParseOpaqueArgb(rule.DifferentialFillColorArgb, out OdfColor fill)
                || !TryGetOdfCellCondition(rule, out string? condition)) return 0;
            conditions.Add((condition!, fill));
        }

        long newCells = 0;
        for (int row = firstRow; row <= lastRow; row++) {
            for (int column = firstColumn; column <= lastColumn; column++) {
                if (materializedCoordinates.Contains((row, column))) {
                    if (targetSheet.Cell(row - 1L, column - 1L).StyleName != null) return 0;
                } else {
                    newCells++;
                }
            }
        }
        if (newCells > options.MaximumExpandedCells - materializedCells) {
            truncated = true;
            return 0;
        }

        OdfStyle baseStyle = document.Styles.CreateAutomatic(OdfStyleFamily.TableCell, "xlCf");
        foreach ((string condition, OdfColor fill) in conditions) {
            string styleName = "xlConditional_" + (worksheet.Index + 1).ToString(CultureInfo.InvariantCulture)
                + "_" + (baseStyle.ConditionalMaps.Count + 1).ToString(CultureInfo.InvariantCulture);
            OdfStyle appliedStyle = document.Styles.CreateNamed(styleName, OdfStyleFamily.TableCell);
            appliedStyle.BackgroundColor = fill;
            baseStyle.AddConditionalMap(condition, appliedStyle.Name);
        }
        for (int row = firstRow; row <= lastRow; row++) {
            for (int column = firstColumn; column <= lastColumn; column++) {
                if (materializedCoordinates.Add((row, column))) materializedCells++;
                targetSheet.Cell(row - 1L, column - 1L).StyleName = baseStyle.Name;
            }
        }
        return ruleCount;
    }

    private static bool TryGetExcelConditionalRange(IReadOnlyList<ExcelConditionalFormattingInfo> rules,
        ExcelWorksheetSnapshot worksheet, ExcelOpenDocumentConversionOptions options,
        out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) {
        firstRow = firstColumn = lastRow = lastColumn = 0;
        string rangeText = rules[0].Range;
        if (!SpreadsheetRangeReference.TryParse(rangeText, SpreadsheetAddressDialect.ExcelA1,
                out SpreadsheetRangeReference? parsed) || !parsed!.Start.IsCell || parsed.Start.SheetName != null
            || parsed.End != null && (!parsed.End.IsCell || parsed.End.SheetName != null)) return false;
        SpreadsheetCellReference end = parsed.End ?? parsed.Start;
        if (parsed.Start.Row > end.Row || parsed.Start.Column > end.Column
            || end.Row > options.MaximumRows || end.Column > options.MaximumColumns) return false;
        firstRow = checked((int)parsed.Start.Row!.Value);
        firstColumn = parsed.Start.Column!.Value;
        lastRow = checked((int)end.Row!.Value);
        lastColumn = end.Column!.Value;
        if (rules.Any(rule => !string.Equals(rule.Range, rangeText, StringComparison.Ordinal))) return false;
        foreach (ExcelMergedRangeSnapshot merge in worksheet.MergedRanges) {
            if (merge.StartRow <= lastRow && merge.EndRow >= firstRow
                && merge.StartColumn <= lastColumn && merge.EndColumn >= firstColumn) return false;
        }
        return true;
    }

    private static bool TryParseOpaqueArgb(string? argb, out OdfColor color) {
        color = default;
        return argb != null && argb.Length == 8 && argb.StartsWith("FF", StringComparison.OrdinalIgnoreCase)
            && OdfColor.TryParse(argb.Substring(2), out color);
    }

    private static bool TryGetOdfCellCondition(ExcelConditionalFormattingInfo rule, out string? condition) {
        condition = null;
        string? operatorText = rule.Operator;
        bool range = string.Equals(operatorText, "Between", StringComparison.OrdinalIgnoreCase)
            || string.Equals(operatorText, "NotBetween", StringComparison.OrdinalIgnoreCase);
        if (rule.Formulas.Count != (range ? 2 : 1)
            || !TryFormatNumericConditionBound(rule.Formulas[0], out string? first)) return false;
        string? second = null;
        if (range && (!TryFormatNumericConditionBound(rule.Formulas[1], out second)
            || decimal.Parse(first!, CultureInfo.InvariantCulture) > decimal.Parse(second!, CultureInfo.InvariantCulture))) return false;
        condition = operatorText?.ToLowerInvariant() switch {
            "greaterthan" => "cell-content()>" + first,
            "greaterthanorequal" => "cell-content()>=" + first,
            "lessthan" => "cell-content()<" + first,
            "lessthanorequal" => "cell-content()<=" + first,
            "equal" => "cell-content()=" + first,
            "notequal" => "cell-content()!=" + first,
            "between" => "cell-content-is-between(" + first + "," + second + ")",
            "notbetween" => "cell-content-is-not-between(" + first + "," + second + ")",
            _ => null
        };
        return condition != null;
    }

    private static bool TryFormatNumericConditionBound(string? text, out string? normalized) {
        normalized = null;
        if (!decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal value)) return false;
        normalized = value.ToString(CultureInfo.InvariantCulture);
        return true;
    }
}
