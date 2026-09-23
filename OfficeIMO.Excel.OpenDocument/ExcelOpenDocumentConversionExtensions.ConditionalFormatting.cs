using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    private const int MaximumConditionalFormattingCellsPerStyle = 4096;

    private sealed class OdsConditionalStylePlan {
        internal OdsConditionalStylePlan(ExcelConditionalFormattingOperator comparison, string threshold, string fillColor) {
            Comparison = comparison;
            Threshold = threshold;
            FillColor = fillColor;
        }

        internal ExcelConditionalFormattingOperator Comparison { get; }
        internal string Threshold { get; }
        internal string FillColor { get; }
    }

    private static void CollectOdsConditionalTarget(OdsDocument source, string styleName,
        int row, int column, Dictionary<string, OdsConditionalStylePlan?> plans,
        Dictionary<string, List<string>> targets, HashSet<string> limits) {
        if (!plans.TryGetValue(styleName, out OdsConditionalStylePlan? plan)) {
            plan = CreateOdsConditionalStylePlan(source, styleName);
            plans.Add(styleName, plan);
        }
        if (plan == null || limits.Contains(styleName)) return;
        if (!targets.TryGetValue(styleName, out List<string>? references)) {
            references = new List<string>();
            targets.Add(styleName, references);
        }
        if (references.Count >= MaximumConditionalFormattingCellsPerStyle) {
            references.Clear();
            limits.Add(styleName);
            return;
        }
        references.Add(SpreadsheetAddressConverter.ToA1(row, column));
    }

    private static OdsConditionalStylePlan? CreateOdsConditionalStylePlan(OdsDocument source, string styleName) {
        OdfStyle? baseStyle = source.Styles.Find(OdfStyleFamily.TableCell, styleName);
        if (baseStyle == null) return null;
        IReadOnlyList<OdfStyleMap> maps = baseStyle.ConditionalMaps;
        if (maps.Count != 1 || !TryParseNumericCellCondition(maps[0].Condition,
                out ExcelConditionalFormattingOperator comparison, out string threshold)) return null;
        OdfStyle[] appliedStyles = source.Styles.Named.Where(style =>
            style.Family == OdfStyleFamily.TableCell &&
            string.Equals(style.Name, maps[0].ApplyStyleName, StringComparison.Ordinal))
            .Take(2).ToArray();
        if (appliedStyles.Length != 1) return null;
        OdfColor? fill;
        try {
            fill = appliedStyles[0].BackgroundColor;
        } catch (FormatException) {
            return null;
        }
        return fill.HasValue
            ? new OdsConditionalStylePlan(comparison, threshold, "FF" + fill.Value.ToString().Substring(1))
            : null;
    }

    private static bool TryParseNumericCellCondition(string condition,
        out ExcelConditionalFormattingOperator comparison, out string threshold) {
        comparison = default;
        threshold = string.Empty;
        const string function = "cell-content()";
        if (!condition.StartsWith(function, StringComparison.Ordinal)) return false;
        string remainder = condition.Substring(function.Length).TrimStart();
        string comparisonToken;
        if (remainder.StartsWith("<=", StringComparison.Ordinal)) {
            comparison = ExcelConditionalFormattingOperator.LessThanOrEqual;
            comparisonToken = "<=";
        } else if (remainder.StartsWith(">=", StringComparison.Ordinal)) {
            comparison = ExcelConditionalFormattingOperator.GreaterThanOrEqual;
            comparisonToken = ">=";
        } else if (remainder.StartsWith("!=", StringComparison.Ordinal)) {
            comparison = ExcelConditionalFormattingOperator.NotEqual;
            comparisonToken = "!=";
        } else if (remainder.StartsWith("<>", StringComparison.Ordinal)) {
            comparison = ExcelConditionalFormattingOperator.NotEqual;
            comparisonToken = "<>";
        } else if (remainder.StartsWith("<", StringComparison.Ordinal)) {
            comparison = ExcelConditionalFormattingOperator.LessThan;
            comparisonToken = "<";
        } else if (remainder.StartsWith(">", StringComparison.Ordinal)) {
            comparison = ExcelConditionalFormattingOperator.GreaterThan;
            comparisonToken = ">";
        } else if (remainder.StartsWith("=", StringComparison.Ordinal)) {
            comparison = ExcelConditionalFormattingOperator.Equal;
            comparisonToken = "=";
        } else return false;
        string operand = remainder.Substring(comparisonToken.Length).Trim();
        if (!decimal.TryParse(operand, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture, out decimal number)) return false;
        threshold = number.ToString(CultureInfo.InvariantCulture);
        return true;
    }

    private static void ApplyOdsConditionalStyles(ExcelSheet sheet,
        Dictionary<string, OdsConditionalStylePlan?> plans,
        Dictionary<string, List<string>> targets, HashSet<string> limits,
        HashSet<string> convertedStyles) {
        foreach (KeyValuePair<string, List<string>> entry in targets) {
            if (limits.Contains(entry.Key) || entry.Value.Count == 0) continue;
            OdsConditionalStylePlan plan = plans[entry.Key]!;
            sheet.AddConditionalRule(string.Join(" ", entry.Value), plan.Comparison,
                plan.Threshold, formula2: null, fillColor: plan.FillColor);
            convertedStyles.Add(entry.Key);
        }
    }
}
