using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    private const int MaximumConditionalFormattingCellsPerStyle = 4096;

    private sealed class OdsConditionalStylePlan {
        internal OdsConditionalStylePlan(ExcelConditionalFormattingOperator comparison, string formula1,
            string? formula2, string? fillColor, string? fontColor,
            bool? bold, bool? italic, bool? underline) {
            Comparison = comparison;
            Formula1 = formula1;
            Formula2 = formula2;
            FillColor = fillColor;
            FontColor = fontColor;
            Bold = bold;
            Italic = italic;
            Underline = underline;
        }

        internal ExcelConditionalFormattingOperator Comparison { get; }
        internal string Formula1 { get; }
        internal string? Formula2 { get; }
        internal string? FillColor { get; }
        internal string? FontColor { get; }
        internal bool? Bold { get; }
        internal bool? Italic { get; }
        internal bool? Underline { get; }
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
                out ExcelConditionalFormattingOperator comparison, out string formula1,
                out string? formula2)) return null;
        OdfStyle[] appliedStyles = source.Styles.Named.Where(style =>
            style.Family == OdfStyleFamily.TableCell &&
            string.Equals(style.Name, maps[0].ApplyStyleName, StringComparison.Ordinal))
            .Take(2).ToArray();
        if (appliedStyles.Length != 1) return null;
        OdfColor? fill;
        OdfColor? fontColor;
        try {
            fill = appliedStyles[0].BackgroundColor;
            fontColor = appliedStyles[0].Color;
        } catch (FormatException) {
            return null;
        }
        bool? bold = appliedStyles[0].Bold == true ? true : (bool?)null;
        bool? italic = appliedStyles[0].Italic == true ? true : (bool?)null;
        bool? underline = appliedStyles[0].Underline == true ? true : (bool?)null;
        if (!fill.HasValue && !fontColor.HasValue && !bold.HasValue && !italic.HasValue && !underline.HasValue)
            return null;
        return new OdsConditionalStylePlan(comparison, formula1, formula2,
            fill.HasValue ? "FF" + fill.Value.ToString().Substring(1) : null,
            fontColor.HasValue ? "FF" + fontColor.Value.ToString().Substring(1) : null,
            bold, italic, underline);
    }

    private static bool TryParseNumericCellCondition(string condition,
        out ExcelConditionalFormattingOperator comparison, out string formula1, out string? formula2) {
        comparison = default;
        formula1 = string.Empty;
        formula2 = null;
        condition = condition.Trim();
        if (TryParseNumericRangeCondition(condition, out comparison, out formula1, out formula2)) {
            return true;
        }
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
        formula1 = number.ToString(CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryParseNumericRangeCondition(string condition,
        out ExcelConditionalFormattingOperator comparison, out string formula1, out string? formula2) {
        comparison = default;
        formula1 = string.Empty;
        formula2 = null;
        const string between = "cell-content-is-between(";
        const string notBetween = "cell-content-is-not-between(";
        string arguments;
        if (condition.StartsWith(between, StringComparison.Ordinal)) {
            comparison = ExcelConditionalFormattingOperator.Between;
            arguments = condition.Substring(between.Length);
        } else if (condition.StartsWith(notBetween, StringComparison.Ordinal)) {
            comparison = ExcelConditionalFormattingOperator.NotBetween;
            arguments = condition.Substring(notBetween.Length);
        } else {
            return false;
        }

        if (!arguments.EndsWith(")", StringComparison.Ordinal)) return false;
        arguments = arguments.Substring(0, arguments.Length - 1);
        int separator = arguments.IndexOf(',');
        if (separator < 0 || arguments.IndexOf(',', separator + 1) >= 0) return false;
        if (!decimal.TryParse(arguments.Substring(0, separator).Trim(),
                NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture, out decimal lower)
            || !decimal.TryParse(arguments.Substring(separator + 1).Trim(),
                NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture, out decimal upper)
            || lower > upper) return false;
        formula1 = lower.ToString(CultureInfo.InvariantCulture);
        formula2 = upper.ToString(CultureInfo.InvariantCulture);
        return true;
    }

    private static void ApplyOdsConditionalStyles(ExcelSheet sheet,
        Dictionary<string, OdsConditionalStylePlan?> plans,
        Dictionary<string, List<string>> targets, HashSet<string> limits,
        HashSet<string> convertedStyles) {
        foreach (KeyValuePair<string, List<string>> entry in targets) {
            if (limits.Contains(entry.Key) || entry.Value.Count == 0) continue;
            OdsConditionalStylePlan plan = plans[entry.Key]!;
            sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                Source = ExcelConditionalFormattingSource.Standard,
                Range = string.Join(" ", entry.Value),
                Type = "CellIs",
                Operator = plan.Comparison.ToString(),
                Formulas = plan.Formula2 == null
                    ? new[] { plan.Formula1 }
                    : new[] { plan.Formula1, plan.Formula2 },
                DifferentialFillColorArgb = plan.FillColor,
                DifferentialFontColorArgb = plan.FontColor,
                DifferentialFontBold = plan.Bold,
                DifferentialFontItalic = plan.Italic,
                DifferentialFontUnderline = plan.Underline
            });
            convertedStyles.Add(entry.Key);
        }
    }
}
