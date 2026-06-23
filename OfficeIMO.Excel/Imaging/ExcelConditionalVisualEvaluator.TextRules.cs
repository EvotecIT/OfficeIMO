using System;
using System.Collections.Generic;

namespace OfficeIMO.Excel {
    internal static partial class ExcelConditionalVisualEvaluator {
        private static bool IsTextRule(ExcelConditionalFormattingInfo rule) =>
            string.Equals(rule.Type, "ContainsText", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(rule.Type, "NotContainsText", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(rule.Type, "BeginsWith", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(rule.Type, "EndsWith", StringComparison.OrdinalIgnoreCase);

        private static bool CanEvaluateTextRule(ExcelConditionalFormattingInfo rule) =>
            !string.IsNullOrEmpty(rule.Text);

        private static void ApplyTextRuleFill(
            IReadOnlyList<ExcelVisualCell> cells,
            ExcelConditionalFormattingInfo rule,
            Dictionary<string, string> fills,
            HashSet<string> stoppedCells) {
            if (string.IsNullOrWhiteSpace(rule.DifferentialFillColorArgb) || !CanEvaluateTextRule(rule)) {
                return;
            }

            foreach (ExcelVisualCell cell in GetRuleCells(cells, rule.Range)) {
                string key = Key(cell.Row, cell.Column);
                if (stoppedCells.Contains(key) || !TextRuleMatches(cell.Text, rule)) {
                    continue;
                }

                if (!fills.ContainsKey(key)) {
                    fills[key] = rule.DifferentialFillColorArgb!;
                }

                if (rule.StopIfTrue) {
                    stoppedCells.Add(key);
                }
            }
        }

        private static bool TextRuleMatches(string? value, ExcelConditionalFormattingInfo rule) {
            string cellText = value ?? string.Empty;
            string text = rule.Text ?? string.Empty;
            if (string.Equals(rule.Type, "ContainsText", StringComparison.OrdinalIgnoreCase)) {
                return cellText.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0;
            }

            if (string.Equals(rule.Type, "NotContainsText", StringComparison.OrdinalIgnoreCase)) {
                return cellText.IndexOf(text, StringComparison.OrdinalIgnoreCase) < 0;
            }

            if (string.Equals(rule.Type, "BeginsWith", StringComparison.OrdinalIgnoreCase)) {
                return cellText.StartsWith(text, StringComparison.OrdinalIgnoreCase);
            }

            return string.Equals(rule.Type, "EndsWith", StringComparison.OrdinalIgnoreCase) &&
                cellText.EndsWith(text, StringComparison.OrdinalIgnoreCase);
        }
    }
}
