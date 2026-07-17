using System;
using System.Collections.Generic;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private List<FormulaDependencyReferenceMatch> CombineFormulaRangeDependencyMatches(
            string formula,
            IReadOnlyList<FormulaDependencyReferenceMatch> matches,
            IReadOnlyList<int> groupingDepths,
            int? sourceRow) {
            var combinedMatches = new List<FormulaDependencyReferenceMatch>(matches.Count);
            int index = 0;
            while (index < matches.Count) {
                FormulaDependencyReferenceMatch current = matches[index++];
                while (index < matches.Count
                    && IsFormulaRangeSeparator(
                        formula,
                        current,
                        matches[index],
                        groupingDepths[current.Index + current.Length])
                    && TryCombineFormulaRangeDependencyMatch(
                        current,
                        matches[index],
                        sourceRow,
                        out FormulaDependencyReferenceMatch combined)) {
                    current = combined;
                    index++;
                }

                combinedMatches.Add(current);
            }

            return combinedMatches;
        }

        private static bool IsFormulaRangeSeparator(
            string formula,
            FormulaDependencyReferenceMatch left,
            FormulaDependencyReferenceMatch right,
            int availableClosingParentheses) {
            int start = left.Index + left.Length;
            if (start >= right.Index) {
                return false;
            }

            int cursor = start;
            for (int consumed = 0; consumed < availableClosingParentheses; consumed++) {
                while (cursor < right.Index && char.IsWhiteSpace(formula[cursor])) {
                    cursor++;
                }

                if (cursor >= right.Index || formula[cursor] != ')') {
                    break;
                }

                cursor++;
            }

            while (cursor < right.Index && char.IsWhiteSpace(formula[cursor])) {
                cursor++;
            }
            if (cursor >= right.Index || formula[cursor] != ':') {
                return false;
            }

            cursor++;
            while (cursor < right.Index && char.IsWhiteSpace(formula[cursor])) {
                cursor++;
            }
            for (; cursor < right.Index; cursor++) {
                if (formula[cursor] != '(' && !char.IsWhiteSpace(formula[cursor])) {
                    return false;
                }
            }

            return true;
        }

        private bool TryCombineFormulaRangeDependencyMatch(
            FormulaDependencyReferenceMatch left,
            FormulaDependencyReferenceMatch right,
            int? sourceRow,
            out FormulaDependencyReferenceMatch combined) {
            combined = default;
            if (!TryResolveFormulaDependencyReference(
                    left.Reference,
                    sourceRow,
                    out ExcelSheet leftSheet,
                    out int leftR1,
                    out int leftC1,
                    out int leftR2,
                    out int leftC2,
                    out _)
                || !TryResolveFormulaDependencyReference(
                    right.Reference,
                    sourceRow,
                    out ExcelSheet rightSheet,
                    out int rightR1,
                    out int rightC1,
                    out int rightR2,
                    out int rightC2,
                    out _)
                || !string.Equals(leftSheet.Name, rightSheet.Name, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            int firstRow = Math.Min(leftR1, rightR1);
            int firstColumn = Math.Min(leftC1, rightC1);
            int lastRow = Math.Max(leftR2, rightR2);
            int lastColumn = Math.Max(leftC2, rightC2);
            string start = A1.CellReference(firstRow, firstColumn);
            string end = A1.CellReference(lastRow, lastColumn);
            string sheetQualifier = "'" + leftSheet.Name.Replace("'", "''") + "'!";
            string reference = sheetQualifier + (start == end ? start : start + ":" + end);
            combined = new FormulaDependencyReferenceMatch(
                left.Index,
                right.Index + right.Length - left.Index,
                reference);
            return true;
        }
    }
}
