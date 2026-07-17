using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private readonly struct FormulaDependencyReferenceMatch {
            internal FormulaDependencyReferenceMatch(int index, int length, string reference) {
                Index = index;
                Length = length;
                Reference = reference;
            }

            internal int Index { get; }
            internal int Length { get; }
            internal string Reference { get; }
        }

        private void AddFormulaDependencies(
            string formula,
            IEnumerable<FormulaDependencyReferenceMatch> matches,
            int? sourceRow,
            ISet<string> dependencies) {
            List<FormulaDependencyReferenceMatch> orderedMatches = GetNonOverlappingFormulaDependencyMatches(matches);
            int[] groupingDepths = BuildOpenGroupingParenthesisDepths(formula);
            int index = 0;
            while (index < orderedMatches.Count) {
                int intersectionEnd = index;
                while (intersectionEnd + 1 < orderedMatches.Count
                    && IsFormulaIntersectionSeparator(
                        formula,
                        orderedMatches[intersectionEnd],
                        orderedMatches[intersectionEnd + 1],
                        groupingDepths[orderedMatches[intersectionEnd].Index + orderedMatches[intersectionEnd].Length])) {
                    intersectionEnd++;
                }

                if (intersectionEnd == index) {
                    dependencies.Add(NormalizeFormulaDependencyReference(orderedMatches[index].Reference, sourceRow));
                } else {
                    AddFormulaIntersectionDependency(orderedMatches, index, intersectionEnd, sourceRow, dependencies);
                }

                index = intersectionEnd + 1;
            }
        }

        private static List<FormulaDependencyReferenceMatch> GetNonOverlappingFormulaDependencyMatches(
            IEnumerable<FormulaDependencyReferenceMatch> matches) {
            var result = new List<FormulaDependencyReferenceMatch>();
            int consumedUntil = -1;
            foreach (FormulaDependencyReferenceMatch match in matches
                .OrderBy(match => match.Index)
                .ThenByDescending(match => match.Length)) {
                if (match.Index < consumedUntil) {
                    continue;
                }

                result.Add(match);
                consumedUntil = match.Index + match.Length;
            }

            return result;
        }

        private static bool IsFormulaIntersectionSeparator(
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
                int whitespaceStart = cursor;
                while (cursor < right.Index && char.IsWhiteSpace(formula[cursor])) {
                    cursor++;
                }

                if (cursor >= right.Index || formula[cursor] != ')') {
                    cursor = whitespaceStart;
                    break;
                }

                cursor++;
            }

            int separatorStart = cursor;
            while (cursor < right.Index && char.IsWhiteSpace(formula[cursor])) {
                cursor++;
            }

            if (cursor == separatorStart) {
                return false;
            }

            for (; cursor < right.Index; cursor++) {
                if (formula[cursor] != '(' && !char.IsWhiteSpace(formula[cursor])) {
                    return false;
                }
            }

            return true;
        }

        private static int[] BuildOpenGroupingParenthesisDepths(string formula) {
            var depths = new int[formula.Length + 1];
            var parentheses = new Stack<bool>();
            int groupingDepth = 0;
            bool inString = false;
            bool inQuotedQualifier = false;
            int structuredReferenceDepth = 0;
            for (int index = 0; index < formula.Length; index++) {
                char character = formula[index];
                if (inString) {
                    if (character == '"') {
                        if (index + 1 < formula.Length && formula[index + 1] == '"') {
                            depths[index + 1] = groupingDepth;
                            index++;
                        } else {
                            inString = false;
                        }
                    }
                    depths[index + 1] = groupingDepth;
                    continue;
                }

                if (inQuotedQualifier) {
                    if (character == '\'') {
                        if (index + 1 < formula.Length && formula[index + 1] == '\'') {
                            depths[index + 1] = groupingDepth;
                            index++;
                        } else {
                            inQuotedQualifier = false;
                        }
                    }
                    depths[index + 1] = groupingDepth;
                    continue;
                }

                if (character == '"') {
                    inString = true;
                    depths[index + 1] = groupingDepth;
                    continue;
                }
                if (character == '\'') {
                    inQuotedQualifier = true;
                    depths[index + 1] = groupingDepth;
                    continue;
                }
                if (character == '[') {
                    structuredReferenceDepth++;
                    depths[index + 1] = groupingDepth;
                    continue;
                }
                if (character == ']' && structuredReferenceDepth > 0) {
                    structuredReferenceDepth--;
                    depths[index + 1] = groupingDepth;
                    continue;
                }
                if (structuredReferenceDepth > 0) {
                    depths[index + 1] = groupingDepth;
                    continue;
                }

                if (character == '(') {
                    int preceding = index - 1;
                    while (preceding >= 0 && char.IsWhiteSpace(formula[preceding])) {
                        preceding--;
                    }
                    bool functionCall = preceding >= 0
                        && (char.IsLetterOrDigit(formula[preceding])
                            || formula[preceding] == '_'
                            || formula[preceding] == '.'
                            || formula[preceding] == ')');
                    bool grouping = !functionCall;
                    parentheses.Push(grouping);
                    if (grouping) {
                        groupingDepth++;
                    }
                } else if (character == ')' && parentheses.Count > 0) {
                    if (parentheses.Pop()) {
                        groupingDepth--;
                    }
                }

                depths[index + 1] = groupingDepth;
            }

            return depths;
        }

        private void AddFormulaIntersectionDependency(
            IReadOnlyList<FormulaDependencyReferenceMatch> matches,
            int startIndex,
            int endIndex,
            int? sourceRow,
            ISet<string> dependencies) {
            ExcelSheet? intersectionSheet = null;
            int intersectionR1 = 1;
            int intersectionC1 = 1;
            int intersectionR2 = A1.MaxRows;
            int intersectionC2 = A1.MaxColumns;
            for (int index = startIndex; index <= endIndex; index++) {
                string reference = matches[index].Reference;
                if (!TryResolveFormulaDependencyReference(
                    reference,
                    sourceRow,
                    out ExcelSheet sheet,
                    out int r1,
                    out int c1,
                    out int r2,
                    out int c2,
                    out _)
                    || intersectionSheet != null
                    && !string.Equals(intersectionSheet.Name, sheet.Name, StringComparison.OrdinalIgnoreCase)) {
                    return;
                }

                intersectionSheet ??= sheet;
                intersectionR1 = Math.Max(intersectionR1, r1);
                intersectionC1 = Math.Max(intersectionC1, c1);
                intersectionR2 = Math.Min(intersectionR2, r2);
                intersectionC2 = Math.Min(intersectionC2, c2);
                if (intersectionR1 > intersectionR2 || intersectionC1 > intersectionC2) {
                    return;
                }
            }

            if (intersectionSheet == null) {
                return;
            }

            string start = A1.CellReference(intersectionR1, intersectionC1);
            string end = A1.CellReference(intersectionR2, intersectionC2);
            dependencies.Add(intersectionR1 == intersectionR2 && intersectionC1 == intersectionC2
                ? $"{intersectionSheet.Name}!{start}"
                : $"{intersectionSheet.Name}!{start}:{end}");
        }
    }
}
