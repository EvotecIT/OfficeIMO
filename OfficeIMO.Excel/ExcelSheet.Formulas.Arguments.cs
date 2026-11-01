using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryResolveTextArgumentValues(IEnumerable<string> tokens, out List<string> values, out string? errorCode) {
            values = new List<string>();
            errorCode = null;
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            foreach (string token in tokens) {
                if (TryResolveFormulaRange(token, out var rangeValues, ref remainingCellBudget, out bool isReference)) {
                    foreach (var rangeValue in rangeValues) {
                        if (rangeValue.IsUnresolvedFormula) { values.Clear(); return false; }
                        if (rangeValue.IsError) { errorCode ??= rangeValue.ErrorCode; continue; }
                        values.Add(FormulaValueToText(rangeValue));
                    }
                    continue;
                }
                if (isReference) { values.Clear(); return false; }
                if (!TryResolveFormulaArgument(token, out FormulaArgumentValue value) || value.IsUnresolvedFormula) {
                    values.Clear();
                    return false;
                }
                if (value.IsError) { errorCode ??= value.ErrorCode; continue; }
                values.Add(FormulaValueToText(value));
            }
            return true;
        }

        private bool TryResolveTextArgument(string token, out string text) {
            text = string.Empty;
            if (!TryResolveFormulaArgument(token, out FormulaArgumentValue value)) {
                return false;
            }

            if (value.IsUnresolvedFormula) {
                return false;
            }
            // Keep the typed error in this expression's scope while the caller
            // qualifies its remaining arguments. It must never become ordinary text.
            if (value.IsError && !_formulaReferenceError.IsError) _formulaReferenceError = value;

            text = FormulaValueToText(value);
            return true;
        }

        private bool TryGetTextWholeNumberArgument(string token, out int number) {
            if (!TryGetLookupIntegerArgument(token, out number, out FormulaArgumentValue error)) return false;
            if (error.IsError) {
                if (!_formulaReferenceError.IsError) _formulaReferenceError = error;
                // A harmless positive selector lets the owner qualify the other
                // operands before the expression publishes the pending error.
                number = 1;
            }
            return true;
        }

        private bool TryResolveBooleanArgument(string token, out bool value) {
            string trimmed = token.Trim();
            if (trimmed.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) {
                value = true;
                return true;
            }

            if (trimmed.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) {
                value = false;
                return true;
            }

            if (TryEvaluateFormulaOrNumeric(trimmed, out double numeric)) {
                value = Math.Abs(numeric) >= double.Epsilon;
                return true;
            }

            value = false;
            return false;
        }

        private static string FormulaValueToText(FormulaArgumentValue value) {
            return value.ErrorCode ?? value.Text ?? (value.Number.HasValue ? InvariantNumberText.Get(value.Number.Value) : string.Empty);
        }

        private bool TryResolveFormulaArgument(string token, out FormulaArgumentValue value, bool allowScalarExpression = true) {
            string trimmed = token.Trim();
            if (allowScalarExpression) return TryEvaluateFormulaValue(trimmed, out value);
            if (ExcelFormulaExpressionParser.TryParseTextLiteral(trimmed, out string literalText)) {
                value = new FormulaArgumentValue(null, literalText);
                return true;
            }

            if (trimmed.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) {
                value = FormulaArgumentValue.Boolean(true);
                return true;
            }

            if (trimmed.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) {
                value = FormulaArgumentValue.Boolean(false);
                return true;
            }

            if (TryParseFormulaErrorLiteral(trimmed, out string errorCode)) {
                value = FormulaArgumentValue.Error(errorCode);
                return true;
            }

            if (TryParseQualifiedFormulaCellReference(trimmed, out ExcelSheet sheet, out int row, out int column)) {
                value = ResolveFormulaDependency(sheet, row, column);
                return true;
            }

            if (TryResolveFormulaRangeReference(trimmed, out ExcelSheet rangeSheet, out int r1, out int c1, out int r2, out int c2)
                && r1 == r2
                && c1 == c2) {
                value = ResolveFormulaDependency(rangeSheet, r1, c1);
                return true;
            }

            if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric)) {
                value = new FormulaArgumentValue(numeric, trimmed);
                return true;
            }

            if (TryResolveDefinedNameConstant(trimmed, out value)) return true;

            if (TryEvaluateFormulaValue(trimmed, out value, allowScalarExpression: false)) {
                return true;
            }

            value = default;
            return false;
        }

        private static bool FormulaValuesEqual(FormulaArgumentValue left, FormulaArgumentValue right) {
            if (left.Number.HasValue && right.Number.HasValue) {
                return left.Number.Value == right.Number.Value;
            }

            string leftText = left.Text ?? (left.Number.HasValue ? InvariantNumberText.Get(left.Number.Value) : string.Empty);
            string rightText = right.Text ?? (right.Number.HasValue ? InvariantNumberText.Get(right.Number.Value) : string.Empty);
            return string.Equals(leftText, rightText, StringComparison.OrdinalIgnoreCase);
        }

        private bool TryResolveFormulaArguments(string args, out List<FormulaArgumentValue> values,
            bool skipBlankScalarReferences = false) {
            values = new List<FormulaArgumentValue>();
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            foreach (string trimmed in SplitFormulaArguments(args)) {
                if (TryResolveFormulaRange(trimmed, out var rangeValues, ref remainingCellBudget, out bool isReference)) {
                    // Integer aggregates ignore an empty direct cell argument.
                    // The same empty position inside a range remains a zero operand.
                    // Resolve and charge the reference before filtering its value.
                    if (skipBlankScalarReferences && rangeValues.Count == 1
                        && !rangeValues[0].HasValue && !rangeValues[0].IsUnresolvedFormula
                        && !rangeValues[0].IsUnevaluatedFormulaCache
                        && TryParseQualifiedFormulaCellReference(trimmed, out _, out _, out _)) continue;
                    values.AddRange(rangeValues);
                    continue;
                }
                // A recognized reference that exceeds the shared budget must
                // not be resolved again through the scalar fallback.
                if (isReference) { values.Clear(); return false; }

                if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric)) {
                    values.Add(new FormulaArgumentValue(numeric, trimmed));
                    continue;
                }

                if (TryResolveFormulaArgument(trimmed, out var argumentValue)) {
                    if (argumentValue.SourceCellKind == null) {
                        argumentValue = NormalizeLookupResult(argumentValue);
                        if (!argumentValue.Number.HasValue && !argumentValue.IsError
                            && TryParseValueFunctionNumber(argumentValue.Text ?? string.Empty, out double scalarNumber)) {
                            argumentValue = new FormulaArgumentValue(scalarNumber, argumentValue.Text,
                                isUnevaluatedFormulaCache: argumentValue.IsUnevaluatedFormulaCache);
                        }
                    }
                    values.Add(argumentValue);
                    continue;
                }

                values.Clear();
                return false;
            }

            return true;
        }

        private static IReadOnlyList<string> SplitFormulaArguments(string args, bool preserveEmpty = false) {
            var tokens = new List<string>();
            var builder = new StringBuilder();
            int depth = 0;
            int bracketDepth = 0;
            int arrayDepth = 0;
            bool inString = false;
            bool inQuotedQualifier = false;

            for (int index = 0; index < args.Length; index++) {
                char ch = args[index];
                if (!inString && bracketDepth > 0 && ch == '\'' && index + 1 < args.Length) {
                    builder.Append(ch);
                    builder.Append(args[++index]);
                    continue;
                }
                if (!inString && bracketDepth == 0 && ch == '\'') {
                    builder.Append(ch);
                    if (inQuotedQualifier && index + 1 < args.Length && args[index + 1] == '\'') {
                        builder.Append(args[++index]);
                    } else {
                        inQuotedQualifier = !inQuotedQualifier;
                    }
                    continue;
                }
                if (inQuotedQualifier) { builder.Append(ch); continue; }
                if (bracketDepth == 0 && ch == '"') {
                    builder.Append(ch);
                    if (inString && index + 1 < args.Length && args[index + 1] == '"') {
                        index++;
                        builder.Append(args[index]);
                        continue;
                    }

                    inString = !inString;
                    continue;
                }

                if (!inString && ch == '(') {
                    depth++;
                    builder.Append(ch);
                    continue;
                }

                if (!inString && ch == ')') {
                    depth--;
                    if (depth < 0) {
                        return Array.Empty<string>();
                    }

                    builder.Append(ch);
                    continue;
                }

                if (!inString && ch == '[') {
                    bracketDepth++;
                    builder.Append(ch);
                    continue;
                }

                if (!inString && ch == ']') {
                    bracketDepth--;
                    if (bracketDepth < 0) {
                        return Array.Empty<string>();
                    }

                    builder.Append(ch);
                    continue;
                }

                if (!inString && bracketDepth == 0 && (ch == '{' || ch == '}')) {
                    arrayDepth += ch == '{' ? 1 : -1;
                    if (arrayDepth < 0 || arrayDepth > 1) return Array.Empty<string>();
                }

                if (!inString && ch == ',' && depth == 0 && bracketDepth == 0 && arrayDepth == 0) {
                    AddToken(tokens, builder, preserveEmpty);
                    continue;
                }

                builder.Append(ch);
            }

            if (depth != 0 || bracketDepth != 0 || arrayDepth != 0 || inString || inQuotedQualifier) {
                return Array.Empty<string>();
            }

            AddToken(tokens, builder, preserveEmpty);
            return tokens;
        }

        private static void AddToken(List<string> tokens, StringBuilder builder, bool preserveEmpty = false) {
            string token = builder.ToString().Trim();
            if (token.Length > 0 || preserveEmpty) {
                tokens.Add(token);
            }

            builder.Clear();
        }

    }
}
