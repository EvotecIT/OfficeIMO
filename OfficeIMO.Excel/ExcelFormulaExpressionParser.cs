using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Excel {
    internal sealed class ExcelFormulaFunctionCallSyntax {
        internal ExcelFormulaFunctionCallSyntax(string name, string arguments, int nameStart, int nameLength) {
            Name = name;
            Arguments = arguments;
            NameStart = nameStart;
            NameLength = nameLength;
        }

        internal string Name { get; }
        internal string Arguments { get; }
        internal int NameStart { get; }
        internal int NameLength { get; }
    }

    internal sealed class ExcelFormulaBinaryExpressionSyntax {
        internal ExcelFormulaBinaryExpressionSyntax(string left, string @operator, string right) {
            Left = left;
            Operator = @operator;
            Right = right;
        }

        internal string Left { get; }
        internal string Operator { get; }
        internal string Right { get; }
    }

    /// <summary>
    /// Bounded lexical parser for the lightweight evaluator's top-level expression shapes.
    /// It deliberately does not claim to be a complete Excel calculation grammar; unsupported
    /// shapes fail closed and are left for Excel to calculate.
    /// </summary>
    internal static class ExcelFormulaExpressionParser {
        private static readonly HashSet<string> SupportedFunctions = new HashSet<string>(
            ("SUM AVERAGE AVERAGEA MIN MINA MAX MAXA COUNT COUNTA COUNTBLANK SUBTOTAL COUNTIF SUMIF AVERAGEIF " +
             "COUNTIFS SUMIFS AVERAGEIFS MINIFS MAXIFS PRODUCT MEDIAN LARGE SMALL MODE.SNGL MODE GEOMEAN HARMEAN " +
             "AVEDEV DEVSQ SUMXMY2 SUMX2MY2 SUMX2PY2 SUMSQ SUMPRODUCT STDEV.S STDEV.P VAR.S VAR.P PERCENTILE.INC " +
             "PERCENTILE.EXC QUARTILE.INC QUARTILE.EXC PERCENTRANK.INC PERCENTRANK.EXC RANK.EQ RANK.AVG COVAR " +
             "COVARIANCE.P COVARIANCE.S CORREL SLOPE INTERCEPT RSQ FORECAST.LINEAR PMT PV FV NPER NPV VLOOKUP HLOOKUP " +
             "XLOOKUP INDEX MATCH XMATCH ABS SIGN ROUND ROUNDUP ROUNDDOWN MROUND TRUNC INT CEILING.MATH FLOOR.MATH " +
             "CEILING FLOOR POWER SQRT LN LOG10 EXP PI RADIANS DEGREES MOD ROW COLUMN ROWS COLUMNS DATE TIME DATEVALUE " +
             "TIMEVALUE TODAY NOW YEAR MONTH DAY HOUR MINUTE SECOND DATEDIF YEARFRAC EDATE EOMONTH DAYS DAYS360 WEEKDAY " +
             "WEEKNUM ISOWEEKNUM NETWORKDAYS WORKDAY.INTL WORKDAY IF IFS SWITCH CHOOSE ISBLANK ISNUMBER ISTEXT ISERROR " +
             "ISERR ISNA ISFORMULA AND OR NOT IFERROR IFNA CONCAT CONCATENATE TEXT TEXTJOIN TEXTBEFORE TEXTAFTER " +
             "FORMULATEXT LEFT RIGHT MID LEN TRIM UPPER LOWER PROPER SUBSTITUTE FIND SEARCH VALUE EXACT REPT")
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries),
            StringComparer.OrdinalIgnoreCase);

        internal static bool TryParseSupportedFunctionCall(string formula, out ExcelFormulaFunctionCallSyntax? call) {
            if (TryParseFunctionCall(formula, out ExcelFormulaFunctionCallSyntax? parsed)
                && SupportedFunctions.Contains(parsed!.Name)) {
                call = parsed;
                return true;
            }
            call = null;
            return false;
        }

        internal static bool TryParseFunctionCall(string formula, out ExcelFormulaFunctionCallSyntax? call) {
            call = null;
            if (formula == null) return false;
            int cursor = 0;
            SkipWhitespace(formula, ref cursor);
            if (cursor < formula.Length && formula[cursor] == '=') {
                cursor++;
                SkipWhitespace(formula, ref cursor);
            }
            int nameStart = cursor;
            if (cursor >= formula.Length || !IsNameStart(formula[cursor])) return false;
            cursor++;
            while (cursor < formula.Length && IsNamePart(formula[cursor])) cursor++;
            int nameEnd = cursor;
            SkipWhitespace(formula, ref cursor);
            if (cursor >= formula.Length || formula[cursor] != '(') return false;

            int argumentsStart = ++cursor;
            int depth = 1;
            int structuredDepth = 0;
            bool inString = false;
            bool inQuotedQualifier = false;
            while (cursor < formula.Length) {
                char current = formula[cursor];
                if (inString) {
                    if (current == '"') {
                        if (cursor + 1 < formula.Length && formula[cursor + 1] == '"') {
                            cursor += 2;
                            continue;
                        }
                        inString = false;
                    }
                    cursor++;
                    continue;
                }
                if (inQuotedQualifier) {
                    if (current == '\'') {
                        if (cursor + 1 < formula.Length && formula[cursor + 1] == '\'') {
                            cursor += 2;
                            continue;
                        }
                        inQuotedQualifier = false;
                    }
                    cursor++;
                    continue;
                }
                if (current == '"') { inString = true; cursor++; continue; }
                if (structuredDepth > 0 && current == '\'') {
                    cursor += cursor + 1 < formula.Length ? 2 : 1;
                    continue;
                }
                if (current == '\'') { inQuotedQualifier = true; cursor++; continue; }
                if (current == '[') { structuredDepth++; cursor++; continue; }
                if (current == ']' && structuredDepth > 0) { structuredDepth--; cursor++; continue; }
                if (structuredDepth == 0 && current == '(') { depth++; cursor++; continue; }
                if (structuredDepth == 0 && current == ')') {
                    depth--;
                    if (depth == 0) break;
                }
                cursor++;
            }
            if (inString || inQuotedQualifier || structuredDepth != 0 || depth != 0) return false;
            int close = cursor++;
            SkipWhitespace(formula, ref cursor);
            if (cursor != formula.Length) return false;

            call = new ExcelFormulaFunctionCallSyntax(
                formula.Substring(nameStart, nameEnd - nameStart),
                formula.Substring(argumentsStart, close - argumentsStart),
                nameStart,
                nameEnd - nameStart);
            return true;
        }

        internal static bool TryParseArithmetic(string formula, out ExcelFormulaBinaryExpressionSyntax? expression) =>
            TryParseBinary(formula, new[] { "+", "-", "*", "/" }, out expression);

        internal static bool TryParseComparison(string formula, out ExcelFormulaBinaryExpressionSyntax? expression) =>
            TryParseBinary(formula, new[] { ">=", "<=", "<>", "=", ">", "<" }, out expression);

        private static bool TryParseBinary(
            string formula,
            IReadOnlyList<string> operators,
            out ExcelFormulaBinaryExpressionSyntax? expression) {
            expression = null;
            if (formula == null) return false;
            int start = 0;
            int end = formula.Length;
            while (start < end && char.IsWhiteSpace(formula[start])) start++;
            if (start < end && formula[start] == '=') {
                start++;
                while (start < end && char.IsWhiteSpace(formula[start])) start++;
            }
            while (end > start && char.IsWhiteSpace(formula[end - 1])) end--;
            if (start >= end) return false;

            int parenthesisDepth = 0;
            int bracketDepth = 0;
            bool inString = false;
            bool inQuotedQualifier = false;
            int operatorStart = -1;
            string? selectedOperator = null;
            for (int cursor = start; cursor < end; cursor++) {
                char current = formula[cursor];
                if (inString) {
                    if (current == '"') {
                        if (cursor + 1 < end && formula[cursor + 1] == '"') { cursor++; continue; }
                        inString = false;
                    }
                    continue;
                }
                if (inQuotedQualifier) {
                    if (current == '\'') {
                        if (cursor + 1 < end && formula[cursor + 1] == '\'') { cursor++; continue; }
                        inQuotedQualifier = false;
                    }
                    continue;
                }
                if (current == '"') { inString = true; continue; }
                if (bracketDepth > 0 && current == '\'') {
                    if (cursor + 1 < end) cursor++;
                    continue;
                }
                if (current == '\'') { inQuotedQualifier = true; continue; }
                if (current == '[') { bracketDepth++; continue; }
                if (current == ']' && bracketDepth > 0) { bracketDepth--; continue; }
                if (bracketDepth > 0) continue;
                if (current == '(') { parenthesisDepth++; continue; }
                if (current == ')') {
                    if (parenthesisDepth == 0) return false;
                    parenthesisDepth--;
                    continue;
                }
                if (parenthesisDepth != 0) continue;

                string? found = null;
                foreach (string candidate in operators) {
                    if (cursor + candidate.Length <= end
                        && string.Compare(formula, cursor, candidate, 0, candidate.Length, StringComparison.Ordinal) == 0) {
                        found = candidate;
                        break;
                    }
                }
                if (found == null || IsUnarySign(formula, start, cursor, found)) continue;
                if (operatorStart >= 0) return false;
                operatorStart = cursor;
                selectedOperator = found;
                cursor += found.Length - 1;
            }
            if (inString || inQuotedQualifier || bracketDepth != 0 || parenthesisDepth != 0
                || operatorStart < 0 || selectedOperator == null) return false;
            string left = formula.Substring(start, operatorStart - start).Trim();
            string right = formula.Substring(operatorStart + selectedOperator.Length, end - operatorStart - selectedOperator.Length).Trim();
            if (left.Length == 0 || right.Length == 0) return false;
            expression = new ExcelFormulaBinaryExpressionSyntax(left, selectedOperator, right);
            return true;
        }

        private static bool IsUnarySign(string formula, int start, int position, string value) {
            if (value != "+" && value != "-") return false;
            int previous = position - 1;
            while (previous >= start && char.IsWhiteSpace(formula[previous])) previous--;
            return previous < start || formula[previous] == '(' || formula[previous] == ','
                || formula[previous] == '+' || formula[previous] == '-' || formula[previous] == '*'
                || formula[previous] == '/' || formula[previous] == '^' || formula[previous] == '='
                || formula[previous] == '<' || formula[previous] == '>';
        }

        private static void SkipWhitespace(string value, ref int cursor) {
            while (cursor < value.Length && char.IsWhiteSpace(value[cursor])) cursor++;
        }

        private static bool IsNameStart(char value) => value == '_' || char.IsLetter(value);
        private static bool IsNamePart(char value) =>
            IsNameStart(value) || char.IsDigit(value) || value == '.' ||
            CharUnicodeInfo.GetUnicodeCategory(value) is UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark;
    }
}