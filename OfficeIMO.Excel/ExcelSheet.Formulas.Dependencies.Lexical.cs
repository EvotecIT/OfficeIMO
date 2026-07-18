using System;
using System.Collections.Generic;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private readonly struct FormulaArgumentSpan {
            internal FormulaArgumentSpan(int start, int end) {
                Start = start;
                End = end;
            }

            internal int Start { get; }
            internal int End { get; }
        }

        private readonly struct FormulaLexicalBinding {
            internal FormulaLexicalBinding(
                string name,
                int declarationStart,
                int declarationEnd,
                int scopeStart,
                int scopeEnd) {
                Name = name;
                DeclarationStart = declarationStart;
                DeclarationEnd = declarationEnd;
                ScopeStart = scopeStart;
                ScopeEnd = scopeEnd;
            }

            internal string Name { get; }
            internal int DeclarationStart { get; }
            internal int DeclarationEnd { get; }
            internal int ScopeStart { get; }
            internal int ScopeEnd { get; }

            internal bool Shadows(string alias, int index, int length) {
                if (!string.Equals(Name, alias, StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }

                int end = index + length;
                return index >= DeclarationStart && end <= DeclarationEnd
                    || index >= ScopeStart && end <= ScopeEnd;
            }
        }

        private static IReadOnlyList<FormulaLexicalBinding> GetFormulaLexicalBindings(string formula) {
            List<FormulaLexicalBinding>? bindings = null;
            for (int index = 0; index < formula.Length; index++) {
                if (!TryGetFormulaLexicalFunctionCall(formula, index, out bool isLet, out int openingParenthesis)
                    || !TryFindClosingFormulaParenthesis(formula, openingParenthesis, out int closingParenthesis)) {
                    continue;
                }

                List<FormulaArgumentSpan> arguments = GetFormulaArgumentSpans(
                    formula,
                    openingParenthesis + 1,
                    closingParenthesis);
                if (isLet) {
                    if (arguments.Count < 3 || arguments.Count % 2 == 0) {
                        continue;
                    }

                    for (int argumentIndex = 0; argumentIndex < arguments.Count - 1; argumentIndex += 2) {
                        AddFormulaLexicalBinding(
                            formula,
                            arguments[argumentIndex],
                            arguments[argumentIndex + 2].Start,
                            closingParenthesis,
                            ref bindings);
                    }
                } else {
                    if (arguments.Count < 2) {
                        continue;
                    }

                    int scopeStart = arguments[arguments.Count - 1].Start;
                    for (int argumentIndex = 0; argumentIndex < arguments.Count - 1; argumentIndex++) {
                        AddFormulaLexicalBinding(
                            formula,
                            arguments[argumentIndex],
                            scopeStart,
                            closingParenthesis,
                            ref bindings);
                    }
                }
            }

            return bindings ?? (IReadOnlyList<FormulaLexicalBinding>)Array.Empty<FormulaLexicalBinding>();
        }

        private static void AddFormulaLexicalBinding(
            string formula,
            FormulaArgumentSpan declaration,
            int scopeStart,
            int scopeEnd,
            ref List<FormulaLexicalBinding>? bindings) {
            int declarationStart = declaration.Start;
            int declarationEnd = declaration.End;
            while (declarationStart < declarationEnd && char.IsWhiteSpace(formula[declarationStart])) {
                declarationStart++;
            }
            while (declarationEnd > declarationStart && char.IsWhiteSpace(formula[declarationEnd - 1])) {
                declarationEnd--;
            }

            if (declarationStart >= declarationEnd) {
                return;
            }

            string name = formula.Substring(declarationStart, declarationEnd - declarationStart);
            if (!IsFormulaDefinedNameToken(name)) {
                return;
            }

            bindings ??= new List<FormulaLexicalBinding>();
            bindings.Add(new FormulaLexicalBinding(
                name,
                declarationStart,
                declarationEnd,
                scopeStart,
                scopeEnd));
        }

        private static bool TryGetFormulaLexicalFunctionCall(
            string formula,
            int index,
            out bool isLet,
            out int openingParenthesis) {
            if (TryMatchFormulaFunctionName(formula, index, "LET", out openingParenthesis)
                || TryMatchFormulaFunctionName(formula, index, "_xlfn.LET", out openingParenthesis)) {
                isLet = true;
                return true;
            }

            if (TryMatchFormulaFunctionName(formula, index, "LAMBDA", out openingParenthesis)
                || TryMatchFormulaFunctionName(formula, index, "_xlfn.LAMBDA", out openingParenthesis)) {
                isLet = false;
                return true;
            }

            isLet = false;
            openingParenthesis = -1;
            return false;
        }

        private string MaskFormulaReferenceShapeArguments(string formula, int? sourceRow, int? sourceColumn) {
            char[]? maskedFormula = null;
            int structuredReferenceDepth = 0;
            bool inQuotedQualifier = false;
            for (int index = 0; index < formula.Length; index++) {
                char character = formula[index];
                if (character == '\'') {
                    if (inQuotedQualifier && index + 1 < formula.Length && formula[index + 1] == '\'') {
                        index++;
                    } else {
                        inQuotedQualifier = !inQuotedQualifier;
                    }
                    continue;
                }
                if (inQuotedQualifier) {
                    continue;
                }
                if (character == '[') {
                    structuredReferenceDepth++;
                    continue;
                }
                if (character == ']' && structuredReferenceDepth > 0) {
                    structuredReferenceDepth--;
                    continue;
                }
                if (structuredReferenceDepth > 0
                    || !TryGetFormulaReferenceShapeFunctionCall(formula, index, out int openingParenthesis)
                    || !TryFindClosingFormulaParenthesis(formula, openingParenthesis, out int closingParenthesis)) {
                    continue;
                }

                List<FormulaArgumentSpan> arguments = GetFormulaArgumentSpans(
                    formula,
                    openingParenthesis + 1,
                    closingParenthesis);
                if (arguments.Count != 1) {
                    continue;
                }

                FormulaArgumentSpan argument = arguments[0];
                string reference = formula.Substring(argument.Start, argument.End - argument.Start).Trim();
                bool resolved = TryResolveFormulaRangeReference(
                    reference,
                    sourceRow,
                    out _,
                    out _,
                    out _,
                    out _,
                    out _);
                if (!resolved && sourceColumn.HasValue) {
                    resolved = TryResolveUnqualifiedCurrentRowTableReferenceRange(
                        reference,
                        sourceRow,
                        sourceColumn,
                        out _,
                        out _,
                        out _,
                        out _,
                        out _);
                }
                if (!resolved) {
                    continue;
                }

                maskedFormula ??= formula.ToCharArray();
                for (int position = argument.Start; position < argument.End; position++) {
                    maskedFormula[position] = ' ';
                }

                index = closingParenthesis;
            }

            return maskedFormula == null ? formula : new string(maskedFormula);
        }

        private static bool TryGetFormulaReferenceShapeFunctionCall(
            string formula,
            int index,
            out int openingParenthesis) {
            return TryMatchFormulaFunctionName(formula, index, "ROW", out openingParenthesis)
                || TryMatchFormulaFunctionName(formula, index, "COLUMN", out openingParenthesis)
                || TryMatchFormulaFunctionName(formula, index, "ROWS", out openingParenthesis)
                || TryMatchFormulaFunctionName(formula, index, "COLUMNS", out openingParenthesis);
        }

        private static bool TryMatchFormulaFunctionName(
            string formula,
            int index,
            string functionName,
            out int openingParenthesis) {
            openingParenthesis = -1;
            if (index < 0
                || index + functionName.Length > formula.Length
                || index > 0 && IsFormulaAliasIdentifierCharacter(formula[index - 1])
                || string.Compare(formula, index, functionName, 0, functionName.Length, StringComparison.OrdinalIgnoreCase) != 0) {
                return false;
            }

            int cursor = index + functionName.Length;
            if (cursor < formula.Length && IsFormulaAliasIdentifierCharacter(formula[cursor])) {
                return false;
            }
            while (cursor < formula.Length && char.IsWhiteSpace(formula[cursor])) {
                cursor++;
            }

            if (cursor >= formula.Length || formula[cursor] != '(') {
                return false;
            }

            openingParenthesis = cursor;
            return true;
        }

        private static bool TryFindClosingFormulaParenthesis(
            string formula,
            int openingParenthesis,
            out int closingParenthesis) {
            int depth = 0;
            int arrayDepth = 0;
            int bracketDepth = 0;
            bool inQuotedQualifier = false;
            for (int index = openingParenthesis; index < formula.Length; index++) {
                char character = formula[index];
                if (character == '\'') {
                    if (inQuotedQualifier && index + 1 < formula.Length && formula[index + 1] == '\'') {
                        index++;
                    } else {
                        inQuotedQualifier = !inQuotedQualifier;
                    }
                } else if (inQuotedQualifier) {
                    continue;
                } else if (character == '[') {
                    bracketDepth++;
                } else if (character == ']' && bracketDepth > 0) {
                    bracketDepth--;
                } else if (bracketDepth == 0 && character == '{') {
                    arrayDepth++;
                } else if (bracketDepth == 0 && character == '}' && arrayDepth > 0) {
                    arrayDepth--;
                } else if (bracketDepth == 0 && arrayDepth == 0 && character == '(') {
                    depth++;
                } else if (bracketDepth == 0 && arrayDepth == 0 && character == ')' && --depth == 0) {
                    closingParenthesis = index;
                    return true;
                }
            }

            closingParenthesis = -1;
            return false;
        }

        private static List<FormulaArgumentSpan> GetFormulaArgumentSpans(
            string formula,
            int start,
            int end) {
            var arguments = new List<FormulaArgumentSpan>();
            int depth = 0;
            int arrayDepth = 0;
            int bracketDepth = 0;
            bool inQuotedQualifier = false;
            int argumentStart = start;
            for (int index = start; index < end; index++) {
                char character = formula[index];
                if (character == '\'') {
                    if (inQuotedQualifier && index + 1 < end && formula[index + 1] == '\'') {
                        index++;
                    } else {
                        inQuotedQualifier = !inQuotedQualifier;
                    }
                } else if (inQuotedQualifier) {
                    continue;
                } else if (character == '[') {
                    bracketDepth++;
                } else if (character == ']' && bracketDepth > 0) {
                    bracketDepth--;
                } else if (bracketDepth == 0 && character == '{') {
                    arrayDepth++;
                } else if (bracketDepth == 0 && character == '}' && arrayDepth > 0) {
                    arrayDepth--;
                } else if (bracketDepth == 0 && arrayDepth == 0 && character == '(') {
                    depth++;
                } else if (bracketDepth == 0 && arrayDepth == 0 && character == ')') {
                    depth--;
                } else if (bracketDepth == 0 && arrayDepth == 0 && character == ',' && depth == 0) {
                    arguments.Add(new FormulaArgumentSpan(argumentStart, index));
                    argumentStart = index + 1;
                }
            }

            arguments.Add(new FormulaArgumentSpan(argumentStart, end));
            return arguments;
        }
    }
}
