using OfficeIMO.Excel.LegacyXls.Write;
using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Encodes expressions that call future functions (_xlfn. names). The established BIFF8 parser has no such calls,
    /// so these expressions are split here by operator precedence and function arguments; each piece without a future
    /// function goes back through the normal path.
    /// </summary>
    internal static partial class XlsbFormulaEncoder {
        private const string FutureFunctionPrefix = "_xlfn.";

        // Binary operators from the lowest precedence level to the highest, with their BIFF12 Ptg (2.5.98.16).
        private static readonly (string Text, byte Token)[][] BinaryOperatorLevels = {
            new[] { ("<>", (byte)0x0E), ("<=", (byte)0x0A), (">=", (byte)0x0C), ("=", (byte)0x0B), ("<", (byte)0x09), (">", (byte)0x0D) },
            new[] { ("&", (byte)0x08) },
            new[] { ("+", (byte)0x03), ("-", (byte)0x04) },
            new[] { ("*", (byte)0x05), ("/", (byte)0x06) },
            new[] { ("^", (byte)0x07) }
        };

        private static bool ContainsFutureFunction(string formulaText) {
            bool quoted = false;
            for (int index = 0; index < formulaText.Length; index++) {
                char character = formulaText[index];
                if (character == '"') quoted = !quoted;
                else if (!quoted && character == '_'
                    && string.Compare(formulaText, index, FutureFunctionPrefix, 0, FutureFunctionPrefix.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                    return true;
                }
            }
            return false;
        }

        private static bool TryEncodeFutureFunctionExpression(string formulaText, int depth, EncodeScope scope,
            out byte[] tokens, out byte[] extra, out string? reason) {
            tokens = Array.Empty<byte>();
            extra = Array.Empty<byte>();
            reason = null;
            if (depth > MaximumControlFunctionDepth) {
                reason = $"Formula nesting exceeds the supported depth of {MaximumControlFunctionDepth}.";
                return false;
            }
            string text = formulaText.Trim();
            if (text.StartsWith("=", StringComparison.Ordinal)) text = text.Substring(1).Trim();
            if (text.IndexOf("_xlpm.", StringComparison.OrdinalIgnoreCase) >= 0) {
                reason = "LAMBDA parameter names are not supported by native XLSB formula generation.";
                return false;
            }

            using var output = new MemoryStream();
            using var extraOutput = new MemoryStream();
            if (TryFindBinarySplit(text, out int split, out int operatorLength, out byte operatorToken)) {
                if (!TryEncodeTokens(text.Substring(0, split), depth + 1, scope, out byte[] left, out byte[] leftExtra, out reason)
                    || !TryEncodeTokens(text.Substring(split + operatorLength), depth + 1, scope, out byte[] right, out byte[] rightExtra, out reason)) {
                    return false;
                }
                output.Write(left, 0, left.Length);
                output.Write(right, 0, right.Length);
                output.WriteByte(operatorToken);
                extraOutput.Write(leftExtra, 0, leftExtra.Length);
                extraOutput.Write(rightExtra, 0, rightExtra.Length);
            } else if (text.Length > 1 && (text[0] == '-' || text[0] == '+')) {
                if (!TryEncodeTokens(text.Substring(1), depth + 1, scope, out byte[] operand, out extra, out reason)) return false;
                output.Write(operand, 0, operand.Length);
                output.WriteByte(text[0] == '-' ? (byte)0x13 : (byte)0x12); // PtgUminus / PtgUplus
                extraOutput.Write(extra, 0, extra.Length);
            } else if (text.Length > 1 && text[text.Length - 1] == '%') {
                if (!TryEncodeTokens(text.Substring(0, text.Length - 1), depth + 1, scope, out byte[] operand, out extra, out reason)) return false;
                output.Write(operand, 0, operand.Length);
                output.WriteByte(0x14); // PtgPercent
                extraOutput.Write(extra, 0, extra.Length);
            } else if (text.Length > 2 && text[0] == '(' && FindClosingParenthesis(text, 0) == text.Length - 1) {
                if (!TryEncodeTokens(text.Substring(1, text.Length - 2), depth + 1, scope, out byte[] operand, out extra, out reason)) return false;
                output.Write(operand, 0, operand.Length);
                output.WriteByte(0x15); // PtgParen
                extraOutput.Write(extra, 0, extra.Length);
            } else if (!TryEncodeFunctionCall(text, depth, scope, output, extraOutput, out reason)) {
                reason ??= "Future functions (_xlfn.) are supported in native XLSB formulas as function calls combined with operators and other functions.";
                return false;
            }

            tokens = output.ToArray();
            extra = extraOutput.ToArray();
            return true;
        }

        private static bool TryEncodeFunctionCall(string text, int depth, EncodeScope scope, Stream output, Stream extraOutput, out string? reason) {
            reason = null;
            int open = text.IndexOf('(');
            if (open <= 0 || FindClosingParenthesis(text, open) != text.Length - 1) return false;
            string name = text.Substring(0, open).Trim();
            if (!TrySplitArguments(text.Substring(open + 1, text.Length - open - 2), out IReadOnlyList<string>? arguments)) return false;
            if (arguments!.Count == 1 && arguments[0].Length == 0) arguments = Array.Empty<string>();

            if (name.StartsWith(FutureFunctionPrefix, StringComparison.OrdinalIgnoreCase)) {
                if (!XlsbFutureFunctions.TryGetName(name, out string canonical) || !XlsbFutureFunctions.IsEncodable(canonical)) {
                    reason = $"Future function '{name}' is not supported by native XLSB formula generation.";
                    return false;
                }
                if (arguments.Count > 254) {
                    reason = $"Future function '{canonical}' has more than 254 arguments.";
                    return false;
                }
                uint? nameIndex = scope.FutureFunctions?.Invoke(canonical);
                if (!nameIndex.HasValue) {
                    reason = $"Native XLSB writing cannot add the defined name that future function '{canonical}' requires here.";
                    return false;
                }
                // user-defined-or-future-function (2.5.98.10, iftab 0x00FF): PtgName to the future-function BrtName,
                // then the arguments, then PtgFuncVar counting the name as its first parameter.
                output.WriteByte(0x23); // PtgName, reference class
                WriteUInt32(output, nameIndex.Value);
                if (!TryEncodeArguments(arguments, depth, scope, output, extraOutput, out reason)) return false;
                output.WriteByte(0x42);
                output.WriteByte(checked((byte)(arguments.Count + 1)));
                WriteUInt16(output, 0x00FF);
                return true;
            }

            if (LegacyXlsFormulaFunctionWriterMetadata.TryGetVariableFunction(text, out ushort variableId, out _)) {
                if (!LegacyXlsFormulaFunctionWriterMetadata.IsSupportedVariableFunctionArgumentCount(variableId, arguments.Count)) {
                    reason = $"Formula function {name.ToUpperInvariant()} does not accept {arguments.Count} arguments.";
                    return false;
                }
                if (!TryEncodeArguments(arguments, depth, scope, output, extraOutput, out reason)) return false;
                output.WriteByte(0x42); // PtgFuncVar
                output.WriteByte(checked((byte)arguments.Count));
                WriteUInt16(output, variableId);
                return true;
            }
            if (LegacyXlsFormulaFunctionWriterMetadata.TryGetFixedFunction(text, out ushort fixedId, out int parameterCount, out _)) {
                if (arguments.Count != parameterCount) {
                    reason = $"Formula function {name.ToUpperInvariant()} requires {parameterCount} arguments.";
                    return false;
                }
                if (!TryEncodeArguments(arguments, depth, scope, output, extraOutput, out reason)) return false;
                output.WriteByte(0x41); // PtgFunc
                WriteUInt16(output, fixedId);
                return true;
            }
            return false;
        }

        /// <summary>Finds the operator to evaluate last: the rightmost top-level binary operator of the lowest level.</summary>
        private static bool TryFindBinarySplit(string text, out int split, out int length, out byte token) {
            foreach ((string Text, byte Token)[] level in BinaryOperatorLevels) {
                split = -1;
                length = 0;
                token = 0;
                int nesting = 0;
                bool quoted = false;
                bool sheetQuoted = false;
                for (int index = 0; index < text.Length; index++) {
                    char character = text[index];
                    if (quoted) {
                        if (character == '"') quoted = false;
                        continue;
                    }
                    if (sheetQuoted) {
                        if (character == '\'') sheetQuoted = false;
                        continue;
                    }
                    switch (character) {
                        case '"': quoted = true; continue;
                        case '\'': sheetQuoted = true; continue;
                        case '(': case '{': case '[': nesting++; continue;
                        case ')': case '}': case ']': nesting--; continue;
                        case '#':
                            // Error literals such as #DIV/0! and #N/A contain operator characters.
                            while (index + 1 < text.Length && (char.IsLetterOrDigit(text[index + 1]) || text[index + 1] is '/' or '!' or '?' or '_')) index++;
                            continue;
                    }
                    if (nesting != 0 || index == 0) continue;
                    foreach ((string operatorText, byte operatorToken) in level) {
                        if (string.CompareOrdinal(text, index, operatorText, 0, operatorText.Length) != 0) continue;
                        if (operatorText.Length == 1 && character is '<' or '>' && index + 1 < text.Length && text[index + 1] is '=' or '>') break;
                        if (operatorText == "=" && text[index - 1] is '<' or '>') break;
                        if (character is '+' or '-' && !IsBinarySign(text, index)) break;
                        split = index;
                        length = operatorText.Length;
                        token = operatorToken;
                        break;
                    }
                    if (length == 2 && split == index) index++;
                }
                if (split > 0 && split + length < text.Length) return true;
            }
            split = -1;
            length = 0;
            token = 0;
            return false;
        }

        // A sign is binary after an operand; after an operator or "(" it is unary, and after a mantissa's E it is an exponent.
        private static bool IsBinarySign(string text, int index) {
            int previous = index - 1;
            while (previous >= 0 && text[previous] == ' ') previous--;
            if (previous < 0) return false;
            char before = text[previous];
            if (!(char.IsLetterOrDigit(before) || before is '_' or '.' or ')' or '"' or '}' or '%' or '\'' or ']' or '!' or '$' or '?')) return false;
            if ((before == 'E' || before == 'e') && previous == index - 1) {
                int start = previous - 1;
                while (start >= 0 && (char.IsDigit(text[start]) || text[start] == '.')) start--;
                bool mantissa = start < previous - 1 && (start < 0 || !(char.IsLetter(text[start]) || text[start] is '_' or '$' or '.'));
                if (mantissa) return false;
            }
            return true;
        }

        private static int FindClosingParenthesis(string text, int open) {
            int depth = 0;
            bool quoted = false;
            bool sheetQuoted = false;
            for (int index = open; index < text.Length; index++) {
                char character = text[index];
                if (quoted) {
                    if (character == '"') quoted = false;
                    continue;
                }
                if (sheetQuoted) {
                    if (character == '\'') sheetQuoted = false;
                    continue;
                }
                if (character == '"') quoted = true;
                else if (character == '\'') sheetQuoted = true;
                else if (character == '(') depth++;
                else if (character == ')' && --depth == 0) return index;
            }
            return -1;
        }
    }
}
