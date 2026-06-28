using System.Globalization;

namespace OfficeIMO.Excel.LegacyXls.Write {
    /// <summary>
    /// Encodes the first native XLS writer formula subset into BIFF8 parsed-expression tokens.
    /// </summary>
    internal static partial class LegacyXlsFormulaEncoder {
        private const string SupportedSubsetDescription = "Only same-sheet, workbook-internal sheet-qualified, or supported external-workbook sheet-qualified cell references and ranges, supported workbook or current-sheet defined names, numeric constants, string constants, Boolean constants, error constants, array constants in formula records, same-sheet reference range/union/intersection operators, unary +/-, +, -, *, /, ^, &, %, comparisons, explicit parentheses, simple aggregate functions SUM, AVERAGE, MIN, MAX, COUNT, CHOOSE, supported lookup/reference functions, supported variable-arity logical functions, and supported fixed functions are supported by native XLS formula saving.";

        internal static bool TryEncode(string formulaText, out byte[] tokens, out string? reason) {
            return TryEncode(formulaText, LegacyXlsFormulaNameIndex.Empty, formulaSheetIndex: -1, out tokens, out reason);
        }

        internal static bool TryEncode(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens,
            out string? reason) {
            tokens = Array.Empty<byte>();
            reason = null;

            string normalized = NormalizeFormula(formulaText);
            if (normalized.Length == 0) {
                reason = "Formula text is empty.";
                return false;
            }

            if (ContainsFutureFunctionAlias(normalized)) {
                reason = "BIFF8 native XLS formulas do not support Excel future-function aliases using the _xlfn. compatibility prefix.";
                return false;
            }

            if (TryEncodeReferenceOperator(normalized, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeParenthesized(normalized, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeIfFunction(normalized, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeChooseFunction(normalized, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeVariableFunction(normalized, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeFixedFunction(normalized, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeBinary(normalized, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeUnary(normalized, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodePostfixPercent(normalized, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeOperand(normalized, allowArea: false, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            reason = SupportedSubsetDescription;
            return false;
        }

        internal static bool TryEncodeListSource(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens,
            out string? reason) {
            tokens = Array.Empty<byte>();
            reason = null;

            string normalized = NormalizeFormula(formulaText);
            if (normalized.Length == 0) {
                reason = "Formula text is empty.";
                return false;
            }

            if (TryEncodeOperand(normalized, allowArea: true, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            return TryEncode(normalized, nameIndex, formulaSheetIndex, out tokens, out reason);
        }

        internal static bool TryEncodeWithRelativeReferenceAnchor(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            ushort anchorRow,
            ushort anchorColumn,
            out byte[] tokens,
            out string? reason) {
            if (!TryEncode(formulaText, nameIndex, formulaSheetIndex, out tokens, out reason)) {
                return false;
            }

            if (!TryConvertReferencesToRelative(tokens, anchorRow, anchorColumn, out byte[] relativeTokens)) {
                reason = "Formula contains references outside the BIFF8 relative-reference subset.";
                tokens = Array.Empty<byte>();
                return false;
            }

            tokens = relativeTokens;
            return true;
        }

        internal static bool TryConvertReferencesToRelative(byte[] formulaTokens, ushort anchorRow, ushort anchorColumn, out byte[] relativeTokens) {
            relativeTokens = (byte[])formulaTokens.Clone();
            int offset = 0;
            while (offset < relativeTokens.Length) {
                int tokenStart = offset;
                byte token = relativeTokens[offset++];
                switch (token) {
                    case 0x03:
                    case 0x04:
                    case 0x05:
                    case 0x06:
                    case 0x07:
                    case 0x08:
                    case 0x09:
                    case 0x0a:
                    case 0x0b:
                    case 0x0c:
                    case 0x0d:
                    case 0x0e:
                    case 0x0f:
                    case 0x10:
                    case 0x11:
                    case 0x12:
                    case 0x13:
                    case 0x14:
                    case 0x15:
                    case 0x16:
                        break;
                    case 0x17:
                        if (offset + 2 > relativeTokens.Length) return false;
                        int characterCount = relativeTokens[offset];
                        byte flags = relativeTokens[offset + 1];
                        offset += checked(2 + (((flags & 0x01) != 0) ? characterCount * 2 : characterCount));
                        break;
                    case 0x19:
                        offset += 3;
                        break;
                    case 0x1c:
                        offset += 1;
                        break;
                    case 0x1d:
                        offset += 1;
                        break;
                    case 0x1e:
                        offset += 2;
                        break;
                    case 0x1f:
                        offset += 8;
                        break;
                    case 0x39:
                        offset += 6;
                        break;
                    case 0x41:
                        offset += 2;
                        break;
                    case 0x42:
                        offset += 3;
                        break;
                    case 0x43:
                        offset += 4;
                        break;
                    case 0x44:
                    case 0x4c:
                        if (offset + 4 > relativeTokens.Length) return false;
                        relativeTokens[tokenStart] = 0x4c;
                        ConvertRelativeReference(relativeTokens, offset, offset + 2, anchorRow, anchorColumn);
                        offset += 4;
                        break;
                    case 0x45:
                    case 0x4d:
                        if (offset + 8 > relativeTokens.Length) return false;
                        relativeTokens[tokenStart] = 0x4d;
                        ConvertRelativeReference(relativeTokens, offset, offset + 4, anchorRow, anchorColumn);
                        ConvertRelativeReference(relativeTokens, offset + 2, offset + 6, anchorRow, anchorColumn);
                        offset += 8;
                        break;
                    case 0x5a:
                        if (offset + 6 > relativeTokens.Length) return false;
                        ConvertRelativeReference(relativeTokens, offset + 2, offset + 4, anchorRow, anchorColumn);
                        offset += 6;
                        break;
                    case 0x5b:
                        if (offset + 10 > relativeTokens.Length) return false;
                        ConvertRelativeReference(relativeTokens, offset + 2, offset + 6, anchorRow, anchorColumn);
                        ConvertRelativeReference(relativeTokens, offset + 4, offset + 8, anchorRow, anchorColumn);
                        offset += 10;
                        break;
                    default:
                        return false;
                }

                if (offset > relativeTokens.Length) {
                    return false;
                }
            }

            return offset == relativeTokens.Length;
        }

        private static void ConvertRelativeReference(byte[] tokens, int rowOffset, int columnOffset, ushort anchorRow, ushort anchorColumn) {
            ushort row = ReadUInt16(tokens, rowOffset);
            ushort columnBits = ReadUInt16(tokens, columnOffset);
            if ((columnBits & 0x8000) != 0) {
                WriteUInt16(tokens, rowOffset, unchecked((ushort)(short)(row - anchorRow)));
            }

            if ((columnBits & 0x4000) != 0) {
                short relativeColumnOffset = unchecked((short)((columnBits & 0x3fff) - anchorColumn));
                ushort updatedColumnBits = (ushort)((columnBits & 0xc000) | (((ushort)relativeColumnOffset) & 0x3fff));
                WriteUInt16(tokens, columnOffset, updatedColumnBits);
            }
        }

        private static string NormalizeFormula(string formulaText) {
            string normalized = formulaText.Trim();
            return normalized.StartsWith("=", StringComparison.Ordinal)
                ? normalized.Substring(1).Trim()
                : normalized;
        }

        private static bool ContainsFutureFunctionAlias(string formulaText) {
            bool inStringLiteral = false;
            for (int i = 0; i < formulaText.Length; i++) {
                char ch = formulaText[i];
                if (!inStringLiteral && TrySkipQuotedSheetName(formulaText, ref i)) {
                    continue;
                }

                if (ch == '"') {
                    if (inStringLiteral && i + 1 < formulaText.Length && formulaText[i + 1] == '"') {
                        i++;
                        continue;
                    }

                    inStringLiteral = !inStringLiteral;
                    continue;
                }

                if (!inStringLiteral
                    && i + 6 <= formulaText.Length
                    && string.Compare(formulaText, i, "_xlfn.", 0, 6, StringComparison.OrdinalIgnoreCase) == 0) {
                    return true;
                }
            }

            return false;
        }

        private static bool TryEncodeVariableFunction(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            if (!LegacyXlsFormulaFunctionWriterMetadata.TryGetVariableFunction(formulaText, out ushort functionId, out int argumentStart)
                || !formulaText.EndsWith(")", StringComparison.Ordinal)) {
                return false;
            }

            string argumentText = formulaText.Substring(argumentStart, formulaText.Length - argumentStart - 1).Trim();
            IReadOnlyList<string> arguments;
            if (argumentText.Length == 0) {
                if (!LegacyXlsFormulaFunctionWriterMetadata.IsSupportedVariableFunctionArgumentCount(functionId, 0)) {
                    return false;
                }

                arguments = Array.Empty<string>();
            } else if (!TrySplitFunctionArguments(argumentText, out IReadOnlyList<string>? parsedArguments)
                || !LegacyXlsFormulaFunctionWriterMetadata.IsSupportedVariableFunctionArgumentCount(functionId, parsedArguments!.Count)) {
                return false;
            } else {
                arguments = parsedArguments;
            }

            using var stream = new MemoryStream();
            foreach (string argument in arguments) {
                if (!TryEncodeFunctionArgument(argument, allowArea: true, nameIndex, formulaSheetIndex, out byte[] argumentTokens)) {
                    return false;
                }

                stream.Write(argumentTokens, 0, argumentTokens.Length);
            }

            stream.WriteByte(0x42);
            stream.WriteByte(checked((byte)arguments.Count));
            WriteUInt16(stream, functionId);
            tokens = stream.ToArray();
            return true;
        }

        private static bool TryEncodeFixedFunction(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            if (!LegacyXlsFormulaFunctionWriterMetadata.TryGetFixedFunction(formulaText, out ushort functionId, out int parameterCount, out int argumentStart)
                || !formulaText.EndsWith(")", StringComparison.Ordinal)) {
                return false;
            }

            string argumentText = formulaText.Substring(argumentStart, formulaText.Length - argumentStart - 1).Trim();
            IReadOnlyList<string> arguments;
            if (parameterCount == 0 && argumentText.Length == 0) {
                arguments = Array.Empty<string>();
            } else if (parameterCount == 1
                && argumentText.Length == 0
                && LegacyXlsFormulaFunctionWriterMetadata.AllowsMissingReferenceArgument(functionId)) {
                arguments = new[] { string.Empty };
            } else if (!TrySplitFunctionArguments(argumentText, out IReadOnlyList<string>? parsedArguments)
                || parsedArguments!.Count != parameterCount) {
                return false;
            } else {
                arguments = parsedArguments;
            }

            using var stream = new MemoryStream();
            foreach (string argument in arguments) {
                if (!TryEncodeFunctionArgument(argument, allowArea: true, nameIndex, formulaSheetIndex, out byte[] argumentTokens)) {
                    return false;
                }

                stream.Write(argumentTokens, 0, argumentTokens.Length);
            }

            if (LegacyXlsFormulaFunctionWriterMetadata.IsVolatileFixedFunction(functionId)) {
                WriteVolatileAttribute(stream);
            }

            stream.WriteByte(0x41);
            WriteUInt16(stream, functionId);
            tokens = stream.ToArray();
            return true;
        }

        private static bool TryEncodeFunctionArgument(
            string argument,
            bool allowArea,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            argument = argument.Trim();
            if (argument.Length == 0) {
                tokens = new[] { (byte)0x16 };
                return true;
            }

            if (TryEncodeOperand(argument, allowArea, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            return TryEncode(argument, nameIndex, formulaSheetIndex, out tokens, out _);
        }

        private static bool TrySplitFunctionArguments(string argumentText, out IReadOnlyList<string>? arguments) {
            arguments = null;
            if (string.IsNullOrWhiteSpace(argumentText)) {
                return false;
            }

            var result = new List<string>();
            var current = new System.Text.StringBuilder(argumentText.Length);
            bool inStringLiteral = false;
            int parenthesisDepth = 0;
            int arrayDepth = 0;
            for (int i = 0; i < argumentText.Length; i++) {
                char ch = argumentText[i];
                if (!inStringLiteral) {
                    int quotedSheetStart = i;
                    if (TrySkipQuotedSheetName(argumentText, ref i)) {
                        current.Append(argumentText, quotedSheetStart, i - quotedSheetStart + 1);
                        continue;
                    }
                }

                if (ch == '"') {
                    current.Append(ch);
                    if (inStringLiteral && i + 1 < argumentText.Length && argumentText[i + 1] == '"') {
                        current.Append(argumentText[++i]);
                        continue;
                    }

                    inStringLiteral = !inStringLiteral;
                    continue;
                }

                if (!inStringLiteral) {
                    if (ch == '(') {
                        parenthesisDepth++;
                    } else if (ch == ')') {
                        if (parenthesisDepth == 0) {
                            return false;
                        }

                        parenthesisDepth--;
                    } else if (ch == '{') {
                        arrayDepth++;
                    } else if (ch == '}') {
                        if (arrayDepth == 0) {
                            return false;
                        }

                        arrayDepth--;
                    } else if (ch == ',' && parenthesisDepth == 0 && arrayDepth == 0) {
                        AddFunctionArgument(result, current);

                        continue;
                    }
                }

                current.Append(ch);
            }

            if (inStringLiteral || parenthesisDepth != 0 || arrayDepth != 0) {
                return false;
            }

            AddFunctionArgument(result, current);
            arguments = result;
            return true;
        }

        private static void AddFunctionArgument(List<string> result, System.Text.StringBuilder current) {
            string argument = current.ToString().Trim();
            current.Clear();
            result.Add(argument);
        }

        private static bool TryEncodeBinary(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            int operatorIndex = FindBinaryOperatorSplit(formulaText, out string? operatorText);
            if (operatorIndex < 0) {
                return false;
            }

            string left = formulaText.Substring(0, operatorIndex).Trim();
            string right = formulaText.Substring(operatorIndex + operatorText!.Length).Trim();
            if (!TryEncodeTerm(left, allowArea: false, nameIndex, formulaSheetIndex, out byte[] leftTokens)
                || !TryEncodeTerm(right, allowArea: false, nameIndex, formulaSheetIndex, out byte[] rightTokens)) {
                return false;
            }

            using var stream = new MemoryStream();
            stream.Write(leftTokens, 0, leftTokens.Length);
            stream.Write(rightTokens, 0, rightTokens.Length);
            stream.WriteByte(GetBinaryOperatorToken(operatorText));
            tokens = stream.ToArray();
            return true;
        }

        private static int FindBinaryOperatorSplit(string formulaText, out string? operatorText) {
            int found = -1;
            int foundPrecedence = int.MaxValue;
            operatorText = null;
            bool inStringLiteral = false;
            int parenthesisDepth = 0;
            int arrayDepth = 0;
            for (int i = 0; i < formulaText.Length; i++) {
                char ch = formulaText[i];
                if (!inStringLiteral && TrySkipQuotedSheetName(formulaText, ref i)) {
                    continue;
                }

                if (ch == '"') {
                    if (inStringLiteral && i + 1 < formulaText.Length && formulaText[i + 1] == '"') {
                        i++;
                        continue;
                    }

                    inStringLiteral = !inStringLiteral;
                    continue;
                }

                if (inStringLiteral) {
                    continue;
                }

                if (ch == '{') {
                    arrayDepth++;
                    continue;
                }

                if (ch == '}') {
                    arrayDepth--;
                    if (arrayDepth < 0) {
                        return -1;
                    }

                    continue;
                }

                if (ch == '(') {
                    parenthesisDepth++;
                    continue;
                }

                if (ch == ')') {
                    parenthesisDepth--;
                    if (parenthesisDepth < 0) {
                        return -1;
                    }

                    continue;
                }

                if (parenthesisDepth > 0 || arrayDepth > 0) {
                    continue;
                }

                string? currentOperator = null;
                if (i + 1 < formulaText.Length) {
                    string twoCharacterOperator = formulaText.Substring(i, 2);
                    if (twoCharacterOperator == "<=" || twoCharacterOperator == ">=" || twoCharacterOperator == "<>") {
                        currentOperator = twoCharacterOperator;
                    }
                }

                currentOperator ??= ch switch {
                    '+' => "+",
                    '-' => "-",
                    '*' => "*",
                    '/' => "/",
                    '^' => "^",
                    '&' => "&",
                    '=' => "=",
                    '<' => "<",
                    '>' => ">",
                    _ => null,
                };

                if (currentOperator == null) {
                    continue;
                }

                if (i == 0 || i == formulaText.Length - currentOperator.Length) {
                    continue;
                }

                char previous = PreviousNonWhiteSpace(formulaText, i);
                if ((currentOperator == "+" || currentOperator == "-")
                    && (previous == '\0' || previous == '+' || previous == '-' || previous == '*' || previous == '/' || previous == '^' || previous == '&' || previous == '=' || previous == '<' || previous == '>')) {
                    continue;
                }

                int currentPrecedence = GetBinaryOperatorPrecedence(currentOperator);
                if (currentPrecedence <= foundPrecedence) {
                    found = i;
                    foundPrecedence = currentPrecedence;
                    operatorText = currentOperator;
                }

                if (currentOperator.Length == 2) {
                    i++;
                }
            }

            return parenthesisDepth == 0 && arrayDepth == 0 ? found : -1;
        }

        private static char PreviousNonWhiteSpace(string text, int index) {
            for (int i = index - 1; i >= 0; i--) {
                if (!char.IsWhiteSpace(text[i])) {
                    return text[i];
                }
            }

            return '\0';
        }

        private static byte GetBinaryOperatorToken(string operatorText) {
            return operatorText switch {
                "+" => (byte)0x03,
                "-" => (byte)0x04,
                "*" => (byte)0x05,
                "/" => (byte)0x06,
                "^" => (byte)0x07,
                "&" => (byte)0x08,
                "<" => (byte)0x09,
                "<=" => (byte)0x0a,
                "=" => (byte)0x0b,
                ">=" => (byte)0x0c,
                ">" => (byte)0x0d,
                "<>" => (byte)0x0e,
                _ => throw new ArgumentOutOfRangeException(nameof(operatorText), operatorText, null),
            };
        }

        private static bool TryEncodeTerm(
            string text,
            bool allowArea,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            if (TryEncodeReferenceOperator(text, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeUnaryTerm(text, allowArea, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodePostfixPercentTerm(text, allowArea, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeParenthesizedTerm(text, allowArea, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeIfFunction(text, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeChooseFunction(text, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeVariableFunction(text, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeFixedFunction(text, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            if (TryEncodeBinary(text, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            return TryEncodeOperand(text, allowArea, nameIndex, formulaSheetIndex, out tokens);
        }

        private static int GetBinaryOperatorPrecedence(string operatorText) {
            return operatorText switch {
                "<" or "<=" or "=" or ">=" or ">" or "<>" => 0,
                "&" => 1,
                "+" or "-" => 2,
                "*" or "/" => 3,
                "^" => 4,
                _ => throw new ArgumentOutOfRangeException(nameof(operatorText), operatorText, null),
            };
        }

        private static bool TryEncodeParenthesized(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            if (!TryGetParenthesizedInnerText(formulaText, out string? innerText)) {
                return false;
            }

            if (!TryEncode(innerText!, nameIndex, formulaSheetIndex, out byte[] innerTokens, out _)) {
                return false;
            }

            tokens = AppendToken(innerTokens, 0x15);
            return true;
        }

        private static bool TryEncodeParenthesizedTerm(
            string text,
            bool allowArea,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            if (!TryGetParenthesizedInnerText(text, out string? innerText)) {
                return false;
            }

            if (!TryEncode(innerText!, nameIndex, formulaSheetIndex, out byte[] innerTokens, out _)) {
                return false;
            }

            tokens = AppendToken(innerTokens, 0x15);
            return true;
        }

        private static bool TryGetParenthesizedInnerText(string text, out string? innerText) {
            innerText = null;
            text = text.Trim();
            if (text.Length < 3 || text[0] != '(' || text[text.Length - 1] != ')') {
                return false;
            }

            bool inStringLiteral = false;
            int parenthesisDepth = 0;
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (!inStringLiteral && TrySkipQuotedSheetName(text, ref i)) {
                    continue;
                }

                if (ch == '"') {
                    if (inStringLiteral && i + 1 < text.Length && text[i + 1] == '"') {
                        i++;
                        continue;
                    }

                    inStringLiteral = !inStringLiteral;
                    continue;
                }

                if (inStringLiteral) {
                    continue;
                }

                if (ch == '(') {
                    parenthesisDepth++;
                    continue;
                }

                if (ch == ')') {
                    parenthesisDepth--;
                    if (parenthesisDepth < 0 || (parenthesisDepth == 0 && i != text.Length - 1)) {
                        return false;
                    }
                }
            }

            if (inStringLiteral || parenthesisDepth != 0) {
                return false;
            }

            innerText = text.Substring(1, text.Length - 2).Trim();
            return innerText.Length > 0;
        }

        private static bool TryEncodeUnary(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            if (!TryGetUnaryOperand(formulaText, out char unaryOperator, out string? operandText)) {
                return false;
            }

            if (!TryEncode(operandText!, nameIndex, formulaSheetIndex, out byte[] operandTokens, out _)) {
                return false;
            }

            tokens = AppendToken(operandTokens, GetUnaryOperatorToken(unaryOperator));
            return true;
        }

        private static bool TryEncodeUnaryTerm(
            string text,
            bool allowArea,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            if (!TryGetUnaryOperand(text, out char unaryOperator, out string? operandText)) {
                return false;
            }

            if (!TryEncodeTerm(operandText!, allowArea, nameIndex, formulaSheetIndex, out byte[] operandTokens)) {
                return false;
            }

            tokens = AppendToken(operandTokens, GetUnaryOperatorToken(unaryOperator));
            return true;
        }

        private static bool TryGetUnaryOperand(string text, out char unaryOperator, out string? operandText) {
            unaryOperator = '\0';
            operandText = null;
            text = text.Trim();
            if (text.Length < 2 || (text[0] != '+' && text[0] != '-')) {
                return false;
            }

            unaryOperator = text[0];
            operandText = text.Substring(1).Trim();
            return operandText.Length > 0;
        }

        private static byte GetUnaryOperatorToken(char unaryOperator) {
            return unaryOperator switch {
                '+' => 0x12,
                '-' => 0x13,
                _ => throw new ArgumentOutOfRangeException(nameof(unaryOperator), unaryOperator, null),
            };
        }

        private static bool TryEncodePostfixPercent(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            if (!TryGetPostfixPercentOperand(formulaText, out string? operandText)) {
                return false;
            }

            if (!TryEncode(operandText!, nameIndex, formulaSheetIndex, out byte[] operandTokens, out _)) {
                return false;
            }

            tokens = AppendToken(operandTokens, 0x14);
            return true;
        }

        private static bool TryEncodePostfixPercentTerm(
            string text,
            bool allowArea,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            if (!TryGetPostfixPercentOperand(text, out string? operandText)) {
                return false;
            }

            if (!TryEncodeTerm(operandText!, allowArea, nameIndex, formulaSheetIndex, out byte[] operandTokens)) {
                return false;
            }

            tokens = AppendToken(operandTokens, 0x14);
            return true;
        }

        private static bool TryGetPostfixPercentOperand(string text, out string? operandText) {
            operandText = null;
            text = text.Trim();
            if (text.Length < 2 || text[text.Length - 1] != '%') {
                return false;
            }

            operandText = text.Substring(0, text.Length - 1).Trim();
            return operandText.Length > 0;
        }

        private static bool TryEncodeOperand(
            string text,
            bool allowArea,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();

            if (TryEncodeSheetQualifiedReference(text, allowArea, nameIndex, out tokens)) {
                return true;
            }

            if (allowArea && TryParseAreaReference(text, out FormulaReference first, out FormulaReference last)) {
                tokens = BuildAreaReferenceToken(first, last);
                return true;
            }

            if (TryParseCellReference(text, out FormulaReference reference)) {
                tokens = BuildCellReferenceToken(reference);
                return true;
            }

            if (TryParseStringLiteral(text, out string? stringValue)) {
                tokens = BuildStringToken(stringValue!);
                return true;
            }

            if (text.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || text.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) {
                tokens = new[] { (byte)0x1d, text.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ? (byte)1 : (byte)0 };
                return true;
            }

            if (LegacyXlsErrorValue.TryGetCode(text, out byte errorCode)) {
                tokens = new[] { (byte)0x1c, errorCode };
                return true;
            }

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                tokens = BuildNumberToken(number);
                return true;
            }

            if (TryEncodeExternalDefinedName(text, nameIndex, out tokens)) {
                return true;
            }

            if (TryEncodeDefinedName(text, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            return false;
        }

        internal static bool TryEncodeFormulaRecord(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens,
            out byte[] extraData,
            out string? reason) {
            extraData = Array.Empty<byte>();

            string normalized = NormalizeFormula(formulaText);
            if (normalized.IndexOf('{') >= 0
                && TryEncodeArrayAwareFormula(normalized, nameIndex, formulaSheetIndex, out tokens, out extraData)) {
                reason = null;
                return true;
            }

            return TryEncode(formulaText, nameIndex, formulaSheetIndex, out tokens, out reason);
        }

        private static bool TryEncodeDefinedName(
            string text,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            text = text.Trim();
            if (text.Length == 0 || text.IndexOf('!') >= 0 || text.IndexOf('[') >= 0) {
                return false;
            }

            if (!nameIndex.TryGetNameIndex(text, formulaSheetIndex, out uint oneBasedNameIndex)) {
                return false;
            }

            tokens = BuildDefinedNameToken(oneBasedNameIndex);
            return true;
        }

        private static bool TryEncodeExternalDefinedName(
            string text,
            LegacyXlsFormulaNameIndex nameIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            text = text.Trim();
            if (TryParseSheetQualifiedExternalDefinedName(text, out string? scopedTarget, out string? scopedSheetName, out string? scopedName)) {
                if (!nameIndex.TryGetExternalNameIndex(scopedTarget!, scopedSheetName, scopedName!, out ushort scopedExternSheetIndex, out uint scopedOneBasedNameIndex)) {
                    return false;
                }

                tokens = BuildExternalDefinedNameToken(scopedExternSheetIndex, scopedOneBasedNameIndex);
                return true;
            }

            if (text.Length < 4 || text[0] != '[') {
                return false;
            }

            int close = text.IndexOf(']');
            if (close <= 1 || close >= text.Length - 1) {
                return false;
            }

            string target = text.Substring(1, close - 1).Trim();
            string name = text.Substring(close + 1).Trim();
            if (target.Length == 0
                || name.Length == 0
                || name.IndexOf('!') >= 0
                || name.IndexOf(':') >= 0
                || name.IndexOf('[') >= 0
                || name.IndexOf(']') >= 0) {
                return false;
            }

            if (!nameIndex.TryGetExternalNameIndex(target, name, out ushort externSheetIndex, out uint oneBasedNameIndex)) {
                return false;
            }

            tokens = BuildExternalDefinedNameToken(externSheetIndex, oneBasedNameIndex);
            return true;
        }

        private static bool TryParseSheetQualifiedExternalDefinedName(string text, out string? target, out string? sheetName, out string? name) {
            target = null;
            sheetName = null;
            name = null;
            int bang = text.IndexOf('!');
            if (bang <= 0 || bang >= text.Length - 1 || text.IndexOf('!', bang + 1) >= 0) {
                return false;
            }

            string sheetToken = text.Substring(0, bang).Trim();
            string nameToken = text.Substring(bang + 1).Trim();
            if (!IsValidExternalNameOperand(nameToken)) {
                return false;
            }

            string unquotedSheet = UnquoteSheetToken(sheetToken);
            if (!LegacyXlsExternSheetTable.TryParseExternalSheetName(unquotedSheet, out target, out sheetName)) {
                return false;
            }

            name = nameToken;
            return true;
        }

        private static bool IsValidExternalNameOperand(string name) {
            return name.Length > 0
                && name.Length <= byte.MaxValue
                && name.IndexOf('!') < 0
                && name.IndexOf(':') < 0
                && name.IndexOf('[') < 0
                && name.IndexOf(']') < 0
                && name.IndexOf('$') < 0;
        }

        private static string UnquoteSheetToken(string sheetToken) {
            string trimmed = sheetToken.Trim();
            if (trimmed.Length >= 2 && trimmed[0] == '\'' && trimmed[trimmed.Length - 1] == '\'') {
                return trimmed.Substring(1, trimmed.Length - 2).Replace("''", "'");
            }

            return trimmed;
        }

        private static bool TryParseAreaReference(string text, out FormulaReference first, out FormulaReference last) {
            first = default;
            last = default;

            int separator = text.IndexOf(':');
            if (separator <= 0 || separator >= text.Length - 1 || text.IndexOf(':', separator + 1) >= 0) {
                return false;
            }

            string firstText = text.Substring(0, separator).Trim();
            string lastText = text.Substring(separator + 1).Trim();
            if (TryParseCellReference(firstText, out first)
                && TryParseCellReference(lastText, out last)) {
                return first.Row <= last.Row && first.Column <= last.Column;
            }

            if (TryParseColumnReference(firstText, out ushort firstColumn, out bool firstColumnRelative)
                && TryParseColumnReference(lastText, out ushort lastColumn, out bool lastColumnRelative)
                && firstColumn <= lastColumn) {
                first = new FormulaReference(0, firstColumn, rowRelative: true, firstColumnRelative);
                last = new FormulaReference(65535, lastColumn, rowRelative: true, lastColumnRelative);
                return true;
            }

            if (TryParseRowReference(firstText, out ushort firstRow, out bool firstRowRelative)
                && TryParseRowReference(lastText, out ushort lastRow, out bool lastRowRelative)
                && firstRow <= lastRow) {
                first = new FormulaReference(firstRow, 0, firstRowRelative, columnRelative: true);
                last = new FormulaReference(lastRow, 255, lastRowRelative, columnRelative: true);
                return true;
            }

            return false;
        }

        private static bool TryParseColumnReference(string text, out ushort zeroBasedColumn, out bool columnRelative) {
            zeroBasedColumn = 0;
            columnRelative = true;
            text = text.Trim();
            if (text.Length == 0) {
                return false;
            }

            int index = 0;
            if (text[index] == '$') {
                columnRelative = false;
                index++;
            }

            int column = 0;
            int columnStart = index;
            while (index < text.Length && IsAsciiLetter(text[index])) {
                char upper = char.ToUpperInvariant(text[index]);
                column = column * 26 + (upper - 'A' + 1);
                if (column > 256) {
                    return false;
                }

                index++;
            }

            if (index != text.Length || index == columnStart || column < 1 || column > 256) {
                return false;
            }

            zeroBasedColumn = checked((ushort)(column - 1));
            return true;
        }

        private static bool TryParseRowReference(string text, out ushort zeroBasedRow, out bool rowRelative) {
            zeroBasedRow = 0;
            rowRelative = true;
            text = text.Trim();
            if (text.Length == 0) {
                return false;
            }

            int index = 0;
            if (text[index] == '$') {
                rowRelative = false;
                index++;
            }

            int row = 0;
            int rowStart = index;
            while (index < text.Length && char.IsDigit(text[index])) {
                row = checked(row * 10 + (text[index] - '0'));
                if (row > 65536) {
                    return false;
                }

                index++;
            }

            if (index != text.Length || index == rowStart || row < 1 || row > 65536) {
                return false;
            }

            zeroBasedRow = checked((ushort)(row - 1));
            return true;
        }

        private static bool TryParseStringLiteral(string text, out string? value) {
            value = null;
            text = text.Trim();
            if (text.Length < 2 || text[0] != '"' || text[text.Length - 1] != '"') {
                return false;
            }

            var builder = new System.Text.StringBuilder();
            for (int i = 1; i < text.Length - 1; i++) {
                char ch = text[i];
                if (ch == '"') {
                    if (i + 1 < text.Length - 1 && text[i + 1] == '"') {
                        builder.Append('"');
                        i++;
                        continue;
                    }

                    return false;
                }

                builder.Append(ch);
            }

            value = builder.ToString();
            return value.Length <= 255;
        }

        private static bool TryParseCellReference(string text, out FormulaReference reference) {
            reference = default;
            text = text.Trim();
            if (text.Length == 0 || text.IndexOf('!') >= 0 || text.IndexOf('[') >= 0) {
                return false;
            }

            int index = 0;
            bool columnRelative = true;
            if (text[index] == '$') {
                columnRelative = false;
                index++;
            }

            int column = 0;
            int columnStart = index;
            while (index < text.Length && IsAsciiLetter(text[index])) {
                char upper = char.ToUpperInvariant(text[index]);
                column = column * 26 + (upper - 'A' + 1);
                if (column > 256) {
                    return false;
                }

                index++;
            }

            if (index == columnStart || column < 1 || column > 256) {
                return false;
            }

            bool rowRelative = true;
            if (index < text.Length && text[index] == '$') {
                rowRelative = false;
                index++;
            }

            int row = 0;
            int rowStart = index;
            while (index < text.Length && char.IsDigit(text[index])) {
                row = checked(row * 10 + (text[index] - '0'));
                index++;
            }

            if (index != text.Length || index == rowStart || row < 1 || row > 65536) {
                return false;
            }

            reference = new FormulaReference(
                checked((ushort)(row - 1)),
                checked((ushort)(column - 1)),
                rowRelative,
                columnRelative);
            return true;
        }

        private static bool IsAsciiLetter(char ch) {
            return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');
        }

        private static byte[] BuildCellReferenceToken(FormulaReference reference) {
            byte[] token = new byte[5];
            token[0] = 0x44;
            WriteUInt16(token, 1, reference.Row);
            WriteUInt16(token, 3, reference.ColumnBits);
            return token;
        }

        private static byte[] BuildAreaReferenceToken(FormulaReference first, FormulaReference last) {
            byte[] token = new byte[9];
            token[0] = 0x45;
            WriteUInt16(token, 1, first.Row);
            WriteUInt16(token, 3, last.Row);
            WriteUInt16(token, 5, first.ColumnBits);
            WriteUInt16(token, 7, last.ColumnBits);
            return token;
        }

        private static byte[] BuildNumberToken(double number) {
            if (number >= 0 && number <= ushort.MaxValue && Math.Truncate(number) == number) {
                byte[] integerToken = new byte[3];
                integerToken[0] = 0x1e;
                WriteUInt16(integerToken, 1, checked((ushort)number));
                return integerToken;
            }

            byte[] numberToken = new byte[9];
            numberToken[0] = 0x1f;
            byte[] numberBytes = BitConverter.GetBytes(number);
            Buffer.BlockCopy(numberBytes, 0, numberToken, 1, numberBytes.Length);
            return numberToken;
        }

        private static byte[] BuildStringToken(string value) {
            byte[] textBytes = EncodeShortUnicodeString(value, out byte flags);
            byte[] token = new byte[checked(3 + textBytes.Length)];
            token[0] = 0x17;
            token[1] = checked((byte)value.Length);
            token[2] = flags;
            Buffer.BlockCopy(textBytes, 0, token, 3, textBytes.Length);
            return token;
        }

        private static byte[] BuildDefinedNameToken(uint oneBasedNameIndex) {
            byte[] token = new byte[5];
            token[0] = 0x43;
            WriteUInt32(token, 1, oneBasedNameIndex);
            return token;
        }

        private static byte[] BuildExternalDefinedNameToken(ushort externSheetIndex, uint oneBasedNameIndex) {
            byte[] token = new byte[7];
            token[0] = 0x39;
            WriteUInt16(token, 1, externSheetIndex);
            WriteUInt32(token, 3, oneBasedNameIndex);
            return token;
        }

        private static byte[] AppendToken(byte[] tokens, byte token) {
            byte[] result = new byte[checked(tokens.Length + 1)];
            Buffer.BlockCopy(tokens, 0, result, 0, tokens.Length);
            result[result.Length - 1] = token;
            return result;
        }

        private static byte[] EncodeShortUnicodeString(string text, out byte flags) {
            if (CanUseCompressedString(text)) {
                flags = 0;
                return System.Text.Encoding.ASCII.GetBytes(text);
            }

            flags = 1;
            return System.Text.Encoding.Unicode.GetBytes(text);
        }

        private static bool CanUseCompressedString(string text) {
            for (int i = 0; i < text.Length; i++) {
                if (text[i] > 0x7f) {
                    return false;
                }
            }

            return true;
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private static void WriteVolatileAttribute(Stream stream) {
            stream.WriteByte(0x19);
            stream.WriteByte(0x01);
            WriteUInt16(stream, 0);
        }

        private static void WriteAttribute(Stream stream, byte attribute, ushort value) {
            stream.WriteByte(0x19);
            stream.WriteByte(attribute);
            WriteUInt16(stream, value);
        }

        private static void WriteUInt16(byte[] buffer, int offset, ushort value) {
            buffer[offset] = (byte)(value & 0xff);
            buffer[offset + 1] = (byte)((value >> 8) & 0xff);
        }

        private static ushort ReadUInt16(byte[] buffer, int offset) {
            return unchecked((ushort)(buffer[offset] | (buffer[offset + 1] << 8)));
        }

        private static void WriteUInt32(byte[] buffer, int offset, uint value) {
            buffer[offset] = (byte)(value & 0xff);
            buffer[offset + 1] = (byte)((value >> 8) & 0xff);
            buffer[offset + 2] = (byte)((value >> 16) & 0xff);
            buffer[offset + 3] = (byte)((value >> 24) & 0xff);
        }

        private readonly struct FormulaReference {
            internal FormulaReference(ushort row, ushort column, bool rowRelative, bool columnRelative) {
                Row = row;
                Column = column;
                RowRelative = rowRelative;
                ColumnRelative = columnRelative;
            }

            internal ushort Row { get; }

            internal ushort Column { get; }

            private bool RowRelative { get; }

            private bool ColumnRelative { get; }

            internal ushort ColumnBits {
                get {
                    ushort bits = Column;
                    if (ColumnRelative) {
                        bits |= 0x4000;
                    }

                    if (RowRelative) {
                        bits |= 0x8000;
                    }

                    return bits;
                }
            }
        }
    }
}
