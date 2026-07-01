using System.Globalization;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsFormulaEncoder {
        private static bool TryEncodeArrayAwareFormula(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens,
            out byte[] extraData) {
            formulaText = formulaText.Trim();
            if (TryEncodeArrayConstant(formulaText, out tokens, out extraData)) {
                return true;
            }

            if (TryEncodeArrayAwareParenthesized(formulaText, nameIndex, formulaSheetIndex, out tokens, out extraData)) {
                return true;
            }

            if (TryEncodeArrayAwareIfFunction(formulaText, nameIndex, formulaSheetIndex, out tokens, out extraData)) {
                return true;
            }

            if (TryEncodeArrayAwareChooseFunction(formulaText, nameIndex, formulaSheetIndex, out tokens, out extraData)) {
                return true;
            }

            if (TryEncodeArrayAwareFunction(formulaText, nameIndex, formulaSheetIndex, out tokens, out extraData)) {
                return true;
            }

            if (TryEncodeArrayAwareBinary(formulaText, nameIndex, formulaSheetIndex, out tokens, out extraData)) {
                return true;
            }

            if (TryEncodeArrayAwareUnary(formulaText, nameIndex, formulaSheetIndex, out tokens, out extraData)) {
                return true;
            }

            if (TryEncodeArrayAwarePostfixPercent(formulaText, nameIndex, formulaSheetIndex, out tokens, out extraData)) {
                return true;
            }

            tokens = Array.Empty<byte>();
            extraData = Array.Empty<byte>();
            return false;
        }

        private static bool TryEncodeArrayAwareParenthesized(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens,
            out byte[] extraData) {
            tokens = Array.Empty<byte>();
            extraData = Array.Empty<byte>();
            if (!TryGetParenthesizedInnerText(formulaText, out string? innerText)
                || !TryEncodeArrayAwareTerm(innerText!, nameIndex, formulaSheetIndex, out byte[] innerTokens, out extraData)) {
                return false;
            }

            tokens = AppendToken(innerTokens, 0x15);
            return true;
        }

        private static bool TryEncodeArrayAwareIfFunction(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens,
            out byte[] extraData) {
            tokens = Array.Empty<byte>();
            extraData = Array.Empty<byte>();
            const string functionName = "IF";
            if (!formulaText.StartsWith(functionName + "(", StringComparison.OrdinalIgnoreCase)
                || !formulaText.EndsWith(")", StringComparison.Ordinal)) {
                return false;
            }

            string argumentText = formulaText.Substring(functionName.Length + 1, formulaText.Length - functionName.Length - 2).Trim();
            if (!TrySplitFunctionArguments(argumentText, out IReadOnlyList<string>? arguments)
                || (arguments!.Count != 2 && arguments.Count != 3)
                || string.IsNullOrWhiteSpace(arguments[0])) {
                return false;
            }

            if (!TryEncodeArrayAwareArgument(arguments[0], nameIndex, formulaSheetIndex, out byte[] conditionTokens, out byte[] conditionExtra)
                || !TryEncodeArrayAwareArgument(arguments[1], nameIndex, formulaSheetIndex, out byte[] trueTokens, out byte[] trueExtra)) {
                return false;
            }

            using var tokenStream = new MemoryStream();
            tokenStream.Write(conditionTokens, 0, conditionTokens.Length);
            WriteAttribute(tokenStream, 0x02, checked((ushort)(trueTokens.Length + 3)));
            tokenStream.Write(trueTokens, 0, trueTokens.Length);

            byte[] falseExtra = Array.Empty<byte>();
            if (arguments.Count == 3) {
                if (!TryEncodeArrayAwareArgument(arguments[2], nameIndex, formulaSheetIndex, out byte[] falseTokens, out falseExtra)) {
                    return false;
                }

                WriteAttribute(tokenStream, 0x08, checked((ushort)falseTokens.Length));
                tokenStream.Write(falseTokens, 0, falseTokens.Length);
            }

            tokenStream.WriteByte(0x42);
            tokenStream.WriteByte(checked((byte)arguments.Count));
            WriteUInt16(tokenStream, 0x0001);
            tokens = tokenStream.ToArray();
            extraData = CombineExtraData(CombineExtraData(conditionExtra, trueExtra), falseExtra);
            return true;
        }

        private static bool TryEncodeArrayAwareChooseFunction(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens,
            out byte[] extraData) {
            tokens = Array.Empty<byte>();
            extraData = Array.Empty<byte>();
            const string functionName = "CHOOSE";
            if (!formulaText.StartsWith(functionName + "(", StringComparison.OrdinalIgnoreCase)
                || !formulaText.EndsWith(")", StringComparison.Ordinal)) {
                return false;
            }

            string argumentText = formulaText.Substring(functionName.Length + 1, formulaText.Length - functionName.Length - 2).Trim();
            if (!TrySplitFunctionArguments(argumentText, out IReadOnlyList<string>? arguments)
                || arguments!.Count < 2
                || arguments.Count > 30
                || string.IsNullOrWhiteSpace(arguments[0])) {
                return false;
            }

            if (!TryEncodeArrayAwareArgument(arguments[0], nameIndex, formulaSheetIndex, out byte[] indexTokens, out byte[] indexExtra)) {
                return false;
            }

            int choiceCount = arguments.Count - 1;
            var choiceTokens = new byte[choiceCount][];
            var choiceExtras = new byte[choiceCount][];
            for (int i = 0; i < choiceCount; i++) {
                if (!TryEncodeArrayAwareArgument(arguments[i + 1], nameIndex, formulaSheetIndex, out choiceTokens[i], out choiceExtras[i])) {
                    return false;
                }
            }

            using var tokenStream = new MemoryStream();
            tokenStream.Write(indexTokens, 0, indexTokens.Length);
            WriteChooseAttribute(tokenStream, choiceTokens);
            for (int i = 0; i < choiceTokens.Length; i++) {
                byte[] optionTokens = choiceTokens[i];
                tokenStream.Write(optionTokens, 0, optionTokens.Length);
                WriteAttribute(tokenStream, 0x08, GetChooseGotoOffset(choiceTokens, i));
            }

            tokenStream.WriteByte(0x42);
            tokenStream.WriteByte(checked((byte)arguments.Count));
            WriteUInt16(tokenStream, 0x0064);
            tokens = tokenStream.ToArray();

            extraData = indexExtra;
            for (int i = 0; i < choiceExtras.Length; i++) {
                extraData = CombineExtraData(extraData, choiceExtras[i]);
            }

            return true;
        }

        private static bool TryEncodeArrayAwareFunction(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens,
            out byte[] extraData) {
            tokens = Array.Empty<byte>();
            extraData = Array.Empty<byte>();
            if (!formulaText.EndsWith(")", StringComparison.Ordinal)) {
                return false;
            }

            bool isVariable = LegacyXlsFormulaFunctionWriterMetadata.TryGetVariableFunction(formulaText, out ushort functionId, out int argumentStart);
            int fixedParameterCount = 0;
            if (!isVariable
                && !LegacyXlsFormulaFunctionWriterMetadata.TryGetFixedFunction(formulaText, out functionId, out fixedParameterCount, out argumentStart)) {
                return false;
            }

            string argumentText = formulaText.Substring(argumentStart, formulaText.Length - argumentStart - 1).Trim();
            IReadOnlyList<string> arguments;
            if (!isVariable && fixedParameterCount == 0 && argumentText.Length == 0) {
                arguments = Array.Empty<string>();
            } else if (!isVariable
                && fixedParameterCount == 1
                && argumentText.Length == 0
                && LegacyXlsFormulaFunctionWriterMetadata.AllowsMissingReferenceArgument(functionId)) {
                arguments = new[] { string.Empty };
            } else if (!TrySplitFunctionArguments(argumentText, out IReadOnlyList<string>? parsedArguments)) {
                return false;
            } else {
                arguments = parsedArguments!;
            }

            if (isVariable) {
                if (arguments.Count == 0
                    || !LegacyXlsFormulaFunctionWriterMetadata.IsSupportedVariableFunctionArgumentCount(functionId, arguments.Count)) {
                    return false;
                }
            } else if (arguments.Count != fixedParameterCount) {
                return false;
            }

            using var tokenStream = new MemoryStream();
            using var extraStream = new MemoryStream();
            foreach (string argument in arguments) {
                if (!TryEncodeArrayAwareArgument(argument, nameIndex, formulaSheetIndex, out byte[] argumentTokens, out byte[] argumentExtra)) {
                    return false;
                }

                tokenStream.Write(argumentTokens, 0, argumentTokens.Length);
                extraStream.Write(argumentExtra, 0, argumentExtra.Length);
            }

            if (!isVariable && LegacyXlsFormulaFunctionWriterMetadata.IsVolatileFixedFunction(functionId)) {
                WriteVolatileAttribute(tokenStream);
            }

            tokenStream.WriteByte(isVariable ? (byte)0x42 : (byte)0x41);
            if (isVariable) {
                tokenStream.WriteByte(checked((byte)arguments.Count));
            }

            WriteUInt16(tokenStream, functionId);
            tokens = tokenStream.ToArray();
            extraData = extraStream.ToArray();
            return true;
        }

        private static bool TryEncodeArrayAwareBinary(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens,
            out byte[] extraData) {
            tokens = Array.Empty<byte>();
            extraData = Array.Empty<byte>();
            int operatorIndex = FindBinaryOperatorSplit(formulaText, out string? operatorText);
            if (operatorIndex < 0) {
                return false;
            }

            string left = formulaText.Substring(0, operatorIndex).Trim();
            string right = formulaText.Substring(operatorIndex + operatorText!.Length).Trim();
            if (!TryEncodeArrayAwareTerm(left, nameIndex, formulaSheetIndex, out byte[] leftTokens, out byte[] leftExtra)
                || !TryEncodeArrayAwareTerm(right, nameIndex, formulaSheetIndex, out byte[] rightTokens, out byte[] rightExtra)) {
                return false;
            }

            using var tokenStream = new MemoryStream();
            tokenStream.Write(leftTokens, 0, leftTokens.Length);
            tokenStream.Write(rightTokens, 0, rightTokens.Length);
            tokenStream.WriteByte(GetBinaryOperatorToken(operatorText));
            tokens = tokenStream.ToArray();
            extraData = CombineExtraData(leftExtra, rightExtra);
            return true;
        }

        private static bool TryEncodeArrayAwareUnary(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens,
            out byte[] extraData) {
            tokens = Array.Empty<byte>();
            extraData = Array.Empty<byte>();
            if (!TryGetUnaryOperand(formulaText, out char unaryOperator, out string? operandText)
                || !TryEncodeArrayAwareTerm(operandText!, nameIndex, formulaSheetIndex, out byte[] operandTokens, out extraData)) {
                return false;
            }

            tokens = AppendToken(operandTokens, GetUnaryOperatorToken(unaryOperator));
            return true;
        }

        private static bool TryEncodeArrayAwarePostfixPercent(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens,
            out byte[] extraData) {
            tokens = Array.Empty<byte>();
            extraData = Array.Empty<byte>();
            if (!TryGetPostfixPercentOperand(formulaText, out string? operandText)
                || !TryEncodeArrayAwareTerm(operandText!, nameIndex, formulaSheetIndex, out byte[] operandTokens, out extraData)) {
                return false;
            }

            tokens = AppendToken(operandTokens, 0x14);
            return true;
        }

        private static bool TryEncodeArrayAwareArgument(
            string argument,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens,
            out byte[] extraData) {
            argument = argument.Trim();
            if (argument.IndexOf('{') >= 0) {
                return TryEncodeArrayAwareFormula(argument, nameIndex, formulaSheetIndex, out tokens, out extraData);
            }

            extraData = Array.Empty<byte>();
            return TryEncodeFunctionArgument(argument, allowArea: true, nameIndex, formulaSheetIndex, out tokens);
        }

        private static bool TryEncodeArrayAwareTerm(
            string text,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens,
            out byte[] extraData) {
            text = text.Trim();
            if (text.IndexOf('{') >= 0) {
                return TryEncodeArrayAwareFormula(text, nameIndex, formulaSheetIndex, out tokens, out extraData);
            }

            extraData = Array.Empty<byte>();
            return TryEncodeTerm(text, allowArea: true, nameIndex, formulaSheetIndex, out tokens);
        }

        private static bool TryEncodeArrayConstant(string text, out byte[] tokens, out byte[] extraData) {
            tokens = Array.Empty<byte>();
            extraData = Array.Empty<byte>();
            text = text.Trim();
            if (text.Length < 3 || text[0] != '{' || text[text.Length - 1] != '}') {
                return false;
            }

            string body = text.Substring(1, text.Length - 2);
            if (!TrySplitArrayConstantParts(body, ';', out IReadOnlyList<string>? rowTexts)
                || rowTexts!.Count == 0
                || rowTexts.Count > ushort.MaxValue + 1) {
                return false;
            }

            var rows = new List<IReadOnlyList<string>>(rowTexts.Count);
            int columnCount = -1;
            foreach (string rowText in rowTexts) {
                if (!TrySplitArrayConstantParts(rowText, ',', out IReadOnlyList<string>? values)
                    || values!.Count == 0
                    || values.Count > 256) {
                    return false;
                }

                if (columnCount < 0) {
                    columnCount = values.Count;
                } else if (values.Count != columnCount) {
                    return false;
                }

                rows.Add(values);
            }

            using var extraStream = new MemoryStream();
            extraStream.WriteByte(checked((byte)(columnCount - 1)));
            WriteUInt16(extraStream, checked((ushort)(rows.Count - 1)));
            foreach (IReadOnlyList<string> row in rows) {
                foreach (string rawValue in row) {
                    if (!TryWriteArrayConstantValue(extraStream, rawValue.Trim())) {
                        return false;
                    }
                }
            }

            tokens = new byte[8];
            tokens[0] = 0x60;
            extraData = extraStream.ToArray();
            return true;
        }

        private static bool TrySplitArrayConstantParts(string text, char separator, out IReadOnlyList<string>? parts) {
            parts = null;
            var result = new List<string>();
            var current = new System.Text.StringBuilder(text.Length);
            bool inStringLiteral = false;
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (ch == '"') {
                    current.Append(ch);
                    if (inStringLiteral && i + 1 < text.Length && text[i + 1] == '"') {
                        current.Append(text[++i]);
                        continue;
                    }

                    inStringLiteral = !inStringLiteral;
                    continue;
                }

                if (!inStringLiteral && ch == separator) {
                    result.Add(current.ToString());
                    current.Clear();
                    continue;
                }

                current.Append(ch);
            }

            if (inStringLiteral) {
                return false;
            }

            result.Add(current.ToString());
            parts = result;
            return true;
        }

        private static bool TryWriteArrayConstantValue(Stream stream, string text) {
            if (text.Length == 0) {
                stream.WriteByte(0x00);
                WritePadding(stream, 8);
                return true;
            }

            if (TryParseStringLiteral(text, out string? stringValue)) {
                if (stringValue!.Length > ushort.MaxValue) {
                    return false;
                }

                stream.WriteByte(0x02);
                WriteUInt16(stream, checked((ushort)stringValue.Length));
                byte[] stringBytes = EncodeShortUnicodeString(stringValue, out byte flags);
                stream.WriteByte(flags);
                stream.Write(stringBytes, 0, stringBytes.Length);
                return true;
            }

            if (text.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || text.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) {
                stream.WriteByte(0x04);
                stream.WriteByte(text.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ? (byte)1 : (byte)0);
                WritePadding(stream, 7);
                return true;
            }

            if (LegacyXlsErrorValue.TryGetCode(text, out byte errorCode)) {
                stream.WriteByte(0x10);
                stream.WriteByte(errorCode);
                WritePadding(stream, 7);
                return true;
            }

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                stream.WriteByte(0x01);
                byte[] numberBytes = BitConverter.GetBytes(number);
                stream.Write(numberBytes, 0, numberBytes.Length);
                return true;
            }

            return false;
        }

        private static byte[] CombineExtraData(byte[] first, byte[] second) {
            if (first.Length == 0) {
                return second;
            }

            if (second.Length == 0) {
                return first;
            }

            byte[] result = new byte[checked(first.Length + second.Length)];
            Buffer.BlockCopy(first, 0, result, 0, first.Length);
            Buffer.BlockCopy(second, 0, result, first.Length, second.Length);
            return result;
        }

        private static void WritePadding(Stream stream, int count) {
            for (int i = 0; i < count; i++) {
                stream.WriteByte(0);
            }
        }
    }
}
