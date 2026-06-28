namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsFormulaEncoder {
        private static bool TryEncodeIfFunction(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
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

            if (!TryEncodeFunctionArgument(arguments[0], allowArea: false, nameIndex, formulaSheetIndex, out byte[] conditionTokens)
                || !TryEncodeFunctionArgument(arguments[1], allowArea: false, nameIndex, formulaSheetIndex, out byte[] trueTokens)) {
                return false;
            }

            using var stream = new MemoryStream();
            stream.Write(conditionTokens, 0, conditionTokens.Length);
            WriteAttribute(stream, 0x02, checked((ushort)(trueTokens.Length + 3)));
            stream.Write(trueTokens, 0, trueTokens.Length);
            if (arguments.Count == 3) {
                if (!TryEncodeFunctionArgument(arguments[2], allowArea: false, nameIndex, formulaSheetIndex, out byte[] falseTokens)) {
                    return false;
                }

                WriteAttribute(stream, 0x08, checked((ushort)falseTokens.Length));
                stream.Write(falseTokens, 0, falseTokens.Length);
            }

            stream.WriteByte(0x42);
            stream.WriteByte(checked((byte)arguments.Count));
            WriteUInt16(stream, 0x0001);
            tokens = stream.ToArray();
            return true;
        }

        private static bool TryEncodeChooseFunction(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
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

            if (!TryEncodeFunctionArgument(arguments[0], allowArea: false, nameIndex, formulaSheetIndex, out byte[] indexTokens)) {
                return false;
            }

            int choiceCount = arguments.Count - 1;
            var choiceTokens = new byte[choiceCount][];
            for (int i = 0; i < choiceCount; i++) {
                if (!TryEncodeFunctionArgument(arguments[i + 1], allowArea: false, nameIndex, formulaSheetIndex, out choiceTokens[i])) {
                    return false;
                }
            }

            using var stream = new MemoryStream();
            stream.Write(indexTokens, 0, indexTokens.Length);
            WriteChooseAttribute(stream, choiceTokens);
            for (int i = 0; i < choiceTokens.Length; i++) {
                byte[] optionTokens = choiceTokens[i];
                stream.Write(optionTokens, 0, optionTokens.Length);
                WriteAttribute(stream, 0x08, GetChooseGotoOffset(choiceTokens, i));
            }

            stream.WriteByte(0x42);
            stream.WriteByte(checked((byte)arguments.Count));
            WriteUInt16(stream, 0x0064);
            tokens = stream.ToArray();
            return true;
        }

        private static void WriteChooseAttribute(Stream stream, IReadOnlyList<byte[]> choiceTokens) {
            stream.WriteByte(0x19);
            stream.WriteByte(0x04);
            WriteUInt16(stream, checked((ushort)choiceTokens.Count));

            WriteUInt16(stream, checked((ushort)((choiceTokens.Count + 1) * 2)));
            int cumulativeChoiceBytes = 0;
            for (int i = 0; i < choiceTokens.Count; i++) {
                cumulativeChoiceBytes = checked(cumulativeChoiceBytes + choiceTokens[i].Length + 4);
                WriteUInt16(stream, checked((ushort)cumulativeChoiceBytes));
            }
        }

        private static ushort GetChooseGotoOffset(IReadOnlyList<byte[]> choiceTokens, int choiceIndex) {
            int remainingBytes = 4;
            for (int i = choiceIndex + 1; i < choiceTokens.Count; i++) {
                remainingBytes = checked(remainingBytes + choiceTokens[i].Length + 4);
            }

            return checked((ushort)(remainingBytes - 1));
        }
    }
}
