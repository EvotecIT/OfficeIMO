namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsFormulaEncoder {
        private static bool TryEncodeReferenceOperator(
            string formulaText,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            if (TryGetParenthesizedInnerText(formulaText, out string? innerText)
                && TryEncodeReferenceOperator(innerText!, nameIndex, formulaSheetIndex, out byte[] innerTokens)) {
                tokens = ShouldPreserveReferenceParentheses(innerTokens)
                    ? AppendToken(innerTokens, 0x15)
                    : innerTokens;
                return true;
            }

            if (!TryFindReferenceOperatorSplit(formulaText, out int operatorIndex, out int operatorLength, out byte operatorToken)) {
                return false;
            }

            string left = formulaText.Substring(0, operatorIndex).Trim();
            string right = formulaText.Substring(operatorIndex + operatorLength).Trim();
            if (!TryEncodeReferenceTerm(left, nameIndex, formulaSheetIndex, out byte[] leftTokens)
                || !TryEncodeReferenceTerm(right, nameIndex, formulaSheetIndex, out byte[] rightTokens)) {
                return false;
            }

            using var stream = new MemoryStream();
            stream.Write(leftTokens, 0, leftTokens.Length);
            stream.Write(rightTokens, 0, rightTokens.Length);
            stream.WriteByte(operatorToken);
            tokens = stream.ToArray();
            return true;
        }

        private static bool TryEncodeReferenceTerm(
            string text,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            text = text.Trim();
            if (TryGetParenthesizedInnerText(text, out string? innerText)
                && TryEncodeReferenceOperator(innerText!, nameIndex, formulaSheetIndex, out byte[] innerTokens)) {
                tokens = ShouldPreserveReferenceParentheses(innerTokens)
                    ? AppendToken(innerTokens, 0x15)
                    : innerTokens;
                return true;
            }

            if (TryEncodeReferenceOperator(text, nameIndex, formulaSheetIndex, out tokens)) {
                return true;
            }

            return TryEncodeReferenceOperand(text, nameIndex, formulaSheetIndex, out tokens);
        }

        private static bool TryEncodeReferenceOperand(
            string text,
            LegacyXlsFormulaNameIndex nameIndex,
            int formulaSheetIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            if (TryEncodeSheetQualifiedReference(text, allowArea: true, nameIndex, out tokens)) {
                return true;
            }

            if (TryParseAreaReference(text, out FormulaReference first, out FormulaReference last)) {
                tokens = BuildAreaReferenceToken(first, last);
                return true;
            }

            if (TryParseCellReference(text, out FormulaReference reference)) {
                tokens = BuildCellReferenceToken(reference);
                return true;
            }

            return TryEncodeDefinedName(text, nameIndex, formulaSheetIndex, out tokens);
        }

        private static bool ShouldPreserveReferenceParentheses(byte[] tokens) {
            return tokens.Length > 0 && tokens[tokens.Length - 1] == 0x11;
        }

        private static bool TryFindReferenceOperatorSplit(
            string formulaText,
            out int operatorIndex,
            out int operatorLength,
            out byte operatorToken) {
            operatorIndex = -1;
            operatorLength = 0;
            operatorToken = 0;
            bool inStringLiteral = false;
            int parenthesisDepth = 0;
            int arrayDepth = 0;
            int rangeOperatorIndex = -1;
            int intersectionOperatorIndex = -1;
            int intersectionOperatorLength = 0;

            for (int i = formulaText.Length - 1; i >= 0; i--) {
                char ch = formulaText[i];
                if (!inStringLiteral && TrySkipQuotedSheetNameBackward(formulaText, ref i)) {
                    continue;
                }

                if (ch == '"') {
                    if (inStringLiteral && i > 0 && formulaText[i - 1] == '"') {
                        i--;
                        continue;
                    }

                    inStringLiteral = !inStringLiteral;
                    continue;
                }

                if (inStringLiteral) {
                    continue;
                }

                if (ch == '}') {
                    arrayDepth++;
                    continue;
                }

                if (ch == '{') {
                    arrayDepth--;
                    if (arrayDepth < 0) {
                        return false;
                    }

                    continue;
                }

                if (ch == ')') {
                    parenthesisDepth++;
                    continue;
                }

                if (ch == '(') {
                    parenthesisDepth--;
                    if (parenthesisDepth < 0) {
                        return false;
                    }

                    continue;
                }

                if (parenthesisDepth != 0 || arrayDepth != 0) {
                    continue;
                }

                if (ch == ',') {
                    operatorIndex = i;
                    operatorLength = 1;
                    operatorToken = 0x10;
                    return true;
                }

                if (ch == ':') {
                    rangeOperatorIndex = i;
                    continue;
                }

                if (char.IsWhiteSpace(ch)) {
                    int start = i;
                    while (start > 0 && char.IsWhiteSpace(formulaText[start - 1])) {
                        start--;
                    }

                    int previous = PreviousNonWhiteSpaceIndex(formulaText, start);
                    int next = NextNonWhiteSpaceIndex(formulaText, i);
                    if (previous >= 0 && next >= 0 && !IsReferenceOperatorBoundary(formulaText[previous]) && !IsReferenceOperatorBoundary(formulaText[next])) {
                        if (intersectionOperatorIndex < 0) {
                            intersectionOperatorIndex = start;
                            intersectionOperatorLength = i - start + 1;
                        }
                    }

                    i = start;
                }
            }

            if (intersectionOperatorIndex >= 0) {
                operatorIndex = intersectionOperatorIndex;
                operatorLength = intersectionOperatorLength;
                operatorToken = 0x0f;
                return true;
            }

            if (rangeOperatorIndex >= 0) {
                operatorIndex = rangeOperatorIndex;
                operatorLength = 1;
                operatorToken = 0x11;
                return true;
            }

            return false;
        }

        private static int PreviousNonWhiteSpaceIndex(string text, int index) {
            for (int i = index - 1; i >= 0; i--) {
                if (!char.IsWhiteSpace(text[i])) {
                    return i;
                }
            }

            return -1;
        }

        private static int NextNonWhiteSpaceIndex(string text, int index) {
            for (int i = index + 1; i < text.Length; i++) {
                if (!char.IsWhiteSpace(text[i])) {
                    return i;
                }
            }

            return -1;
        }

        private static bool IsReferenceOperatorBoundary(char ch) {
            return ch == '+'
                || ch == '-'
                || ch == '*'
                || ch == '/'
                || ch == '^'
                || ch == '&'
                || ch == '='
                || ch == '<'
                || ch == '>'
                || ch == ','
                || ch == ':'
                || ch == '('
                || ch == ')';
        }
    }
}
