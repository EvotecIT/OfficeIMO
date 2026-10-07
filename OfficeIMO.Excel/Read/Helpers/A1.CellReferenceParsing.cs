namespace OfficeIMO.Excel {
    public static partial class A1 {
        internal static bool TryParseCellReferenceFast(ReadOnlySpan<char> cellRef, out int row, out int col) {
            row = 0;
            col = 0;
            if (cellRef.IsEmpty) return false;

            ReadOnlySpan<char> text = cellRef;
            int length = text.Length;
            char first = text[0];
            char last = text[length - 1];
            if (!char.IsWhiteSpace(first) && last >= '0' && last <= '9') {
                int index = 0;
                for (; index < length; index++) {
                    char ch = ToUpperAscii(text[index]);
                    if (ch < 'A' || ch > 'Z') {
                        break;
                    }

                    int value = ch - 'A' + 1;
                    if (col > (int.MaxValue - value) / 26) {
                        row = 0;
                        col = 0;
                        return false;
                    }

                    col = (col * 26) + value;
                }

                if (index == 0 || index == length) {
                    row = 0;
                    col = 0;
                    return false;
                }

                for (; index < length; index++) {
                    char ch = text[index];
                    if (ch < '0' || ch > '9') {
                        row = 0;
                        col = 0;
                        return false;
                    }

                    int digit = ch - '0';
                    if (row > (int.MaxValue - digit) / 10) {
                        row = 0;
                        col = 0;
                        return false;
                    }

                    row = (row * 10) + digit;
                }

                if (row <= 0 || row > MaxRows || col <= 0 || col > MaxColumns) {
                    row = 0;
                    col = 0;
                    return false;
                }

                return true;
            }

            return TryParseCellRef(cellRef, 0, length, out row, out col);
        }

        internal static int ParseColumnIndexFromCellReferenceWithKnownRowFast(ReadOnlySpan<char> cellRef) {
            if (cellRef.IsEmpty) return 0;

            ReadOnlySpan<char> text = cellRef;
            int length = text.Length;
            char first = text[0];
            char last = text[length - 1];
            if (!char.IsWhiteSpace(first) && last >= '0' && last <= '9') {
                char firstColumn = ToUpperAscii(first);
                if (firstColumn >= 'A' && firstColumn <= 'Z' && length >= 2) {
                    char second = text[1];
                    if (second >= '0' && second <= '9') {
                        return HasNonZeroDigitSuffix(text, 1, length)
                            ? firstColumn - 'A' + 1
                            : 0;
                    }

                    char secondColumn = ToUpperAscii(second);
                    if (secondColumn >= 'A' && secondColumn <= 'Z'
                        && length >= 3
                        && text[2] >= '0'
                        && text[2] <= '9') {
                        return HasNonZeroDigitSuffix(text, 2, length)
                            ? (((firstColumn - 'A' + 1) * 26) + (secondColumn - 'A' + 1))
                            : 0;
                    }
                }

                int commonIndex = 0;
                int commonCol = 0;
                for (; commonIndex < length; commonIndex++) {
                    char ch = ToUpperAscii(text[commonIndex]);
                    if (ch < 'A' || ch > 'Z') {
                        break;
                    }

                    int value = ch - 'A' + 1;
                    if (commonCol > (int.MaxValue - value) / 26) {
                        return 0;
                    }

                    commonCol = (commonCol * 26) + value;
                }

                if (commonIndex == 0 || commonIndex == length) {
                    return 0;
                }

                bool commonHasNonZeroRowDigit = false;
                for (; commonIndex < length; commonIndex++) {
                    char ch = text[commonIndex];
                    if (ch < '0' || ch > '9') {
                        return 0;
                    }

                    commonHasNonZeroRowDigit |= ch != '0';
                }

                return commonHasNonZeroRowDigit ? commonCol : 0;
            }

            int index = 0;
            while (index < length && char.IsWhiteSpace(text[index])) {
                index++;
            }

            int col = 0;
            int letterStart = index;
            for (; index < length; index++) {
                char ch = ToUpperAscii(text[index]);
                if (ch < 'A' || ch > 'Z') {
                    break;
                }

                int value = ch - 'A' + 1;
                if (col > (int.MaxValue - value) / 26) {
                    return 0;
                }

                col = (col * 26) + value;
            }

            if (index == letterStart || index == length) {
                return 0;
            }

            bool hasNonZeroRowDigit = false;
            for (; index < length; index++) {
                char ch = text[index];
                if (ch >= '0' && ch <= '9') {
                    hasNonZeroRowDigit |= ch != '0';
                    continue;
                }

                if (!char.IsWhiteSpace(ch)) {
                    return 0;
                }

                while (++index < length) {
                    if (!char.IsWhiteSpace(text[index])) {
                        return 0;
                    }
                }

                break;
            }

            return hasNonZeroRowDigit ? col : 0;
        }

        private static bool HasNonZeroDigitSuffix(ReadOnlySpan<char> text, int start, int end) {
            bool hasNonZeroDigit = false;
            for (int i = start; i < end; i++) {
                char ch = text[i];
                if (ch < '0' || ch > '9') {
                    return false;
                }

                hasNonZeroDigit |= ch != '0';
            }

            return hasNonZeroDigit;
        }

        private static bool TryParseCellRef(ReadOnlySpan<char> text, int start, int length,
            out int row, out int col, bool enforceWorksheetBounds = true) {
            row = 0;
            col = 0;
            if (text.IsEmpty || length <= 0) {
                return false;
            }

            ReadOnlySpan<char> source = text;
            if (!IsValidSlice(source, start, length)) {
                return false;
            }

            TrimBounds(source, ref start, ref length);
            if (length <= 0) {
                return false;
            }

            int end = start + length;
            int i = start;
            for (; i < end; i++) {
                char ch = ToUpperAscii(source[i]);
                if (ch < 'A' || ch > 'Z') {
                    break;
                }

                int value = ch - 'A' + 1;
                if (col > (int.MaxValue - value) / 26) {
                    row = 0;
                    col = 0;
                    return false;
                }

                col = col * 26 + value;
            }

            if (i == start || i == end) {
                row = 0;
                col = 0;
                return false;
            }

            for (; i < end; i++) {
                char ch = source[i];
                if (ch < '0' || ch > '9') {
                    row = 0;
                    col = 0;
                    return false;
                }

                int digit = ch - '0';
                if (row > (int.MaxValue - digit) / 10) {
                    row = 0;
                    col = 0;
                    return false;
                }

                row = row * 10 + digit;
            }

            if (row <= 0 || col <= 0
                || (enforceWorksheetBounds
                    && (row > MaxRows || col > MaxColumns))) {
                row = 0;
                col = 0;
                return false;
            }

            return true;
        }

        private static bool IsValidSlice(ReadOnlySpan<char> text, int start, int length) {
            return (uint)start <= (uint)text.Length
                && (uint)length <= (uint)(text.Length - start);
        }

    }
}
