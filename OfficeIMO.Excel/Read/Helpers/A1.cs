namespace OfficeIMO.Excel {
    /// <summary>
    /// Utility helpers for parsing and converting Excel A1 references. Public so examples and
    /// consumers can reuse consistent logic without re-implementing regexes or math.
    /// </summary>
    public static class A1 {
        /// <summary>Maximum row index supported by the Excel worksheet grid.</summary>
        public const int MaxRows = 1048576;

        /// <summary>Maximum column index supported by the Excel worksheet grid.</summary>
        public const int MaxColumns = 16384;

        private const int ColumnLettersBufferLength = 7;
        private const int CellReferenceBufferLength = 24;

        /// <summary>
        /// Parses a single A1 cell reference (e.g., "B5") into a 1-based (row, column) tuple.
        /// Returns (0,0) when the input does not match a valid simple cell reference.
        /// </summary>
        /// <param name="cellRef">A1 cell reference, without sheet prefix.</param>
        /// <returns>Tuple of row and column (1-based). Returns (0,0) if invalid.</returns>
        public static (int Row, int Col) ParseCellRef(string cellRef) {
            return TryParseCellRef(cellRef, 0, cellRef?.Length ?? 0, out int row, out int col) ? (row, col) : (0, 0);
        }

        /// <summary>
        /// Tries to parse an A1 range (e.g., "A1:B10") into 1-based, normalized bounds.
        /// Returns false when the input is not a valid A1 range.
        /// </summary>
        public static bool TryParseRange(string a1Range, out int r1, out int c1, out int r2, out int c2) {
            r1 = c1 = r2 = c2 = 0;
            if (string.IsNullOrWhiteSpace(a1Range)) return false;
            int start = 0;
            int length = a1Range.Length;
            TrimBounds(a1Range, ref start, ref length);

            int separator = a1Range.IndexOf(':', start, length);
            if (separator < 0) return false;

            if (!TryParseCellRef(a1Range, start, separator - start, out r1, out c1)
                || !TryParseCellRef(a1Range, separator + 1, start + length - separator - 1, out r2, out c2)) {
                r1 = c1 = r2 = c2 = 0;
                return false;
            }

            if (c1 > c2) (c1, c2) = (c2, c1);
            if (r1 > r2) (r1, r2) = (r2, r1);
            return true;
        }

        internal static bool TryParseWholeColumnRange(string reference, out int c1, out int c2) {
            c1 = c2 = 0;
            if (!TrySplitWholeRange(reference, out string start, out string end)
                || !IsAsciiColumnName(start)
                || !IsAsciiColumnName(end)) {
                return false;
            }

            c1 = ColumnLettersToIndex(start);
            c2 = ColumnLettersToIndex(end);
            if (c1 <= 0 || c1 > MaxColumns || c2 <= 0 || c2 > MaxColumns) {
                c1 = c2 = 0;
                return false;
            }

            if (c1 > c2) (c1, c2) = (c2, c1);
            return true;
        }

        internal static bool TryParseWholeRowRange(string reference, out int r1, out int r2) {
            r1 = r2 = 0;
            if (!TrySplitWholeRange(reference, out string start, out string end)
                || !int.TryParse(start, out r1)
                || !int.TryParse(end, out r2)
                || r1 <= 0
                || r1 > MaxRows
                || r2 <= 0
                || r2 > MaxRows) {
                r1 = r2 = 0;
                return false;
            }

            if (r1 > r2) (r1, r2) = (r2, r1);
            return true;
        }

        private static bool TrySplitWholeRange(string reference, out string start, out string end) {
            start = end = string.Empty;
            if (string.IsNullOrWhiteSpace(reference)) return false;

            string normalized = reference.Replace("$", string.Empty).Trim();
            int separator = normalized.IndexOf(':');
            if (separator <= 0
                || separator != normalized.LastIndexOf(':')
                || separator == normalized.Length - 1) {
                return false;
            }

            start = normalized.Substring(0, separator).Trim();
            end = normalized.Substring(separator + 1).Trim();
            return start.Length > 0 && end.Length > 0;
        }

        private static bool IsAsciiColumnName(string value) {
            if (value.Length == 0 || value.Length > 3) return false;
            foreach (char character in value) {
                char upper = ToUpperAscii(character);
                if (upper < 'A' || upper > 'Z') return false;
            }

            return true;
        }

        internal static bool TryParseStrictRange(string a1Range, out int r1, out int c1, out int r2, out int c2) {
            r1 = c1 = r2 = c2 = 0;
            if (string.IsNullOrWhiteSpace(a1Range)) return false;
            int start = 0;
            int length = a1Range.Length;
            TrimBounds(a1Range, ref start, ref length);

            int separator = a1Range.IndexOf(':', start, length);
            if (separator < 0 || separator != a1Range.LastIndexOf(':', start + length - 1, length)) {
                return false;
            }

            if (!TryParseCellRef(a1Range, start, separator - start, out r1, out c1)
                || !TryParseCellRef(a1Range, separator + 1, start + length - separator - 1, out r2, out c2)) {
                r1 = c1 = r2 = c2 = 0;
                return false;
            }

            return true;
        }

        /// <summary>
        /// Parses an A1 range (e.g., "A1:B10") into 1-based, normalized bounds.
        /// If the bounds are inverted, they are swapped so that r1 &lt;= r2 and c1 &lt;= c2.
        /// </summary>
        /// <param name="a1Range">A1 range string, without sheet prefix.</param>
        /// <returns>(r1, c1, r2, c2) 1-based coordinates.</returns>
        /// <exception cref="ArgumentException">Thrown when the input is not a valid A1 range.</exception>
        /// <example>
        /// var (r1, c1, r2, c2) = A1.ParseRange("B2:D10");
        /// // r1=2, c1=2, r2=10, c2=4
        /// </example>
        public static (int r1, int c1, int r2, int c2) ParseRange(string a1Range) {
            if (!TryParseRange(a1Range, out int r1, out int c1, out int r2, out int c2)) {
                throw new ArgumentException($"Invalid A1 range '{a1Range}'.");
            }

            return (r1, c1, r2, c2);
        }

        /// <summary>
        /// Converts column letters (e.g., "A", "AA") to a 1-based column index.
        /// Non-letter characters are ignored; returns 0 for empty/invalid input.
        /// </summary>
        /// <param name="letters">Excel column letters.</param>
        /// <returns>1-based column index, or 0 when input yields no letters.</returns>
        public static int ColumnLettersToIndex(string letters) {
            int res = 0;
            foreach (char character in letters) {
                char ch = ToUpperAscii(character);
                if (ch < 'A' || ch > 'Z') continue;
                int value = ch - 'A' + 1;
                if (res > (int.MaxValue - value) / 26) {
                    return 0;
                }
                res = res * 26 + value;
            }
            return res;
        }

        internal static int ParseColumnIndexFromCellReference(string? cellRef) {
            if (string.IsNullOrEmpty(cellRef)) return 0;
            string text = cellRef!;
            int start = char.IsWhiteSpace(text[0]) ? FindFirstNonWhiteSpace(text) : 0;
            int end = char.IsWhiteSpace(text[text.Length - 1]) ? FindLastNonWhiteSpace(text, start) : text.Length;
            if (start >= end) return 0;

            return ParseColumnIndexFromTrimmedCellReference(text, start, end);
        }

        internal static int ParseColumnIndexFromCellReferenceFast(string? cellRef) {
            if (string.IsNullOrEmpty(cellRef)) return 0;

            string text = cellRef!;
            int length = text.Length;
            char first = text[0];
            char last = text[length - 1];
            if (!char.IsWhiteSpace(first) && last >= '0' && last <= '9') {
                char firstColumn = ToUpperAscii(first);
                if (firstColumn >= 'A' && firstColumn <= 'Z' && length >= 2) {
                    char second = text[1];
                    if (second >= '0' && second <= '9') {
                        return HasPositiveInt32DigitSuffix(text, 1, length)
                            ? firstColumn - 'A' + 1
                            : 0;
                    }

                    char secondColumn = ToUpperAscii(second);
                    if (secondColumn >= 'A' && secondColumn <= 'Z'
                        && length >= 3
                        && text[2] >= '0'
                        && text[2] <= '9') {
                        return HasPositiveInt32DigitSuffix(text, 2, length)
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
                int commonRow = 0;
                for (; commonIndex < length; commonIndex++) {
                    char ch = text[commonIndex];
                    if (ch < '0' || ch > '9') {
                        return 0;
                    }

                    int digit = ch - '0';
                    if (commonRow > (int.MaxValue - digit) / 10) {
                        return 0;
                    }

                    commonRow = (commonRow * 10) + digit;
                    commonHasNonZeroRowDigit |= digit != 0;
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
            int row = 0;
            for (; index < length; index++) {
                char ch = text[index];
                if (ch >= '0' && ch <= '9') {
                    int digit = ch - '0';
                    if (row > (int.MaxValue - digit) / 10) {
                        return 0;
                    }

                    row = (row * 10) + digit;
                    hasNonZeroRowDigit |= digit != 0;
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
            }

            return hasNonZeroRowDigit ? col : 0;
        }

        internal static bool TryParseCellReferenceFast(string? cellRef, out int row, out int col) {
            row = 0;
            col = 0;
            if (string.IsNullOrEmpty(cellRef)) return false;

            string text = cellRef!;
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

        internal static int ParseColumnIndexFromCellReferenceWithKnownRowFast(string? cellRef) {
            if (string.IsNullOrEmpty(cellRef)) return 0;

            string text = cellRef!;
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

        private static bool HasNonZeroDigitSuffix(string text, int start, int end) {
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

        private static bool HasPositiveInt32DigitSuffix(string text, int start, int end) {
            bool hasNonZeroDigit = false;
            int value = 0;
            for (int i = start; i < end; i++) {
                char ch = text[i];
                if (ch < '0' || ch > '9') {
                    return false;
                }

                int digit = ch - '0';
                if (value > (int.MaxValue - digit) / 10) {
                    return false;
                }

                value = (value * 10) + digit;
                hasNonZeroDigit |= digit != 0;
            }

            return hasNonZeroDigit;
        }

        private static int ParseColumnIndexFromTrimmedCellReference(string text, int start, int end) {
            int col = 0;

            int i = start;
            for (; i < end; i++) {
                char ch = ToUpperAscii(text[i]);
                if (ch < 'A' || ch > 'Z') {
                    break;
                }

                int value = ch - 'A' + 1;
                if (col > (int.MaxValue - value) / 26) {
                    return 0;
                }

                col = col * 26 + value;
            }

            if (i == start || i == end) {
                return 0;
            }

            int row = 0;
            for (; i < end; i++) {
                char ch = text[i];
                if (ch < '0' || ch > '9') {
                    return 0;
                }

                int digit = ch - '0';
                if (row > (int.MaxValue - digit) / 10) {
                    return 0;
                }

                row = (row * 10) + digit;
            }

            return row > 0 ? col : 0;
        }

        private static int FindFirstNonWhiteSpace(string text) {
            int index = 0;
            while (index < text.Length && char.IsWhiteSpace(text[index])) {
                index++;
            }

            return index;
        }

        private static int FindLastNonWhiteSpace(string text, int start) {
            int end = text.Length;
            while (end > start && char.IsWhiteSpace(text[end - 1])) {
                end--;
            }

            return end;
        }

        /// <summary>
        /// Converts a 1-based column index to Excel column letters (e.g., 1→"A", 27→"AA").
        /// </summary>
        /// <param name="index">1-based column index.</param>
        /// <returns>Excel column letters; returns "A" for non-positive inputs.</returns>
        /// <example>
        /// string col = A1.ColumnIndexToLetters(28); // "AB"
        /// int idx = A1.ColumnLettersToIndex("AB"); // 28
        /// </example>
        public static string ColumnIndexToLetters(int index) {
            if (index <= 0) return "A";
            char[] buffer = new char[ColumnLettersBufferLength];
            int position = 0;
            AppendColumnLetters(index, buffer, ref position);
            return new string(buffer, 0, position);
        }

        /// <summary>
        /// Builds an A1 cell reference from 1-based row and column indexes.
        /// </summary>
        public static string CellReference(int row, int column) {
            return BuildCellReference(row, column, absolute: false);
        }

        internal static string AbsoluteCellReference(int row, int column) {
            return BuildCellReference(row, column, absolute: true);
        }

        private static bool TryParseCellRef(string? text, int start, int length, out int row, out int col) {
            row = 0;
            col = 0;
            if (string.IsNullOrEmpty(text) || length <= 0) {
                return false;
            }

            string source = text!;
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

            if (row <= 0 || row > MaxRows || col <= 0 || col > MaxColumns) {
                row = 0;
                col = 0;
                return false;
            }

            return true;
        }

        private static bool IsValidSlice(string text, int start, int length) {
            return (uint)start <= (uint)text.Length
                && (uint)length <= (uint)(text.Length - start);
        }

        private static void TrimBounds(string text, ref int start, ref int length) {
            TrimStart(text, ref start, ref length);
            while (length > 0 && char.IsWhiteSpace(text[start + length - 1])) {
                length--;
            }
        }

        private static void TrimStart(string text, ref int start, ref int length) {
            while (length > 0 && char.IsWhiteSpace(text[start])) {
                start++;
                length--;
            }
        }

        private static char ToUpperAscii(char character) {
            return character >= 'a' && character <= 'z'
                ? (char)(character - ('a' - 'A'))
                : character;
        }

        private static string BuildCellReference(int row, int column, bool absolute) {
            if (row <= 0) {
                throw new ArgumentOutOfRangeException(nameof(row), "A1 references are 1-based and require a positive row.");
            }
            if (column <= 0) {
                throw new ArgumentOutOfRangeException(nameof(column), "A1 references are 1-based and require a positive column.");
            }

            char[] buffer = new char[CellReferenceBufferLength];
            int position = 0;
            if (absolute) {
                buffer[position++] = '$';
            }
            AppendColumnLetters(column, buffer, ref position);
            if (absolute) {
                buffer[position++] = '$';
            }
            AppendPositiveInt(row, buffer, ref position);
            return new string(buffer, 0, position);
        }

        private static void AppendColumnLetters(int index, char[] buffer, ref int position) {
            int letters = 0;
            int n = index;
            while (n > 0) {
                letters++;
                n = (n - 1) / 26;
            }

            int start = position;
            position += letters;
            n = index;
            for (int i = position - 1; i >= start; i--) {
                int rem = (n - 1) % 26;
                buffer[i] = (char)('A' + rem);
                n = (n - 1) / 26;
            }
        }

        private static void AppendPositiveInt(int value, char[] buffer, ref int position) {
            int digits = 1;
            int n = value;
            while (n >= 10) {
                digits++;
                n /= 10;
            }

            int start = position;
            position += digits;
            n = value;
            for (int i = position - 1; i >= start; i--) {
                buffer[i] = (char)('0' + (n % 10));
                n /= 10;
            }
        }
    }
}
