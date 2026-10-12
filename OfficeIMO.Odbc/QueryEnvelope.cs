namespace OfficeIMO.Odbc {

    /// <summary>Restricts application commands to a single SELECT. SQLite remains the SQL parser.</summary>
    internal static class QueryEnvelope {
        internal static string Prepare(string sql, out int parameterCount) {
            Validate(sql);
            parameterCount = 0;
            System.Text.StringBuilder result = new System.Text.StringBuilder(sql.Length);
            int index = 0;
            while (index < sql.Length) {
                int start = index;
                SkipTrivia(sql, ref index);
                if (index != start) { result.Append(sql.AsSpan(start, index - start)); continue; }
                char character = sql[index++];
                result.Append(character);
                if (character is '\'' or '"' or '`' or '[') {
                    char close = character == '[' ? ']' : character;
                    while (index < sql.Length) {
                        char next = sql[index++];
                        result.Append(next);
                        if (next != close) continue;
                        if (index < sql.Length && sql[index] == close) { result.Append(sql[index++]); continue; }
                        break;
                    }
                } else if (character == '?') {
                    if (++parameterCount > 256 || index < sql.Length && char.IsAsciiDigit(sql[index]))
                        throw new NotSupportedException("Use at most 256 unnumbered ODBC parameter markers.");
                    result.Length--;
                    result.Append("@p").Append(parameterCount);
                }
            }
            return result.ToString();
        }

        internal static void Validate(string sql) {
            if (sql.Length > 64 * 1024) throw new ArgumentException("SQL text exceeds 64 Ki characters.");
            int index = 0;
            SkipTrivia(sql, ref index);
            int start = index;
            while (index < sql.Length && char.IsAsciiLetter(sql[index])) index++;
            if (!sql.AsSpan(start, index - start).Equals("SELECT", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("Only a single read-only SELECT statement is supported.");
            while (index < sql.Length) {
                if (sql[index] is '\'' or '"' or '`' or '[') {
                    char close = sql[index] == '[' ? ']' : sql[index];
                    index++;
                    bool closed = false;
                    while (index < sql.Length) {
                        if (sql[index++] != close) continue;
                        if (index < sql.Length && sql[index] == close) { index++; continue; }
                        closed = true;
                        break;
                    }
                    if (!closed) throw new ArgumentException("Unterminated SQL literal or identifier.");
                } else if (sql[index] == ';') {
                    index++;
                    SkipTrivia(sql, ref index);
                    if (index != sql.Length) throw new NotSupportedException("SQL batches are not supported.");
                } else {
                    int before = index;
                    SkipTrivia(sql, ref index);
                    if (before == index) index++;
                }
            }
        }

        private static void SkipTrivia(string sql, ref int index) {
            while (index < sql.Length) {
                if (char.IsWhiteSpace(sql[index])) { index++; continue; }
                if (index + 1 < sql.Length && sql[index] == '-' && sql[index + 1] == '-') {
                    index += 2;
                    while (index < sql.Length && sql[index] != '\n') index++;
                    continue;
                }
                if (index + 1 < sql.Length && sql[index] == '/' && sql[index + 1] == '*') {
                    int end = sql.IndexOf("*/", index + 2, StringComparison.Ordinal);
                    if (end < 0) throw new ArgumentException("Unterminated SQL comment.");
                    index = end + 2;
                    continue;
                }
                break;
            }
        }
    }
}
