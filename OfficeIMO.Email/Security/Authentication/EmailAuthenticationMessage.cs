namespace OfficeIMO.Email;

/// <summary>
/// Raw RFC 5322 header fields and body bytes used by DKIM and ARC. Header text is held as one char per byte so
/// canonicalization reproduces the exact transmitted bytes, including 8-bit content.
/// </summary>
internal sealed class EmailAuthenticationMessage {
    private EmailAuthenticationMessage(List<Field> fields, byte[] body, bool lineEndingsNormalized) {
        Fields = fields;
        Body = body;
        LineEndingsNormalized = lineEndingsNormalized;
    }

    internal IReadOnlyList<Field> Fields { get; }

    internal byte[] Body { get; }

    internal bool LineEndingsNormalized { get; }

    internal static EmailAuthenticationMessage Parse(byte[] raw, int maxHeaderFields, out bool headerLimitExceeded) {
        byte[] data = NormalizeLineEndings(raw, out bool normalized);
        var fields = new List<Field>();
        headerLimitExceeded = false;
        int position = 0;
        int bodyStart = data.Length;
        while (position < data.Length) {
            int lineEnd = IndexOfCrLf(data, position);
            int next = lineEnd < 0 ? data.Length : lineEnd + 2;
            if (lineEnd == position) {
                bodyStart = next;
                break;
            }

            bool continuation = data[position] == (byte)' ' || data[position] == (byte)'\t';
            if (continuation && fields.Count > 0) {
                Field last = fields[fields.Count - 1];
                fields[fields.Count - 1] = new Field(last.Name, last.Raw + ToText(data, position, next - position));
            } else {
                if (fields.Count >= maxHeaderFields) {
                    headerLimitExceeded = true;
                    bodyStart = data.Length;
                    break;
                }

                string text = ToText(data, position, next - position);
                int colon = text.IndexOf(':');
                string name = colon < 0 ? string.Empty : text.Substring(0, colon).TrimEnd(' ', '\t');
                fields.Add(new Field(name, text));
            }

            position = next;
            if (lineEnd < 0) bodyStart = data.Length;
        }

        byte[] body = new byte[Math.Max(0, data.Length - bodyStart)];
        if (body.Length > 0) Buffer.BlockCopy(data, bodyStart, body, 0, body.Length);
        return new EmailAuthenticationMessage(fields, body, normalized);
    }

    internal static string ToText(byte[] data, int offset, int count) {
        var chars = new char[count];
        for (int index = 0; index < count; index++) chars[index] = (char)data[offset + index];
        return new string(chars);
    }

    internal static byte[] ToBytes(string text) {
        var bytes = new byte[text.Length];
        for (int index = 0; index < text.Length; index++) bytes[index] = unchecked((byte)text[index]);
        return bytes;
    }

    private static int IndexOfCrLf(byte[] data, int start) {
        for (int index = start; index + 1 < data.Length; index++) {
            if (data[index] == (byte)'\r' && data[index + 1] == (byte)'\n') return index;
        }

        return -1;
    }

    private static byte[] NormalizeLineEndings(byte[] raw, out bool normalized) {
        int bareLineFeeds = 0;
        for (int index = 0; index < raw.Length; index++) {
            if (raw[index] == (byte)'\n' && (index == 0 || raw[index - 1] != (byte)'\r')) bareLineFeeds++;
        }

        normalized = bareLineFeeds > 0;
        if (!normalized) return raw;
        var result = new byte[raw.Length + bareLineFeeds];
        int target = 0;
        for (int index = 0; index < raw.Length; index++) {
            if (raw[index] == (byte)'\n' && (index == 0 || raw[index - 1] != (byte)'\r')) result[target++] = (byte)'\r';
            result[target++] = raw[index];
        }

        return result;
    }

    /// <summary>One unfolded-on-demand header field with its exact source text, including the final CRLF.</summary>
    internal readonly struct Field {
        internal Field(string name, string raw) {
            Name = name;
            Raw = raw;
        }

        internal string Name { get; }

        internal string Raw { get; }

        /// <summary>Text after the first colon, still folded, without the final line break.</summary>
        internal string Value {
            get {
                int colon = Raw.IndexOf(':');
                string value = colon < 0 ? string.Empty : Raw.Substring(colon + 1);
                return value.EndsWith("\r\n", StringComparison.Ordinal) ? value.Substring(0, value.Length - 2) : value;
            }
        }

        internal bool Is(string name) => Name.Equals(name, StringComparison.OrdinalIgnoreCase);
    }
}
