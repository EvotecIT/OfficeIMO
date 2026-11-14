namespace OfficeIMO.Email;

/// <summary>DKIM canonicalization algorithms.</summary>
public enum EmailDkimCanonicalization {
    /// <summary>The "simple" algorithm: bytes are used as transmitted.</summary>
    Simple = 0,

    /// <summary>The "relaxed" algorithm: whitespace and header-name case are normalized.</summary>
    Relaxed = 1
}

/// <summary>RFC 6376 section 3.4 header and body canonicalization over raw message text.</summary>
internal static class EmailDkimCanonicalizer {
    /// <summary>Canonicalizes one header field. The input includes its final CRLF when present.</summary>
    internal static string CanonicalizeHeader(string rawField, EmailDkimCanonicalization mode) {
        if (mode == EmailDkimCanonicalization.Simple) return rawField;
        int colon = rawField.IndexOf(':');
        string name = colon < 0 ? rawField : rawField.Substring(0, colon);
        string value = colon < 0 ? string.Empty : rawField.Substring(colon + 1);
        var builder = new StringBuilder(rawField.Length);
        builder.Append(name.TrimEnd(' ', '\t').ToLowerInvariant());
        builder.Append(':');
        bool pendingSpace = false;
        bool started = false;
        foreach (char character in value) {
            if (character == '\r' || character == '\n') continue;
            if (character == ' ' || character == '\t') {
                pendingSpace = started;
                continue;
            }

            if (pendingSpace) builder.Append(' ');
            pendingSpace = false;
            started = true;
            builder.Append(character);
        }

        builder.Append("\r\n");
        return builder.ToString();
    }

    /// <summary>Canonicalizes a complete body and returns the canonical bytes.</summary>
    internal static byte[] CanonicalizeBody(byte[] body, EmailDkimCanonicalization mode) {
        using var output = new MemoryStream(body.Length + 2);
        int position = 0;
        int pendingEmptyLines = 0;
        bool anyContent = false;
        while (position < body.Length) {
            int lineEnd = position;
            while (lineEnd + 1 < body.Length && !(body[lineEnd] == (byte)'\r' && body[lineEnd + 1] == (byte)'\n')) lineEnd++;
            bool hasCrLf = lineEnd + 1 < body.Length;
            if (!hasCrLf) lineEnd = body.Length;
            byte[] line = mode == EmailDkimCanonicalization.Relaxed
                ? RelaxLine(body, position, lineEnd - position)
                : Slice(body, position, lineEnd - position);
            if (line.Length == 0) {
                if (hasCrLf) pendingEmptyLines++;
            } else {
                for (int index = 0; index < pendingEmptyLines; index++) {
                    output.WriteByte((byte)'\r');
                    output.WriteByte((byte)'\n');
                }

                pendingEmptyLines = 0;
                output.Write(line, 0, line.Length);
                output.WriteByte((byte)'\r');
                output.WriteByte((byte)'\n');
                anyContent = true;
            }

            position = hasCrLf ? lineEnd + 2 : body.Length;
        }

        if (!anyContent && mode == EmailDkimCanonicalization.Simple) return new[] { (byte)'\r', (byte)'\n' };
        return output.ToArray();
    }

    private static byte[] Slice(byte[] data, int offset, int count) {
        var result = new byte[count];
        Buffer.BlockCopy(data, offset, result, 0, count);
        return result;
    }

    private static byte[] RelaxLine(byte[] data, int offset, int count) {
        var result = new List<byte>(count);
        bool pendingSpace = false;
        for (int index = offset; index < offset + count; index++) {
            byte value = data[index];
            if (value == (byte)' ' || value == (byte)'\t') {
                pendingSpace = true;
                continue;
            }

            if (pendingSpace) result.Add((byte)' ');
            pendingSpace = false;
            result.Add(value);
        }

        return result.ToArray();
    }
}
