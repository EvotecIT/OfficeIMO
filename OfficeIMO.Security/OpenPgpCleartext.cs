using System.Text;

namespace OfficeIMO.Security;

/// <summary>RFC 4880 section 7 cleartext signature framework: extracts the signed text and the signature block.</summary>
internal static class OpenPgpCleartext {
    private const string Begin = "-----BEGIN PGP SIGNED MESSAGE-----";
    private const string SignatureBegin = "-----BEGIN PGP SIGNATURE-----";
    private const string SignatureEnd = "-----END PGP SIGNATURE-----";

    /// <summary>
    /// Returns the exact signed bytes: dash-escaping removed, trailing spaces and tabs removed from every line,
    /// lines joined with CRLF and no line break after the last line.
    /// </summary>
    internal static bool TryParse(byte[] message, out byte[]? signedText, out byte[]? signatureBlock, out string? error) {
        signedText = null;
        signatureBlock = null;
        error = null;
        List<(int Start, int Length)> lines = SplitLines(message);
        int index = 0;
        while (index < lines.Count && IsBlank(message, lines[index])) index++;
        if (index >= lines.Count || !LineEquals(message, lines[index], Begin)) {
            error = "The message does not start with a cleartext signature header.";
            return false;
        }

        index++;
        while (index < lines.Count && !IsBlank(message, lines[index])) {
            string header = Text(message, lines[index]);
            if (header.IndexOf(": ", StringComparison.Ordinal) <= 0) {
                error = "The cleartext signature has a malformed armor header.";
                return false;
            }

            index++;
        }

        if (index >= lines.Count) {
            error = "The cleartext signature has no body.";
            return false;
        }

        index++;
        int bodyStart = index;
        while (index < lines.Count && !LineEquals(message, lines[index], SignatureBegin)) index++;
        if (index >= lines.Count) {
            error = "The cleartext signature has no signature block.";
            return false;
        }

        int signatureStart = index;
        int signatureEnd = signatureStart;
        while (signatureEnd < lines.Count && !LineEquals(message, lines[signatureEnd], SignatureEnd)) signatureEnd++;
        if (signatureEnd >= lines.Count) {
            error = "The signature block is not terminated.";
            return false;
        }

        using var signed = new MemoryStream();
        for (int line = bodyStart; line < signatureStart; line++) {
            (int start, int length) = lines[line];
            if (length >= 2 && message[start] == (byte)'-' && message[start + 1] == (byte)' ') {
                start += 2;
                length -= 2;
            } else if (length >= 1 && message[start] == (byte)'-') {
                error = "A cleartext line starting with '-' is not dash-escaped.";
                return false;
            }

            while (length > 0 && (message[start + length - 1] == (byte)' ' || message[start + length - 1] == (byte)'\t')) length--;
            if (line > bodyStart) {
                signed.WriteByte((byte)'\r');
                signed.WriteByte((byte)'\n');
            }

            signed.Write(message, start, length);
        }

        using var block = new MemoryStream();
        for (int line = signatureStart; line <= signatureEnd; line++) {
            block.Write(message, lines[line].Start, lines[line].Length);
            block.WriteByte((byte)'\n');
        }

        signedText = signed.ToArray();
        signatureBlock = block.ToArray();
        return true;
    }

    private static List<(int Start, int Length)> SplitLines(byte[] data) {
        var lines = new List<(int, int)>();
        int start = 0;
        for (int index = 0; index <= data.Length; index++) {
            if (index == data.Length || data[index] == (byte)'\n') {
                int end = index;
                if (end > start && data[end - 1] == (byte)'\r') end--;
                if (index < data.Length || end > start) lines.Add((start, end - start));
                start = index + 1;
            }
        }

        return lines;
    }

    private static bool IsBlank(byte[] data, (int Start, int Length) line) {
        for (int index = line.Start; index < line.Start + line.Length; index++) {
            if (data[index] != (byte)' ' && data[index] != (byte)'\t') return false;
        }

        return true;
    }

    private static bool LineEquals(byte[] data, (int Start, int Length) line, string expected) {
        int length = line.Length;
        while (length > 0 && (data[line.Start + length - 1] == (byte)' ' || data[line.Start + length - 1] == (byte)'\t')) length--;
        if (length != expected.Length) return false;
        for (int index = 0; index < length; index++) {
            if (data[line.Start + index] != (byte)expected[index]) return false;
        }

        return true;
    }

    private static string Text(byte[] data, (int Start, int Length) line) => Encoding.ASCII.GetString(data, line.Start, line.Length);
}
