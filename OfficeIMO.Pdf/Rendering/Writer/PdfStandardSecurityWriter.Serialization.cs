using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfStandardSecurityWriter {
    // Encryption receives the canonical envelopes emitted by PdfObjectBytes, not arbitrary PDF files.
    // Anchor those envelopes so payload tokens can never become object boundaries.
    private static byte[] ReadSerializedObjectBody(byte[] bytes, int objectNumber) {
        string header = objectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj\n";
        const string footer = "endobj\n";
        if (!MatchesToken(bytes, 0, header.Length, header) ||
            !MatchesToken(bytes, bytes.Length - footer.Length, bytes.Length, footer) ||
            bytes.Length < header.Length + footer.Length) {
            throw new InvalidOperationException("PDF encryption requires a canonical indirect-object envelope.");
        }
        return Slice(bytes, header.Length, bytes.Length - footer.Length);
    }

    private sealed class StreamLayout {
        internal int DataStart { get; set; }
        internal int DataEnd { get; set; }
        internal int LengthStart { get; set; }
        internal int LengthEnd { get; set; }
        internal bool IsMetadata { get; set; }
    }

    /// <summary>Finds the root dictionary's stream and direct length without interpreting string or stream payloads.</summary>
    private static StreamLayout? ReadStreamLayout(byte[] body, CancellationToken cancellationToken) {
        int index = 0;
        var token = ReadSecurityToken(body, ref index, cancellationToken);
        if (!MatchesToken(body, token.Start, token.End, "<<")) return null;

        var result = new StreamLayout();
        int length = -1;
        while (true) {
            var key = ReadSecurityToken(body, ref index, cancellationToken);
            if (MatchesToken(body, key.Start, key.End, ">>")) break;
            if (key.Start == key.End || body[key.Start] != '/') throw InvalidStreamSyntax();
            var value = ReadSecurityToken(body, ref index, cancellationToken);
            if (value.Start == value.End) throw InvalidStreamSyntax();
            bool directValue = SkipSecurityValue(body, value, ref index, cancellationToken);
            if (MatchesToken(body, key.Start, key.End, "/Length")) {
                length = directValue && TryReadStreamLength(body, value.Start, value.End, out int declaredLength) ? declaredLength : -1;
                result.LengthStart = value.Start;
                result.LengthEnd = value.End;
            }
            if (MatchesToken(body, key.Start, key.End, "/Type")) {
                result.IsMetadata = directValue && MatchesToken(body, value.Start, value.End, "/Metadata");
            }
        }

        token = ReadSecurityToken(body, ref index, cancellationToken);
        if (token.Start == token.End) return null;
        if (!MatchesToken(body, token.Start, token.End, "stream") || length < 0) throw InvalidStreamSyntax();
        // Canonical writers use LF; accept CRLF as well. The EOL is not part of the payload.
        if (index < body.Length && body[index] == '\r') index++;
        if (index >= body.Length || body[index++] != '\n') throw InvalidStreamSyntax();
        result.DataStart = index;
        if (length > body.Length - index) throw InvalidStreamSyntax();
        result.DataEnd = index + length;
        index = result.DataEnd;
        token = ReadSecurityToken(body, ref index, cancellationToken);
        if (!MatchesToken(body, token.Start, token.End, "endstream")) throw InvalidStreamSyntax();
        token = ReadSecurityToken(body, ref index, cancellationToken);
        if (token.Start != token.End) throw InvalidStreamSyntax();
        return result;
    }

    // Skip one dictionary value, including nested containers and indirect references. No recursive descent
    // or reader repair is appropriate here: malformed generated syntax must fail before publishing ciphertext.
    private static bool SkipSecurityValue(byte[] bytes, (int Start, int End) token, ref int index, CancellationToken cancellationToken) {
        bool dictionary = MatchesToken(bytes, token.Start, token.End, "<<");
        bool array = MatchesToken(bytes, token.Start, token.End, "[");
        if (dictionary || array) {
            var closes = new Stack<string>();
            closes.Push(dictionary ? ">>" : "]");
            while (closes.Count > 0) {
                token = ReadSecurityToken(bytes, ref index, cancellationToken);
                if (token.Start == token.End) throw InvalidStreamSyntax();
                if (MatchesToken(bytes, token.Start, token.End, "<<")) closes.Push(">>");
                else if (MatchesToken(bytes, token.Start, token.End, "[")) closes.Push("]");
                else if (MatchesToken(bytes, token.Start, token.End, ">>") || MatchesToken(bytes, token.Start, token.End, "]")) {
                    if (!MatchesToken(bytes, token.Start, token.End, closes.Pop())) throw InvalidStreamSyntax();
                }
            }
            return false;
        }
        if (TryReadStreamLength(bytes, token.Start, token.End, out _)) {
            int lookahead = index;
            var generation = ReadSecurityToken(bytes, ref lookahead, cancellationToken);
            if (TryReadStreamLength(bytes, generation.Start, generation.End, out _)) {
                var reference = ReadSecurityToken(bytes, ref lookahead, cancellationToken);
                if (MatchesToken(bytes, reference.Start, reference.End, "R")) {
                    index = lookahead;
                    return false;
                }
            }
        }
        return true;
    }

    private static (int Start, int End) ReadSecurityToken(byte[] bytes, ref int index, CancellationToken cancellationToken) {
        while (index < bytes.Length) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (IsWhiteSpace(bytes[index])) { index++; continue; }
            if (bytes[index] != '%') break;
            while (index < bytes.Length && bytes[index] != '\r' && bytes[index] != '\n') {
                if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                index++;
            }
        }
        int start = index;
        if (index == bytes.Length) return (start, index);
        byte first = bytes[index++];
        if (first == '(') {
            if (!TryReadLiteralString(bytes, start, out int end, out _, cancellationToken)) throw InvalidStreamSyntax();
            index = end + 1;
        } else if (first == '<') {
            if (index < bytes.Length && bytes[index] == '<') index++;
            else {
                if (!TryReadHexString(bytes, start, out int end, out _, cancellationToken)) throw InvalidStreamSyntax();
                index = end + 1;
            }
        } else if (first == '>' && index < bytes.Length && bytes[index] == '>') {
            index++;
        } else if (first != '[' && first != ']' && first != '>' && first != ')') {
            while (index < bytes.Length && !IsWhiteSpace(bytes[index]) && !IsSecurityDelimiter(bytes[index])) {
                if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                index++;
            }
        }
        return (start, index);
    }

    private static bool IsSecurityDelimiter(byte value) => value == '(' || value == ')' || value == '<' || value == '>' ||
        value == '[' || value == ']' || value == '/' || value == '%';

    private static bool MatchesToken(byte[] bytes, int start, int end, string token) {
        if (start < 0 || end > bytes.Length || end - start != token.Length) return false;
        for (int i = 0; i < token.Length; i++) if (bytes[start + i] != token[i]) return false;
        return true;
    }

    private static bool TryReadStreamLength(byte[] bytes, int start, int end, out int length) {
        length = 0;
        if (start == end) return false;
        for (int i = start; i < end; i++) {
            int digit = bytes[i] - '0';
            if (digit < 0 || digit > 9 || length > (int.MaxValue - digit) / 10) return false;
            length = length * 10 + digit;
        }
        return true;
    }

    private static InvalidOperationException InvalidStreamSyntax() =>
        new InvalidOperationException("PDF encryption requires a valid serialized stream dictionary with a matching direct /Length.");
}
