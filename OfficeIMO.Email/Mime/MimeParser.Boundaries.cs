namespace OfficeIMO.Email;

internal static partial class MimeParser {
    internal static List<ArraySegment<byte>> SplitMultipart(byte[] data, int offset, int count, string boundary,
        MimeParserState state, string location) {
        byte[] marker = Encoding.ASCII.GetBytes(string.Concat("--", boundary));
        int end = offset + count;
        int lineStart = offset;
        int partStart = -1;
        bool closed = false;
        List<ArraySegment<byte>> parts = new List<ArraySegment<byte>>();

        while (lineStart < end) {
            state.ThrowIfCancellationRequested();
#if NET8_0_OR_GREATER
            lineStart = FindBoundaryCandidate(data, lineStart, offset, end, state);
            if (lineStart == end) break;
#endif
            int lineEnd = lineStart;
#if NET8_0_OR_GREATER
            int newline = data.AsSpan(lineStart, end - lineStart).IndexOfAny((byte)'\r', (byte)'\n');
            lineEnd = newline < 0 ? end : lineStart + newline;
#else
            while (lineEnd < end && data[lineEnd] != '\r' && data[lineEnd] != '\n') lineEnd++;
#endif
            if (IsBoundaryLine(data, lineStart, lineEnd, marker, out bool closing)) {
                if (partStart >= 0) {
                    int partEnd = TrimLineEndingBefore(data, partStart, lineStart);
                    state.EnsurePendingPartCount(parts.Count + 1);
                    parts.Add(new ArraySegment<byte>(data, partStart, Math.Max(0, partEnd - partStart)));
                }
                if (closing) {
                    closed = true;
                    break;
                }
                partStart = SkipLineEnding(data, lineEnd, end);
            }
            lineStart = SkipLineEnding(data, lineEnd, end);
        }

        if (!closed) {
            if (partStart >= 0 && partStart < end) {
                int partEnd = TrimLineEndingBefore(data, partStart, end);
                state.EnsurePendingPartCount(parts.Count + 1);
                parts.Add(new ArraySegment<byte>(data, partStart, Math.Max(0, partEnd - partStart)));
            }
            state.Diagnostics.Add(new EmailDiagnostic(BoundaryNotClosedDiagnosticCode,
                string.Concat("Multipart boundary '", boundary, "' has no closing delimiter."),
                EmailDiagnosticSeverity.Warning, location));
        }
        if (partStart < 0) {
            state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BOUNDARY_NOT_FOUND",
                string.Concat("Multipart boundary '", boundary, "' was not found."),
                EmailDiagnosticSeverity.Error, location));
        }
        return parts;
    }

#if NET8_0_OR_GREATER
    // Skip runs without delimiter candidates instead of visiting every wrapped line,
    // then retain the existing complete-line predicate.
    private static int FindBoundaryCandidate(byte[] data, int position, int start, int end, MimeParserState state) {
        while (position < end) {
            state.ThrowIfCancellationRequested();
            int count = Math.Min(64 * 1024, end - position);
            int dash = data.AsSpan(position, count).IndexOf((byte)'-');
            if (dash < 0) {
                position += count;
                continue;
            }
            position += dash;
            if (position == start || data[position - 1] == '\r' || data[position - 1] == '\n') return position;
            int newline = data.AsSpan(position, end - position).IndexOfAny((byte)'\r', (byte)'\n');
            if (newline < 0) return end;
            position = SkipLineEnding(data, position + newline, end);
        }
        return end;
    }
#endif

    private static bool IsBoundaryLine(byte[] data, int start, int end, byte[] marker, out bool closing) {
        closing = false;
        if (end - start < marker.Length) return false;
        for (int i = 0; i < marker.Length; i++) {
            if (data[start + i] != marker[i]) return false;
        }
        int position = start + marker.Length;
        if (position + 1 < end && data[position] == '-' && data[position + 1] == '-') {
            closing = true;
            position += 2;
        }
        while (position < end && (data[position] == ' ' || data[position] == '\t')) position++;
        return position == end;
    }

    private static int SkipLineEnding(byte[] data, int position, int end) {
        if (position < end && data[position] == '\r') position++;
        if (position < end && data[position] == '\n') position++;
        return position;
    }

    private static int TrimLineEndingBefore(byte[] data, int minimum, int position) {
        int result = position;
        if (result > minimum && data[result - 1] == '\n') result--;
        if (result > minimum && data[result - 1] == '\r') result--;
        return result;
    }
}
