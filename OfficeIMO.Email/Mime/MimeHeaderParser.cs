namespace OfficeIMO.Email;

internal static class MimeHeaderParser {
    internal static int Parse(byte[] data, int offset, int count, EmailReaderOptions options,
        IList<EmailHeader> headers, IList<EmailDiagnostic> diagnostics, string location) {
        int end = offset + count;
        if (offset + 3 <= end && data[offset] == 0xEF && data[offset + 1] == 0xBB && data[offset + 2] == 0xBF) {
            offset += 3;
        }
        int headerEnd = FindHeaderEnd(data, offset, end, out int separatorLength);
        int headerBytes = headerEnd - offset;
        if (headerBytes > options.MaxHeaderBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxHeaderBytes), headerBytes, options.MaxHeaderBytes);
        }

        string block = Encoding.UTF8.GetString(data, offset, headerBytes);
        string normalized = block.Replace("\r\n", "\n").Replace('\r', '\n');
        string[] lines = normalized.Split('\n');
        string? currentName = null;
        StringBuilder currentValue = new StringBuilder();
        int malformedIndex = 0;

        Action flush = () => {
            if (currentName == null) return;
            if (headers.Count >= options.MaxHeaderCount) {
                throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxHeaderCount), headers.Count + 1, options.MaxHeaderCount);
            }
            string rawValue = currentValue.ToString().Trim();
            headers.Add(new EmailHeader(currentName, MimeTextCodec.DecodeHeader(rawValue, diagnostics, location), rawValue));
            currentName = null;
            currentValue.Clear();
        };

        foreach (string line in lines) {
            if (line.Length > 0 && (line[0] == ' ' || line[0] == '\t') && currentName != null) {
                if (currentValue.Length > 0) currentValue.Append(' ');
                currentValue.Append(line.Trim());
                continue;
            }

            flush();
            int colon = line.IndexOf(':');
            if (colon <= 0) {
                if (line.Length > 0) {
                    diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_HEADER_MALFORMED",
                        "A header line without a valid field name was ignored.", EmailDiagnosticSeverity.Warning,
                        string.Concat(location, "/header[", malformedIndex.ToString(CultureInfo.InvariantCulture), "]")));
                    malformedIndex++;
                }
                continue;
            }
            currentName = line.Substring(0, colon).Trim();
            currentValue.Append(line.Substring(colon + 1).Trim());
        }
        flush();

        return headerEnd + separatorLength;
    }

    internal static string? GetValue(IEnumerable<EmailHeader> headers, string name) {
        EmailHeader? header = headers.FirstOrDefault(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
        return header?.Value;
    }

    internal static IEnumerable<string> GetValues(IEnumerable<EmailHeader> headers, string name) {
        return headers.Where(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)).Select(item => item.Value);
    }

    /// <summary>Returns the unfolded source value before encoded-word decoding.</summary>
    internal static string? GetRawValue(IEnumerable<EmailHeader> headers, string name) {
        EmailHeader? header = headers.FirstOrDefault(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
        return header?.RawValue ?? header?.Value;
    }

    /// <summary>Returns unfolded source values before encoded-word decoding.</summary>
    internal static IEnumerable<string> GetRawValues(IEnumerable<EmailHeader> headers, string name) {
        return headers.Where(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase))
            .Select(item => item.RawValue ?? item.Value);
    }

    private static int FindHeaderEnd(byte[] data, int offset, int end, out int separatorLength) {
        if (offset < end && data[offset] == '\r') {
            separatorLength = offset + 1 < end && data[offset + 1] == '\n' ? 2 : 1;
            return offset;
        }
        if (offset < end && data[offset] == '\n') {
            separatorLength = 1;
            return offset;
        }
        for (int i = offset; i < end; i++) {
            if (i + 3 < end && data[i] == '\r' && data[i + 1] == '\n' && data[i + 2] == '\r' && data[i + 3] == '\n') {
                separatorLength = 4;
                return i;
            }
            if (i + 1 < end && data[i] == '\n' && data[i + 1] == '\n') {
                separatorLength = 2;
                return i;
            }
            if (i + 1 < end && data[i] == '\r' && data[i + 1] == '\r') {
                separatorLength = 2;
                return i;
            }
        }
        separatorLength = 0;
        return end;
    }
}
