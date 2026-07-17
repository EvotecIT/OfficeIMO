namespace OfficeIMO.Email;

/// <summary>
/// Builds a metadata/body skeleton while decoding ordinary attachment entities directly to file-backed sources.
/// The established MIME parser then applies the complete semantic projection to the bounded skeleton.
/// </summary>
internal sealed class MimeStreamingParser {
    private readonly Stream _input;
    private readonly EmailReaderOptions _options;
    private readonly IList<EmailDiagnostic> _diagnostics;
    private readonly CancellationToken _cancellationToken;
    private readonly EmailReadWorkspace _workspace;
    private readonly List<ExternalPart> _externalParts = new List<ExternalPart>();
    private int _analyzedPartCount;
    private long _totalAttachmentBytes;

    private MimeStreamingParser(Stream input, EmailReaderOptions options,
        IList<EmailDiagnostic> diagnostics, CancellationToken cancellationToken, EmailReadWorkspace workspace) {
        _input = input;
        _options = options;
        _diagnostics = diagnostics;
        _cancellationToken = cancellationToken;
        _workspace = workspace;
    }

    internal static EmailDocument Parse(Stream input, EmailReaderOptions options,
        IList<EmailDiagnostic> diagnostics, CancellationToken cancellationToken, EmailReadWorkspace workspace) {
        if (!input.CanSeek) throw new ArgumentException("Streaming MIME parsing requires a seekable source.", nameof(input));
        var parser = new MimeStreamingParser(input, options, diagnostics, cancellationToken, workspace);
        long start = input.Position;
        long end = input.Length;
        parser.AnalyzeMessage(start, end, 0, "message");
        parser.DecodeExternalParts();
        byte[] skeleton = parser.CreateSkeleton(start, end);
        EmailDocument document = MimeParser.Parse(skeleton, options, diagnostics, cancellationToken);
        parser.ApplyExternalContent(document);
        return document;
    }

    private void AnalyzeMessage(long start, long end, int depth, string location) {
        if (depth > _options.MaxNestedMessageDepth) return;
        HeaderSection section = ReadHeaders(start, end, location);
        AnalyzeEntity(section.Headers, section.BodyStart, end, 0, location);
    }

    private void AnalyzeEntity(IReadOnlyList<EmailHeader> headers, long bodyStart, long end,
        int depth, string location) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (depth > _options.MaxMimeDepth) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxMimeDepth), depth,
                _options.MaxMimeDepth);
        }
        _analyzedPartCount++;
        if (_analyzedPartCount > _options.MaxPartCount) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxPartCount), _analyzedPartCount,
                _options.MaxPartCount);
        }

        MimeValue contentType = MimeValueParser.Parse(MimeHeaderParser.GetValue(headers, "Content-Type"),
            "text/plain", _diagnostics, location);
        MimeValue disposition = MimeValueParser.Parse(MimeHeaderParser.GetValue(headers, "Content-Disposition"),
            string.Empty, _diagnostics, location);
        string? fileName = disposition.GetParameter("filename") ?? contentType.GetParameter("name");
        bool attachmentDisposition = string.Equals(disposition.Value, "attachment",
            StringComparison.OrdinalIgnoreCase);

        if (contentType.Value.StartsWith("multipart/", StringComparison.OrdinalIgnoreCase) &&
            !attachmentDisposition && string.IsNullOrWhiteSpace(fileName)) {
            string? boundary = contentType.GetParameter("boundary");
            if (boundary == null) return;
            IReadOnlyList<Segment> parts = SplitMultipart(bodyStart, end, boundary, location);
            for (int index = 0; index < parts.Count; index++) {
                Segment part = parts[index];
                string partLocation = string.Concat(location, "/part[",
                    index.ToString(CultureInfo.InvariantCulture), "]");
                HeaderSection child = ReadHeaders(part.Start, part.End, partLocation);
                AnalyzeEntity(child.Headers, child.BodyStart, part.End, depth + 1, partLocation);
            }
            return;
        }

        bool embedded = string.Equals(contentType.Value, "message/rfc822", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(contentType.Value, "message/global", StringComparison.OrdinalIgnoreCase);
        bool semantic = string.Equals(contentType.Value, "text/calendar", StringComparison.OrdinalIgnoreCase) ||
            VCardCodec.IsVCardContentType(contentType.Value, contentType.GetParameter("profile"));
        if ((!attachmentDisposition && string.IsNullOrWhiteSpace(fileName)) || embedded || semantic) return;

        _externalParts.Add(new ExternalPart(bodyStart, end,
            MimeHeaderParser.GetValue(headers, "Content-Transfer-Encoding"), fileName,
            contentType.Value, MimeParser.TrimAngleBrackets(MimeHeaderParser.GetValue(headers, "Content-ID")),
            MimeHeaderParser.GetValue(headers, "Content-Location"), location));
    }

    private HeaderSection ReadHeaders(long start, long end, string location) {
        _input.Position = start;
        var capture = new byte[1];
        long headerEnd = end;
        while (TryReadLine(_input, end, capture, out LineInfo line)) {
            long headerBytes = line.Next - start;
            if (headerBytes > _options.MaxHeaderBytes) {
                throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxHeaderBytes), headerBytes,
                    _options.MaxHeaderBytes);
            }
            if (line.Length == 0) {
                headerEnd = line.Next;
                break;
            }
        }
        int length = checked((int)(headerEnd - start));
        byte[] bytes = ReadRange(start, length);
        var headers = new List<EmailHeader>();
        int bodyOffset = MimeHeaderParser.Parse(bytes, 0, bytes.Length, _options, headers, _diagnostics, location);
        return new HeaderSection(headers, checked(start + bodyOffset));
    }

    private IReadOnlyList<Segment> SplitMultipart(long start, long end, string boundary, string location) {
        byte[] marker = Encoding.ASCII.GetBytes(string.Concat("--", boundary));
        var capture = new byte[checked(marker.Length + 1024)];
        var parts = new List<Segment>();
        long partStart = -1;
        long previousContentEnd = start;
        bool found = false;
        bool closed = false;
        _input.Position = start;
        while (TryReadLine(_input, end, capture, out LineInfo line)) {
            if (IsBoundary(line, capture, marker, out bool closing)) {
                found = true;
                if (partStart >= 0) {
                    if (parts.Count >= _options.MaxPartCount) {
                        throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxPartCount),
                            parts.Count + 1, _options.MaxPartCount);
                    }
                    parts.Add(new Segment(partStart, Math.Max(partStart, previousContentEnd)));
                }
                if (closing) {
                    closed = true;
                    break;
                }
                partStart = line.Next;
                previousContentEnd = partStart;
            } else if (partStart >= 0) {
                previousContentEnd = line.ContentEnd;
            }
        }
        if (!closed && partStart >= 0) {
            parts.Add(new Segment(partStart, Math.Max(partStart, previousContentEnd)));
            _diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BOUNDARY_NOT_CLOSED",
                string.Concat("Multipart boundary '", boundary, "' has no closing delimiter."),
                EmailDiagnosticSeverity.Warning, location));
        }
        if (!found) {
            _diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BOUNDARY_NOT_FOUND",
                string.Concat("Multipart boundary '", boundary, "' was not found."),
                EmailDiagnosticSeverity.Error, location));
        }
        return parts;
    }

    private void DecodeExternalParts() {
        for (int index = 0; index < _externalParts.Count; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            ExternalPart part = _externalParts[index];
            string? path = _options.IncludeAttachmentContent ? _workspace.CreateContentPath() : null;
            try {
                long decodedLength;
                if (path != null) {
                    using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read,
                               81920, FileOptions.SequentialScan)) {
                        decodedLength = DecodePart(part, output);
                    }
                } else {
                    decodedLength = DecodePart(part, Stream.Null);
                }
                EnsureAttachmentLimits(decodedLength);
                part.Length = decodedLength;
                if (path != null) {
                    part.Source = _workspace.RegisterContent(
                        string.Concat("mime/", index.ToString("D8", CultureInfo.InvariantCulture)),
                        path, decodedLength);
                }
            } catch {
                if (path != null && File.Exists(path)) File.Delete(path);
                throw;
            }
        }
    }

    private long DecodePart(ExternalPart part, Stream output) {
        string encoding = (part.TransferEncoding ?? string.Empty).Trim().ToLowerInvariant();
        switch (encoding) {
            case "base64":
                Base64Analysis analysis = AnalyzeBase64(part.Start, part.End);
                if (!analysis.IsValid) {
                    _diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BASE64_INVALID",
                        "The invalid Base64 payload was preserved without decoding.",
                        EmailDiagnosticSeverity.Error, part.Location));
                    return CopyRaw(part.Start, part.End, output);
                }
                if (analysis.RecoveredPadding) {
                    _diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BASE64_PADDING_RECOVERED",
                        "Missing Base64 padding was recovered.", EmailDiagnosticSeverity.Warning, part.Location));
                }
                return DecodeBase64(part.Start, part.End, output);
            case "quoted-printable":
                return DecodeQuotedPrintable(part.Start, part.End, output, part.Location);
            case "7bit":
            case "8bit":
            case "binary":
            case "":
                return CopyRaw(part.Start, part.End, output);
            default:
                _diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_TRANSFER_ENCODING_UNKNOWN",
                    string.Concat("Transfer encoding '", encoding, "' was preserved without decoding."),
                    EmailDiagnosticSeverity.Warning, part.Location));
                return CopyRaw(part.Start, part.End, output);
        }
    }

    private Base64Analysis AnalyzeBase64(long start, long end) {
        _input.Position = start;
        long compact = 0;
        int padding = 0;
        bool sawPadding = false;
        bool valid = true;
        while (_input.Position < end) {
            if ((compact & 0xFFFF) == 0) _cancellationToken.ThrowIfCancellationRequested();
            int current = _input.ReadByte();
            if (current < 0) break;
            byte value = (byte)current;
            if (IsWhiteSpace(value)) continue;
            compact++;
            if (value == '=') {
                sawPadding = true;
                padding++;
            } else if (sawPadding || DecodeBase64Value(value) < 0) {
                valid = false;
                break;
            }
        }
        int remainder = (int)(compact % 4);
        if (padding > 2 || padding > 0 && remainder != 0 || padding == 0 && remainder == 1) valid = false;
        return new Base64Analysis(valid, valid && padding == 0 && (remainder == 2 || remainder == 3));
    }

    private long DecodeBase64(long start, long end, Stream output) {
        _input.Position = start;
        var quartet = new byte[4];
        int count = 0;
        long written = 0;
        while (_input.Position < end) {
            if ((_input.Position - start & 0xFFFF) == 0) _cancellationToken.ThrowIfCancellationRequested();
            int current = _input.ReadByte();
            if (current < 0) break;
            byte value = (byte)current;
            if (IsWhiteSpace(value)) continue;
            quartet[count++] = value;
            if (count == 4) {
                WriteBase64Quartet(quartet, 4, output, ref written);
                count = 0;
            }
        }
        if (count > 0) WriteBase64Quartet(quartet, count, output, ref written);
        return written;
    }

    private void WriteBase64Quartet(byte[] quartet, int count, Stream output, ref long written) {
        int first = DecodeBase64Value(quartet[0]);
        int second = DecodeBase64Value(quartet[1]);
        int third = count > 2 && quartet[2] != '=' ? DecodeBase64Value(quartet[2]) : 0;
        int fourth = count > 3 && quartet[3] != '=' ? DecodeBase64Value(quartet[3]) : 0;
        WriteDecoded(output, (byte)((first << 2) | (second >> 4)), ref written);
        if (count > 2 && quartet[2] != '=') {
            WriteDecoded(output, (byte)((second << 4) | (third >> 2)), ref written);
        }
        if (count > 3 && quartet[3] != '=') {
            WriteDecoded(output, (byte)((third << 6) | fourth), ref written);
        }
    }

    private long DecodeQuotedPrintable(long start, long end, Stream output, string location) {
        _input.Position = start;
        long written = 0;
        bool invalid = false;
        while (_input.Position < end) {
            if ((_input.Position - start & 0xFFFF) == 0) _cancellationToken.ThrowIfCancellationRequested();
            int current = _input.ReadByte();
            if (current < 0) break;
            if (current != '=') {
                WriteDecoded(output, (byte)current, ref written);
                continue;
            }
            long afterEquals = _input.Position;
            int first = _input.Position < end ? _input.ReadByte() : -1;
            if (first == '\n') continue;
            if (first == '\r' && _input.Position < end) {
                int second = _input.ReadByte();
                if (second == '\n') continue;
            }
            _input.Position = afterEquals;
            first = _input.Position < end ? _input.ReadByte() : -1;
            int secondHex = _input.Position < end ? _input.ReadByte() : -1;
            if (first >= 0 && secondHex >= 0 && TryHex((byte)first, out int high) &&
                TryHex((byte)secondHex, out int low)) {
                WriteDecoded(output, (byte)((high << 4) | low), ref written);
                continue;
            }
            _input.Position = afterEquals;
            WriteDecoded(output, (byte)'=', ref written);
            invalid = true;
        }
        if (invalid) {
            _diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_QUOTED_PRINTABLE_INVALID",
                "An invalid quoted-printable escape was preserved.", EmailDiagnosticSeverity.Warning, location));
        }
        return written;
    }

    private long CopyRaw(long start, long end, Stream output) {
        _input.Position = start;
        var buffer = new byte[81920];
        long remaining = end - start;
        long written = 0;
        while (remaining > 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            int read = _input.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
            if (read == 0) throw new EndOfStreamException("The MIME entity ended unexpectedly.");
            written = checked(written + read);
            EnsureCurrentAttachmentLimit(written);
            output.Write(buffer, 0, read);
            remaining -= read;
        }
        return written;
    }

    private void WriteDecoded(Stream output, byte value, ref long written) {
        written = checked(written + 1);
        EnsureCurrentAttachmentLimit(written);
        output.WriteByte(value);
    }

    private void EnsureCurrentAttachmentLimit(long currentLength) {
        if (currentLength > _options.MaxAttachmentBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxAttachmentBytes), currentLength,
                _options.MaxAttachmentBytes);
        }
        long total = checked(_totalAttachmentBytes + currentLength);
        if (total > _options.MaxTotalAttachmentBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxTotalAttachmentBytes), total,
                _options.MaxTotalAttachmentBytes);
        }
    }

    private void EnsureAttachmentLimits(long length) {
        if (length > _options.MaxAttachmentBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxAttachmentBytes), length,
                _options.MaxAttachmentBytes);
        }
        long total = checked(_totalAttachmentBytes + length);
        if (total > _options.MaxTotalAttachmentBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxTotalAttachmentBytes), total,
                _options.MaxTotalAttachmentBytes);
        }
        _totalAttachmentBytes = total;
    }

    private byte[] CreateSkeleton(long start, long end) {
        _externalParts.Sort((left, right) => left.Start.CompareTo(right.Start));
        using (var output = new MemoryStream()) {
            long cursor = start;
            foreach (ExternalPart part in _externalParts) {
                if (part.Start < cursor) throw new InvalidDataException("Streaming MIME attachment ranges overlap.");
                CopyRange(cursor, part.Start, output);
                cursor = part.End;
            }
            CopyRange(cursor, end, output);
            if (output.Length > _options.MaxInputBytes) {
                throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxInputBytes), output.Length,
                    _options.MaxInputBytes);
            }
            return output.ToArray();
        }
    }

    private void CopyRange(long start, long end, Stream output) {
        _input.Position = start;
        var buffer = new byte[81920];
        long remaining = end - start;
        while (remaining > 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            int read = _input.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
            if (read == 0) throw new EndOfStreamException("The MIME artifact ended unexpectedly.");
            output.Write(buffer, 0, read);
            remaining -= read;
        }
    }

    private void ApplyExternalContent(EmailDocument document) {
        int searchStart = 0;
        foreach (ExternalPart part in _externalParts) {
            int match = -1;
            for (int index = searchStart; index < document.Attachments.Count; index++) {
                EmailAttachment candidate = document.Attachments[index];
                if (string.Equals(candidate.FileName, part.FileName, StringComparison.Ordinal) &&
                    string.Equals(candidate.ContentType, part.ContentType, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(candidate.ContentId, part.ContentId, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(candidate.ContentLocation, part.ContentLocation, StringComparison.Ordinal)) {
                    match = index;
                    break;
                }
            }
            if (match < 0) {
                _diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_STREAMING_ATTACHMENT_MAP_FAILED",
                    "A streamed MIME attachment could not be mapped back to the projected document.",
                    EmailDiagnosticSeverity.Error, part.Location));
                continue;
            }
            EmailAttachment attachment = document.Attachments[match];
            attachment.Content = null;
            attachment.ContentSource = part.Source;
            attachment.Length = part.Length;
            searchStart = match + 1;
        }
    }

    private byte[] ReadRange(long start, int count) {
        var result = new byte[count];
        _input.Position = start;
        int total = 0;
        while (total < count) {
            int read = _input.Read(result, total, count - total);
            if (read == 0) throw new EndOfStreamException("The MIME header is truncated.");
            total += read;
        }
        return result;
    }

    private bool TryReadLine(Stream input, long end, byte[] capture, out LineInfo line) {
        if (input.Position >= end) {
            line = default;
            return false;
        }
        long start = input.Position;
        int length = 0;
        int captured = 0;
        while (input.Position < end) {
            if ((length & 0xFFFF) == 0) _cancellationToken.ThrowIfCancellationRequested();
            int current = input.ReadByte();
            if (current < 0) break;
            if (current == '\r' || current == '\n') {
                long contentEnd = input.Position - 1;
                if (current == '\r' && input.Position < end) {
                    int next = input.ReadByte();
                    if (next != '\n') input.Position--;
                }
                line = new LineInfo(start, contentEnd, input.Position, length, captured);
                return true;
            }
            if (captured < capture.Length) capture[captured++] = (byte)current;
            length++;
        }
        line = new LineInfo(start, input.Position, input.Position, length, captured);
        return true;
    }

    private static bool IsBoundary(LineInfo line, byte[] captured, byte[] marker, out bool closing) {
        closing = false;
        if (line.Length < marker.Length || line.CapturedLength != line.Length) return false;
        for (int index = 0; index < marker.Length; index++) {
            if (captured[index] != marker[index]) return false;
        }
        int position = marker.Length;
        if (position + 2 <= line.Length && captured[position] == '-' && captured[position + 1] == '-') {
            closing = true;
            position += 2;
        }
        while (position < line.Length && (captured[position] == ' ' || captured[position] == '\t')) position++;
        return position == line.Length;
    }

    private static int DecodeBase64Value(byte value) {
        if (value >= 'A' && value <= 'Z') return value - 'A';
        if (value >= 'a' && value <= 'z') return value - 'a' + 26;
        if (value >= '0' && value <= '9') return value - '0' + 52;
        if (value == '+') return 62;
        if (value == '/') return 63;
        return -1;
    }

    private static bool IsWhiteSpace(byte value) =>
        value == ' ' || value == '\t' || value == '\r' || value == '\n' || value == '\f' || value == '\v';

    private static bool TryHex(byte value, out int decoded) {
        if (value >= '0' && value <= '9') { decoded = value - '0'; return true; }
        if (value >= 'A' && value <= 'F') { decoded = value - 'A' + 10; return true; }
        if (value >= 'a' && value <= 'f') { decoded = value - 'a' + 10; return true; }
        decoded = 0;
        return false;
    }

    private readonly struct HeaderSection {
        internal HeaderSection(IReadOnlyList<EmailHeader> headers, long bodyStart) {
            Headers = headers;
            BodyStart = bodyStart;
        }
        internal IReadOnlyList<EmailHeader> Headers { get; }
        internal long BodyStart { get; }
    }

    private readonly struct Segment {
        internal Segment(long start, long end) { Start = start; End = end; }
        internal long Start { get; }
        internal long End { get; }
    }

    private readonly struct LineInfo {
        internal LineInfo(long start, long contentEnd, long next, int length, int capturedLength) {
            Start = start;
            ContentEnd = contentEnd;
            Next = next;
            Length = length;
            CapturedLength = capturedLength;
        }
        internal long Start { get; }
        internal long ContentEnd { get; }
        internal long Next { get; }
        internal int Length { get; }
        internal int CapturedLength { get; }
    }

    private sealed class ExternalPart {
        internal ExternalPart(long start, long end, string? transferEncoding, string? fileName,
            string contentType, string? contentId, string? contentLocation, string location) {
            Start = start;
            End = end;
            TransferEncoding = transferEncoding;
            FileName = fileName;
            ContentType = contentType;
            ContentId = contentId;
            ContentLocation = contentLocation;
            Location = location;
        }
        internal long Start { get; }
        internal long End { get; }
        internal string? TransferEncoding { get; }
        internal string? FileName { get; }
        internal string ContentType { get; }
        internal string? ContentId { get; }
        internal string? ContentLocation { get; }
        internal string Location { get; }
        internal long Length { get; set; }
        internal IEmailContentSource? Source { get; set; }
    }

    private readonly struct Base64Analysis {
        internal Base64Analysis(bool isValid, bool recoveredPadding) {
            IsValid = isValid;
            RecoveredPadding = recoveredPadding;
        }
        internal bool IsValid { get; }
        internal bool RecoveredPadding { get; }
    }
}
