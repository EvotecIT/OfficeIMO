namespace OfficeIMO.Email;

/// <summary>Externalizes ordinary TNEF attachment-data attributes while retaining semantic parsing.</summary>
internal sealed class TnefStreamingParser {
    private readonly Stream _input;
    private readonly EmailReaderOptions _options;
    private readonly IList<EmailDiagnostic> _diagnostics;
    private readonly CancellationToken _cancellationToken;
    private readonly EmailReadWorkspace _workspace;
    private readonly List<ExternalAttribute> _external = new List<ExternalAttribute>();
    private long _totalAttachmentBytes;

    private TnefStreamingParser(Stream input, EmailReaderOptions options,
        IList<EmailDiagnostic> diagnostics, CancellationToken cancellationToken, EmailReadWorkspace workspace) {
        _input = input;
        _options = options;
        _diagnostics = diagnostics;
        _cancellationToken = cancellationToken;
        _workspace = workspace;
    }

    internal static EmailDocument Parse(Stream input, EmailReaderOptions options,
        IList<EmailDiagnostic> diagnostics, CancellationToken cancellationToken, EmailReadWorkspace workspace) {
        var parser = new TnefStreamingParser(input, options, diagnostics, cancellationToken, workspace);
        byte[] skeleton = parser.CreateSkeleton();
        EmailDocument document = TnefReader.Read(skeleton, options, diagnostics, cancellationToken);
        parser.ApplyExternalContent(document);
        return document;
    }

    private byte[] CreateSkeleton() {
        long start = _input.Position;
        long end = _input.Length;
        if (end - start < 6) return ReadRange(start, checked((int)(end - start)));
        using (var output = new MemoryStream()) {
            CopyRange(start, start + 6, output);
            _input.Position = start + 6;
            int attributeCount = 0;
            int attachmentIndex = -1;
            while (_input.Position < end) {
                _cancellationToken.ThrowIfCancellationRequested();
                long attributeStart = _input.Position;
                if (end - attributeStart < 9) {
                    CopyRange(attributeStart, end, output);
                    break;
                }
                var header = ReadRange(attributeStart, 9);
                byte level = header[0];
                uint tag = ReadUInt32(header, 1);
                uint length = ReadUInt32(header, 5);
                long dataStart = checked(attributeStart + 9);
                long checksumOffset = checked(dataStart + length);
                if (checksumOffset > end - 2) {
                    CopyRange(attributeStart, end, output);
                    break;
                }
                attributeCount++;
                if (attributeCount > _options.MaxTnefAttributeCount) {
                    throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxTnefAttributeCount),
                        attributeCount, _options.MaxTnefAttributeCount);
                }
                if (level == (byte)TnefAttributeLevel.Attachment && tag == TnefConstants.AttachRendData) {
                    attachmentIndex++;
                } else if (level == (byte)TnefAttributeLevel.Attachment && attachmentIndex < 0) {
                    attachmentIndex = 0;
                }

                if (level == (byte)TnefAttributeLevel.Attachment && tag == TnefConstants.AttachData) {
                    Externalize(output, attachmentIndex, dataStart, length, checksumOffset);
                } else {
                    CopyRange(attributeStart, checksumOffset + 2, output);
                }
                _input.Position = checksumOffset + 2;
            }
            return output.ToArray();
        }
    }

    private void Externalize(Stream skeleton, int attachmentIndex, long dataStart, uint length,
        long checksumOffset) {
        EnsureAttachmentLimits(length);
        string? path = _options.IncludeAttachmentContent ? _workspace.CreateContentPath() : null;
        ushort checksum = 0;
        try {
            Stream output = path == null
                ? Stream.Null
                : new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read,
                    81920, FileOptions.SequentialScan);
            using (output) {
                _input.Position = dataStart;
                var buffer = new byte[81920];
                long remaining = length;
                uint sum = 0;
                while (remaining > 0) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    int read = _input.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
                    if (read == 0) throw new EndOfStreamException("The TNEF attachment data is truncated.");
                    for (int index = 0; index < read; index++) sum += buffer[index];
                    output.Write(buffer, 0, read);
                    remaining -= read;
                }
                checksum = unchecked((ushort)sum);
            }
            byte[] storedBytes = ReadRange(checksumOffset, 2);
            ushort storedChecksum = (ushort)(storedBytes[0] | (storedBytes[1] << 8));
            if (storedChecksum != checksum) {
                _diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_CHECKSUM_MISMATCH",
                    string.Concat("Attribute 0x", TnefConstants.AttachData.ToString("X8", CultureInfo.InvariantCulture),
                        " has an invalid checksum."), EmailDiagnosticSeverity.Warning, "tnef"));
            }

            IEmailContentSource? source = path == null ? null : _workspace.RegisterContent(
                string.Concat("tnef/", _external.Count.ToString("D8", CultureInfo.InvariantCulture)),
                path, length);
            _external.Add(new ExternalAttribute(attachmentIndex, length, source));

            skeleton.WriteByte((byte)TnefAttributeLevel.Attachment);
            WriteUInt32(skeleton, TnefConstants.AttachData);
            WriteUInt32(skeleton, 0);
            skeleton.WriteByte(0);
            skeleton.WriteByte(0);
        } catch {
            if (path != null && File.Exists(path)) File.Delete(path);
            throw;
        }
    }

    private void ApplyExternalContent(EmailDocument document) {
        foreach (ExternalAttribute external in _external) {
            if (external.AttachmentIndex < 0 || external.AttachmentIndex >= document.Attachments.Count) {
                _diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_STREAMING_ATTACHMENT_MAP_FAILED",
                    "A streamed TNEF attachment could not be mapped back to the projected document.",
                    EmailDiagnosticSeverity.Error, "tnef"));
                continue;
            }
            EmailAttachment attachment = document.Attachments[external.AttachmentIndex];
            attachment.Content = null;
            attachment.ContentSource = external.Source;
            attachment.Length = external.Length;
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

    private void CopyRange(long start, long end, Stream output) {
        _input.Position = start;
        var buffer = new byte[81920];
        long remaining = end - start;
        while (remaining > 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            int read = _input.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
            if (read == 0) throw new EndOfStreamException("The TNEF artifact ended unexpectedly.");
            output.Write(buffer, 0, read);
            remaining -= read;
        }
    }

    private byte[] ReadRange(long start, int count) {
        _input.Position = start;
        var result = new byte[count];
        int total = 0;
        while (total < count) {
            int read = _input.Read(result, total, count - total);
            if (read == 0) throw new EndOfStreamException("The TNEF artifact ended unexpectedly.");
            total += read;
        }
        return result;
    }

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24);

    private static void WriteUInt32(Stream output, uint value) {
        output.WriteByte((byte)value);
        output.WriteByte((byte)(value >> 8));
        output.WriteByte((byte)(value >> 16));
        output.WriteByte((byte)(value >> 24));
    }

    private sealed class ExternalAttribute {
        internal ExternalAttribute(int attachmentIndex, long length, IEmailContentSource? source) {
            AttachmentIndex = attachmentIndex;
            Length = length;
            Source = source;
        }
        internal int AttachmentIndex { get; }
        internal long Length { get; }
        internal IEmailContentSource? Source { get; }
    }
}
