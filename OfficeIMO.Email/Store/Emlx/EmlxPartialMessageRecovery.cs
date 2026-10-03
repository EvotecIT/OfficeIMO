using OfficeIMO.Core.Internal;

namespace OfficeIMO.Email.Store;

/// <summary>Restores identified, empty MIME parts from bounded Apple Mail sibling storage as normalized Base64.</summary>
internal sealed class EmlxPartialMessageRecovery : IDisposable {
    private FileStream? _message;
    internal Stream? Message => _message;
    internal int RecoveredParts { get; private set; }
    internal int UnresolvedParts { get; private set; }

    internal static EmlxPartialMessageRecovery Restore(Stream input, EmailStoreReaderOptions options, long propertyLimit,
        Func<string, Stream?>? openPart, IList<EmailStoreDiagnostic> diagnostics, CancellationToken cancellationToken) {
        var result = new EmlxPartialMessageRecovery();
        try {
            result.RestoreCore(input, options, propertyLimit, openPart, diagnostics, cancellationToken);
            return result;
        } catch (EmailLimitExceededException exception) {
            result.Dispose();
            throw new EmailStoreLimitExceededException(
                exception.LimitName == nameof(EmailWriterOptions.MaxOutputBytes)
                    ? nameof(EmailStoreReaderOptions.MaxMessageBytes) : EmailStoreMessageReader.ConvertLimit(exception).LimitName,
                exception.ActualValue, exception.MaximumValue);
        } catch { result.Dispose(); throw; }
    }

    private void RestoreCore(Stream input, EmailStoreReaderOptions options, long propertyLimit, Func<string, Stream?>? openPart,
        IList<EmailStoreDiagnostic> diagnostics, CancellationToken cancellationToken) {
        var mimeDiagnostics = new List<EmailDiagnostic>();
        IReadOnlyList<MimeStreamingParser.MimeSourcePart> parts = MimeStreamingParser.InspectParts(input,
            EmailStoreMessageReader.CreateOptions(options, false, propertyLimit), mimeDiagnostics, cancellationToken);
        // Rewriting ambiguous headers could erase the evidence or reinterpret a protected entity.
        // Keep the original MIME intact whenever its structure cannot establish one safe interpretation.
        bool recoveryBlocked = mimeDiagnostics.Any(diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error ||
            diagnostic.Code == MimeHeaderParser.DuplicateSingletonHeaderDiagnosticCode ||
            diagnostic.Code == MimeValueParser.DuplicateSecurityParameterDiagnosticCode) ||
            parts.Any(part => MimeProtectionProjection.Classify(part.ContentType, string.Empty) != EmailProtectionKind.None);
        long cursor = 0;
        long totalDecoded = 0;
        var buffer = new byte[64 * 1024];
        foreach (MimeStreamingParser.MimeSourcePart part in parts.OrderBy(part => part.BodyStart)) {
            cancellationToken.ThrowIfCancellationRequested();
            EmailHeader[] lengths = part.Headers.Where(header => header.Name.Equals("X-Apple-Content-Length", StringComparison.OrdinalIgnoreCase)).ToArray();
            if (lengths.Length == 0 || !IsEmpty(input, part.BodyStart, part.End, buffer, cancellationToken)) continue;
            if (lengths.Length == 1 && long.TryParse(lengths[0].Value, NumberStyles.None, CultureInfo.InvariantCulture, out long declared) && declared == 0) continue;
            string path = GetPartPath(part.Location);
            string? encoding = MimeHeaderParser.GetValue(part.Headers, "Content-Transfer-Encoding");
            using Stream? content = !recoveryBlocked && lengths.Length == 1 &&
                long.TryParse(lengths[0].Value, NumberStyles.None, CultureInfo.InvariantCulture, out long expected) && expected > 0 &&
                IsSupportedEncoding(encoding) ? openPart?.Invoke(path) : null;
            if (content == null) {
                UnresolvedParts++;
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_EMLX_PART_UNAVAILABLE",
                    "The empty Apple Mail MIME part has no unambiguous supported sibling payload inside the selected root. Protected MIME content is not reconstructed.",
                    EmailStoreDiagnosticSeverity.Warning, path));
                continue;
            }
            long length = content.Length;
            if (length == 0) {
                UnresolvedParts++;
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_EMLX_PART_UNAVAILABLE",
                    "An empty sibling file cannot establish the missing Apple Mail part's content.", EmailStoreDiagnosticSeverity.Warning, path));
                continue;
            }
            if (length > options.MaxAttachmentBytes)
                throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxAttachmentBytes), length, options.MaxAttachmentBytes);
            totalDecoded = checked(totalDecoded + length);
            if (totalDecoded > options.MaxTotalAttachmentBytes)
                throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes), totalDecoded, options.MaxTotalAttachmentBytes);
            if (_message == null) _message = OfficeTemporaryFile.Create("OfficeIMO.Email.PartialEmlx.", ".eml", FileOptions.SequentialScan, out _);
            using (var output = new EmailBoundedWriteStream(_message, options.MaxMessageBytes)) {
                CopyRange(input, cursor, part.HeaderStart, output, buffer, cancellationToken);
                MimeWriter.WriteBase64PartHeaders(output, part.Headers);
                // Apple stores decoded bytes. Its X-Apple-Content-Length is retained as provenance, not treated as a verified decoded size.
                MimeWriter.WriteBase64(output, content, 76, length, cancellationToken);
                if (content.Position != length) throw new InvalidDataException("The Apple Mail sibling payload changed while it was read.");
            }
            cursor = part.End;
            RecoveredParts++;
            diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_EMLX_PART_RECOVERED",
                "The identified decoded sibling payload was restored as Base64; original transfer formatting and payload-dependent integrity headers are not retained.",
                EmailStoreDiagnosticSeverity.Information, path));
        }
        if (_message != null) {
            using (var output = new EmailBoundedWriteStream(_message, options.MaxMessageBytes))
                CopyRange(input, cursor, input.Length, output, buffer, cancellationToken);
            _message.Position = 0;
        }
        input.Position = 0;
    }

    private static bool IsSupportedEncoding(string? encoding) {
        string value = (encoding ?? string.Empty).Trim().ToLowerInvariant();
        return value == "base64" || value == "quoted-printable" || value == "7bit" || value == "8bit" || value == "binary" || value == string.Empty;
    }

    private static bool IsEmpty(Stream input, long start, long end, byte[] buffer, CancellationToken cancellationToken) {
        input.Position = start;
        while (input.Position < end) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = input.Read(buffer, 0, (int)Math.Min(buffer.Length, end - input.Position));
            if (read == 0) throw new EndOfStreamException("The partial MIME entity ended unexpectedly.");
            for (int index = 0; index < read; index++) {
                byte value = buffer[index];
                if (value != '\r' && value != '\n' && value != ' ' && value != '\t') return false;
            }
        }
        return true;
    }

    private static void CopyRange(Stream input, long start, long end, Stream output, byte[] buffer, CancellationToken cancellationToken) {
        if (start > end) throw new InvalidDataException("Apple Mail MIME recovery ranges overlap.");
        input.Position = start;
        while (input.Position < end) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = input.Read(buffer, 0, (int)Math.Min(buffer.Length, end - input.Position));
            if (read == 0) throw new EndOfStreamException("The partial MIME entity ended unexpectedly.");
            output.Write(buffer, 0, read);
        }
    }

    private static string GetPartPath(string location) {
        if (location == "message") return "1";
        return string.Join(".", location.Split('/').Skip(1).Select(part => {
            if (!part.StartsWith("part[", StringComparison.Ordinal) || !part.EndsWith("]", StringComparison.Ordinal) ||
                !int.TryParse(part.Substring(5, part.Length - 6), NumberStyles.None, CultureInfo.InvariantCulture, out int index))
                throw new InvalidDataException("The MIME part identity cannot be mapped to Apple Mail storage.");
            return checked(index + 1).ToString(CultureInfo.InvariantCulture);
        }));
    }

    public void Dispose() { _message?.Dispose(); _message = null; }
}
