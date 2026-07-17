using OfficeIMO.Drawing.Internal;
using OfficeIMO.Email;
using System.Xml;

namespace OfficeIMO.Email.Store;

/// <summary>Writes Apple Mail EMLX envelopes over the OfficeIMO.Email EML engine.</summary>
public sealed class EmailStoreEmlxWriter {
    private readonly EmailStoreEmlxWriterOptions _options;

    /// <summary>Creates a writer with the default policy.</summary>
    public EmailStoreEmlxWriter() : this(EmailStoreEmlxWriterOptions.Default) { }

    /// <summary>Creates a writer with an explicit policy.</summary>
    public EmailStoreEmlxWriter(EmailStoreEmlxWriterOptions options) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    /// <summary>Writer policy.</summary>
    public EmailStoreEmlxWriterOptions Options => _options;

    /// <summary>Serializes one message to complete EMLX bytes.</summary>
    public byte[] ToBytes(EmailDocument document) {
        byte[] bytes = Create(document, out EmailWriteResult result);
        if (!result.HasErrors) return bytes;
        EmailDiagnostic error = result.Diagnostics.First(diagnostic =>
            diagnostic.Severity == EmailDiagnosticSeverity.Error);
        throw new InvalidDataException("The EMLX artifact could not be serialized: " +
            error.Code + ": " + error.Message);
    }

    /// <summary>Atomically writes one EMLX file.</summary>
    public EmailWriteResult Write(EmailDocument document, string filePath) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        byte[] bytes = Create(document, out EmailWriteResult result);
        if (result.HasErrors) return result;
        OfficeFileCommit.WriteAllBytes(filePath, bytes);
        return result;
    }

    /// <summary>Writes one EMLX artifact to a caller-owned stream without closing it.</summary>
    public EmailWriteResult Write(EmailDocument document, Stream stream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        byte[] bytes = Create(document, out EmailWriteResult result);
        if (result.HasErrors) return result;
        OfficeStreamWriter.WriteAllBytes(stream, bytes);
        return result;
    }

    /// <summary>Asynchronously atomically writes one EMLX file.</summary>
    public async Task<EmailWriteResult> WriteAsync(EmailDocument document, string filePath,
        CancellationToken cancellationToken = default) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = Create(document, out EmailWriteResult result);
        if (result.HasErrors) return result;
        await OfficeFileCommit.WriteAllBytesAsync(filePath, bytes, cancellationToken: cancellationToken)
            .ConfigureAwait(false);
        return result;
    }

    /// <summary>Asynchronously writes to a caller-owned stream without closing it.</summary>
    public async Task<EmailWriteResult> WriteAsync(EmailDocument document, Stream stream,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = Create(document, out EmailWriteResult result);
        if (result.HasErrors) return result;
        await OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
        return result;
    }

    private byte[] Create(EmailDocument document, out EmailWriteResult result) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        byte[] metadata = _options.IncludeMetadata
            ? CreateMetadata(document, _options.MaxOutputBytes)
            : Array.Empty<byte>();
        long fixedBytes = checked(metadata.LongLength + (metadata.Length > 0 ? 1L : 0L) + 2L);
        if (fixedBytes >= _options.MaxOutputBytes) {
            throw new EmailLimitExceededException(
                nameof(EmailStoreEmlxWriterOptions.MaxOutputBytes),
                checked(fixedBytes + 1L), _options.MaxOutputBytes);
        }
        long messageBudget = Math.Min(_options.MessageOptions.MaxOutputBytes,
            Math.Min(_options.MaxOutputBytes - fixedBytes, int.MaxValue));
        EmailWriterOptions sourceOptions = _options.MessageOptions;
        var boundedMessageOptions = new EmailWriterOptions(
            sourceOptions.ConversionLossPolicy,
            sourceOptions.UsePreservedRawSource,
            sourceOptions.IncludeBccHeader,
            sourceOptions.Base64LineLength,
            sourceOptions.MaxNestedMessageDepth,
            messageBudget);
        var messageWriter = new EmailDocumentWriter(boundedMessageOptions);
        byte[] message;
        EmailWriteResult messageResult;
        try {
            message = messageWriter.ToBytes(document, EmailFileFormat.Eml, out messageResult);
        } catch (EmailLimitExceededException exception) when (
            exception.LimitName == nameof(EmailWriterOptions.MaxOutputBytes)) {
            throw new EmailLimitExceededException(nameof(EmailStoreEmlxWriterOptions.MaxOutputBytes),
                exception.ActualValue, _options.MaxOutputBytes);
        }
        if (messageResult.HasErrors) {
            result = messageResult;
            return Array.Empty<byte>();
        }
        byte[] prefix = Encoding.ASCII.GetBytes(message.LongLength.ToString(CultureInfo.InvariantCulture) + "\n");
        long total = checked(prefix.LongLength + message.LongLength + metadata.LongLength +
            (metadata.Length > 0 ? 1L : 0L));
        if (total > _options.MaxOutputBytes || total > int.MaxValue)
            throw new EmailLimitExceededException(nameof(EmailStoreEmlxWriterOptions.MaxOutputBytes), total,
                Math.Min(_options.MaxOutputBytes, int.MaxValue));
        var output = new byte[(int)total];
        int offset = 0;
        Buffer.BlockCopy(prefix, 0, output, offset, prefix.Length); offset += prefix.Length;
        Buffer.BlockCopy(message, 0, output, offset, message.Length); offset += message.Length;
        if (metadata.Length > 0) {
            output[offset++] = (byte)'\n';
            Buffer.BlockCopy(metadata, 0, output, offset, metadata.Length);
        }
        result = new EmailWriteResult(total, messageResult.Diagnostics, messageResult.UsedPreservedSource);
        return output;
    }

    private static byte[] CreateMetadata(EmailDocument document, long maxOutputBytes) {
        try {
            using (var output = new EmailBoundedMemoryStream(maxOutputBytes)) {
                var settings = new XmlWriterSettings {
                    Encoding = new UTF8Encoding(false),
                    Indent = true,
                    NewLineChars = "\n",
                    NewLineHandling = NewLineHandling.Replace,
                    CloseOutput = false
                };
                using (XmlWriter writer = XmlWriter.Create(output, settings)) {
                    writer.WriteStartDocument();
                    writer.WriteDocType("plist", "-//Apple//DTD PLIST 1.0//EN",
                        "http://www.apple.com/DTDs/PropertyList-1.0.dtd", null);
                    writer.WriteStartElement("plist");
                    writer.WriteAttributeString("version", "1.0");
                    writer.WriteStartElement("dict");
                    WriteInteger(writer, "flags", CreateFlags(document));
                    WriteDate(writer, "date-received", document.ReceivedDate);
                    WriteDate(writer, "date-sent", document.Date);
                    WriteString(writer, "subject", document.Subject);
                    WriteString(writer, "message-id", document.MessageId);
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
                return output.ToArray();
            }
        } catch (EmailLimitExceededException exception) when (
            exception.LimitName == nameof(EmailWriterOptions.MaxOutputBytes)) {
            throw new EmailLimitExceededException(nameof(EmailStoreEmlxWriterOptions.MaxOutputBytes),
                exception.ActualValue, maxOutputBytes);
        } catch (Exception exception) when (exception is ArgumentException || exception is XmlException) {
            throw new InvalidDataException(
                "The EMLX property-list metadata contains text that XML cannot represent.", exception);
        }
    }

    private static long CreateFlags(EmailDocument document) {
        long flags = 0;
        if (document.MessageMetadata.IsRead == true) flags |= 1L << 0;
        if (PropertyFlag(document, "Emlx:Flag:Deleted")) flags |= 1L << 1;
        if (PropertyFlag(document, "Emlx:Flag:Answered")) flags |= 1L << 2;
        if (PropertyFlag(document, "Emlx:Flag:Encrypted")) flags |= 1L << 3;
        if (PropertyFlag(document, "Emlx:Flag:Flagged")) flags |= 1L << 4;
        if (PropertyFlag(document, "Emlx:Flag:Recent")) flags |= 1L << 5;
        if (document.MessageMetadata.IsDraft) flags |= 1L << 6;
        if (PropertyFlag(document, "Emlx:Flag:Initial")) flags |= 1L << 7;
        if (PropertyFlag(document, "Emlx:Flag:Forwarded")) flags |= 1L << 8;
        if (PropertyFlag(document, "Emlx:Flag:Redirected")) flags |= 1L << 9;
        int attachmentCount = PropertyFlag(document, "Emlx:IsPartial") &&
            TryGetIntegerProperty(document, "Emlx:Flag:AttachmentCount", out int storedAttachmentCount)
            ? storedAttachmentCount
            : document.Attachments.Count;
        flags |= (long)Math.Max(0, Math.Min(attachmentCount, 63)) << 10;
        if (TryGetIntegerProperty(document, "Emlx:Flag:PriorityLevel", out int priorityLevel))
            flags |= (long)Math.Max(0, Math.Min(priorityLevel, 127)) << 16;
        if (PropertyFlag(document, "Emlx:Flag:Signed")) flags |= 1L << 23;
        if (PropertyFlag(document, "Emlx:Flag:IsJunk")) flags |= 1L << 24;
        if (PropertyFlag(document, "Emlx:Flag:IsNotJunk")) flags |= 1L << 25;
        return flags;
    }

    private static bool PropertyFlag(EmailDocument document, string name) =>
        document.Properties.TryGetValue(name, out object? value) && value is bool enabled && enabled;

    private static bool TryGetIntegerProperty(EmailDocument document, string name, out int result) {
        if (document.Properties.TryGetValue(name, out object? value)) {
            if (value is int number) { result = number; return true; }
            if (value is short shortNumber) { result = shortNumber; return true; }
            if (value is byte byteNumber) { result = byteNumber; return true; }
        }
        result = 0;
        return false;
    }

    private static void WriteInteger(XmlWriter writer, string key, long value) {
        writer.WriteElementString("key", key);
        writer.WriteElementString("integer", value.ToString(CultureInfo.InvariantCulture));
    }

    private static void WriteDate(XmlWriter writer, string key, DateTimeOffset? value) {
        if (!value.HasValue) return;
        long seconds = (long)(value.Value.ToUniversalTime() -
            new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero)).TotalSeconds;
        WriteInteger(writer, key, seconds);
    }

    private static void WriteString(XmlWriter writer, string key, string? value) {
        if (string.IsNullOrWhiteSpace(value)) return;
        writer.WriteElementString("key", key);
        writer.WriteElementString("string", value);
    }
}
