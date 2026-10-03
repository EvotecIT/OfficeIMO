using OfficeIMO.Core.Internal;
using OfficeIMO.Email;
using System.Xml;

namespace OfficeIMO.Email.Store;

/// <summary>Writes Apple Mail EMLX envelopes over the OfficeIMO.Email EML engine.</summary>
public sealed class EmailStoreEmlxWriter {
    private const string MetadataPropertyPrefix = "Emlx:Metadata:";
    private static readonly HashSet<string> DerivedMetadataKeys = new HashSet<string>(
        new[] { "flags", "date-received", "date-sent", "subject", "message-id" },
        StringComparer.Ordinal);
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
        using EmlxArtifact artifact = Stage(document);
        if (artifact.Result.HasErrors) {
            EmailDiagnostic error = artifact.Result.Diagnostics.First(diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
            throw new InvalidDataException("The EMLX artifact could not be serialized: " + error.Code + ": " + error.Message);
        }
        using var output = new EmailBoundedMemoryStream(Math.Min(_options.MaxOutputBytes, int.MaxValue));
        artifact.CopyTo(output);
        return output.ToArray();
    }

    /// <summary>Atomically writes one EMLX file using a bounded, owner-only staging file.</summary>
    public EmailWriteResult Write(EmailDocument document, string filePath) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        using EmlxArtifact artifact = Stage(document);
        if (!artifact.Result.HasErrors) OfficeFileCommit.Write(filePath, artifact.CopyTo);
        return artifact.Result;
    }

    /// <summary>Writes one EMLX artifact to a caller-owned stream without closing it.</summary>
    public EmailWriteResult Write(EmailDocument document, Stream stream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        using EmlxArtifact artifact = Stage(document);
        if (!artifact.Result.HasErrors) OfficeStreamWriter.Write(stream, artifact.CopyTo);
        return artifact.Result;
    }

    /// <summary>Asynchronously stages message and attachment I/O and atomically writes one EMLX file.</summary>
    public async Task<EmailWriteResult> WriteAsync(EmailDocument document, string filePath,
        CancellationToken cancellationToken = default) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        using EmlxArtifact artifact = await StageAsync(document, cancellationToken).ConfigureAwait(false);
        if (!artifact.Result.HasErrors) {
            await OfficeFileCommit.WriteAsync(filePath, artifact.CopyToAsync,
                cancellationToken: cancellationToken).ConfigureAwait(false);
        }
        return artifact.Result;
    }

    /// <summary>Asynchronously writes to a caller-owned stream without closing it.</summary>
    public async Task<EmailWriteResult> WriteAsync(EmailDocument document, Stream stream,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        using EmlxArtifact artifact = await StageAsync(document, cancellationToken).ConfigureAwait(false);
        if (!artifact.Result.HasErrors) {
            await OfficeStreamWriter.WriteAsync(stream, artifact.CopyToAsync, cancellationToken).ConfigureAwait(false);
        }
        return artifact.Result;
    }

    private EmlxArtifact Stage(EmailDocument document) {
        byte[] metadata = PrepareMetadata(document, out List<EmailDiagnostic> diagnostics);
        if (diagnostics.Any(item => item.Severity == EmailDiagnosticSeverity.Error)) return Blocked(document, diagnostics);
        FileStream message = OfficeTemporaryFile.Create("OfficeIMO.Email.Emlx.", ".eml", FileOptions.SequentialScan, out _);
        try {
            EmailWriteResult result = CreateMessageWriter(metadata.Length).Write(document, message);
            return Complete(document, message, metadata, diagnostics, result);
        } catch (EmailLimitExceededException exception) when (exception.LimitName == nameof(EmailWriterOptions.MaxOutputBytes)) {
            message.Dispose();
            throw new EmailLimitExceededException(nameof(EmailStoreEmlxWriterOptions.MaxOutputBytes),
                exception.ActualValue, _options.MaxOutputBytes);
        } catch { message.Dispose(); throw; }
    }

    private async Task<EmlxArtifact> StageAsync(EmailDocument document, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] metadata = PrepareMetadata(document, out List<EmailDiagnostic> diagnostics);
        cancellationToken.ThrowIfCancellationRequested();
        if (diagnostics.Any(item => item.Severity == EmailDiagnosticSeverity.Error)) return Blocked(document, diagnostics);
        FileStream message = OfficeTemporaryFile.Create("OfficeIMO.Email.Emlx.", ".eml",
            FileOptions.Asynchronous | FileOptions.SequentialScan, out _);
        try {
            EmailWriteResult result = await CreateMessageWriter(metadata.Length).WriteAsync(document, message,
                cancellationToken: cancellationToken).ConfigureAwait(false);
            return Complete(document, message, metadata, diagnostics, result);
        } catch (EmailLimitExceededException exception) when (exception.LimitName == nameof(EmailWriterOptions.MaxOutputBytes)) {
            message.Dispose();
            throw new EmailLimitExceededException(nameof(EmailStoreEmlxWriterOptions.MaxOutputBytes),
                exception.ActualValue, _options.MaxOutputBytes);
        } catch { message.Dispose(); throw; }
    }

    private EmailDocumentWriter CreateMessageWriter(int metadataLength) {
        long fixedBytes = metadataLength + (metadataLength > 0 ? 1L : 0L) + 2L;
        if (fixedBytes >= _options.MaxOutputBytes) {
            throw new EmailLimitExceededException(nameof(EmailStoreEmlxWriterOptions.MaxOutputBytes),
                fixedBytes + 1L, _options.MaxOutputBytes);
        }
        EmailWriterOptions source = _options.MessageOptions;
        return new EmailDocumentWriter(new EmailWriterOptions(source.ConversionLossPolicy, source.UsePreservedRawSource,
            source.IncludeBccHeader, source.Base64LineLength, source.MaxNestedMessageDepth,
            Math.Min(source.MaxOutputBytes, _options.MaxOutputBytes - fixedBytes)), preservesAppleMailMetadata: true);
    }

    private EmlxArtifact Complete(EmailDocument document, FileStream message, byte[] metadata,
        List<EmailDiagnostic> diagnostics, EmailWriteResult messageResult) {
        diagnostics.AddRange(messageResult.Diagnostics);
        if (messageResult.HasErrors) {
            message.Dispose();
            return Blocked(document, diagnostics);
        }
        byte[] prefix = Encoding.ASCII.GetBytes(message.Length.ToString(CultureInfo.InvariantCulture) + "\n");
        long total = checked(prefix.LongLength + message.Length + metadata.LongLength + (metadata.Length > 0 ? 1L : 0L));
        if (total > _options.MaxOutputBytes) {
            throw new EmailLimitExceededException(nameof(EmailStoreEmlxWriterOptions.MaxOutputBytes), total, _options.MaxOutputBytes);
        }
        bool loss = messageResult.LossDisposition == EmailConversionLossDisposition.Accepted ||
            diagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
        return new EmlxArtifact(message, prefix, metadata, new EmailWriteResult(total, document.Format, EmailFileFormat.Emlx,
            diagnostics.AsReadOnly(), EmailArtifactSourceSelection.Regenerated,
            loss ? EmailConversionLossDisposition.Accepted : EmailConversionLossDisposition.None,
            messageResult.AttachmentContentLifetime));
    }

    private static EmlxArtifact Blocked(EmailDocument document, List<EmailDiagnostic> diagnostics) =>
        new EmlxArtifact(null, Array.Empty<byte>(), Array.Empty<byte>(),
            new EmailWriteResult(0, document.Format, EmailFileFormat.Emlx, diagnostics.AsReadOnly(),
                EmailArtifactSourceSelection.None, EmailConversionLossDisposition.Blocked, EmailAttachmentContentLifetime.NotAccessed));

    private byte[] PrepareMetadata(EmailDocument document, out List<EmailDiagnostic> diagnostics) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        diagnostics = new List<EmailDiagnostic>();
        bool hasMetadata = document.Properties.ContainsKey("Emlx:RawMetadata") || document.Properties.ContainsKey("Emlx:Metadata") ||
            document.Properties.Keys.Any(key => key.StartsWith(MetadataPropertyPrefix, StringComparison.OrdinalIgnoreCase));
        bool opaque = PropertyFlag(document, "Emlx:MetadataOpaque");
        if (hasMetadata && (!_options.IncludeMetadata || opaque)) {
            EmailConversionLossPolicy policy = _options.MessageOptions.ConversionLossPolicy;
            diagnostics.Add(new EmailDiagnostic(opaque && _options.IncludeMetadata ? "EMAIL_EMLX_METADATA_OPAQUE" : "EMAIL_EMLX_METADATA_OMITTED",
                opaque && _options.IncludeMetadata
                    ? "The original opaque metadata trailer is retained without reconciling edited message fields or flags."
                    : "The source Apple Mail metadata trailer is omitted by the writer policy.",
                policy == EmailConversionLossPolicy.Block ? EmailDiagnosticSeverity.Error :
                    policy == EmailConversionLossPolicy.Allow ? EmailDiagnosticSeverity.Information : EmailDiagnosticSeverity.Warning,
                "Emlx:Metadata", opaque && _options.IncludeMetadata ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.Omission));
            if (policy == EmailConversionLossPolicy.Block) return Array.Empty<byte>();
        }
        if (!_options.IncludeMetadata) return Array.Empty<byte>();
        if (opaque && document.Properties.TryGetValue("Emlx:RawMetadata", out object? raw) && raw is byte[] bytes) {
            if (bytes.LongLength > _options.MaxOutputBytes) {
                throw new EmailLimitExceededException(nameof(EmailStoreEmlxWriterOptions.MaxOutputBytes), bytes.LongLength, _options.MaxOutputBytes);
            }
            return bytes;
        }
        return CreateMetadata(document, _options.MaxOutputBytes, _options.MaxMetadataDepth, _options.MaxMetadataProperties);
    }

    private sealed class EmlxArtifact : IDisposable {
        private readonly FileStream? _message;
        private readonly byte[] _prefix;
        private readonly byte[] _metadata;
        internal EmlxArtifact(FileStream? message, byte[] prefix, byte[] metadata, EmailWriteResult result) {
            _message = message;
            _prefix = prefix;
            _metadata = metadata;
            Result = result;
        }
        internal EmailWriteResult Result { get; }
        internal void CopyTo(Stream output) {
            output.Write(_prefix, 0, _prefix.Length);
            _message!.Position = 0;
            _message.CopyTo(output, 81920);
            if (_metadata.Length == 0) return;
            output.WriteByte((byte)'\n');
            output.Write(_metadata, 0, _metadata.Length);
        }
        internal async Task CopyToAsync(Stream output, CancellationToken cancellationToken) {
            await output.WriteAsync(_prefix, 0, _prefix.Length, cancellationToken).ConfigureAwait(false);
            _message!.Position = 0;
            await _message.CopyToAsync(output, 81920, cancellationToken).ConfigureAwait(false);
            if (_metadata.Length == 0) return;
            await output.WriteAsync(new byte[] { (byte)'\n' }, 0, 1, cancellationToken).ConfigureAwait(false);
            await output.WriteAsync(_metadata, 0, _metadata.Length, cancellationToken).ConfigureAwait(false);
        }
        public void Dispose() => _message?.Dispose();
    }

    private static byte[] CreateMetadata(EmailDocument document, long maxOutputBytes,
        int maxMetadataDepth, int maxMetadataProperties) {
        try {
            using (var output = new EmailBoundedMemoryStream(maxOutputBytes)) {
                int metadataPropertyCount = 0;
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
                    foreach (KeyValuePair<string, object?> pair in GetRetainedMetadata(document)) {
                        WriteMetadataKey(writer, pair.Key, ref metadataPropertyCount, maxMetadataProperties);
                        WritePlistValue(writer, pair.Value, pair.Key, depth: 1, maxMetadataDepth,
                            ref metadataPropertyCount, maxMetadataProperties);
                    }
                    WriteInteger(writer, "flags", CreateFlags(document),
                        ref metadataPropertyCount, maxMetadataProperties);
                    WriteDate(writer, "date-received", document.ReceivedDate,
                        ref metadataPropertyCount, maxMetadataProperties);
                    WriteDate(writer, "date-sent", document.Date,
                        ref metadataPropertyCount, maxMetadataProperties);
                    WriteString(writer, "subject", document.Subject,
                        ref metadataPropertyCount, maxMetadataProperties);
                    WriteString(writer, "message-id", document.MessageId,
                        ref metadataPropertyCount, maxMetadataProperties);
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

    private static IEnumerable<KeyValuePair<string, object?>> GetRetainedMetadata(EmailDocument document) {
        var retained = new Dictionary<string, object?>(StringComparer.Ordinal);
        if (document.Properties.TryGetValue("Emlx:Metadata", out object? catalog) &&
            catalog is IReadOnlyDictionary<string, object?> exact) {
            foreach (KeyValuePair<string, object?> pair in exact) {
                if (!DerivedMetadataKeys.Contains(pair.Key)) retained.Add(pair.Key, pair.Value);
            }
            return retained.OrderBy(pair => pair.Key, StringComparer.Ordinal);
        }
        foreach (KeyValuePair<string, object?> pair in document.Properties) {
            if (!pair.Key.StartsWith(MetadataPropertyPrefix, StringComparison.OrdinalIgnoreCase)) continue;
            string key = pair.Key.Substring(MetadataPropertyPrefix.Length);
            if (key.Length == 0)
                throw new InvalidDataException("A retained EMLX metadata property has an empty plist key.");
            if (DerivedMetadataKeys.Contains(key)) continue;
            retained[key] = pair.Value;
        }
        return retained.OrderBy(pair => pair.Key, StringComparer.Ordinal);
    }

    private static void WritePlistValue(XmlWriter writer, object? value, string path, int depth,
        int maxMetadataDepth, ref int metadataPropertyCount, int maxMetadataProperties) {
        if (depth > maxMetadataDepth) {
            throw new InvalidDataException(
                "Retained EMLX metadata exceeds the supported plist nesting depth at '" + path + "'.");
        }
        if (value is string text) {
            writer.WriteElementString("string", text);
        } else if (value is bool boolean) {
            writer.WriteStartElement(boolean ? "true" : "false");
            writer.WriteEndElement();
        } else if (value is long integer) {
            writer.WriteElementString("integer", integer.ToString(CultureInfo.InvariantCulture));
        } else if (value is double real) {
            writer.WriteElementString("real", real.ToString("R", CultureInfo.InvariantCulture));
        } else if (value is DateTimeOffset date) {
            writer.WriteElementString("date", date.ToUniversalTime().ToString(
                "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF'Z'", CultureInfo.InvariantCulture));
        } else if (value is byte[] data) {
            writer.WriteElementString("data", Convert.ToBase64String(data));
        } else if (value is IReadOnlyDictionary<string, object?> dictionary) {
            writer.WriteStartElement("dict");
            var keys = new HashSet<string>(StringComparer.Ordinal);
            foreach (KeyValuePair<string, object?> pair in dictionary.OrderBy(
                pair => pair.Key, StringComparer.Ordinal)) {
                if (!keys.Add(pair.Key)) {
                    throw new InvalidDataException(
                        "Retained EMLX metadata contains duplicate nested plist keys at '" + path + "'.");
                }
                WriteMetadataKey(writer, pair.Key, ref metadataPropertyCount, maxMetadataProperties);
                WritePlistValue(writer, pair.Value, path + "." + pair.Key, depth + 1,
                    maxMetadataDepth, ref metadataPropertyCount, maxMetadataProperties);
            }
            writer.WriteEndElement();
        } else if (value is object?[] array) {
            writer.WriteStartElement("array");
            for (int index = 0; index < array.Length; index++) {
                IncrementMetadataPropertyCount(ref metadataPropertyCount, maxMetadataProperties);
                WritePlistValue(writer, array[index], path + "[" +
                    index.ToString(CultureInfo.InvariantCulture) + "]", depth + 1,
                    maxMetadataDepth, ref metadataPropertyCount, maxMetadataProperties);
            }
            writer.WriteEndElement();
        } else {
            throw new InvalidDataException("The retained EMLX metadata value at '" + path +
                "' has unsupported type '" + (value?.GetType().FullName ?? "null") + "'.");
        }
    }

    private static long CreateFlags(EmailDocument document) {
        long original = 0;
        if (document.Properties.TryGetValue("Emlx:Metadata", out object? catalog) &&
            catalog is IReadOnlyDictionary<string, object?> exact) {
            if (exact.TryGetValue("flags", out object? stored) && stored is long retainedFlags) original = retainedFlags;
        } else if (document.Properties.TryGetValue("Emlx:Metadata:flags", out object? flat) && flat is long flatFlags) original = flatFlags;
        long flags = original & ~((1L << 26) - 1);
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

    private static void WriteInteger(XmlWriter writer, string key, long value,
        ref int metadataPropertyCount, int maxMetadataProperties) {
        WriteMetadataKey(writer, key, ref metadataPropertyCount, maxMetadataProperties);
        writer.WriteElementString("integer", value.ToString(CultureInfo.InvariantCulture));
    }

    private static void WriteDate(XmlWriter writer, string key, DateTimeOffset? value,
        ref int metadataPropertyCount, int maxMetadataProperties) {
        if (!value.HasValue) return;
        long seconds = (long)(value.Value.ToUniversalTime() -
            new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero)).TotalSeconds;
        WriteInteger(writer, key, seconds, ref metadataPropertyCount, maxMetadataProperties);
    }

    private static void WriteString(XmlWriter writer, string key, string? value,
        ref int metadataPropertyCount, int maxMetadataProperties) {
        if (string.IsNullOrWhiteSpace(value)) return;
        WriteMetadataKey(writer, key, ref metadataPropertyCount, maxMetadataProperties);
        writer.WriteElementString("string", value);
    }

    private static void WriteMetadataKey(XmlWriter writer, string key,
        ref int metadataPropertyCount, int maxMetadataProperties) {
        IncrementMetadataPropertyCount(ref metadataPropertyCount, maxMetadataProperties);
        writer.WriteElementString("key", key);
    }

    private static void IncrementMetadataPropertyCount(ref int metadataPropertyCount,
        int maxMetadataProperties) {
        metadataPropertyCount++;
        if (metadataPropertyCount > maxMetadataProperties) {
            throw new EmailStoreLimitExceededException(
                nameof(EmailStoreEmlxWriterOptions.MaxMetadataProperties),
                metadataPropertyCount, maxMetadataProperties);
        }
    }
}
