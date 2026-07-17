using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Email;

internal static class TnefWriter {
    private static readonly Guid IidMessage = new Guid("00020307-0000-0000-C000-000000000046");
    private static readonly Guid IidStorage = new Guid("0000000B-0000-0000-C000-000000000046");
    private static readonly HashSet<uint> ManagedMessageAttributes = new HashSet<uint> {
        TnefConstants.TnefVersion, TnefConstants.OemCodePage, TnefConstants.MessageClass,
        TnefConstants.Subject, TnefConstants.Body, TnefConstants.DateSent, TnefConstants.DateReceived,
        TnefConstants.MessageId, TnefConstants.MessageProperties, TnefConstants.RecipientTable
    };
    private static readonly HashSet<uint> ManagedAttachmentAttributes = new HashSet<uint> {
        TnefConstants.AttachRendData, TnefConstants.AttachTitle, TnefConstants.AttachTransportFilename,
        TnefConstants.AttachData, TnefConstants.AttachmentProperties
    };

    internal static byte[] Write(EmailDocument document, EmailWriterOptions options, IList<EmailDiagnostic> diagnostics) {
        using (var output = new EmailBoundedMemoryStream(options.MaxOutputBytes)) {
            Write(output, document, options, diagnostics);
            return output.ToArray();
        }
    }

    internal static void Write(Stream output, EmailDocument document, EmailWriterOptions options,
        IList<EmailDiagnostic> diagnostics) => WriteMessage(output, document, options, diagnostics, 0);

    /// <summary>Returns whether raw TNEF attributes exist that only the TNEF writer can reproduce.</summary>
    internal static bool HasUnmanagedRawAttributes(EmailDocument document) =>
        document.TnefAttributes.Any(attribute => !ManagedMessageAttributes.Contains(attribute.Tag)) ||
        document.Attachments.Any(attachment => attachment.TnefAttributes.Any(attribute =>
            !ManagedAttachmentAttributes.Contains(attribute.Tag)));

    private static void WriteMessage(Stream output, EmailDocument document, EmailWriterOptions options,
        IList<EmailDiagnostic> diagnostics, int depth) {
        if (depth > options.MaxNestedMessageDepth) throw new InvalidOperationException("The embedded-message write depth exceeds the configured maximum.");
        int codePage = MapiStringEncodingContext.FromCodePage(document.OutlookCodePage ?? 65001).PrimaryCodePage;
        WriteUInt32(output, TnefConstants.Signature);
            WriteUInt16(output, 1);
            WriteAttribute(output, TnefAttributeLevel.Message, TnefConstants.TnefVersion, EncodeUInt32(TnefConstants.Version));
            byte[] codePageBytes = new byte[8];
            MsgBinary.WriteUInt32(codePageBytes, 0, unchecked((uint)codePage));
            WriteAttribute(output, TnefAttributeLevel.Message, TnefConstants.OemCodePage, codePageBytes);
            WriteStringAttribute(output, TnefAttributeLevel.Message, TnefConstants.MessageClass, codePage,
                document.MessageClass ?? DefaultMessageClass(document.OutlookItemKind), diagnostics, "tnef/message-class");
            WriteStringAttribute(output, TnefAttributeLevel.Message, TnefConstants.Subject, codePage, document.Subject,
                diagnostics, "tnef/subject");
            WriteStringAttribute(output, TnefAttributeLevel.Message, TnefConstants.Body, codePage, document.Body.Text,
                diagnostics, "tnef/body");
            if (document.Date.HasValue) WriteAttribute(output, TnefAttributeLevel.Message, TnefConstants.DateSent, EncodeDate(document.Date.Value));
            if (document.ReceivedDate.HasValue) WriteAttribute(output, TnefAttributeLevel.Message, TnefConstants.DateReceived, EncodeDate(document.ReceivedDate.Value));
            WriteStringAttribute(output, TnefAttributeLevel.Message, TnefConstants.MessageId, codePage, document.MessageId,
                diagnostics, "tnef/message-id");

            IReadOnlyList<MapiProperty>[] rows = document.Recipients
                .Where(recipient => recipient.Kind != EmailRecipientKind.ReplyTo)
                .Select((recipient, index) => MsgWriter.CreateRecipientProperties(recipient, index).Properties)
                .ToArray();
            if (rows.Length > 0) {
                WriteAttribute(output, TnefAttributeLevel.Message, TnefConstants.RecipientTable,
                    TnefMapiCodec.WriteRecipientTable(rows, codePage, diagnostics, "tnef/recipients"));
            }

            MapiProperty[] messageProperties = MsgWriter.CreateMessageProperties(document, diagnostics, "tnef", options).Properties
                .Where(property => !IsMessageAttributeProperty(property.PropertyId)).ToArray();
            if (messageProperties.Length > 0) {
                WriteAttribute(output, TnefAttributeLevel.Message, TnefConstants.MessageProperties,
                    TnefMapiCodec.WriteProperties(messageProperties, codePage, diagnostics, "tnef/mapi"));
            }
            foreach (TnefAttribute attribute in document.TnefAttributes.Where(attribute => !ManagedMessageAttributes.Contains(attribute.Tag))) {
                WriteAttribute(output, TnefAttributeLevel.Message, attribute.Tag, attribute.Data);
            }

            EmailAttachment[] writableAttachments = document.Attachments.Where(attachment =>
                !attachment.IsProjectedSemanticContent).ToArray();
            for (int index = 0; index < writableAttachments.Length; index++) {
                WriteAttachment(output, writableAttachments[index], index, codePage, options, diagnostics, depth);
            }
    }

    private static void WriteAttachment(Stream output, EmailAttachment attachment, int index, int codePage,
        EmailWriterOptions options, IList<EmailDiagnostic> diagnostics, int depth) {
        int method = attachment.MapiAttachMethod ?? (attachment.EmbeddedDocument != null ? 5 :
            attachment.StructuredStorageStreams.Count > 0 ? 6 : 1);
        bool hasContent = attachment.Content != null || attachment.ContentSource != null ||
            EmailAttachmentStreamScope.HasStagedContent(attachment);
        byte[]? content = method == 1 ? null : EmailAttachmentContent.ReadOrNull(attachment, options.MaxOutputBytes);
        byte[] rendition = new byte[14];
        MsgBinary.WriteUInt16(rendition, 0, method == 6 ? (ushort)2 : (ushort)1);
        MsgBinary.WriteUInt32(rendition, 2, attachment.IsInline ? 0U : 0xffffffffU);
        WriteAttribute(output, TnefAttributeLevel.Attachment, TnefConstants.AttachRendData, rendition);
        string attachmentLocation = string.Concat("tnef/attachment[", index.ToString(CultureInfo.InvariantCulture), "]");
        WriteStringAttribute(output, TnefAttributeLevel.Attachment, TnefConstants.AttachTitle, codePage,
            attachment.FileName, diagnostics, string.Concat(attachmentLocation, "/title"));
        WriteStringAttribute(output, TnefAttributeLevel.Attachment, TnefConstants.AttachTransportFilename, codePage,
            attachment.FileName, diagnostics, string.Concat(attachmentLocation, "/transport-filename"));
        if (method == 1 && hasContent) {
            using (ContentLease lease = OpenContent(attachment, options.MaxOutputBytes)) {
                WriteAttribute(output, TnefAttributeLevel.Attachment, TnefConstants.AttachData,
                    lease.Stream, lease.Length, options.MaxOutputBytes);
            }
        }

        MsgPropertyBuilder builder = MsgWriter.CreateAttachmentProperties(attachment, index, method, diagnostics,
            attachmentLocation, attachment.EmbeddedDocument != null || hasContent, content);
        var properties = builder.Properties.Where(property => !IsAttachmentAttributeProperty(property.PropertyId)).
            Select(Clone).ToList();
        if (method == 5 && attachment.EmbeddedDocument != null) {
            byte[] nested = WriteMessageToBytes(attachment.EmbeddedDocument, options, diagnostics, depth + 1);
            properties.RemoveAll(property => property.PropertyId == 0x3701);
            properties.Add(new MapiProperty(0x3701, MapiPropertyType.Object, Combine(IidMessage.ToByteArray(), nested)));
        } else if (method == 5 && content != null) {
            byte[] opaque = content.Length >= 16 &&
                new Guid(MsgBinary.Slice(content, 0, 16)) == IidMessage
                    ? (byte[])content.Clone()
                    : Combine(IidMessage.ToByteArray(), content);
            properties.RemoveAll(property => property.PropertyId == 0x3701);
            properties.Add(new MapiProperty(0x3701, MapiPropertyType.Object, opaque));
        } else if (method == 6 && attachment.StructuredStorageStreams.Count > 0) {
            OfficeCompoundStream[] compoundStreams = attachment.StructuredStorageStreams
                .Select(stream => new OfficeCompoundStream(stream.Key, stream.Value)).ToArray();
            byte[] compound = OfficeCompoundFileWriter.Write(compoundStreams);
            properties.RemoveAll(property => property.PropertyId == 0x3701);
            properties.Add(new MapiProperty(0x3701, MapiPropertyType.Object, Combine(IidStorage.ToByteArray(), compound)));
        } else if (method == 6 && content != null) {
            byte[] opaque = content.Length >= 16 &&
                new Guid(MsgBinary.Slice(content, 0, 16)) == IidStorage
                    ? (byte[])content.Clone()
                    : Combine(IidStorage.ToByteArray(), content);
            properties.RemoveAll(property => property.PropertyId == 0x3701);
            properties.Add(new MapiProperty(0x3701, MapiPropertyType.Object, opaque));
        }
        if (properties.Count > 0) {
            WriteAttribute(output, TnefAttributeLevel.Attachment, TnefConstants.AttachmentProperties,
                TnefMapiCodec.WriteProperties(properties, codePage, diagnostics,
                    string.Concat("tnef/attachment[", index.ToString(CultureInfo.InvariantCulture), "]/mapi")));
        }
        foreach (TnefAttribute attribute in attachment.TnefAttributes.Where(attribute => !ManagedAttachmentAttributes.Contains(attribute.Tag))) {
            WriteAttribute(output, TnefAttributeLevel.Attachment, attribute.Tag, attribute.Data);
        }
    }

    private static bool IsMessageAttributeProperty(ushort id) {
        return id == 0x001A || id == 0x0037 || id == 0x1000 || id == 0x0039 || id == 0x0E06;
    }

    private static bool IsAttachmentAttributeProperty(ushort id) {
        return id == 0x3701 || id == 0x3704 || id == 0x3707;
    }

    private static MapiProperty Clone(MapiProperty property) {
        return new MapiProperty(property.PropertyId, property.PropertyType, property.Value, property.Flags, property.Name) {
            RawData = property.RawData == null ? null : (byte[])property.RawData.Clone()
        };
    }

    private static void WriteStringAttribute(Stream output, TnefAttributeLevel level, uint tag, int codePage,
        string? value, IList<EmailDiagnostic> diagnostics, string location) {
        if (value == null) return;
        byte[] bytes;
        try {
            bytes = MsgValueWriter.EncodeString8(string.Concat(value, "\0"), codePage);
        } catch (EncoderFallbackException) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_STRING8_CHARACTER_UNENCODABLE",
                string.Concat("Text contains characters that code page ", codePage.ToString(CultureInfo.InvariantCulture),
                    " cannot represent; replacement encoding was used."),
                EmailDiagnosticSeverity.Warning, location));
            bytes = MsgValueWriter.EncodeString8WithReplacement(string.Concat(value, "\0"), codePage);
        }
        WriteAttribute(output, level, tag, bytes);
    }

    private static void WriteAttribute(Stream output, TnefAttributeLevel level, uint tag, byte[] data) {
        output.WriteByte((byte)level);
        WriteUInt32(output, tag);
        WriteUInt32(output, unchecked((uint)data.Length));
        output.Write(data, 0, data.Length);
        WriteUInt16(output, CalculateChecksum(data));
    }

    private static void WriteAttribute(Stream output, TnefAttributeLevel level, uint tag, Stream input,
        long length, long maximumInputBytes) {
        if (length < 0 || length > uint.MaxValue) {
            throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), length,
                Math.Min(maximumInputBytes, uint.MaxValue));
        }
        if (length > maximumInputBytes) {
            throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), length,
                maximumInputBytes);
        }
        output.WriteByte((byte)level);
        WriteUInt32(output, tag);
        WriteUInt32(output, unchecked((uint)length));
        var buffer = new byte[81920];
        long remaining = length;
        uint checksum = 0;
        while (remaining > 0) {
            int read = input.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
            if (read == 0) throw new EndOfStreamException("The attachment content ended before its declared length.");
            output.Write(buffer, 0, read);
            for (int index = 0; index < read; index++) checksum += buffer[index];
            remaining -= read;
        }
        if (input.ReadByte() >= 0) {
            throw new InvalidDataException("The attachment content exceeds its declared length.");
        }
        WriteUInt16(output, unchecked((ushort)checksum));
    }

    private static byte[] WriteMessageToBytes(EmailDocument document, EmailWriterOptions options,
        IList<EmailDiagnostic> diagnostics, int depth) {
        using (var output = new EmailBoundedMemoryStream(options.MaxOutputBytes)) {
            WriteMessage(output, document, options, diagnostics, depth);
            return output.ToArray();
        }
    }

    private static ContentLease OpenContent(EmailAttachment attachment, long maximumInputBytes) {
        long? length = EmailAttachmentStreamScope.GetLength(attachment);
        if (length.HasValue) return new ContentLease(EmailAttachmentStreamScope.OpenRead(attachment), length.Value);

        string path = Path.Combine(Path.GetTempPath(),
            string.Concat("OfficeIMO.Email.Tnef.", Guid.NewGuid().ToString("N"), ".content"));
        try {
            long copied = 0;
            using (Stream input = EmailAttachmentStreamScope.OpenRead(attachment))
            using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read,
                       81920, FileOptions.SequentialScan)) {
                var buffer = new byte[81920];
                while (true) {
                    int read = input.Read(buffer, 0, buffer.Length);
                    if (read == 0) break;
                    copied = checked(copied + read);
                    if (copied > maximumInputBytes) {
                        throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), copied,
                            maximumInputBytes);
                    }
                    output.Write(buffer, 0, read);
                }
            }
            return new ContentLease(new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
                81920, FileOptions.SequentialScan), copied, path);
        } catch {
            OfficeFileCommit.DeleteIfExists(path);
            throw;
        }
    }

    private sealed class ContentLease : IDisposable {
        private readonly string? _temporaryPath;
        internal ContentLease(Stream stream, long length, string? temporaryPath = null) {
            Stream = stream;
            Length = length;
            _temporaryPath = temporaryPath;
        }
        internal Stream Stream { get; }
        internal long Length { get; }
        public void Dispose() {
            Stream.Dispose();
            OfficeFileCommit.DeleteIfExists(_temporaryPath);
        }
    }

    private static byte[] EncodeDate(DateTimeOffset value) {
        DateTime utc = value.UtcDateTime;
        byte[] result = new byte[14];
        MsgBinary.WriteUInt16(result, 0, unchecked((ushort)utc.Year));
        MsgBinary.WriteUInt16(result, 2, unchecked((ushort)utc.Month));
        MsgBinary.WriteUInt16(result, 4, unchecked((ushort)utc.Day));
        MsgBinary.WriteUInt16(result, 6, unchecked((ushort)utc.Hour));
        MsgBinary.WriteUInt16(result, 8, unchecked((ushort)utc.Minute));
        MsgBinary.WriteUInt16(result, 10, unchecked((ushort)utc.Second));
        MsgBinary.WriteUInt16(result, 12, unchecked((ushort)utc.DayOfWeek));
        return result;
    }

    private static byte[] EncodeUInt32(uint value) {
        byte[] result = new byte[4];
        MsgBinary.WriteUInt32(result, 0, value);
        return result;
    }

    private static byte[] Combine(byte[] first, byte[] second) {
        byte[] result = new byte[first.Length + second.Length];
        Buffer.BlockCopy(first, 0, result, 0, first.Length);
        Buffer.BlockCopy(second, 0, result, first.Length, second.Length);
        return result;
    }

    private static ushort CalculateChecksum(byte[] data) {
        uint result = 0;
        for (int index = 0; index < data.Length; index++) result += data[index];
        return unchecked((ushort)result);
    }

    private static void WriteUInt16(Stream output, ushort value) {
        byte[] bytes = new byte[2];
        MsgBinary.WriteUInt16(bytes, 0, value);
        output.Write(bytes, 0, bytes.Length);
    }

    private static void WriteUInt32(Stream output, uint value) {
        byte[] bytes = new byte[4];
        MsgBinary.WriteUInt32(bytes, 0, value);
        output.Write(bytes, 0, bytes.Length);
    }

    private static string DefaultMessageClass(OutlookItemKind kind) {
        switch (kind) {
            case OutlookItemKind.Appointment: return "IPM.Appointment";
            case OutlookItemKind.Contact: return "IPM.Contact";
            case OutlookItemKind.Task: return "IPM.Task";
            case OutlookItemKind.Journal: return "IPM.Activity";
            case OutlookItemKind.Note: return "IPM.StickyNote";
            default: return "IPM.Note";
        }
    }
}
