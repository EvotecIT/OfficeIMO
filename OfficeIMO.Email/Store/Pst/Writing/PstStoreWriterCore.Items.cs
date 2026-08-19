using OfficeIMO.Email;
using System.Security.Cryptography;

namespace OfficeIMO.Email.Store;

internal sealed partial class PstStoreWriterCore {
    private WrittenMessage WriteMessage(EmailDocument document, uint messageNid,
        int depth, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (depth > _options.MaxNestedMessageDepth) {
            throw new InvalidOperationException("The embedded-message depth exceeds the configured PST limit.");
        }

        var emailDiagnostics = new List<EmailDiagnostic>();
        MsgPropertyBuilder messageBuilder = MsgWriter.CreateMessageProperties(
            document, emailDiagnostics, string.Concat("message/", FormatId(messageNid)),
            new EmailWriterOptions(EmailConversionLossPolicy.Allow,
                maxNestedMessageDepth: _options.MaxNestedMessageDepth,
                maxOutputBytes: uint.MaxValue));
        // PidTagMessageSize is calculated from the serialized PC and all of its
        // recursively referenced subnode data after those objects are written.
        messageBuilder.Set(MapiKnownProperties.PidTag.MessageSize, 0);
        messageBuilder.Set(MapiKnownProperties.PidTag.MessageStatus,
            document.MapiProperties.GetNullableMapiValue(MapiKnownProperties.PidTag.MessageStatus) ?? 0);
        messageBuilder.Set(MapiKnownProperties.PidTag.SearchKey,
            document.MapiProperties.GetMapiValueOrDefault(MapiKnownProperties.PidTag.SearchKey) ??
                CreateObjectKey(messageNid));
        messageBuilder.Set(MapiKnownProperties.PidTag.DisplayTo,
            JoinDisplayRecipients(document, EmailRecipientKind.To));
        messageBuilder.Set(MapiKnownProperties.PidTag.DisplayCc,
            JoinDisplayRecipients(document, EmailRecipientKind.Cc));
        messageBuilder.Set(MapiKnownProperties.PidTag.DisplayBcc,
            JoinDisplayRecipients(document, EmailRecipientKind.Bcc));
        // PST-required defaults are added after the shared MSG projection, so restore
        // the public patch as the final preservation authority before serialization.
        document.MapiWritePatch.Apply(messageBuilder);
        // MessageSize is structural PST evidence and is always regenerated from the
        // final serialized PC/subnode graph, even when a preservation patch removes it.
        messageBuilder.Set(MapiKnownProperties.PidTag.MessageSize, 0);
        TranslateDiagnostics(emailDiagnostics);

        IReadOnlyList<MapiProperty> messageProperties =
            _namedProperties.Map(messageBuilder.Properties, Report,
                string.Concat("message/", FormatId(messageNid)));
        int codePage = ResolveCodePage(document);
        var messageSubnodes = new List<PstWriterSubnode>();

        EmailRecipient[] recipients = document.Recipients
            .Where(item => item.Kind != EmailRecipientKind.ReplyTo).ToArray();
        var recipientRows = new List<PstWriterTableRow>(recipients.Length);
        for (int index = 0; index < recipients.Length; index++) {
            MsgPropertyBuilder recipientBuilder = MsgWriter.CreateRecipientProperties(recipients[index], index);
            recipientBuilder.Set(MapiKnownProperties.PidTag.Responsibility, false);
            recipientBuilder.Set(MapiKnownProperties.PidTag.RecordKey,
                recipients[index].MapiProperties.GetMapiValueOrDefault(MapiKnownProperties.PidTag.RecordKey) ??
                    CreateObjectKey(checked((uint)index + 1)));
            recipientBuilder.Set(MapiKnownProperties.PidTag.DisplayNamePrintable, MapiPropertyType.String8,
                recipients[index].Address.DisplayName ?? recipients[index].Address.Address);
            recipientBuilder.Set(MapiKnownProperties.PidTag.SendRichInfo,
                recipients[index].MapiProperties.GetNullableMapiValue(MapiKnownProperties.PidTag.SendRichInfo) ?? false);
            recipientRows.Add(new PstWriterTableRow(checked((uint)index + 1),
                _namedProperties.Map(recipientBuilder.Properties, Report,
                    string.Concat("message/", FormatId(messageNid), "/recipient/",
                        index.ToString(CultureInfo.InvariantCulture)))));
        }
        PstWriterContextResult recipientTable = PstTableContextWriter.Write(_file,
            recipientRows, codePage, RecipientColumns, Report,
            string.Concat("message/", FormatId(messageNid), "/recipients"));
        messageSubnodes.Add(new PstWriterSubnode(0x692,
            recipientTable.DataBid, recipientTable.SubnodeBid));

        EmailAttachment[] attachments = document.Attachments
            .Where(item => !item.IsProjectedSemanticContent).ToArray();
        var attachmentRows = new List<PstWriterTableRow>(attachments.Length);
        for (int index = 0; index < attachments.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            uint attachmentNid = checked(((uint)(index + 0x10) << 5) | 0x05U);
            WrittenAttachment written = WriteAttachment(attachments[index], attachmentNid,
                codePage, depth, cancellationToken);
            messageSubnodes.Add(new PstWriterSubnode(attachmentNid,
                written.Context.DataBid, written.Context.SubnodeBid));
            attachmentRows.Add(new PstWriterTableRow(attachmentNid,
                SelectTableProperties(written.TableProperties, AttachmentColumns)));
        }
        if (attachmentRows.Count > 0) {
            PstWriterContextResult attachmentTable = PstTableContextWriter.Write(_file,
                attachmentRows, codePage, AttachmentColumns, Report,
                string.Concat("message/", FormatId(messageNid), "/attachments"));
            messageSubnodes.Add(new PstWriterSubnode(0x671,
                attachmentTable.DataBid, attachmentTable.SubnodeBid));
        }

        PstWriterContextResult context = PstPropertyContextWriter.WriteWithCalculatedSize(_file,
            messageProperties, codePage, messageSubnodes, null, null,
            Report, string.Concat("message/", FormatId(messageNid)),
            MapiKnownProperties.PidTag.MessageSize);
        var tableProperties = new List<MapiProperty>(SelectTableProperties(
            messageProperties, ContentsColumns.Concat(AssociatedColumns).ToArray()));
        tableProperties.RemoveAll(item => item.PropertyId ==
            MapiKnownProperties.PidTag.MessageSize.GetStandardPropertyId());
        tableProperties.Add(Property(MapiKnownProperties.PidTag.MessageSize,
            checked((int)Math.Min(context.SerializedDataLength, int.MaxValue))));
        tableProperties.Add(Property(0x0E30, MapiPropertyType.Binary,
            CreateReplicaId(messageNid)));
        tableProperties.Add(Property(MapiKnownProperties.PidTag.ReplChangenum,
            checked((long)messageNid)));
        tableProperties.Add(Property(MapiKnownProperties.PidTag.ReplVersionHistory,
            CreateReplVersionHistory()));
        tableProperties.Add(Property(MapiKnownProperties.PidTag.ReplFlags, 0));
        tableProperties.Add(Property(MapiKnownProperties.PidTag.LtpRowVer, 1));
        tableProperties.RemoveAll(item => item.PropertyId ==
            MapiKnownProperties.PidTag.ConversationId.GetStandardPropertyId());
        byte[]? conversationId = CreateConversationId(messageProperties);
        if (conversationId != null) {
            tableProperties.Add(Property(MapiKnownProperties.PidTag.ConversationId,
                conversationId));
        }
        return new WrittenMessage(context, tableProperties);
    }

    private WrittenAttachment WriteAttachment(EmailAttachment attachment, uint attachmentNid,
        int codePage, int parentDepth, CancellationToken cancellationToken) {
        var diagnostics = new List<EmailDiagnostic>();
        int method = attachment.MapiAttachMethod ??
            (attachment.EmbeddedDocument != null ? 5 :
                attachment.StructuredStorageStreams.Count > 0 ? 6 : 1);
        bool hasSource = attachment.Content != null || attachment.ContentSource != null ||
            attachment.EmbeddedDocument != null || attachment.StructuredStorageStreams.Count > 0;
        MsgPropertyBuilder builder = MsgWriter.CreateAttachmentProperties(
            attachment, checked((int)(attachmentNid >> 5)), method, diagnostics,
            string.Concat("attachment/0x", attachmentNid.ToString("X8", CultureInfo.InvariantCulture)),
            hasRetainedObjectContent: hasSource, materializedContent: null);
        var subnodes = new List<PstWriterSubnode>();
        var valueReferences = new Dictionary<ushort, PstWriterValueReference>();
        var objectReferences = new Dictionary<ushort, PstWriterObjectReference>();
        long contentLength = Math.Max(0, attachment.Length);

        if (method == 5 && attachment.EmbeddedDocument != null) {
            if (parentDepth >= _options.MaxNestedMessageDepth) {
                Report(new EmailStoreDiagnostic(
                    "EMAIL_STORE_PST_WRITE_EMBEDDED_DEPTH_LIMIT",
                    "An embedded item could not be written because the configured nesting limit was reached.",
                    EmailStoreDiagnosticSeverity.Error,
                    string.Concat("attachment/0x", attachmentNid.ToString("X8", CultureInfo.InvariantCulture))));
            } else {
                const uint embeddedNid = 0x224;
                WrittenMessage embedded = WriteMessage(attachment.EmbeddedDocument,
                    embeddedNid, parentDepth + 1, cancellationToken);
                contentLength = embedded.Context.SerializedDataLength;
                subnodes.Add(new PstWriterSubnode(embeddedNid,
                    embedded.Context.DataBid, embedded.Context.SubnodeBid));
                objectReferences[MapiKnownProperties.PidTag.AttachData.GetStandardPropertyId()] =
                    new PstWriterObjectReference(embeddedNid, contentLength);
                builder.Set(MapiKnownProperties.PidTag.AttachData, MapiPropertyType.Object, null);
            }
        } else if (method == 5) {
            Report(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_WRITE_EMBEDDED_CONTENT_UNAVAILABLE",
                "An embedded attachment has no projected embedded item and was retained as metadata only.",
                EmailStoreDiagnosticSeverity.Error,
                string.Concat("attachment/0x", attachmentNid.ToString("X8", CultureInfo.InvariantCulture))));
        } else if (method == 6) {
            if (TryWriteAttachmentPayload(attachment, out ulong contentBid,
                out contentLength, cancellationToken)) {
                const uint contentNid = 0x3F;
                subnodes.Add(new PstWriterSubnode(contentNid, contentBid));
                objectReferences[MapiKnownProperties.PidTag.AttachData.GetStandardPropertyId()] =
                    new PstWriterObjectReference(contentNid, contentLength);
                builder.Set(MapiKnownProperties.PidTag.AttachData, MapiPropertyType.Object, null);
            } else if (attachment.StructuredStorageStreams.Count > 0) {
                Report(new EmailStoreDiagnostic(
                    "EMAIL_STORE_PST_WRITE_STRUCTURED_STORAGE_OMITTED",
                    "Structured-storage attachment streams require an original compound payload and were retained as metadata only.",
                    EmailStoreDiagnosticSeverity.Error,
                    string.Concat("attachment/0x", attachmentNid.ToString("X8", CultureInfo.InvariantCulture))));
            }
        } else if (TryWriteAttachmentPayload(attachment, out ulong contentBid,
            out contentLength, cancellationToken)) {
            const uint contentNid = 0x3F;
            valueReferences[MapiKnownProperties.PidTag.AttachData.GetStandardPropertyId()] =
                new PstWriterValueReference(contentNid, contentBid);
            builder.Set(MapiKnownProperties.PidTag.AttachData, Array.Empty<byte>());
        } else if (attachment.Length > 0) {
            Report(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_WRITE_ATTACHMENT_CONTENT_UNAVAILABLE",
                "Attachment content was unavailable and only its metadata could be written.",
                EmailStoreDiagnosticSeverity.Error,
                string.Concat("attachment/0x", attachmentNid.ToString("X8", CultureInfo.InvariantCulture))));
        }
        builder.Set(MapiKnownProperties.PidTag.AttachSize, 0);
        TranslateDiagnostics(diagnostics);
        string location = string.Concat("attachment/0x",
            attachmentNid.ToString("X8", CultureInfo.InvariantCulture));
        IReadOnlyList<MapiProperty> properties = _namedProperties.Map(
            builder.Properties, Report, location);
        var finalProperties = new MsgPropertyBuilder(properties);
        finalProperties.Set(MapiKnownProperties.PidTag.AttachSize,
            CalculateAttachmentObjectSize(properties, codePage, contentLength,
                valueReferences, objectReferences));
        properties = finalProperties.Properties;
        PstWriterContextResult context = PstPropertyContextWriter.Write(_file,
            properties, codePage, subnodes, valueReferences, objectReferences,
            Report, string.Concat("attachment/0x", attachmentNid.ToString("X8", CultureInfo.InvariantCulture)));
        return new WrittenAttachment(context, properties);
    }

    private bool TryWriteAttachmentPayload(EmailAttachment attachment,
        out ulong dataBid, out long length, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (attachment.Content != null) {
            length = attachment.Content.LongLength;
            dataBid = _file.WriteDataTree(attachment.Content);
            return true;
        }
        if (attachment.ContentSource != null) {
            using (Stream stream = attachment.OpenContentStream()) {
                cancellationToken.ThrowIfCancellationRequested();
                dataBid = _file.WriteDataTree(stream, out length);
                return true;
            }
        }
        dataBid = 0;
        length = 0;
        return false;
    }

    private static int CalculateAttachmentObjectSize(IReadOnlyList<MapiProperty> properties,
        int codePage, long contentLength,
        IReadOnlyDictionary<ushort, PstWriterValueReference> valueReferences,
        IReadOnlyDictionary<ushort, PstWriterObjectReference> objectReferences) {
        ushort attachDataId = MapiKnownProperties.PidTag.AttachData.GetStandardPropertyId();
        long total = 0;
        foreach (MapiProperty property in properties.GroupBy(item => item.PropertyId)
            .Select(group => group.Last())) {
            long valueSize;
            if (property.PropertyId == attachDataId && valueReferences.ContainsKey(attachDataId)) {
                valueSize = contentLength;
            } else if (property.PropertyId == attachDataId &&
                objectReferences.TryGetValue(attachDataId, out PstWriterObjectReference? reference)) {
                valueSize = reference.Size;
            } else {
                valueSize = PstPropertyValueWriter.GetLogicalValueSize(property, codePage);
            }
            total = checked(total + Math.Max(0, valueSize));
        }
        return checked((int)Math.Min(total, int.MaxValue));
    }

    private void TranslateDiagnostics(IEnumerable<EmailDiagnostic> diagnostics) {
        foreach (EmailDiagnostic diagnostic in diagnostics) {
            Report(new EmailStoreDiagnostic(diagnostic.Code, diagnostic.Message,
                diagnostic.Severity == EmailDiagnosticSeverity.Error
                    ? EmailStoreDiagnosticSeverity.Error
                    : diagnostic.Severity == EmailDiagnosticSeverity.Information
                        ? EmailStoreDiagnosticSeverity.Information
                        : EmailStoreDiagnosticSeverity.Warning,
                diagnostic.Location));
        }
    }

    private static int ResolveCodePage(EmailDocument document) =>
        document.OutlookCodePage.GetValueOrDefault(65001) > 0
            ? document.OutlookCodePage.GetValueOrDefault(65001)
            : 65001;

    private static string? JoinDisplayRecipients(EmailDocument document, EmailRecipientKind kind) {
        string[] values = document.Recipients.Where(recipient => recipient.Kind == kind)
            .Select(recipient => string.IsNullOrWhiteSpace(recipient.Address.DisplayName)
                ? recipient.Address.Address
                : recipient.Address.DisplayName!)
            .Where(value => !string.IsNullOrWhiteSpace(value)).Select(value => value!).ToArray();
        return values.Length == 0 ? null : string.Join("; ", values);
    }

    private static byte[]? CreateConversationId(IReadOnlyList<MapiProperty> properties) {
        byte[]? retained = properties.GetMapiValueOrDefault(MapiKnownProperties.PidTag.ConversationId);
        if (retained != null && retained.Length == 16) return (byte[])retained.Clone();
        bool tracking = properties.GetNullableMapiValue(
            MapiKnownProperties.PidTag.ConversationIndexTracking) == true;
        byte[]? conversationIndex = properties.GetMapiValueOrDefault(
            MapiKnownProperties.PidTag.ConversationIndex);
        if (tracking && conversationIndex != null && conversationIndex.Length >= 22 &&
            conversationIndex[0] == 0x01) {
            var fromIndex = new byte[16];
            Buffer.BlockCopy(conversationIndex, 6, fromIndex, 0, fromIndex.Length);
            return fromIndex;
        }
        string? topic = properties.GetMapiValueOrDefault(
            MapiKnownProperties.PidTag.ConversationTopic);
        if (string.IsNullOrEmpty(topic)) return null;
        int terminator = topic!.IndexOf('\0');
        if (terminator >= 0) topic = topic.Substring(0, terminator);
        if (topic.Length == 0) return null;
        string normalized = topic.Length > 255 ? topic.Substring(0, 255) : topic;
        byte[] bytes = Encoding.Unicode.GetBytes(normalized.ToUpperInvariant());
        using (MD5 md5 = MD5.Create()) return md5.ComputeHash(bytes);
    }

    private byte[] CreateObjectKey(uint value) {
        byte[] bytes = _providerUid.ToByteArray();
        PstBinary.WriteUInt32(bytes, 0, PstBinary.UInt32(bytes, 0) ^ value);
        return bytes;
    }

    private readonly struct WrittenAttachment {
        internal WrittenAttachment(PstWriterContextResult context,
            IReadOnlyList<MapiProperty> tableProperties) {
            Context = context;
            TableProperties = tableProperties;
        }
        internal PstWriterContextResult Context { get; }
        internal IReadOnlyList<MapiProperty> TableProperties { get; }
    }
}
