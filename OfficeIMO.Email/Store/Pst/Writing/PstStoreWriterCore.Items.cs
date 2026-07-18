using OfficeIMO.Email;

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
        messageBuilder.Set(MapiKnownProperties.PidTag.MessageSize, EstimateMessageSize(document));
        messageBuilder.Set(MapiKnownProperties.PidTag.MessageStatus,
            document.MapiProperties.GetNullableMapiValue(MapiKnownProperties.PidTag.MessageStatus) ?? 0);
        messageBuilder.Set(MapiKnownProperties.PidTag.SearchKey,
            document.MapiProperties.GetMapiValueOrDefault(MapiKnownProperties.PidTag.SearchKey) ??
                CreateObjectKey(messageNid));
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

        PstWriterContextResult context = PstPropertyContextWriter.Write(_file,
            messageProperties, codePage, messageSubnodes, null, null,
            Report, string.Concat("message/", FormatId(messageNid)));
        return new WrittenMessage(context,
            SelectTableProperties(messageProperties, ContentsColumns.Concat(AssociatedColumns).ToArray()));
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
                subnodes.Add(new PstWriterSubnode(embeddedNid,
                    embedded.Context.DataBid, embedded.Context.SubnodeBid));
                objectReferences[MapiKnownProperties.PidTag.AttachData.GetStandardPropertyId()] =
                    new PstWriterObjectReference(embeddedNid, 0);
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
        builder.Set(MapiKnownProperties.PidTag.AttachSize,
            checked((int)Math.Min(contentLength, int.MaxValue)));
        TranslateDiagnostics(diagnostics);
        string location = string.Concat("attachment/0x",
            attachmentNid.ToString("X8", CultureInfo.InvariantCulture));
        IReadOnlyList<MapiProperty> properties = _namedProperties.Map(
            builder.Properties, Report, location);
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

    private static int EstimateMessageSize(EmailDocument document) {
        long length = 0;
        if (document.Body.Text != null) length += Encoding.Unicode.GetByteCount(document.Body.Text);
        if (document.Body.Html != null) length += Encoding.UTF8.GetByteCount(document.Body.Html);
        if (document.Body.Rtf != null) length += Encoding.UTF8.GetByteCount(document.Body.Rtf);
        foreach (EmailAttachment attachment in document.Attachments) {
            length = checked(length + Math.Max(0, attachment.Content?.LongLength ?? attachment.Length));
        }
        return checked((int)Math.Min(length, int.MaxValue));
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
