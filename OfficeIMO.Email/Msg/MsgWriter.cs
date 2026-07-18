using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Email;

internal static partial class MsgWriter {
    private static readonly Guid MessageStorageClassId = new Guid("00020D0B-0000-0000-C000-000000000046");
    private static readonly Guid StorageInterfaceId = new Guid("0000000B-0000-0000-C000-000000000046");
    private static readonly DateTimeOffset FallbackCreationTime =
        new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero);

    internal static byte[] Write(EmailDocument document, EmailWriterOptions options,
        IList<EmailDiagnostic> diagnostics, bool asTemplate = false) {
        using (var output = new EmailBoundedMemoryStream(options.MaxOutputBytes)) {
            Write(output, document, options, diagnostics, asTemplate);
            return output.ToArray();
        }
    }

    internal static void Write(Stream output, EmailDocument document, EmailWriterOptions options,
        IList<EmailDiagnostic> diagnostics, bool asTemplate = false) {
        var streams = new List<OfficeCompoundStream>();
        var resources = new List<IDisposable>();
        var names = new MsgNamedPropertyWriter();
        try {
            BuildMessage(document, string.Empty, MsgPropertyStreamKind.TopLevel, names, streams, resources,
                diagnostics, options, 0, asTemplate);
            names.WriteStreams(streams);
            long outputLength = OfficeCompoundFileWriter.GetSerializedLength(streams);
            if (outputLength > options.MaxOutputBytes) {
                throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), outputLength,
                    options.MaxOutputBytes);
            }
            OfficeCompoundFileWriter.Write(output, streams, MessageStorageClassId);
        } finally {
            foreach (IDisposable resource in resources) resource.Dispose();
        }
    }

    private static void BuildMessage(EmailDocument document, string prefix, MsgPropertyStreamKind kind,
        MsgNamedPropertyWriter names, IList<OfficeCompoundStream> streams, IList<IDisposable> resources,
        IList<EmailDiagnostic> diagnostics, EmailWriterOptions options, int depth, bool asTemplate = false) {
        if (depth > options.MaxNestedMessageDepth) throw new InvalidOperationException("The embedded-message write depth exceeds the configured maximum.");
        EmailRecipient[] storageRecipients = document.Recipients
            .Where(recipient => recipient.Kind != EmailRecipientKind.ReplyTo)
            .ToArray();
        EmailAttachment[] writableAttachments = OutlookTaskCommunicationAttachmentProjection
            .GetWritableAttachments(document);
        int codePage = MapiStringEncodingContext.FromCodePage(document.OutlookCodePage ?? 65001).PrimaryCodePage;
        MsgPropertyBuilder messageProperties = CreateMessageProperties(document, diagnostics, prefix, options,
            asTemplate);
        MsgPropertyWriter.Write(prefix, kind, messageProperties.Properties, storageRecipients.Length,
            writableAttachments.Length, names, streams, diagnostics, codePage);

        for (int index = 0; index < storageRecipients.Length; index++) {
            EmailRecipient recipient = storageRecipients[index];
            string storage = MsgBinary.CombinePath(prefix,
                string.Concat("__recip_version1.0_#", index.ToString("X8", CultureInfo.InvariantCulture)));
            MsgPropertyBuilder properties = CreateRecipientProperties(recipient, index);
            MsgPropertyWriter.Write(storage, MsgPropertyStreamKind.ChildObject, properties.Properties,
                0, 0, names, streams, diagnostics, codePage);
        }

        for (int index = 0; index < writableAttachments.Length; index++) {
            EmailAttachment attachment = writableAttachments[index];
            string storage = MsgBinary.CombinePath(prefix,
                string.Concat("__attach_version1.0_#", index.ToString("X8", CultureInfo.InvariantCulture)));
            int method = attachment.MapiAttachMethod ?? (attachment.EmbeddedDocument != null ? 5 :
                attachment.StructuredStorageStreams.Count > 0 ? 6 : 1);
            bool hasContent = attachment.Content != null || attachment.ContentSource != null ||
                EmailAttachmentStreamScope.HasStagedContent(attachment);
            byte[]? content = method == 1 ? null : EmailAttachmentContent.ReadOrNull(attachment, options.MaxOutputBytes);
            EmailDocument? embeddedDocument = attachment.EmbeddedDocument ??
                TryReadOpaqueEmbeddedTnef(attachment, content, diagnostics, storage);
            MsgPropertyBuilder properties = CreateAttachmentProperties(attachment, index, method, diagnostics, storage,
                embeddedDocument != null || attachment.StructuredStorageStreams.Count > 0 || hasContent, content,
                EmailAttachmentStreamScope.GetLength(attachment));
            IReadOnlyDictionary<uint, OfficeCompoundStream>? streamOverrides = null;
            if (method == 1 && hasContent) {
                string contentStreamName = MsgBinary.CombinePath(storage, "__substg1.0_37010102");
                AttachmentContentRegistration registration = AttachmentContentRegistration.Create(
                    attachment, contentStreamName, options.MaxOutputBytes);
                resources.Add(registration);
                streamOverrides = new Dictionary<uint, OfficeCompoundStream> {
                    [((uint)MapiKnownProperties.PidTag.AttachData.GetStandardPropertyId() << 16) |
                        (ushort)MapiPropertyType.Binary] = registration.Stream
                };
            }
            MsgPropertyWriter.Write(storage, MsgPropertyStreamKind.ChildObject, properties.Properties,
                0, 0, names, streams, diagnostics, codePage,
                method == 5 ? 1U : method == 6 ? 4U : 0U, streamOverrides);

            string objectStorage = MsgBinary.CombinePath(storage, "__substg1.0_3701000D");
            if (method == 5 && embeddedDocument != null) {
                BuildMessage(embeddedDocument, objectStorage, MsgPropertyStreamKind.EmbeddedMessage,
                    names, streams, resources, diagnostics, options, depth + 1);
            } else if (method == 5 && attachment.StructuredStorageStreams.Count > 0) {
                foreach (KeyValuePair<string, byte[]> stream in attachment.StructuredStorageStreams
                    .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase)) {
                    streams.Add(new OfficeCompoundStream(MsgBinary.CombinePath(objectStorage, stream.Key), stream.Value));
                }
            } else if (method == 5 && content != null) {
                streams.Add(new OfficeCompoundStream(
                    MsgBinary.CombinePath(objectStorage, "CONTENTS"), content));
                diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_OPAQUE_EMBEDDED_CONTENT_WRAPPED",
                    "Opaque embedded attachment bytes were retained in the MSG object storage because the payload could not be projected as a nested message.",
                    EmailDiagnosticSeverity.Warning, objectStorage));
            } else if (method == 6) {
                WriteStructuredAttachment(attachment, content, objectStorage, streams, diagnostics);
            }
        }
    }

    private static void WriteStructuredAttachment(EmailAttachment attachment, byte[]? content, string objectStorage,
        IList<OfficeCompoundStream> streams, IList<EmailDiagnostic> diagnostics) {
        if (attachment.StructuredStorageStreams.Count > 0) {
            foreach (KeyValuePair<string, byte[]> stream in attachment.StructuredStorageStreams
                .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase)) {
                streams.Add(new OfficeCompoundStream(MsgBinary.CombinePath(objectStorage, stream.Key), stream.Value));
            }
            return;
        }

        if (content == null) return;
        byte[] retained = content;
        byte[] compoundCandidate = retained.Length > 16 &&
            new Guid(MsgBinary.Slice(retained, 0, 16)) == StorageInterfaceId
                ? MsgBinary.Slice(retained, 16, retained.Length - 16)
                : retained;
        if (OfficeCompoundFileReader.TryRead(compoundCandidate, out OfficeCompoundFile? compound, out _) &&
            compound != null) {
            foreach (KeyValuePair<string, byte[]> stream in compound.Streams
                .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase)) {
                streams.Add(new OfficeCompoundStream(MsgBinary.CombinePath(objectStorage, stream.Key), stream.Value));
            }
            return;
        }

        streams.Add(new OfficeCompoundStream(MsgBinary.CombinePath(objectStorage, "CONTENTS"), retained));
        diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_OPAQUE_STRUCTURED_CONTENT_WRAPPED",
            "Opaque structured attachment bytes were retained in the MSG object storage because the original compound payload could not be expanded.",
            EmailDiagnosticSeverity.Warning, objectStorage));
    }

    internal static MsgPropertyBuilder CreateMessageProperties(EmailDocument document,
        IList<EmailDiagnostic> diagnostics, string location, EmailWriterOptions? options = null,
        bool asTemplate = false) {
        var properties = new MsgPropertyBuilder(document.MapiProperties);
        int codePage = MapiStringEncodingContext.FromCodePage(document.OutlookCodePage ?? 65001).PrimaryCodePage;
        EmailMessageMetadata metadata = document.MessageMetadata;
        string messageClass = ResolveMessageClass(document);
        properties.Set(MapiKnownProperties.PidTag.MessageClass, messageClass);
        properties.SetDefault(MapiKnownProperties.PidLid.SideEffects, 0);
        properties.SetDefault(MapiKnownProperties.PidName.AcceptLanguage, ResolveAcceptLanguage(document));
        properties.Set(MapiKnownProperties.PidTag.StoreSupportMask, 0x00040E79);
        properties.Set(MapiKnownProperties.PidTag.AlternateRecipientAllowed, true);
        const int managedMessageFlags = 0x0001 | 0x0002 | 0x0008 | 0x0010 | 0x0400;
        int messageFlags = document.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.MessageFlags) &
            ~managedMessageFlags;
        bool hasAttachments = document.Attachments.Any(attachment => !attachment.IsProjectedSemanticContent);
        messageFlags |= 0x0002;
        if (hasAttachments) messageFlags |= 0x0010;
        if (metadata.IsDraft || asTemplate) messageFlags |= 0x0008;
        if (metadata.IsRead == true) messageFlags |= 0x0001 | 0x0400;
        properties.Set(MapiKnownProperties.PidTag.MessageFlags, messageFlags);
        properties.Set(MapiKnownProperties.PidTag.HasAttachments, hasAttachments);
        properties.Set(MapiKnownProperties.PidTag.Subject, document.Subject);
        ResolveSubject(document.Subject, metadata, out string? subjectPrefix, out string? normalizedSubject);
        properties.Set(MapiKnownProperties.PidTag.SubjectPrefix, subjectPrefix);
        properties.Set(MapiKnownProperties.PidTag.NormalizedSubject, normalizedSubject);
        properties.Set(MapiKnownProperties.PidTag.ConversationTopic, metadata.ConversationTopic ?? normalizedSubject);
        properties.Set(MapiKnownProperties.PidTag.ConversationIndex, metadata.ConversationIndex);
        properties.Set(MapiKnownProperties.PidTag.Body, document.Body.Text);
        properties.Set(MapiKnownProperties.PidTag.Html, null);
        properties.Set(MapiKnownProperties.PidTag.NativeBodyInfo, null);
        properties.Set(MapiKnownProperties.PidTag.RtfCompressed, null);
        properties.Set(MapiKnownProperties.PidTag.RtfInSync, null);
        if (document.Body.Html != null) {
            properties.Set(MapiKnownProperties.PidTag.Html,
                EncodeString8(document.Body.Html, codePage, diagnostics, string.Concat(location, "/html")));
            properties.Set(MapiKnownProperties.PidTag.NativeBodyInfo, 3);
        } else if (document.Body.Rtf != null) {
            properties.Set(MapiKnownProperties.PidTag.NativeBodyInfo, 2);
        } else if (document.Body.Text != null) {
            properties.Set(MapiKnownProperties.PidTag.NativeBodyInfo, 1);
        }
        if (document.Body.Rtf != null) {
            if (EmailRtfByteCodec.TryEncode(document.Body.Rtf, out byte[] rtfBytes)) {
                properties.Set(MapiKnownProperties.PidTag.RtfCompressed, MapiCompressedRtfCodec.Compress(rtfBytes));
                properties.Set(MapiKnownProperties.PidTag.RtfInSync, true);
            } else {
                diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_RTF_CHARACTER_UNENCODABLE",
                    "The RTF source contains a character above U+00FF. Serialize it through OfficeIMO.Rtf so the character is represented by an RTF escape.",
                    EmailDiagnosticSeverity.Error, location));
            }
        }
        properties.Set(MapiKnownProperties.PidTag.InternetMessageId,
            string.IsNullOrWhiteSpace(document.MessageId) ? null : string.Concat("<", document.MessageId!.Trim().Trim('<', '>'), ">"));
        properties.Set(MapiKnownProperties.PidTag.ClientSubmitTime, document.Date);
        properties.Set(MapiKnownProperties.PidTag.MessageDeliveryTime, document.ReceivedDate);
        properties.Set(MapiKnownProperties.PidTag.InternetReferences, metadata.InternetReferences);
        properties.Set(MapiKnownProperties.PidTag.InReplyToId, metadata.InReplyToId);
        properties.Set(MapiKnownProperties.PidTag.Importance, (int)(metadata.Importance ?? EmailMessageImportance.Normal));
        properties.Set(MapiKnownProperties.PidTag.Priority, (int)(metadata.Priority ?? EmailMessagePriority.Normal));
        properties.Set(MapiKnownProperties.PidTag.IconIndex,
            metadata.IconIndex ?? (metadata.IsDraft || asTemplate
                ? 0x00000103
                : metadata.IsRead == true ? 0x00000100 : 0x00000101));
        properties.Set(MapiKnownProperties.PidTag.ReadReceiptRequested, metadata.ReadReceiptRequested);
        properties.Set(MapiKnownProperties.PidTag.OriginatorDeliveryReportRequested, metadata.DeliveryReceiptRequested);
        properties.Set(MapiKnownProperties.PidTag.Sensitivity, metadata.Sensitivity);
        properties.Set(MapiKnownProperties.PidTag.OriginalSensitivity, metadata.OriginalSensitivity);
        DateTimeOffset created = metadata.CreatedDate ?? document.Date ?? document.ReceivedDate ?? FallbackCreationTime;
        properties.Set(MapiKnownProperties.PidTag.CreationTime, created);
        properties.Set(MapiKnownProperties.PidTag.LastModificationTime, metadata.ModifiedDate ?? created);
        properties.Set(MapiKnownProperties.PidTag.InternetCodepage, codePage);
        properties.Set(MapiKnownProperties.PidTag.MessageCodepage, codePage);
        properties.Set(MapiKnownProperties.PidTag.MessageLocaleId, metadata.LocaleId ?? 1033);
        properties.Set(MapiKnownProperties.PidTag.LastModifierName, metadata.LastModifierName);
        properties.Set(MapiKnownProperties.PidTag.ConversationId, metadata.ConversationId);
        properties.Set(MapiKnownProperties.PidTag.MessageEditorFormat, metadata.EditorFormat);
        properties.Set(MapiKnownProperties.PidName.ReactionsSummary, metadata.ReactionsSummary);
        properties.Set(MapiKnownProperties.PidName.OwnerReactionHistory, metadata.OwnerReactionHistory);
        properties.Set(MapiKnownProperties.PidName.OwnerReactionType, metadata.OwnerReactionType);
        properties.Set(MapiKnownProperties.PidName.OwnerReactionTime, metadata.OwnerReactionTime);
        properties.Set(MapiKnownProperties.PidName.ReactionsCount, metadata.ReactionsCount);
        properties.Set(MapiKnownProperties.PidTag.SentRepresentingName, document.From?.DisplayName);
        properties.Set(MapiKnownProperties.PidTag.SentRepresentingEmailAddress, document.From?.Address);
        properties.Set(MapiKnownProperties.PidTag.SentRepresentingAddressType, document.From?.AddressType ?? "SMTP");
        properties.Set(MapiKnownProperties.PidTag.SentRepresentingSmtpAddress, document.From?.Address);
        EmailAddress? sender = document.Sender ?? document.From;
        properties.Set(MapiKnownProperties.PidTag.SenderName, sender?.DisplayName);
        properties.Set(MapiKnownProperties.PidTag.SenderEmailAddress, sender?.Address);
        properties.Set(MapiKnownProperties.PidTag.SenderAddressType, sender?.AddressType ?? "SMTP");
        properties.Set(MapiKnownProperties.PidTag.SenderSmtpAddress, sender?.Address);
        properties.Set(MapiKnownProperties.PidTag.SenderEntryId,
            sender == null ? null : MsgIdentity.CreateOneOffEntryId(sender));
        SetReceivedAddress(properties, document.ReceivedBy, MapiKnownProperties.PidTag.ReceivedByName,
            MapiKnownProperties.PidTag.ReceivedByAddressType, MapiKnownProperties.PidTag.ReceivedByEmailAddress,
            MapiKnownProperties.PidTag.ReceivedByEntryId);
        SetReceivedAddress(properties, document.ReceivedRepresenting,
            MapiKnownProperties.PidTag.ReceivedRepresentingName,
            MapiKnownProperties.PidTag.ReceivedRepresentingAddressType,
            MapiKnownProperties.PidTag.ReceivedRepresentingEmailAddress,
            MapiKnownProperties.PidTag.ReceivedRepresentingEntryId);
        properties.Set(MapiKnownProperties.PidTag.DisplayTo, JoinRecipients(document, EmailRecipientKind.To));
        properties.Set(MapiKnownProperties.PidTag.DisplayCc, JoinRecipients(document, EmailRecipientKind.Cc));
        properties.Set(MapiKnownProperties.PidTag.DisplayBcc, JoinRecipients(document, EmailRecipientKind.Bcc));
        properties.Set(MapiKnownProperties.PidTag.ReplyRecipientNames, JoinRecipients(document, EmailRecipientKind.ReplyTo));
        properties.Set(MapiKnownProperties.PidName.Keywords,
            metadata.Categories.Count == 0 ? null : metadata.Categories.Cast<object>().ToArray());
        string? headers = MimeWriter.CreateTransportHeaders(document, options ?? EmailWriterOptions.Default);
        properties.Set(MapiKnownProperties.PidTag.TransportMessageHeaders, headers);
        AddTypedProperties(properties, document);
        OutlookMessageSemanticsWriter.Apply(properties, document, codePage, diagnostics, location);
        document.MapiWritePatch.Apply(properties);
        return properties;
    }

    private static string ResolveAcceptLanguage(EmailDocument document) {
        EmailHeader? header = document.Headers.FirstOrDefault(item =>
            string.Equals(item.Name, "Accept-Language", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(item.Name, "X-Accept-Language", StringComparison.OrdinalIgnoreCase));
        string? value = header?.Value;
        if (!string.IsNullOrWhiteSpace(value)) return value!;
        int localeId = document.MessageMetadata.LocaleId ?? 1033;
        try {
            return CultureInfo.GetCultureInfo(localeId).Name;
        } catch (CultureNotFoundException) {
            return "en-US";
        }
    }

    private static void SetReceivedAddress(MsgPropertyBuilder properties, EmailAddress? address,
        MapiPropertyKey<string> displayNameKey, MapiPropertyKey<string> addressTypeKey,
        MapiPropertyKey<string> addressKey, MapiPropertyKey<byte[]> entryKey) {
        properties.Set(displayNameKey, address?.DisplayName);
        properties.Set(addressTypeKey, address?.AddressType ?? (address == null ? null : "SMTP"));
        properties.Set(addressKey, address?.Address);
        properties.Set(entryKey,
            address == null ? null : MsgIdentity.CreateOneOffEntryId(address));
    }

    internal static MsgPropertyBuilder CreateRecipientProperties(EmailRecipient recipient, int index) {
        var properties = new MsgPropertyBuilder(recipient.MapiProperties);
        int type = recipient.Kind == EmailRecipientKind.To ? 1 : recipient.Kind == EmailRecipientKind.Cc ? 2 :
            recipient.Kind == EmailRecipientKind.Bcc ? 3 :
            recipient.Kind == EmailRecipientKind.Resource || recipient.Kind == EmailRecipientKind.Room ? 4 : 0;
        string addressType = string.IsNullOrWhiteSpace(recipient.Address.AddressType) ? "SMTP" : recipient.Address.AddressType!;
        string? address = recipient.Address.Address;
        properties.Set(MapiKnownProperties.PidTag.RowId, recipient.MapiRowId ?? index);
        properties.Set(MapiKnownProperties.PidTag.EntryId, MsgIdentity.CreateOneOffEntryId(recipient.Address));
        properties.Set(MapiKnownProperties.PidTag.RecipientType, type);
        properties.Set(MapiKnownProperties.PidTag.DisplayName, recipient.Address.DisplayName ?? recipient.Address.Address);
        properties.Set(MapiKnownProperties.PidTag.AddressType, addressType);
        properties.Set(MapiKnownProperties.PidTag.EmailAddress, address);
        properties.Set(MapiKnownProperties.PidTag.SmtpAddress, address);
        properties.Set(MapiKnownProperties.PidTag.SearchKey, MsgIdentity.CreateSearchKey(addressType, address));
        properties.Set(MapiKnownProperties.PidTag.ObjectType, recipient.MapiObjectType ?? 6);
        properties.Set(MapiKnownProperties.PidTag.DisplayType, recipient.MapiDisplayType ?? 0);
        properties.Set(MapiKnownProperties.PidTag.DisplayTypeEx,
            recipient.MapiDisplayTypeEx ?? (recipient.Kind == EmailRecipientKind.Room ? 7 : 0));
        return properties;
    }

    internal static MsgPropertyBuilder CreateAttachmentProperties(EmailAttachment attachment, int index, int method,
        IList<EmailDiagnostic> diagnostics, string location, bool hasRetainedObjectContent = false,
        byte[]? materializedContent = null, long? retainedContentLength = null) {
        var properties = new MsgPropertyBuilder(attachment.MapiProperties);
        properties.Set(MapiKnownProperties.PidTag.ObjectType, 7);
        properties.Set(MapiKnownProperties.PidTag.AttachMethod, method);
        properties.Set(MapiKnownProperties.PidTag.AttachNumber, index);
        properties.Set(MapiKnownProperties.PidTag.AttachLongFilename, attachment.FileName);
        properties.Set(MapiKnownProperties.PidTag.AttachFilename, attachment.FileName);
        properties.Set(MapiKnownProperties.PidTag.DisplayName, attachment.FileName);
        properties.Set(MapiKnownProperties.PidTag.AttachMimeTag, attachment.ContentType);
        properties.Set(MapiKnownProperties.PidTag.AttachContentId, attachment.ContentId);
        properties.Set(MapiKnownProperties.PidTag.AttachContentLocation, attachment.ContentLocation);
        properties.Set(MapiKnownProperties.PidTag.AttachExtension, GetLogicalExtension(attachment.FileName));
        properties.Set(MapiKnownProperties.PidTag.RenderingPosition,
            attachment.RenderingPosition >= 0 ? attachment.RenderingPosition : attachment.IsInline ? 0 : -1);
        properties.Set(MapiKnownProperties.PidTag.AttachFlags, attachment.IsInline ? 0x00000004 : 0);
        properties.Set(MapiKnownProperties.PidTag.AttachmentHidden, attachment.IsHidden || attachment.IsInline);
        properties.Set(MapiKnownProperties.PidTag.AttachmentContactPhoto, attachment.IsContactPhoto);
        properties.Set(MapiKnownProperties.PidTag.CreationTime, attachment.CreatedDate);
        properties.Set(MapiKnownProperties.PidTag.LastModificationTime,
            attachment.ModifiedDate ?? attachment.CreatedDate);
        properties.Set(MapiKnownProperties.PidTag.AttachLongPathname, attachment.LinkedPath);
        if (method == 5 || method == 6) {
            properties.Set(MapiKnownProperties.PidTag.AttachData, MapiPropertyType.Object, null);
            if (method == 5 && attachment.EmbeddedDocument == null && !hasRetainedObjectContent) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE",
                    "An embedded MSG attachment has no retained embedded document.",
                    EmailDiagnosticSeverity.Error, location));
            } else if (method == 6 && attachment.StructuredStorageStreams.Count == 0 && attachment.Length > 0 &&
                !hasRetainedObjectContent) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE",
                    "A structured MSG attachment has a declared length but no retained storage streams.",
                    EmailDiagnosticSeverity.Error, location));
            }
        } else if (materializedContent != null || attachment.Content != null) {
            byte[] content = materializedContent ?? attachment.Content!;
            properties.Set(MapiKnownProperties.PidTag.AttachData, content);
            properties.Set(MapiKnownProperties.PidTag.AttachSize, content.Length);
        } else if (hasRetainedObjectContent) {
            // Keep the variable-property row so a caller can provide a reopenable stream override.
            properties.Set(MapiKnownProperties.PidTag.AttachData, Array.Empty<byte>());
            properties.Set(MapiKnownProperties.PidTag.AttachSize,
                checked((int)Math.Min(retainedContentLength ?? attachment.Length, int.MaxValue)));
        } else {
            properties.Set(MapiKnownProperties.PidTag.AttachData, null);
            properties.Set(MapiKnownProperties.PidTag.AttachSize,
                checked((int)Math.Min(attachment.Length, int.MaxValue)));
            if (attachment.Length > 0) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE",
                    "An MSG attachment has a declared length but no retained content.",
                    EmailDiagnosticSeverity.Error, location));
            }
        }
        return properties;
    }

    private static EmailDocument? TryReadOpaqueEmbeddedTnef(EmailAttachment attachment, byte[]? content,
        IList<EmailDiagnostic> diagnostics, string location) {
        if (attachment.MapiAttachMethod != 5 || content == null || content.Length < 20 ||
            new Guid(MsgBinary.Slice(content, 0, 16)) != new Guid("00020307-0000-0000-C000-000000000046")) {
            return null;
        }
        byte[] nested = MsgBinary.Slice(content, 16, content.Length - 16);
        if (MsgBinary.ReadUInt32(nested, 0) != TnefConstants.Signature) return null;
        diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_EMBEDDED_MESSAGE_REPROJECTED",
            "Opaque embedded TNEF content was projected while writing MSG storage.",
            EmailDiagnosticSeverity.Warning, location));
        return TnefReader.Read(nested, EmailReaderOptions.Default, diagnostics, CancellationToken.None);
    }

    private sealed class AttachmentContentRegistration : IDisposable {
        private readonly string? _temporaryPath;

        private AttachmentContentRegistration(OfficeCompoundStream stream, string? temporaryPath) {
            Stream = stream;
            _temporaryPath = temporaryPath;
        }

        internal OfficeCompoundStream Stream { get; }

        internal static AttachmentContentRegistration Create(EmailAttachment attachment, string streamName,
            long maximumBytes) {
            long? length = EmailAttachmentStreamScope.GetLength(attachment);
            if (length.HasValue) {
                if (length.Value > maximumBytes) {
                    throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), length.Value,
                        maximumBytes);
                }
                return new AttachmentContentRegistration(new OfficeCompoundStream(streamName, length.Value,
                    () => EmailAttachmentStreamScope.OpenRead(attachment)), null);
            }

            string path = Path.Combine(Path.GetTempPath(),
                string.Concat("OfficeIMO.Email.Msg.", Guid.NewGuid().ToString("N"), ".content"));
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
                        if (copied > maximumBytes) {
                            throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), copied,
                                maximumBytes);
                        }
                        output.Write(buffer, 0, read);
                    }
                }
                return new AttachmentContentRegistration(new OfficeCompoundStream(streamName, copied,
                    () => new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
                        81920, FileOptions.SequentialScan)), path);
            } catch {
                OfficeFileCommit.DeleteIfExists(path);
                throw;
            }
        }

        public void Dispose() => OfficeFileCommit.DeleteIfExists(_temporaryPath);
    }

    private static string GetLogicalExtension(string? fileName) {
        if (string.IsNullOrEmpty(fileName)) return string.Empty;
        int separator = Math.Max(fileName!.LastIndexOf('/'), fileName.LastIndexOf('\\'));
        int dot = fileName.LastIndexOf('.');
        return dot > separator && dot + 1 < fileName.Length ? fileName.Substring(dot) : string.Empty;
    }

    private static byte[] EncodeString8(string value, int codePage, IList<EmailDiagnostic> diagnostics,
        string location) {
        try {
            return MsgValueWriter.EncodeString8(value, codePage);
        } catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException ||
            exception is EncoderFallbackException) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_STRING8_ENCODING_INVALID",
                string.Concat("Text could not be encoded with MAPI code page ",
                    codePage.ToString(CultureInfo.InvariantCulture), ": ", exception.Message),
                EmailDiagnosticSeverity.Error, location));
            return Array.Empty<byte>();
        }
    }

    internal static string ResolveMessageClass(EmailDocument document) {
        if (document.MessageClass != null) return document.MessageClass;
        if (document.MeetingCommunication != null) {
            switch (document.MeetingCommunication.Kind) {
                case OutlookMeetingCommunicationKind.RequestOrUpdate: return "IPM.Schedule.Meeting.Request";
                case OutlookMeetingCommunicationKind.Cancellation: return "IPM.Schedule.Meeting.Canceled";
                case OutlookMeetingCommunicationKind.ResponseAccepted: return "IPM.Schedule.Meeting.Resp.Pos";
                case OutlookMeetingCommunicationKind.ResponseTentative: return "IPM.Schedule.Meeting.Resp.Tent";
                case OutlookMeetingCommunicationKind.ResponseDeclined: return "IPM.Schedule.Meeting.Resp.Neg";
                case OutlookMeetingCommunicationKind.ForwardNotification: return "IPM.Schedule.Meeting.Forward.Notification";
            }
        }
        if (document.TaskCommunication != null) {
            switch (document.TaskCommunication.Kind) {
                case OutlookTaskCommunicationKind.Request: return "IPM.TaskRequest";
                case OutlookTaskCommunicationKind.Accept: return "IPM.TaskRequest.Accept";
                case OutlookTaskCommunicationKind.Decline: return "IPM.TaskRequest.Decline";
                case OutlookTaskCommunicationKind.Update: return "IPM.TaskRequest.Update";
            }
        }
        switch (document.OutlookItemKind) {
            case OutlookItemKind.Appointment: return "IPM.Appointment";
            case OutlookItemKind.Contact: return "IPM.Contact";
            case OutlookItemKind.Task: return "IPM.Task";
            case OutlookItemKind.Journal: return "IPM.Activity";
            case OutlookItemKind.Note: return "IPM.StickyNote";
            case OutlookItemKind.DistributionList: return "IPM.DistList";
            default: return "IPM.Note";
        }
    }

    private static string? JoinRecipients(EmailDocument document, EmailRecipientKind kind) {
        string[] values = document.Recipients.Where(recipient => recipient.Kind == kind)
            .Select(recipient => recipient.Address.ToString()).Where(value => value.Length > 0).ToArray();
        return values.Length == 0 ? null : string.Join("; ", values);
    }

    private static void ResolveSubject(string? subject, EmailMessageMetadata metadata,
        out string? prefix, out string? normalized) {
        if (subject == null) {
            prefix = null;
            normalized = null;
            return;
        }
        prefix = metadata.SubjectPrefix ?? string.Empty;
        normalized = metadata.NormalizedSubject ?? subject;
        if (metadata.NormalizedSubject != null || subject.Length == 0) return;

        int colon = subject!.IndexOf(':');
        if (colon > 0 && colon <= 3 && colon + 1 < subject.Length) {
            prefix = subject.Substring(0, colon + 1);
            if (colon + 1 < subject.Length && subject[colon + 1] == ' ') prefix += " ";
            normalized = subject.Substring(prefix.Length);
        }
    }
}
