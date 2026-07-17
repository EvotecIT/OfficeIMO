using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Email;

internal static class MsgWriter {
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
        EmailAttachment[] writableAttachments = document.Attachments.Where(attachment =>
            !attachment.IsProjectedSemanticContent).ToArray();
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
                    [0x37010102U] = registration.Stream
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
        string messageClass = document.MessageClass ?? DefaultMessageClass(document.OutlookItemKind);
        properties.Set(0x001A, MapiPropertyType.Unicode, messageClass);
        properties.SetNamedDefault(MsgProjection.PsetidCommon, 0x8510, MapiPropertyType.Integer32, 0);
        properties.SetNamedDefault(MsgProjection.PsInternetHeaders, "acceptlanguage",
            MapiPropertyType.Unicode, ResolveAcceptLanguage(document));
        properties.Set(0x340D, MapiPropertyType.Integer32, 0x00040E79);
        properties.Set(0x0002, MapiPropertyType.Boolean, true);
        const int managedMessageFlags = 0x0001 | 0x0002 | 0x0008 | 0x0010 | 0x0400;
        int messageFlags = (MsgProjection.GetInt(document.MapiProperties, 0x0E07) ?? 0) & ~managedMessageFlags;
        bool hasAttachments = document.Attachments.Any(attachment => !attachment.IsProjectedSemanticContent);
        messageFlags |= 0x0002;
        if (hasAttachments) messageFlags |= 0x0010;
        if (metadata.IsDraft || asTemplate) messageFlags |= 0x0008;
        if (metadata.IsRead == true) messageFlags |= 0x0001 | 0x0400;
        properties.Set(0x0E07, MapiPropertyType.Integer32, messageFlags);
        properties.Set(0x0E1B, MapiPropertyType.Boolean, hasAttachments);
        properties.Set(0x0037, MapiPropertyType.Unicode, document.Subject);
        ResolveSubject(document.Subject, metadata, out string? subjectPrefix, out string? normalizedSubject);
        properties.Set(0x003D, MapiPropertyType.Unicode, subjectPrefix);
        properties.Set(0x0E1D, MapiPropertyType.Unicode, normalizedSubject);
        properties.Set(0x0070, MapiPropertyType.Unicode, metadata.ConversationTopic ?? normalizedSubject);
        properties.Set(0x0071, MapiPropertyType.Binary, metadata.ConversationIndex);
        properties.Set(0x1000, MapiPropertyType.Unicode, document.Body.Text);
        properties.Set(0x1013, MapiPropertyType.Binary, null);
        properties.Set(0x1016, MapiPropertyType.Integer32, null);
        properties.Set(0x1009, MapiPropertyType.Binary, null);
        properties.Set(0x0E1F, MapiPropertyType.Boolean, null);
        if (document.Body.Html != null) {
            properties.Set(0x1013, MapiPropertyType.Binary,
                EncodeString8(document.Body.Html, codePage, diagnostics, string.Concat(location, "/html")));
            properties.Set(0x1016, MapiPropertyType.Integer32, 3);
        } else if (document.Body.Rtf != null) {
            properties.Set(0x1016, MapiPropertyType.Integer32, 2);
        } else if (document.Body.Text != null) {
            properties.Set(0x1016, MapiPropertyType.Integer32, 1);
        }
        if (document.Body.Rtf != null) {
            if (EmailRtfByteCodec.TryEncode(document.Body.Rtf, out byte[] rtfBytes)) {
                properties.Set(0x1009, MapiPropertyType.Binary, MapiCompressedRtfCodec.Compress(rtfBytes));
                properties.Set(0x0E1F, MapiPropertyType.Boolean, true);
            } else {
                diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_RTF_CHARACTER_UNENCODABLE",
                    "The RTF source contains a character above U+00FF. Serialize it through OfficeIMO.Rtf so the character is represented by an RTF escape.",
                    EmailDiagnosticSeverity.Error, location));
            }
        }
        properties.Set(0x1035, MapiPropertyType.Unicode,
            string.IsNullOrWhiteSpace(document.MessageId) ? null : string.Concat("<", document.MessageId!.Trim().Trim('<', '>'), ">"));
        properties.Set(0x0039, MapiPropertyType.Time, document.Date);
        properties.Set(0x0E06, MapiPropertyType.Time, document.ReceivedDate);
        properties.Set(0x1039, MapiPropertyType.Unicode, metadata.InternetReferences);
        properties.Set(0x1042, MapiPropertyType.Unicode, metadata.InReplyToId);
        properties.Set(0x0017, MapiPropertyType.Integer32, (int)(metadata.Importance ?? EmailMessageImportance.Normal));
        properties.Set(0x0026, MapiPropertyType.Integer32, (int)(metadata.Priority ?? EmailMessagePriority.Normal));
        properties.Set(0x1080, MapiPropertyType.Integer32,
            metadata.IconIndex ?? (metadata.IsDraft || asTemplate
                ? 0x00000103
                : metadata.IsRead == true ? 0x00000100 : 0x00000101));
        properties.Set(0x0029, MapiPropertyType.Boolean, metadata.ReadReceiptRequested);
        properties.Set(0x0023, MapiPropertyType.Boolean, metadata.DeliveryReceiptRequested);
        properties.Set(0x0036, MapiPropertyType.Integer32, metadata.Sensitivity);
        properties.Set(0x002E, MapiPropertyType.Integer32, metadata.OriginalSensitivity);
        DateTimeOffset created = metadata.CreatedDate ?? document.Date ?? document.ReceivedDate ?? FallbackCreationTime;
        properties.Set(0x3007, MapiPropertyType.Time, created);
        properties.Set(0x3008, MapiPropertyType.Time, metadata.ModifiedDate ?? created);
        properties.Set(0x3FDE, MapiPropertyType.Integer32, codePage);
        properties.Set(0x3FFD, MapiPropertyType.Integer32, codePage);
        properties.Set(0x3FF1, MapiPropertyType.Integer32, metadata.LocaleId ?? 1033);
        properties.Set(0x3FFA, MapiPropertyType.Unicode, metadata.LastModifierName);
        properties.Set(0x3013, MapiPropertyType.Binary, metadata.ConversationId);
        properties.Set(0x5909, MapiPropertyType.Integer32, metadata.EditorFormat);
        properties.SetNamed(MsgProjection.PsetidReactions, "ReactionsSummary", MapiPropertyType.Binary,
            metadata.ReactionsSummary);
        properties.SetNamed(MsgProjection.PsetidReactions, "OwnerReactionHistory", MapiPropertyType.Binary,
            metadata.OwnerReactionHistory);
        properties.SetNamed(MsgProjection.PsetidReactions, "OwnerReactionType", MapiPropertyType.Unicode,
            metadata.OwnerReactionType);
        properties.SetNamed(MsgProjection.PsetidReactions, "OwnerReactionTime", MapiPropertyType.Time,
            metadata.OwnerReactionTime);
        properties.SetNamed(MsgProjection.PsetidReactions, "ReactionsCount", MapiPropertyType.Integer32,
            metadata.ReactionsCount);
        properties.Set(0x0042, MapiPropertyType.Unicode, document.From?.DisplayName);
        properties.Set(0x0065, MapiPropertyType.Unicode, document.From?.Address);
        properties.Set(0x0064, MapiPropertyType.Unicode, document.From?.AddressType ?? "SMTP");
        properties.Set(0x5D02, MapiPropertyType.Unicode, document.From?.Address);
        EmailAddress? sender = document.Sender ?? document.From;
        properties.Set(0x0C1A, MapiPropertyType.Unicode, sender?.DisplayName);
        properties.Set(0x0C1F, MapiPropertyType.Unicode, sender?.Address);
        properties.Set(0x0C1E, MapiPropertyType.Unicode, sender?.AddressType ?? "SMTP");
        properties.Set(0x5D01, MapiPropertyType.Unicode, sender?.Address);
        properties.Set(0x0C19, MapiPropertyType.Binary,
            sender == null ? null : MsgIdentity.CreateOneOffEntryId(sender));
        SetReceivedAddress(properties, document.ReceivedBy, 0x0040, 0x0075, 0x0076, 0x003F);
        SetReceivedAddress(properties, document.ReceivedRepresenting, 0x0044, 0x0077, 0x0078, 0x0043);
        properties.Set(0x0E04, MapiPropertyType.Unicode, JoinRecipients(document, EmailRecipientKind.To));
        properties.Set(0x0E03, MapiPropertyType.Unicode, JoinRecipients(document, EmailRecipientKind.Cc));
        properties.Set(0x0E02, MapiPropertyType.Unicode, JoinRecipients(document, EmailRecipientKind.Bcc));
        properties.Set(0x0050, MapiPropertyType.Unicode, JoinRecipients(document, EmailRecipientKind.ReplyTo));
        properties.SetNamed(MsgProjection.PsPublicStrings, "Keywords", MapiPropertyType.MultipleUnicode,
            metadata.Categories.Count == 0 ? null : metadata.Categories.Cast<object>().ToArray());
        string? headers = MimeWriter.CreateTransportHeaders(document, options ?? EmailWriterOptions.Default);
        properties.Set(0x007D, MapiPropertyType.Unicode, headers);
        AddTypedProperties(properties, document);
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
        ushort displayNameId, ushort addressTypeId, ushort addressId, ushort entryId) {
        properties.Set(displayNameId, MapiPropertyType.Unicode, address?.DisplayName);
        properties.Set(addressTypeId, MapiPropertyType.Unicode, address?.AddressType ?? (address == null ? null : "SMTP"));
        properties.Set(addressId, MapiPropertyType.Unicode, address?.Address);
        properties.Set(entryId, MapiPropertyType.Binary,
            address == null ? null : MsgIdentity.CreateOneOffEntryId(address));
    }

    internal static MsgPropertyBuilder CreateRecipientProperties(EmailRecipient recipient, int index) {
        var properties = new MsgPropertyBuilder(recipient.MapiProperties);
        int type = recipient.Kind == EmailRecipientKind.To ? 1 : recipient.Kind == EmailRecipientKind.Cc ? 2 :
            recipient.Kind == EmailRecipientKind.Bcc ? 3 :
            recipient.Kind == EmailRecipientKind.Resource || recipient.Kind == EmailRecipientKind.Room ? 4 : 0;
        string addressType = string.IsNullOrWhiteSpace(recipient.Address.AddressType) ? "SMTP" : recipient.Address.AddressType!;
        string? address = recipient.Address.Address;
        properties.Set(0x3000, MapiPropertyType.Integer32, recipient.MapiRowId ?? index);
        properties.Set(0x0FFF, MapiPropertyType.Binary, MsgIdentity.CreateOneOffEntryId(recipient.Address));
        properties.Set(0x0C15, MapiPropertyType.Integer32, type);
        properties.Set(0x3001, MapiPropertyType.Unicode, recipient.Address.DisplayName ?? recipient.Address.Address);
        properties.Set(0x3002, MapiPropertyType.Unicode, addressType);
        properties.Set(0x3003, MapiPropertyType.Unicode, address);
        properties.Set(0x39FE, MapiPropertyType.Unicode, address);
        properties.Set(0x300B, MapiPropertyType.Binary, MsgIdentity.CreateSearchKey(addressType, address));
        properties.Set(0x0FFE, MapiPropertyType.Integer32, recipient.MapiObjectType ?? 6);
        properties.Set(0x3900, MapiPropertyType.Integer32, recipient.MapiDisplayType ?? 0);
        properties.Set(0x3905, MapiPropertyType.Integer32,
            recipient.MapiDisplayTypeEx ?? (recipient.Kind == EmailRecipientKind.Room ? 7 : 0));
        return properties;
    }

    internal static MsgPropertyBuilder CreateAttachmentProperties(EmailAttachment attachment, int index, int method,
        IList<EmailDiagnostic> diagnostics, string location, bool hasRetainedObjectContent = false,
        byte[]? materializedContent = null, long? retainedContentLength = null) {
        var properties = new MsgPropertyBuilder(attachment.MapiProperties);
        properties.Set(0x0FFE, MapiPropertyType.Integer32, 7);
        properties.Set(0x3705, MapiPropertyType.Integer32, method);
        properties.Set(0x0E21, MapiPropertyType.Integer32, index);
        properties.Set(0x3707, MapiPropertyType.Unicode, attachment.FileName);
        properties.Set(0x3704, MapiPropertyType.Unicode, attachment.FileName);
        properties.Set(0x3001, MapiPropertyType.Unicode, attachment.FileName);
        properties.Set(0x370E, MapiPropertyType.Unicode, attachment.ContentType);
        properties.Set(0x3712, MapiPropertyType.Unicode, attachment.ContentId);
        properties.Set(0x3713, MapiPropertyType.Unicode, attachment.ContentLocation);
        properties.Set(0x3703, MapiPropertyType.Unicode, GetLogicalExtension(attachment.FileName));
        properties.Set(0x370B, MapiPropertyType.Integer32,
            attachment.RenderingPosition >= 0 ? attachment.RenderingPosition : attachment.IsInline ? 0 : -1);
        properties.Set(0x3714, MapiPropertyType.Integer32, attachment.IsInline ? 0x00000004 : 0);
        properties.Set(0x7FFE, MapiPropertyType.Boolean, attachment.IsHidden || attachment.IsInline);
        properties.Set(0x7FFF, MapiPropertyType.Boolean, attachment.IsContactPhoto);
        properties.Set(0x3007, MapiPropertyType.Time, attachment.CreatedDate);
        properties.Set(0x3008, MapiPropertyType.Time, attachment.ModifiedDate ?? attachment.CreatedDate);
        properties.Set(0x370D, MapiPropertyType.Unicode, attachment.LinkedPath);
        if (method == 5 || method == 6) {
            properties.Set(0x3701, MapiPropertyType.Object, null);
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
            properties.Set(0x3701, MapiPropertyType.Binary, content);
            properties.Set(0x0E20, MapiPropertyType.Integer32, content.Length);
        } else if (hasRetainedObjectContent) {
            // Keep the variable-property row so a caller can provide a reopenable stream override.
            properties.Set(0x3701, MapiPropertyType.Binary, Array.Empty<byte>());
            properties.Set(0x0E20, MapiPropertyType.Integer32,
                checked((int)Math.Min(retainedContentLength ?? attachment.Length, int.MaxValue)));
        } else {
            properties.Set(0x3701, MapiPropertyType.Binary, null);
            properties.Set(0x0E20, MapiPropertyType.Integer32, checked((int)Math.Min(attachment.Length, int.MaxValue)));
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

    private static void AddTypedProperties(MsgPropertyBuilder properties, EmailDocument document) {
        if (document.Appointment != null) {
            OutlookAppointment item = document.Appointment;
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x820D, MapiPropertyType.Time, item.Start);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x820E, MapiPropertyType.Time, item.End);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8516, MapiPropertyType.Time, item.Start);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8517, MapiPropertyType.Time, item.End);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8208, MapiPropertyType.Unicode, item.Location);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8215, MapiPropertyType.Boolean, item.IsAllDay);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8205, MapiPropertyType.Integer32, item.BusyStatus);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8217, MapiPropertyType.Integer32, item.MeetingStatus);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8218, MapiPropertyType.Integer32, item.ResponseStatus);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8201, MapiPropertyType.Integer32, item.Sequence);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8213, MapiPropertyType.Integer32,
                item.DurationMinutes ?? GetDurationMinutes(item.Start, item.End));
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8238, MapiPropertyType.Unicode,
                item.AllAttendees ?? JoinAppointmentAttendees(document, null));
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x823B, MapiPropertyType.Unicode,
                item.RequiredAttendees ?? JoinAppointmentAttendees(document, EmailRecipientKind.To));
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x823C, MapiPropertyType.Unicode,
                item.OptionalAttendees ?? JoinAppointmentAttendees(document, EmailRecipientKind.Cc));
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x825A, MapiPropertyType.Boolean, item.NotAllowPropose);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8231, MapiPropertyType.Integer32, item.RecurrenceType);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8232, MapiPropertyType.Unicode, item.RecurrencePattern);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8216, MapiPropertyType.Binary, item.RecurrenceState);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8223, MapiPropertyType.Boolean,
                item.IsRecurring ?? (item.RecurrenceState != null || item.RecurrenceType.GetValueOrDefault() != 0));
            properties.SetNamed(MsgProjection.PsetidCalendarAssistant, 0x0015, MapiPropertyType.Integer32, item.ClientIntentFlags);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8501, MapiPropertyType.Integer32, item.ReminderDeltaMinutes);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8502, MapiPropertyType.Time, item.ReminderTime);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8503, MapiPropertyType.Boolean, item.ReminderIsSet);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8560, MapiPropertyType.Time, item.ReminderSignalTime);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8233, MapiPropertyType.Binary, item.TimeZoneStructure);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8234, MapiPropertyType.Unicode, item.TimeZoneDescription);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x825E, MapiPropertyType.Binary, item.StartTimeZoneDefinition);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x825F, MapiPropertyType.Binary, item.EndTimeZoneDefinition);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8260, MapiPropertyType.Binary, item.RecurrenceTimeZoneDefinition);
        }
        if (document.Contact != null) {
            OutlookContact item = document.Contact;
            properties.Set(0x3001, MapiPropertyType.Unicode, item.DisplayName);
            properties.Set(0x3A45, MapiPropertyType.Unicode, item.Prefix);
            properties.Set(0x3A0A, MapiPropertyType.Unicode, item.Initials);
            properties.Set(0x3A06, MapiPropertyType.Unicode, item.GivenName);
            properties.Set(0x3A44, MapiPropertyType.Unicode, item.MiddleName);
            properties.Set(0x3A11, MapiPropertyType.Unicode, item.Surname);
            properties.Set(0x3A05, MapiPropertyType.Unicode, item.Generation);
            properties.Set(0x3A16, MapiPropertyType.Unicode, item.CompanyName);
            properties.Set(0x3A17, MapiPropertyType.Unicode, item.JobTitle);
            properties.Set(0x3A18, MapiPropertyType.Unicode, item.Department);
            properties.Set(0x3A4F, MapiPropertyType.Unicode, item.NickName);
            properties.Set(0x3A4E, MapiPropertyType.Unicode, item.ManagerName);
            properties.Set(0x3A30, MapiPropertyType.Unicode, item.AssistantName);
            properties.Set(0x3A48, MapiPropertyType.Unicode, item.SpouseName);
            properties.Set(0x3A58, MapiPropertyType.Unicode,
                item.Children.Count == 0 ? null : string.Join(", ", item.Children));
            properties.Set(0x3A46, MapiPropertyType.Unicode, item.Profession);
            properties.Set(0x3A0C, MapiPropertyType.Unicode, item.Language);
            properties.Set(0x3A0D, MapiPropertyType.Unicode, item.Location);
            properties.Set(0x3A19, MapiPropertyType.Unicode, item.OfficeLocation);
            properties.Set(0x3A42, MapiPropertyType.Time, item.Birthday);
            properties.Set(0x3A41, MapiPropertyType.Time, item.WeddingAnniversary);
            properties.Set(0x3A51, MapiPropertyType.Unicode, item.BusinessHomePage);
            properties.Set(0x3A50, MapiPropertyType.Unicode, item.PersonalHomePage);
            properties.SetNamed(MsgProjection.PsetidAddress, 0x8005, MapiPropertyType.Unicode, item.FileAs);
            properties.SetNamed(MsgProjection.PsetidAddress, 0x8062, MapiPropertyType.Unicode, item.InstantMessagingAddress);
            properties.SetNamed(MsgProjection.PsetidAddress, 0x80DE, MapiPropertyType.Time, item.Birthday);
            properties.SetNamed(MsgProjection.PsetidAddress, 0x80DF, MapiPropertyType.Time, item.WeddingAnniversary);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8506, MapiPropertyType.Boolean, item.IsPrivate);
            properties.SetNamed(MsgProjection.PsetidAddress, 0x8015, MapiPropertyType.Boolean,
                item.HasPicture ?? document.Attachments.Any(attachment => attachment.IsContactPhoto));
            properties.SetNamed(MsgProjection.PsetidAddress, 0x802B, MapiPropertyType.Unicode, item.Html);
            AddContactAddressProperties(properties, item);
            AddContactPhoneProperties(properties, item.Phones);
            AddContactEmailProperties(properties, item.Email1, 0x8080, 0x8082, 0x8083, 0x8084, 0x8085);
            AddContactEmailProperties(properties, item.Email2, 0x8090, 0x8092, 0x8093, 0x8094, 0x8095);
            AddContactEmailProperties(properties, item.Email3, 0x80A0, 0x80A2, 0x80A3, 0x80A4, 0x80A5);
        }
        if (document.Task != null) {
            OutlookTask item = document.Task;
            properties.SetNamed(MsgProjection.PsetidTask, 0x8104, MapiPropertyType.Time, item.Start);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8105, MapiPropertyType.Time, item.Due);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8101, MapiPropertyType.Integer32, item.Status);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8102, MapiPropertyType.Floating64, item.PercentComplete);
            properties.SetNamed(MsgProjection.PsetidTask, 0x811C, MapiPropertyType.Boolean, item.IsComplete);
            properties.SetNamed(MsgProjection.PsetidTask, 0x811F, MapiPropertyType.Unicode, item.Owner);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8110, MapiPropertyType.Integer32, ToMinutes(item.ActualEffort));
            properties.SetNamed(MsgProjection.PsetidTask, 0x8111, MapiPropertyType.Integer32, ToMinutes(item.EstimatedEffort));
            properties.SetNamed(MsgProjection.PsetidTask, 0x811B, MapiPropertyType.Boolean, item.SendUpdates);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8119, MapiPropertyType.Boolean, item.SendStatusOnComplete);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8129, MapiPropertyType.Integer32, item.Ownership);
            properties.SetNamed(MsgProjection.PsetidTask, 0x812A, MapiPropertyType.Integer32, item.AcceptanceState);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8112, MapiPropertyType.Integer32, item.Version);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8113, MapiPropertyType.Integer32, item.State);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8121, MapiPropertyType.Unicode, item.Assigner);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8103, MapiPropertyType.Boolean, item.IsTeamTask);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8123, MapiPropertyType.Integer32, item.Ordinal);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8223, MapiPropertyType.Boolean, item.IsRecurring);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8501, MapiPropertyType.Integer32, item.ReminderDeltaMinutes);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8502, MapiPropertyType.Time, item.ReminderTime);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8503, MapiPropertyType.Boolean, item.ReminderIsSet);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8560, MapiPropertyType.Time, item.ReminderSignalTime);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8516, MapiPropertyType.Time, item.CommonStart);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8517, MapiPropertyType.Time, item.CommonEnd);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8518, MapiPropertyType.Integer32, item.Mode);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x85A0, MapiPropertyType.Time, item.ToDoOrdinalDate);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x85A1, MapiPropertyType.Unicode, item.ToDoSubOrdinal);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x853A, MapiPropertyType.MultipleUnicode,
                ToObjectArray(item.Contacts));
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8539, MapiPropertyType.MultipleUnicode,
                ToObjectArray(item.Companies));
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8535, MapiPropertyType.Unicode, item.BillingInformation);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8534, MapiPropertyType.Unicode, item.Mileage);
            properties.Set(0x1091, MapiPropertyType.Time, item.CompletedAt);
        }
        if (document.Journal != null) {
            OutlookJournal item = document.Journal;
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8516, MapiPropertyType.Time, item.Start);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8517, MapiPropertyType.Time, item.End);
            properties.SetNamed(MsgProjection.PsetidLog, 0x8700, MapiPropertyType.Unicode, item.Type);
            properties.SetNamed(MsgProjection.PsetidLog, 0x8706, MapiPropertyType.Time, item.Start);
            properties.SetNamed(MsgProjection.PsetidLog, 0x8707, MapiPropertyType.Integer32, item.DurationMinutes);
            properties.SetNamed(MsgProjection.PsetidLog, 0x8708, MapiPropertyType.Time, item.End);
            properties.SetNamed(MsgProjection.PsetidLog, 0x870C, MapiPropertyType.Integer32, item.Flags);
            properties.SetNamed(MsgProjection.PsetidLog, 0x870E, MapiPropertyType.Boolean, item.DocumentPrinted);
            properties.SetNamed(MsgProjection.PsetidLog, 0x870F, MapiPropertyType.Boolean, item.DocumentSaved);
            properties.SetNamed(MsgProjection.PsetidLog, 0x8710, MapiPropertyType.Boolean, item.DocumentRouted);
            properties.SetNamed(MsgProjection.PsetidLog, 0x8711, MapiPropertyType.Boolean, item.DocumentPosted);
            properties.SetNamed(MsgProjection.PsetidLog, 0x8712, MapiPropertyType.Unicode, item.TypeDescription);
        }
        if (document.Note != null) {
            OutlookNote item = document.Note;
            properties.SetNamed(MsgProjection.PsetidNote, 0x8B00, MapiPropertyType.Integer32, item.Color);
            properties.SetNamed(MsgProjection.PsetidNote, 0x8B02, MapiPropertyType.Integer32, item.Width);
            properties.SetNamed(MsgProjection.PsetidNote, 0x8B03, MapiPropertyType.Integer32, item.Height);
            properties.SetNamed(MsgProjection.PsetidNote, 0x8B04, MapiPropertyType.Integer32, item.X);
            properties.SetNamed(MsgProjection.PsetidNote, 0x8B05, MapiPropertyType.Integer32, item.Y);
        }
    }

    private static void AddContactAddressProperties(MsgPropertyBuilder properties, OutlookContact contact) {
        AddFixedAddress(properties, contact.BusinessAddress, 0x3A29, 0x3A27, 0x3A28, 0x3A2A, 0x3A26, 0x3A2B);
        AddFixedAddress(properties, contact.HomeAddress, 0x3A5D, 0x3A59, 0x3A5C, 0x3A5B, 0x3A5A, 0x3A5E);
        AddFixedAddress(properties, contact.OtherAddress, 0x3A63, 0x3A5F, 0x3A62, 0x3A61, 0x3A60, 0x3A64);
        properties.Set(0x3A15, MapiPropertyType.Unicode, contact.BusinessAddress.Formatted);
        properties.SetNamed(MsgProjection.PsetidAddress, 0x801A, MapiPropertyType.Unicode, contact.HomeAddress.Formatted);
        properties.SetNamed(MsgProjection.PsetidAddress, 0x801B, MapiPropertyType.Unicode,
            contact.WorkAddress.Formatted ?? contact.BusinessAddress.Formatted);
        properties.SetNamed(MsgProjection.PsetidAddress, 0x801C, MapiPropertyType.Unicode, contact.OtherAddress.Formatted);
        properties.SetNamed(MsgProjection.PsetidAddress, 0x8045, MapiPropertyType.Unicode,
            contact.WorkAddress.Street ?? contact.BusinessAddress.Street);
        properties.SetNamed(MsgProjection.PsetidAddress, 0x8046, MapiPropertyType.Unicode,
            contact.WorkAddress.City ?? contact.BusinessAddress.City);
        properties.SetNamed(MsgProjection.PsetidAddress, 0x8047, MapiPropertyType.Unicode,
            contact.WorkAddress.StateOrProvince ?? contact.BusinessAddress.StateOrProvince);
        properties.SetNamed(MsgProjection.PsetidAddress, 0x8048, MapiPropertyType.Unicode,
            contact.WorkAddress.PostalCode ?? contact.BusinessAddress.PostalCode);
        properties.SetNamed(MsgProjection.PsetidAddress, 0x8049, MapiPropertyType.Unicode,
            contact.WorkAddress.Country ?? contact.BusinessAddress.Country);
        properties.SetNamed(MsgProjection.PsetidAddress, 0x804A, MapiPropertyType.Unicode,
            contact.WorkAddress.PostOfficeBox ?? contact.BusinessAddress.PostOfficeBox);
        properties.SetNamed(MsgProjection.PsetidAddress, 0x80DB, MapiPropertyType.Unicode, contact.WorkAddress.CountryCode);
    }

    private static void AddFixedAddress(MsgPropertyBuilder properties, OutlookPostalAddress address,
        ushort streetId, ushort cityId, ushort stateId, ushort postalId, ushort countryId, ushort postOfficeBoxId) {
        properties.Set(streetId, MapiPropertyType.Unicode, address.Street);
        properties.Set(cityId, MapiPropertyType.Unicode, address.City);
        properties.Set(stateId, MapiPropertyType.Unicode, address.StateOrProvince);
        properties.Set(postalId, MapiPropertyType.Unicode, address.PostalCode);
        properties.Set(countryId, MapiPropertyType.Unicode, address.Country);
        properties.Set(postOfficeBoxId, MapiPropertyType.Unicode, address.PostOfficeBox);
    }

    private static void AddContactPhoneProperties(MsgPropertyBuilder properties, OutlookContactPhones phones) {
        properties.Set(0x3A08, MapiPropertyType.Unicode, phones.Business);
        properties.Set(0x3A1B, MapiPropertyType.Unicode, phones.Business2);
        properties.Set(0x3A09, MapiPropertyType.Unicode, phones.Home);
        properties.Set(0x3A2F, MapiPropertyType.Unicode, phones.Home2);
        properties.Set(0x3A1C, MapiPropertyType.Unicode, phones.Mobile);
        properties.Set(0x3A1F, MapiPropertyType.Unicode, phones.Other);
        properties.Set(0x3A1A, MapiPropertyType.Unicode, phones.Primary);
        properties.Set(0x3A24, MapiPropertyType.Unicode, phones.BusinessFax);
        properties.Set(0x3A25, MapiPropertyType.Unicode, phones.HomeFax);
        properties.Set(0x3A23, MapiPropertyType.Unicode, phones.PrimaryFax);
        properties.Set(0x3A2E, MapiPropertyType.Unicode, phones.Assistant);
        properties.Set(0x3A57, MapiPropertyType.Unicode, phones.CompanyMain);
        properties.Set(0x3A1E, MapiPropertyType.Unicode, phones.Car);
        properties.Set(0x3A1D, MapiPropertyType.Unicode, phones.Radio);
        properties.Set(0x3A21, MapiPropertyType.Unicode, phones.Pager);
        properties.Set(0x3A02, MapiPropertyType.Unicode, phones.Callback);
        properties.Set(0x3A2C, MapiPropertyType.Unicode, phones.Telex);
        properties.Set(0x3A4B, MapiPropertyType.Unicode, phones.TextTelephone);
        properties.Set(0x3A2D, MapiPropertyType.Unicode, phones.Isdn);
    }

    private static void AddContactEmailProperties(MsgPropertyBuilder properties, OutlookContactEmailAddress email,
        uint displayId, uint addressTypeId, uint addressId, uint originalDisplayId, uint entryId) {
        properties.SetNamed(MsgProjection.PsetidAddress, displayId, MapiPropertyType.Unicode, email.DisplayName);
        properties.SetNamed(MsgProjection.PsetidAddress, addressTypeId, MapiPropertyType.Unicode,
            email.AddressType ?? (email.Address == null ? null : "SMTP"));
        properties.SetNamed(MsgProjection.PsetidAddress, addressId, MapiPropertyType.Unicode, email.Address);
        properties.SetNamed(MsgProjection.PsetidAddress, originalDisplayId, MapiPropertyType.Unicode,
            email.OriginalDisplayName ?? email.Address);
        byte[]? originalEntryId = email.OriginalEntryId;
        if (originalEntryId == null && email.Address != null) {
            originalEntryId = MsgIdentity.CreateOneOffEntryId(new EmailAddress(email.Address, email.DisplayName) {
                AddressType = email.AddressType ?? "SMTP"
            });
        }
        properties.SetNamed(MsgProjection.PsetidAddress, entryId, MapiPropertyType.Binary, originalEntryId);
    }

    private static int? GetDurationMinutes(DateTimeOffset? start, DateTimeOffset? end) {
        if (!start.HasValue || !end.HasValue) return null;
        double minutes = Math.Round((end.Value - start.Value).TotalMinutes);
        return minutes >= int.MinValue && minutes <= int.MaxValue ? (int)minutes : (int?)null;
    }

    private static string? JoinAppointmentAttendees(EmailDocument document, EmailRecipientKind? kind) {
        string[] attendees = document.Recipients
            .Where(recipient => recipient.Kind != EmailRecipientKind.ReplyTo &&
                (!kind.HasValue || recipient.Kind == kind.Value))
            .Select(recipient => recipient.Address.DisplayName ?? recipient.Address.Address)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Cast<string>()
            .ToArray();
        return attendees.Length == 0 ? null : string.Join("; ", attendees);
    }

    private static int? ToMinutes(TimeSpan? value) {
        if (!value.HasValue) return null;
        double minutes = Math.Round(value.Value.TotalMinutes);
        return minutes >= int.MinValue && minutes <= int.MaxValue ? (int)minutes : (int?)null;
    }

    private static object[]? ToObjectArray(IList<string> values) =>
        values.Count == 0 ? null : values.Cast<object>().ToArray();

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
