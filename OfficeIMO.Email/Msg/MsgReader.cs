using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Email;

internal static class MsgReader {
    internal static bool TryRead(byte[] data, EmailReaderOptions options, IList<EmailDiagnostic> diagnostics,
        CancellationToken cancellationToken, out EmailDocument document) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeCompoundReadOptions compoundOptions = EmailCompoundReadPolicy.Create(options);
        OfficeCompoundFile? compound;
        string? error;
        try {
            if (!OfficeCompoundFileReader.TryRead(data, compoundOptions, out compound, out error) || compound == null) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_COMPOUND_INVALID", error ?? "The MSG compound file is invalid.",
                    EmailDiagnosticSeverity.Error));
                document = new EmailDocument { Format = EmailFileFormat.Unknown, OutlookItemKind = OutlookItemKind.Unknown };
                return false;
            }
        } catch (OfficeCompoundStreamLimitExceededException exception) {
            throw new EmailLimitExceededException(exception.LimitName, exception.ActualValue, exception.MaximumValue);
        }
        if (!compound.Streams.ContainsKey("__properties_version1.0")) {
            document = new EmailDocument { Format = EmailFileFormat.Unknown, OutlookItemKind = OutlookItemKind.Unknown };
            return false;
        }

        return TryRead(compound, options, diagnostics, cancellationToken, null, out document);
    }

    internal static bool TryRead(OfficeCompoundFile compound, EmailReaderOptions options,
        IList<EmailDiagnostic> diagnostics, CancellationToken cancellationToken,
        IReadOnlyDictionary<string, IEmailContentSource>? externalContent,
        out EmailDocument document) {
        if (!compound.Streams.ContainsKey("__properties_version1.0")) {
            document = new EmailDocument { Format = EmailFileFormat.Unknown, OutlookItemKind = OutlookItemKind.Unknown };
            return false;
        }

        MsgParserState state = new MsgParserState(options, diagnostics, cancellationToken);
        MsgNamedPropertyMap names = MsgNamedPropertyMap.Read(compound, diagnostics, state);
        document = ReadMessage(compound, string.Empty, MsgPropertyStreamKind.TopLevel, names, state, 0, null,
            externalContent);
        return true;
    }

    private static EmailDocument ReadMessage(OfficeCompoundFile compound, string prefix, MsgPropertyStreamKind kind,
        MsgNamedPropertyMap names, MsgParserState state, int nestedDepth,
        MapiStringEncodingContext? inheritedEncoding,
        IReadOnlyDictionary<string, IEmailContentSource>? externalContent) {
        state.ThrowIfCancellationRequested();
        var document = new EmailDocument { Format = EmailFileFormat.OutlookMsg };
        List<MapiProperty> messageProperties = MsgPropertyReader.Read(
            compound, prefix, kind, names, state, inheritedEncoding, out MapiStringEncodingContext encoding);
        document.OutlookCodePage = encoding.PrimaryCodePage;
        foreach (MapiProperty property in messageProperties) {
            document.MapiProperties.Add(property);
        }
        MsgProjection.Apply(document, state, string.IsNullOrEmpty(prefix) ? "msg" : prefix, encoding);

        foreach (string recipientPath in GetDirectChildStorages(compound, prefix, "__recip_version1.0_#")) {
            state.ThrowIfCancellationRequested();
            ReadRecipient(compound, recipientPath, names, state, document, encoding);
        }
        MsgProjection.ApplyTransportHeaderRecipients(document, state,
            string.IsNullOrEmpty(prefix) ? "msg" : prefix);
        foreach (string attachmentPath in GetDirectChildStorages(compound, prefix, "__attach_version1.0_#")) {
            state.ThrowIfCancellationRequested();
            ReadAttachment(compound, attachmentPath, names, state, document, nestedDepth, encoding,
                externalContent);
        }
        EmailProtectionProjection.Apply(document, state.Diagnostics,
            string.IsNullOrEmpty(prefix) ? "msg" : prefix);
        return document;
    }

    private static void ReadRecipient(OfficeCompoundFile compound, string path, MsgNamedPropertyMap names,
        MsgParserState state, EmailDocument document, MapiStringEncodingContext inheritedEncoding) {
        List<MapiProperty> properties = MsgPropertyReader.Read(
            compound, path, MsgPropertyStreamKind.ChildObject, names, state, inheritedEncoding, out _);
        EmailAddress? address = MsgAddressProjection.ReadAddress(
            properties,
            displayNameId: 0x3001,
            smtpAddressId: 0x39FE,
            emailAddressId: 0x3003,
            addressTypeId: 0x3002,
            originalAddressId: 0x403E);
        var recipient = new EmailRecipient(
            MsgAddressProjection.ReadRecipientKind(properties),
            address ?? new EmailAddress(null)) {
            MapiRowId = MsgProjection.GetInt(properties, 0x3000),
            MapiObjectType = MsgProjection.GetInt(properties, 0x0FFE),
            MapiDisplayType = MsgProjection.GetInt(properties, 0x3900),
            MapiDisplayTypeEx = MsgProjection.GetInt(properties, 0x3905)
        };
        foreach (MapiProperty property in properties) recipient.MapiProperties.Add(property);
        document.Recipients.Add(recipient);
    }

    private static void ReadAttachment(OfficeCompoundFile compound, string path, MsgNamedPropertyMap names,
        MsgParserState state, EmailDocument document, int nestedDepth, MapiStringEncodingContext inheritedEncoding,
        IReadOnlyDictionary<string, IEmailContentSource>? externalContent) {
        List<MapiProperty> properties = MsgPropertyReader.Read(
            compound, path, MsgPropertyStreamKind.ChildObject, names, state, inheritedEncoding, out _);
        int method = MsgProjection.GetInt(properties, 0x3705) ?? 1;
        byte[]? content = properties.FirstOrDefault(property => property.PropertyId == 0x3701)?.Value as byte[];
        string contentPath = MsgBinary.CombinePath(path, "__substg1.0_37010102");
        IEmailContentSource? externalSource = null;
        if (externalContent?.TryGetValue(contentPath, out externalSource) == true) content = null;
        var attachment = new EmailAttachment {
            FileName = MsgProjection.GetString(properties, 0x3707) ?? MsgProjection.GetString(properties, 0x3704) ??
                MsgProjection.GetString(properties, 0x3001),
            ContentType = MsgProjection.GetString(properties, 0x370E),
            ContentId = TrimAngle(MsgProjection.GetString(properties, 0x3712)),
            ContentLocation = MsgProjection.GetString(properties, 0x3713),
            IsInline = !string.IsNullOrWhiteSpace(MsgProjection.GetString(properties, 0x3712)) ||
                ((MsgProjection.GetInt(properties, 0x3714) ?? 0) & 0x00000004) != 0,
            IsHidden = MsgProjection.GetBool(properties, 0x7FFE) ?? false,
            IsContactPhoto = MsgProjection.GetBool(properties, 0x7FFF) ?? false,
            RenderingPosition = MsgProjection.GetInt(properties, 0x370B) ?? -1,
            CreatedDate = MsgProjection.GetDate(properties, 0x3007),
            ModifiedDate = MsgProjection.GetDate(properties, 0x3008),
            LinkedPath = MsgProjection.GetString(properties, 0x370D),
            MapiAttachMethod = method,
            Length = externalSource?.Length ?? content?.LongLength ??
                Math.Max(0, MsgProjection.GetInt(properties, 0x0E20) ?? 0)
        };
        foreach (MapiProperty property in properties) attachment.MapiProperties.Add(property);
        if (externalSource != null) {
            foreach (MapiProperty property in attachment.MapiProperties.Where(property => property.PropertyId == 0x3701)) {
                property.Value = null;
                property.RawData = null;
            }
        }

        string objectStorage = MsgBinary.CombinePath(path, "__substg1.0_3701000D");
        bool hasObjectStorage = compound.Entries.Any(entry => entry.IsStorage &&
            string.Equals(entry.Path, objectStorage, StringComparison.OrdinalIgnoreCase));
        if (method == 5 && hasObjectStorage) {
            long total = GetStorageLength(compound, objectStorage);
            attachment.Length = total;
            state.CountAttachment(total);
            if (nestedDepth >= state.Options.MaxNestedMessageDepth) {
                RetainStorageStreams(compound, objectStorage, attachment, state);
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_NESTED_MESSAGE_LIMIT",
                    state.Options.IncludeAttachmentContent
                        ? "The embedded MSG was retained as opaque storage but not projected because the nested-message limit was reached."
                        : "The embedded MSG was not projected because the nested-message limit was reached and attachment content retention is disabled.",
                    EmailDiagnosticSeverity.Warning, objectStorage));
            } else if (!compound.Streams.ContainsKey(
                MsgBinary.CombinePath(objectStorage, "__properties_version1.0"))) {
                RetainStorageStreams(compound, objectStorage, attachment, state);
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_EMBEDDED_STORAGE_INVALID",
                    state.Options.IncludeAttachmentContent
                        ? "The embedded MSG storage was retained but not projected because its property stream is missing."
                        : "The embedded MSG storage was not projected because its property stream is missing and attachment content retention is disabled.",
                    EmailDiagnosticSeverity.Warning, objectStorage));
            } else {
                attachment.EmbeddedDocument = ReadMessage(compound, objectStorage, MsgPropertyStreamKind.EmbeddedMessage,
                    names, state, nestedDepth + 1, inheritedEncoding, externalContent);
            }
        } else if (method == 6 && hasObjectStorage) {
            long total = GetStorageLength(compound, objectStorage);
            string storagePrefix = string.Concat(objectStorage, "/");
            foreach (KeyValuePair<string, byte[]> stream in compound.Streams.Where(item =>
                item.Key.StartsWith(storagePrefix, StringComparison.OrdinalIgnoreCase))) {
                state.ThrowIfCancellationRequested();
                string relative = stream.Key.Substring(storagePrefix.Length);
                if (state.Options.IncludeAttachmentContent) {
                    attachment.StructuredStorageStreams[relative] = stream.Value;
                }
            }
            attachment.Length = total;
            state.CountAttachment(total);
        } else {
            long length = content?.LongLength ?? attachment.Length;
            state.CountAttachment(length);
            attachment.Content = state.Options.IncludeAttachmentContent && content != null ? (byte[])content.Clone() : null;
            attachment.ContentSource = state.Options.IncludeAttachmentContent ? externalSource : null;
            if (content != null && IsTnef(content) && nestedDepth < state.Options.MaxNestedMessageDepth) {
                attachment.EmbeddedDocument = TnefReader.Read(content, state, nestedDepth + 1,
                    string.Concat(path, "/tnef"));
            } else if (content != null && IsTnef(content)) {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_NESTED_MESSAGE_LIMIT",
                    "The encapsulated TNEF attachment was retained but not projected because the nested-message limit was reached.",
                    EmailDiagnosticSeverity.Warning, path));
            }
        }

        if (!state.Options.IncludeAttachmentContent) {
            foreach (MapiProperty property in attachment.MapiProperties.Where(property => property.PropertyId == 0x3701)) {
                property.Value = null;
                property.RawData = null;
            }
        }
        document.Attachments.Add(attachment);
    }

    private static void RetainStorageStreams(OfficeCompoundFile compound, string storagePath,
        EmailAttachment attachment, MsgParserState state) {
        if (!state.Options.IncludeAttachmentContent) return;
        string storagePrefix = string.Concat(storagePath, "/");
        foreach (KeyValuePair<string, byte[]> stream in compound.Streams.Where(item =>
            item.Key.StartsWith(storagePrefix, StringComparison.OrdinalIgnoreCase))) {
            state.ThrowIfCancellationRequested();
            attachment.StructuredStorageStreams[stream.Key.Substring(storagePrefix.Length)] = stream.Value;
        }
    }

    private static long GetStorageLength(OfficeCompoundFile compound, string storagePath) {
        string storagePrefix = string.Concat(storagePath, "/");
        long total = 0;
        foreach (KeyValuePair<string, byte[]> stream in compound.Streams.Where(item =>
            item.Key.StartsWith(storagePrefix, StringComparison.OrdinalIgnoreCase))) {
            total = checked(total + stream.Value.LongLength);
        }
        return total;
    }

    private static IEnumerable<string> GetDirectChildStorages(OfficeCompoundFile compound, string parentPath, string prefix) {
        string pathPrefix = string.IsNullOrEmpty(parentPath) ? string.Empty : string.Concat(parentPath, "/");
        return compound.Entries
            .Where(entry => entry.IsStorage && entry.Path.StartsWith(pathPrefix, StringComparison.OrdinalIgnoreCase))
            .Where(entry => !entry.IsFallback)
            .Select(entry => entry.Path.Substring(pathPrefix.Length))
            .Where(relative => relative.IndexOf('/') < 0 && relative.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .OrderBy(relative => relative, StringComparer.OrdinalIgnoreCase)
            .Select(relative => string.Concat(pathPrefix, relative))
            .ToArray();
    }

    private static string? TrimAngle(string? value) {
        return string.IsNullOrWhiteSpace(value) ? value : value!.Trim().Trim('<', '>');
    }

    private static bool IsTnef(byte[] bytes) => bytes.Length >= 4 && MsgBinary.ReadUInt32(bytes, 0) == 0x223E9F78;
}
