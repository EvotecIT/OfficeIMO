using OfficeIMO.Shared;

namespace OfficeIMO.Email;

internal static class MsgReader {
    internal static bool TryRead(byte[] data, EmailReaderOptions options, IList<EmailDiagnostic> diagnostics,
        CancellationToken cancellationToken, out EmailDocument document) {
        cancellationToken.ThrowIfCancellationRequested();
        var compoundOptions = new OfficeCompoundReadOptions(
            options.MaxCompoundDirectoryEntries,
            options.MaxCompoundDirectoryEntries,
            Math.Min(options.MaxInputBytes, int.MaxValue),
            options.MaxDecodedPropertyBytes);
        if (!OfficeCompoundFileReader.TryRead(data, compoundOptions, out OfficeCompoundFile? compound, out string? error) || compound == null) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_COMPOUND_INVALID", error ?? "The MSG compound file is invalid.",
                EmailDiagnosticSeverity.Error));
            document = new EmailDocument { Format = EmailFileFormat.Unknown, OutlookItemKind = OutlookItemKind.Unknown };
            return false;
        }
        if (!compound.Streams.ContainsKey("__properties_version1.0")) {
            document = new EmailDocument { Format = EmailFileFormat.Unknown, OutlookItemKind = OutlookItemKind.Unknown };
            return false;
        }

        MsgParserState state = new MsgParserState(options, diagnostics, cancellationToken);
        MsgNamedPropertyMap names = MsgNamedPropertyMap.Read(compound, diagnostics, state);
        document = ReadMessage(compound, string.Empty, MsgPropertyStreamKind.TopLevel, names, state, 0);
        return true;
    }

    private static EmailDocument ReadMessage(OfficeCompoundFile compound, string prefix, MsgPropertyStreamKind kind,
        MsgNamedPropertyMap names, MsgParserState state, int nestedDepth) {
        state.ThrowIfCancellationRequested();
        var document = new EmailDocument { Format = EmailFileFormat.OutlookMsg };
        foreach (MapiProperty property in MsgPropertyReader.Read(compound, prefix, kind, names, state)) {
            document.MapiProperties.Add(property);
        }
        MsgProjection.Apply(document, state.Options, state.Diagnostics, string.IsNullOrEmpty(prefix) ? "msg" : prefix);

        foreach (string recipientPath in GetDirectChildStorages(compound, prefix, "__recip_version1.0_#")) {
            state.ThrowIfCancellationRequested();
            ReadRecipient(compound, recipientPath, names, state, document);
        }
        foreach (string attachmentPath in GetDirectChildStorages(compound, prefix, "__attach_version1.0_#")) {
            state.ThrowIfCancellationRequested();
            ReadAttachment(compound, attachmentPath, names, state, document, nestedDepth);
        }
        return document;
    }

    private static void ReadRecipient(OfficeCompoundFile compound, string path, MsgNamedPropertyMap names,
        MsgParserState state, EmailDocument document) {
        List<MapiProperty> properties = MsgPropertyReader.Read(compound, path, MsgPropertyStreamKind.ChildObject, names, state);
        int recipientType = MsgProjection.GetInt(properties, 0x0C15) ?? 0;
        EmailRecipientKind kind = recipientType == 1 ? EmailRecipientKind.To :
            recipientType == 2 ? EmailRecipientKind.Cc : recipientType == 3 ? EmailRecipientKind.Bcc : EmailRecipientKind.Unknown;
        string? displayName = MsgProjection.GetString(properties, 0x3001);
        string? address = MsgProjection.GetString(properties, 0x39FE) ?? MsgProjection.GetString(properties, 0x3003);
        string? addressType = MsgProjection.GetString(properties, 0x3002);
        var recipient = new EmailRecipient(kind, new EmailAddress(address, displayName,
            addressType == null ? address : string.Concat(addressType, ":", address)));
        foreach (MapiProperty property in properties) recipient.MapiProperties.Add(property);
        document.Recipients.Add(recipient);
    }

    private static void ReadAttachment(OfficeCompoundFile compound, string path, MsgNamedPropertyMap names,
        MsgParserState state, EmailDocument document, int nestedDepth) {
        List<MapiProperty> properties = MsgPropertyReader.Read(compound, path, MsgPropertyStreamKind.ChildObject, names, state);
        int method = MsgProjection.GetInt(properties, 0x3705) ?? 1;
        byte[]? content = properties.FirstOrDefault(property => property.PropertyId == 0x3701)?.Value as byte[];
        var attachment = new EmailAttachment {
            FileName = MsgProjection.GetString(properties, 0x3707) ?? MsgProjection.GetString(properties, 0x3704) ??
                MsgProjection.GetString(properties, 0x3001),
            ContentType = MsgProjection.GetString(properties, 0x370E),
            ContentId = TrimAngle(MsgProjection.GetString(properties, 0x3712)),
            ContentLocation = MsgProjection.GetString(properties, 0x3713),
            IsInline = !string.IsNullOrWhiteSpace(MsgProjection.GetString(properties, 0x3712)),
            MapiAttachMethod = method,
            Length = content?.LongLength ?? Math.Max(0, MsgProjection.GetInt(properties, 0x0E20) ?? 0)
        };
        foreach (MapiProperty property in properties) attachment.MapiProperties.Add(property);

        string objectStorage = MsgBinary.CombinePath(path, "__substg1.0_3701000D");
        bool hasObjectStorage = compound.Entries.Any(entry => entry.IsStorage &&
            string.Equals(entry.Path, objectStorage, StringComparison.OrdinalIgnoreCase));
        if (method == 5 && hasObjectStorage) {
            state.CountAttachment(0);
            if (nestedDepth >= state.Options.MaxNestedMessageDepth) {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_NESTED_MESSAGE_LIMIT",
                    "The embedded MSG was retained but not projected because the nested-message limit was reached.",
                    EmailDiagnosticSeverity.Warning, objectStorage));
            } else {
                attachment.EmbeddedDocument = ReadMessage(compound, objectStorage, MsgPropertyStreamKind.EmbeddedMessage,
                    names, state, nestedDepth + 1);
            }
        } else if (method == 6 && hasObjectStorage) {
            long total = 0;
            string storagePrefix = string.Concat(objectStorage, "/");
            foreach (KeyValuePair<string, byte[]> stream in compound.Streams.Where(item =>
                item.Key.StartsWith(storagePrefix, StringComparison.OrdinalIgnoreCase))) {
                state.ThrowIfCancellationRequested();
                string relative = stream.Key.Substring(storagePrefix.Length);
                attachment.StructuredStorageStreams[relative] = stream.Value;
                total = checked(total + stream.Value.LongLength);
            }
            attachment.Length = total;
            state.CountAttachment(total);
        } else {
            long length = content?.LongLength ?? attachment.Length;
            state.CountAttachment(length);
            attachment.Content = state.Options.IncludeAttachmentContent && content != null ? (byte[])content.Clone() : null;
        }

        if (!state.Options.IncludeAttachmentContent) {
            foreach (MapiProperty property in attachment.MapiProperties.Where(property => property.PropertyId == 0x3701)) {
                property.Value = null;
                property.RawData = null;
            }
        }
        document.Attachments.Add(attachment);
    }

    private static IEnumerable<string> GetDirectChildStorages(OfficeCompoundFile compound, string parentPath, string prefix) {
        string pathPrefix = string.IsNullOrEmpty(parentPath) ? string.Empty : string.Concat(parentPath, "/");
        return compound.Entries
            .Where(entry => entry.IsStorage && entry.Path.StartsWith(pathPrefix, StringComparison.OrdinalIgnoreCase))
            .Select(entry => entry.Path.Substring(pathPrefix.Length))
            .Where(relative => relative.IndexOf('/') < 0 && relative.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .OrderBy(relative => relative, StringComparer.OrdinalIgnoreCase)
            .Select(relative => string.Concat(pathPrefix, relative))
            .ToArray();
    }

    private static string? TrimAngle(string? value) {
        return string.IsNullOrWhiteSpace(value) ? value : value!.Trim().Trim('<', '>');
    }
}
