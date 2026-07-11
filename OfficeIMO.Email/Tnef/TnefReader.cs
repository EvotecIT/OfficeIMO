using OfficeIMO.Shared;

namespace OfficeIMO.Email;

internal static class TnefReader {
    private static readonly Guid IidMessage = new Guid("00020307-0000-0000-C000-000000000046");
    private static readonly Guid IidStorage = new Guid("0000000B-0000-0000-C000-000000000046");

    internal static EmailDocument Read(byte[] data, EmailReaderOptions options, IList<EmailDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        var state = new MsgParserState(options, diagnostics, cancellationToken);
        return ReadMessage(data, state, 0, "tnef");
    }

    internal static EmailDocument Read(byte[] data, MsgParserState state, int nestedDepth, string location) {
        return ReadMessage(data, state, nestedDepth, location);
    }

    private static EmailDocument ReadMessage(byte[] data, MsgParserState state, int nestedDepth, string location) {
        state.ThrowIfCancellationRequested();
        var document = new EmailDocument { Format = EmailFileFormat.Tnef, OutlookItemKind = OutlookItemKind.Message };
        if (data.Length < 6 || MsgBinary.ReadUInt32(data, 0) != TnefConstants.Signature) {
            state.Diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_SIGNATURE_INVALID",
                "The TNEF signature is missing or invalid.", EmailDiagnosticSeverity.Error, location));
            return document;
        }
        ushort key = MsgBinary.ReadUInt16(data, 4);
        if (key == 0) {
            state.Diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_KEY_INVALID",
                "The TNEF attachment key is zero.", EmailDiagnosticSeverity.Warning, location));
        }

        List<ParsedAttribute> attributes = ParseAttributes(data, state, location);
        int codePage = ReadCodePage(attributes);
        document.OutlookCodePage = codePage;
        EmailAttachment? currentAttachment = null;
        string? subject = null;
        string? body = null;
        string? messageClass = null;
        string? messageId = null;
        DateTimeOffset? date = null;
        DateTimeOffset? received = null;

        foreach (ParsedAttribute attribute in attributes) {
            state.ThrowIfCancellationRequested();
            var rawAttribute = new TnefAttribute(attribute.Level, attribute.Tag, attribute.Data, attribute.ChecksumIsValid);
            if (attribute.Level == TnefAttributeLevel.Message) {
                document.TnefAttributes.Add(rawAttribute);
                switch (attribute.Tag) {
                    case TnefConstants.Subject: subject = DecodeString(attribute.Data, codePage, state, location); break;
                    case TnefConstants.Body: body = DecodeString(attribute.Data, codePage, state, location); break;
                    case TnefConstants.MessageClass: messageClass = DecodeString(attribute.Data, codePage, state, location); break;
                    case TnefConstants.MessageId: messageId = DecodeString(attribute.Data, codePage, state, location); break;
                    case TnefConstants.DateSent: date = DecodeDate(attribute.Data, state.Diagnostics, location); break;
                    case TnefConstants.DateReceived: received = DecodeDate(attribute.Data, state.Diagnostics, location); break;
                    case TnefConstants.MessageProperties:
                        AddProperties(document.MapiProperties,
                            TnefMapiCodec.ReadProperties(attribute.Data, codePage, state, string.Concat(location, "/mapi")));
                        break;
                    case TnefConstants.RecipientTable:
                        AddRecipients(document, TnefMapiCodec.ReadRecipientTable(attribute.Data, codePage, state,
                            string.Concat(location, "/recipients")));
                        break;
                }
            } else {
                if (attribute.Tag == TnefConstants.AttachRendData || currentAttachment == null) {
                    currentAttachment = new EmailAttachment();
                    document.Attachments.Add(currentAttachment);
                    if (attribute.Tag != TnefConstants.AttachRendData) {
                        state.Diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_ATTACHMENT_BOUNDARY_MISSING",
                            "An attachment attribute appeared before attAttachRendData.", EmailDiagnosticSeverity.Warning, location));
                    }
                }
                currentAttachment.TnefAttributes.Add(rawAttribute);
                ApplyAttachmentAttribute(currentAttachment, attribute, codePage, state, location);
            }
        }

        MsgProjection.Apply(document, state, location, MapiStringEncodingContext.FromCodePage(codePage));
        MsgProjection.ApplyTransportHeaderRecipients(document, state, location);
        document.Format = EmailFileFormat.Tnef;
        document.Subject = subject ?? document.Subject;
        document.Body.Text = body ?? document.Body.Text;
        document.MessageClass = messageClass ?? document.MessageClass;
        document.OutlookItemKind = MsgProjection.Classify(document.MessageClass);
        MsgProjection.ApplyTyped(document);
        document.MessageId = string.IsNullOrWhiteSpace(messageId) ? document.MessageId : messageId!.Trim().Trim('<', '>');
        document.Date = date ?? document.Date;
        document.ReceivedDate = received ?? document.ReceivedDate;

        foreach (EmailAttachment attachment in document.Attachments) {
            state.ThrowIfCancellationRequested();
            ProjectAttachment(attachment, data, state, nestedDepth, location);
        }
        EmailProtectionProjection.Apply(document, state.Diagnostics, location);
        return document;
    }

    private static List<ParsedAttribute> ParseAttributes(byte[] data, MsgParserState state, string location) {
        var result = new List<ParsedAttribute>();
        int offset = 6;
        long completedAttachmentBytes = 0;
        long currentAttachmentBytes = 0;
        long pendingDecodedPropertyBytes = 0;
        while (offset < data.Length) {
            state.CountTnefAttribute();
            if (offset + 9 > data.Length) {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_ATTRIBUTE_TRUNCATED",
                    "A TNEF attribute header is truncated.", EmailDiagnosticSeverity.Error, location));
                break;
            }
            byte rawLevel = data[offset++];
            uint tag = MsgBinary.ReadUInt32(data, offset);
            offset += 4;
            uint rawLength = MsgBinary.ReadUInt32(data, offset);
            offset += 4;
            if (rawLength > int.MaxValue || offset > data.Length - (int)rawLength - 2) {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_ATTRIBUTE_LENGTH_INVALID",
                    "A TNEF attribute length exceeds the remaining input.", EmailDiagnosticSeverity.Error, location));
                break;
            }
            long attachmentMapiPayloadLength = 0;
            if (tag == TnefConstants.MessageProperties || tag == TnefConstants.RecipientTable ||
                tag == TnefConstants.AttachmentProperties) {
                if (!TnefMapiCodec.TryPreflightProperties(
                    data, offset, (int)rawLength, state, tag == TnefConstants.RecipientTable,
                    out long decodedPropertyBytes, out attachmentMapiPayloadLength)) {
                    state.Diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_MAPI_PREFLIGHT_INVALID",
                        string.Concat("TNEF MAPI attribute 0x", tag.ToString("X8", CultureInfo.InvariantCulture),
                            " is malformed and was rejected before buffering."),
                        EmailDiagnosticSeverity.Error, location));
                    break;
                }
                pendingDecodedPropertyBytes = checked(pendingDecodedPropertyBytes + decodedPropertyBytes);
                state.EnsureDecodedPropertyBytesWithinLimits(pendingDecodedPropertyBytes);
            }
            if (rawLevel == (byte)TnefAttributeLevel.Attachment) {
                if (tag == TnefConstants.AttachRendData) {
                    completedAttachmentBytes = checked(completedAttachmentBytes + currentAttachmentBytes);
                    currentAttachmentBytes = 0;
                }
                long candidateLength = tag == TnefConstants.AttachData
                    ? rawLength
                    : tag == TnefConstants.AttachmentProperties
                        ? attachmentMapiPayloadLength
                        : 0;
                if (candidateLength > currentAttachmentBytes) {
                    currentAttachmentBytes = candidateLength;
                    state.EnsureAttachmentBytesWithinLimits(currentAttachmentBytes, completedAttachmentBytes);
                }
            }
            byte[] bytes = MsgBinary.Slice(data, offset, (int)rawLength);
            offset += (int)rawLength;
            ushort storedChecksum = MsgBinary.ReadUInt16(data, offset);
            offset += 2;
            ushort actualChecksum = CalculateChecksum(bytes);
            bool validChecksum = storedChecksum == actualChecksum;
            if (!validChecksum) {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_CHECKSUM_MISMATCH",
                    string.Concat("Attribute 0x", tag.ToString("X8", CultureInfo.InvariantCulture), " has an invalid checksum."),
                    EmailDiagnosticSeverity.Warning, location));
            }
            TnefAttributeLevel level = rawLevel == 2 ? TnefAttributeLevel.Attachment : TnefAttributeLevel.Message;
            if (rawLevel != 1 && rawLevel != 2) {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_LEVEL_UNKNOWN",
                    string.Concat("Unknown TNEF attribute level ", rawLevel.ToString(CultureInfo.InvariantCulture), " was treated as message-level."),
                    EmailDiagnosticSeverity.Warning, location));
            }
            result.Add(new ParsedAttribute(level, tag, bytes, validChecksum));
        }
        return result;
    }

    private static void ApplyAttachmentAttribute(EmailAttachment attachment, ParsedAttribute attribute, int codePage,
        MsgParserState state, string location) {
        switch (attribute.Tag) {
            case TnefConstants.AttachRendData:
                if (attribute.Data.Length >= 12) {
                    ushort type = MsgBinary.ReadUInt16(attribute.Data, 0);
                    attachment.MapiAttachMethod = type == 2 ? 6 : 1;
                    attachment.IsInline = MsgBinary.ReadInt32(attribute.Data, 2) >= 0;
                }
                break;
            case TnefConstants.AttachTitle:
            case TnefConstants.AttachTransportFilename:
                attachment.FileName = DecodeString(attribute.Data, codePage, state, location);
                break;
            case TnefConstants.AttachData:
                attachment.Length = attribute.Data.LongLength;
                attachment.Content = state.Options.IncludeAttachmentContent ? (byte[])attribute.Data.Clone() : null;
                break;
            case TnefConstants.AttachmentProperties:
                AddProperties(attachment.MapiProperties,
                    TnefMapiCodec.ReadProperties(attribute.Data, codePage, state, string.Concat(location, "/attachment-mapi")));
                break;
        }
    }

    private static void ProjectAttachment(EmailAttachment attachment, byte[] source, MsgParserState state,
        int nestedDepth, string location) {
        attachment.FileName = MsgProjection.GetString(attachment.MapiProperties, 0x3707) ??
            MsgProjection.GetString(attachment.MapiProperties, 0x3704) ?? attachment.FileName;
        attachment.ContentType = MsgProjection.GetString(attachment.MapiProperties, 0x370E) ?? attachment.ContentType;
        attachment.ContentId = MsgProjection.GetString(attachment.MapiProperties, 0x3712)?.Trim().Trim('<', '>') ?? attachment.ContentId;
        attachment.ContentLocation = MsgProjection.GetString(attachment.MapiProperties, 0x3713) ?? attachment.ContentLocation;
        attachment.IsHidden = MsgProjection.GetBool(attachment.MapiProperties, 0x7FFE) ?? attachment.IsHidden;
        attachment.IsContactPhoto = MsgProjection.GetBool(attachment.MapiProperties, 0x7FFF) ?? attachment.IsContactPhoto;
        attachment.RenderingPosition = MsgProjection.GetInt(attachment.MapiProperties, 0x370B) ?? attachment.RenderingPosition;
        attachment.CreatedDate = MsgProjection.GetDate(attachment.MapiProperties, 0x3007) ?? attachment.CreatedDate;
        attachment.ModifiedDate = MsgProjection.GetDate(attachment.MapiProperties, 0x3008) ?? attachment.ModifiedDate;
        attachment.LinkedPath = MsgProjection.GetString(attachment.MapiProperties, 0x370D) ?? attachment.LinkedPath;
        attachment.IsInline = attachment.IsInline || !string.IsNullOrWhiteSpace(attachment.ContentId) ||
            ((MsgProjection.GetInt(attachment.MapiProperties, 0x3714) ?? 0) & 0x00000004) != 0;
        int method = MsgProjection.GetInt(attachment.MapiProperties, 0x3705) ?? attachment.MapiAttachMethod ?? 1;
        attachment.MapiAttachMethod = method;
        MapiProperty? dataProperty = attachment.MapiProperties.FirstOrDefault(property => property.PropertyId == 0x3701);
        byte[]? objectBytes = dataProperty?.Value as byte[];
        if (method == 5 && objectBytes != null && objectBytes.Length >= 20 && new Guid(MsgBinary.Slice(objectBytes, 0, 16)) == IidMessage) {
            int nestedLength = objectBytes.Length - 16;
            state.CountAttachment(nestedLength);
            byte[] nested = MsgBinary.Slice(objectBytes, 16, nestedLength);
            if (nestedDepth < state.Options.MaxNestedMessageDepth && nested.Length >= 4 && MsgBinary.ReadUInt32(nested, 0) == TnefConstants.Signature) {
                attachment.EmbeddedDocument = ReadMessage(nested, state, nestedDepth + 1, string.Concat(location, "/embedded"));
            } else {
                if (state.Options.IncludeAttachmentContent) attachment.Content = (byte[])objectBytes.Clone();
                bool depthLimited = nestedDepth >= state.Options.MaxNestedMessageDepth;
                state.Diagnostics.Add(new EmailDiagnostic(
                    depthLimited ? "EMAIL_TNEF_NESTED_MESSAGE_LIMIT" : "EMAIL_TNEF_EMBEDDED_MESSAGE_INVALID",
                    depthLimited
                        ? "The embedded TNEF message was retained as opaque content but not projected because the nested-message limit was reached."
                        : "The embedded TNEF message was retained as opaque content but could not be projected.",
                    EmailDiagnosticSeverity.Warning, location));
            }
        } else if (method == 6 && objectBytes != null && objectBytes.Length > 16 && new Guid(MsgBinary.Slice(objectBytes, 0, 16)) == IidStorage) {
            int compoundLength = objectBytes.Length - 16;
            state.EnsureAttachmentBytesWithinLimits(compoundLength);
            byte[] compoundBytes = MsgBinary.Slice(objectBytes, 16, compoundLength);
            OfficeCompoundFile? compound;
            string? compoundError;
            bool compoundRead;
            try {
                compoundRead = OfficeCompoundFileReader.TryRead(compoundBytes,
                    EmailCompoundReadPolicy.CreateForAttachment(state.Options, state.TotalAttachmentBytes),
                    out compound, out compoundError);
            } catch (OfficeCompoundStreamLimitExceededException exception) {
                throw new EmailLimitExceededException(
                    exception.LimitName, exception.ActualValue, exception.MaximumValue);
            }
            if (compoundRead && compound != null) {
                long total = 0;
                foreach (KeyValuePair<string, byte[]> stream in compound.Streams) {
                    state.ThrowIfCancellationRequested();
                    if (state.Options.IncludeAttachmentContent) {
                        attachment.StructuredStorageStreams[stream.Key] = stream.Value;
                    }
                    total = checked(total + stream.Value.LongLength);
                }
                attachment.Length = total;
                state.CountAttachment(total);
            } else {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_COMPOUND_ATTACHMENT_INVALID",
                    compoundError ?? "The TNEF compound attachment could not be read.",
                    EmailDiagnosticSeverity.Warning, location));
                state.CountAttachment(compoundBytes.LongLength);
                attachment.Content = state.Options.IncludeAttachmentContent ? compoundBytes : null;
            }
        } else {
            byte[]? mapiContent = dataProperty?.Value as byte[];
            if (mapiContent != null) {
                attachment.Length = mapiContent.LongLength;
                if (state.Options.IncludeAttachmentContent) {
                    attachment.Content = (byte[])mapiContent.Clone();
                }
            }
            state.CountAttachment(mapiContent?.LongLength ?? attachment.Content?.LongLength ?? attachment.Length);
        }

        if (!state.Options.IncludeAttachmentContent && dataProperty != null) {
            dataProperty.Value = null;
            dataProperty.RawData = null;
        }
    }

    private static void AddRecipients(EmailDocument document, IEnumerable<List<MapiProperty>> rows) {
        foreach (List<MapiProperty> properties in rows) {
            EmailAddress? address = MsgAddressProjection.ReadAddress(
                properties, 0x3001, 0x39FE, 0x3003, 0x3002, 0x403E);
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
    }

    private static void AddProperties(IList<MapiProperty> target, IEnumerable<MapiProperty> source) {
        foreach (MapiProperty property in source) target.Add(property);
    }

    private static int ReadCodePage(IEnumerable<ParsedAttribute> attributes) {
        ParsedAttribute? attribute = attributes.FirstOrDefault(item => item.Tag == TnefConstants.OemCodePage);
        return attribute != null && attribute.Data.Length >= 4 ? MsgBinary.ReadInt32(attribute.Data, 0) : 1252;
    }

    private static string DecodeString(byte[] bytes, int codePage, MsgParserState state, string location) {
        return MimeTextCodec.DecodeText(bytes, codePage, state.Diagnostics, location).TrimEnd('\0');
    }

    private static DateTimeOffset? DecodeDate(byte[] bytes, IList<EmailDiagnostic> diagnostics, string location) {
        if (bytes.Length < 14) return null;
        try {
            return new DateTimeOffset(MsgBinary.ReadUInt16(bytes, 0), MsgBinary.ReadUInt16(bytes, 2),
                MsgBinary.ReadUInt16(bytes, 4), MsgBinary.ReadUInt16(bytes, 6), MsgBinary.ReadUInt16(bytes, 8),
                MsgBinary.ReadUInt16(bytes, 10), TimeSpan.Zero);
        } catch (ArgumentOutOfRangeException ex) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_DATE_INVALID", ex.Message, EmailDiagnosticSeverity.Warning, location));
            return null;
        }
    }

    private static ushort CalculateChecksum(byte[] bytes) {
        uint checksum = 0;
        for (int index = 0; index < bytes.Length; index++) checksum += bytes[index];
        return unchecked((ushort)checksum);
    }

    private sealed class ParsedAttribute {
        internal ParsedAttribute(TnefAttributeLevel level, uint tag, byte[] data, bool checksumIsValid) {
            Level = level; Tag = tag; Data = data; ChecksumIsValid = checksumIsValid;
        }
        internal TnefAttributeLevel Level { get; }
        internal uint Tag { get; }
        internal byte[] Data { get; }
        internal bool ChecksumIsValid { get; }
    }
}
