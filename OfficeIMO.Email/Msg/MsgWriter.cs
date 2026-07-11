using OfficeIMO.Shared;

namespace OfficeIMO.Email;

internal static class MsgWriter {
    internal static byte[] Write(EmailDocument document, EmailWriterOptions options, IList<EmailDiagnostic> diagnostics) {
        var streams = new List<OfficeCompoundStream>();
        var names = new MsgNamedPropertyWriter();
        BuildMessage(document, string.Empty, MsgPropertyStreamKind.TopLevel, names, streams, diagnostics, options, 0);
        names.WriteStreams(streams);
        return OfficeCompoundFileWriter.Write(streams);
    }

    private static void BuildMessage(EmailDocument document, string prefix, MsgPropertyStreamKind kind,
        MsgNamedPropertyWriter names, IList<OfficeCompoundStream> streams, IList<EmailDiagnostic> diagnostics,
        EmailWriterOptions options, int depth) {
        if (depth > options.MaxNestedMessageDepth) throw new InvalidOperationException("The embedded-message write depth exceeds the configured maximum.");
        EmailRecipient[] storageRecipients = document.Recipients
            .Where(recipient => recipient.Kind != EmailRecipientKind.ReplyTo)
            .ToArray();
        MsgPropertyBuilder messageProperties = CreateMessageProperties(document, diagnostics, prefix);
        MsgPropertyWriter.Write(prefix, kind, messageProperties.Properties, storageRecipients.Length,
            document.Attachments.Count, names, streams, diagnostics);

        for (int index = 0; index < storageRecipients.Length; index++) {
            EmailRecipient recipient = storageRecipients[index];
            string storage = MsgBinary.CombinePath(prefix,
                string.Concat("__recip_version1.0_#", index.ToString("X8", CultureInfo.InvariantCulture)));
            MsgPropertyBuilder properties = CreateRecipientProperties(recipient, index);
            MsgPropertyWriter.Write(storage, MsgPropertyStreamKind.ChildObject, properties.Properties,
                0, 0, names, streams, diagnostics);
        }

        for (int index = 0; index < document.Attachments.Count; index++) {
            EmailAttachment attachment = document.Attachments[index];
            string storage = MsgBinary.CombinePath(prefix,
                string.Concat("__attach_version1.0_#", index.ToString("X8", CultureInfo.InvariantCulture)));
            int method = attachment.MapiAttachMethod ?? (attachment.EmbeddedDocument != null ? 5 :
                attachment.StructuredStorageStreams.Count > 0 ? 6 : 1);
            MsgPropertyBuilder properties = CreateAttachmentProperties(attachment, index, method, diagnostics, storage);
            MsgPropertyWriter.Write(storage, MsgPropertyStreamKind.ChildObject, properties.Properties,
                0, 0, names, streams, diagnostics, method == 5 ? 1U : method == 6 ? 4U : 0U);

            string objectStorage = MsgBinary.CombinePath(storage, "__substg1.0_3701000D");
            if (method == 5 && attachment.EmbeddedDocument != null) {
                BuildMessage(attachment.EmbeddedDocument, objectStorage, MsgPropertyStreamKind.EmbeddedMessage,
                    names, streams, diagnostics, options, depth + 1);
            } else if (method == 6) {
                foreach (KeyValuePair<string, byte[]> stream in attachment.StructuredStorageStreams
                    .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase)) {
                    streams.Add(new OfficeCompoundStream(MsgBinary.CombinePath(objectStorage, stream.Key), stream.Value));
                }
            }
        }
    }

    internal static MsgPropertyBuilder CreateMessageProperties(EmailDocument document,
        IList<EmailDiagnostic> diagnostics, string location) {
        var properties = new MsgPropertyBuilder(document.MapiProperties);
        EmailMessageMetadata metadata = document.MessageMetadata;
        string messageClass = document.MessageClass ?? DefaultMessageClass(document.OutlookItemKind);
        properties.Set(0x001A, MapiPropertyType.Unicode, messageClass);
        properties.Set(0x0FFF, MapiPropertyType.Binary,
            MsgIdentity.CreateStableBytes("message-entry", 16, document.MessageId, document.Subject));
        properties.Set(0x0FF6, MapiPropertyType.Binary,
            MsgIdentity.CreateStableBytes("message-instance", 4, document.MessageId, document.Subject));
        properties.Set(0x0FF9, MapiPropertyType.Binary,
            MsgIdentity.CreateStableBytes("message-record", 16, document.MessageId, document.Subject));
        properties.Set(0x340D, MapiPropertyType.Integer32, 0x00040000);
        properties.Set(0x340F, MapiPropertyType.Integer32, 0x00040000);
        properties.Set(0x0002, MapiPropertyType.Boolean, true);
        properties.Set(0x0FF4, MapiPropertyType.Integer32, 0x00000007);
        properties.Set(0x0FF7, MapiPropertyType.Integer32, 0x00000001);
        properties.Set(0x0FFE, MapiPropertyType.Integer32, 5);
        int messageFlags = 0x0002;
        if (document.Attachments.Count > 0) messageFlags |= 0x0010;
        if (metadata.IsDraft) messageFlags |= 0x0008;
        if (metadata.IsRead == true) messageFlags |= 0x0001 | 0x0400;
        properties.Set(0x0E07, MapiPropertyType.Integer32, messageFlags);
        properties.Set(0x0E1B, MapiPropertyType.Boolean, document.Attachments.Count > 0);
        properties.Set(0x0037, MapiPropertyType.Unicode, document.Subject);
        ResolveSubject(document.Subject, metadata, out string subjectPrefix, out string normalizedSubject);
        properties.Set(0x003D, MapiPropertyType.Unicode, subjectPrefix);
        properties.Set(0x0E1D, MapiPropertyType.Unicode, normalizedSubject);
        properties.Set(0x0070, MapiPropertyType.Unicode, metadata.ConversationTopic ?? normalizedSubject);
        properties.Set(0x0071, MapiPropertyType.Binary, metadata.ConversationIndex);
        properties.Set(0x1000, MapiPropertyType.Unicode, document.Body.Text);
        if (document.Body.Html != null) {
            properties.Set(0x1013, MapiPropertyType.Binary, Encoding.UTF8.GetBytes(document.Body.Html));
            properties.Set(0x1016, MapiPropertyType.Integer32, 3);
        } else if (document.Body.Rtf != null) {
            properties.Set(0x1016, MapiPropertyType.Integer32, 2);
        } else if (document.Body.Text != null) {
            properties.Set(0x1016, MapiPropertyType.Integer32, 1);
        }
        if (document.Body.Rtf != null && TryGetBytePreservingRtf(document.Body.Rtf, diagnostics, location,
            out byte[] rtfBytes)) {
            properties.Set(0x1009, MapiPropertyType.Binary, MapiCompressedRtfCodec.Compress(rtfBytes));
            properties.Set(0x0E1F, MapiPropertyType.Boolean, true);
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
            metadata.IconIndex ?? (metadata.IsDraft ? 0x00000103 : metadata.IsRead == true ? 0x00000100 : 0x00000101));
        properties.Set(0x0029, MapiPropertyType.Boolean, metadata.ReadReceiptRequested);
        properties.Set(0x3007, MapiPropertyType.Time, metadata.CreatedDate ?? document.Date);
        properties.Set(0x3008, MapiPropertyType.Time, metadata.ModifiedDate ?? metadata.CreatedDate ?? document.Date);
        properties.Set(0x3FDE, MapiPropertyType.Integer32, document.OutlookCodePage ?? 65001);
        properties.Set(0x3FFD, MapiPropertyType.Integer32, document.OutlookCodePage ?? 65001);
        properties.Set(0x3FF1, MapiPropertyType.Integer32, 1033);
        properties.Set(0x0042, MapiPropertyType.Unicode, document.From?.DisplayName);
        properties.Set(0x0065, MapiPropertyType.Unicode, document.From?.Address);
        properties.Set(0x0064, MapiPropertyType.Unicode, document.From?.AddressType ?? "SMTP");
        properties.Set(0x5D02, MapiPropertyType.Unicode, document.From?.Address);
        EmailAddress? sender = document.Sender ?? document.From;
        properties.Set(0x0C1A, MapiPropertyType.Unicode, sender?.DisplayName);
        properties.Set(0x0C1F, MapiPropertyType.Unicode, sender?.Address);
        properties.Set(0x0C1E, MapiPropertyType.Unicode, sender?.AddressType ?? "SMTP");
        properties.Set(0x5D01, MapiPropertyType.Unicode, sender?.Address);
        if (sender != null) properties.Set(0x0C19, MapiPropertyType.Binary, MsgIdentity.CreateOneOffEntryId(sender));
        properties.Set(0x0E04, MapiPropertyType.Unicode, JoinRecipients(document, EmailRecipientKind.To));
        properties.Set(0x0E03, MapiPropertyType.Unicode, JoinRecipients(document, EmailRecipientKind.Cc));
        properties.Set(0x0E02, MapiPropertyType.Unicode, JoinRecipients(document, EmailRecipientKind.Bcc));
        properties.Set(0x0050, MapiPropertyType.Unicode, JoinRecipients(document, EmailRecipientKind.ReplyTo));
        if (metadata.Categories.Count > 0) {
            properties.SetNamed(MsgProjection.PsPublicStrings, "Keywords", MapiPropertyType.MultipleUnicode,
                metadata.Categories.Cast<object>().ToArray());
        }
        if (document.Headers.Count > 0) {
            string headers = string.Join("\r\n", document.Headers.Select(header =>
                string.Concat(header.Name, ": ", header.RawValue ?? header.Value)));
            properties.Set(0x007D, MapiPropertyType.Unicode, headers);
        }
        AddTypedProperties(properties, document);
        return properties;
    }

    private static bool TryGetBytePreservingRtf(string rtf, IList<EmailDiagnostic> diagnostics,
        string location, out byte[] bytes) {
        bytes = new byte[rtf.Length];
        for (int index = 0; index < rtf.Length; index++) {
            if (rtf[index] > byte.MaxValue) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_RTF_CHARACTER_UNENCODABLE",
                    "The RTF source contains a character above U+00FF. Serialize it through OfficeIMO.Rtf so the character is represented by an RTF escape.",
                    EmailDiagnosticSeverity.Error, location));
                bytes = Array.Empty<byte>();
                return false;
            }
            bytes[index] = unchecked((byte)rtf[index]);
        }
        return true;
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
        properties.Set(0x0FF6, MapiPropertyType.Binary,
            MsgIdentity.CreateStableBytes("recipient-instance", 4, index.ToString(CultureInfo.InvariantCulture), address));
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
        IList<EmailDiagnostic> diagnostics, string location) {
        var properties = new MsgPropertyBuilder(attachment.MapiProperties);
        properties.Set(0x0FF6, MapiPropertyType.Binary,
            MsgIdentity.CreateStableBytes("attachment-instance", 4, index.ToString(CultureInfo.InvariantCulture), attachment.FileName));
        properties.Set(0x0FF9, MapiPropertyType.Binary,
            MsgIdentity.CreateStableBytes("attachment-record", 16, index.ToString(CultureInfo.InvariantCulture), attachment.FileName));
        properties.Set(0x0FFE, MapiPropertyType.Integer32, 7);
        properties.Set(0x3705, MapiPropertyType.Integer32, method);
        properties.Set(0x0E21, MapiPropertyType.Integer32, index);
        properties.Set(0x3707, MapiPropertyType.Unicode, attachment.FileName);
        properties.Set(0x3704, MapiPropertyType.Unicode, attachment.FileName);
        properties.Set(0x3001, MapiPropertyType.Unicode, attachment.FileName);
        properties.Set(0x370E, MapiPropertyType.Unicode, attachment.ContentType);
        properties.Set(0x3712, MapiPropertyType.Unicode, attachment.ContentId);
        properties.Set(0x3713, MapiPropertyType.Unicode, attachment.ContentLocation);
        properties.Set(0x3703, MapiPropertyType.Unicode, Path.GetExtension(attachment.FileName));
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
        } else if (attachment.Content != null) {
            properties.Set(0x3701, MapiPropertyType.Binary, attachment.Content);
            properties.Set(0x0E20, MapiPropertyType.Integer32, attachment.Content.Length);
        } else {
            properties.Set(0x0E20, MapiPropertyType.Integer32, checked((int)Math.Min(attachment.Length, int.MaxValue)));
            if (attachment.Length > 0) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE",
                    "An MSG attachment has a declared length but no retained content.",
                    EmailDiagnosticSeverity.Error, location));
            }
        }
        return properties;
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
        return checked((int)Math.Round((end.Value - start.Value).TotalMinutes));
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
        return checked((int)Math.Round(value.Value.TotalMinutes));
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
        out string prefix, out string normalized) {
        prefix = metadata.SubjectPrefix ?? string.Empty;
        normalized = metadata.NormalizedSubject ?? subject ?? string.Empty;
        if (metadata.NormalizedSubject != null || string.IsNullOrEmpty(subject)) return;

        int colon = subject!.IndexOf(':');
        if (colon > 0 && colon <= 3 && colon + 1 < subject.Length) {
            prefix = subject.Substring(0, colon + 1);
            if (colon + 1 < subject.Length && subject[colon + 1] == ' ') prefix += " ";
            normalized = subject.Substring(prefix.Length);
        }
    }
}
