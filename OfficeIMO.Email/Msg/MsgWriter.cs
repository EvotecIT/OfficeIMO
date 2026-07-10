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
        MsgPropertyBuilder messageProperties = CreateMessageProperties(document, diagnostics, prefix);
        MsgPropertyWriter.Write(prefix, kind, messageProperties.Properties, document.Recipients.Count,
            document.Attachments.Count, names, streams, diagnostics);

        for (int index = 0; index < document.Recipients.Count; index++) {
            EmailRecipient recipient = document.Recipients[index];
            string storage = MsgBinary.CombinePath(prefix,
                string.Concat("__recip_version1.0_#", index.ToString("X8", CultureInfo.InvariantCulture)));
            MsgPropertyBuilder properties = CreateRecipientProperties(recipient);
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
        string messageClass = document.MessageClass ?? DefaultMessageClass(document.OutlookItemKind);
        properties.Set(0x001A, MapiPropertyType.Unicode, messageClass);
        properties.Set(0x340D, MapiPropertyType.Integer32, 0x00040000);
        properties.Set(0x0E07, MapiPropertyType.Integer32, document.Attachments.Count > 0 ? 0x10 : 0);
        properties.Set(0x0E1B, MapiPropertyType.Boolean, document.Attachments.Count > 0);
        properties.Set(0x0037, MapiPropertyType.Unicode, document.Subject);
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
        properties.Set(0x0042, MapiPropertyType.Unicode, document.From?.DisplayName);
        properties.Set(0x0065, MapiPropertyType.Unicode, document.From?.Address);
        properties.Set(0x5D02, MapiPropertyType.Unicode, document.From?.Address);
        properties.Set(0x0C1A, MapiPropertyType.Unicode, document.Sender?.DisplayName);
        properties.Set(0x0C1F, MapiPropertyType.Unicode, document.Sender?.Address);
        properties.Set(0x5D01, MapiPropertyType.Unicode, document.Sender?.Address);
        properties.Set(0x0E04, MapiPropertyType.Unicode, JoinRecipients(document, EmailRecipientKind.To));
        properties.Set(0x0E03, MapiPropertyType.Unicode, JoinRecipients(document, EmailRecipientKind.Cc));
        properties.Set(0x0E02, MapiPropertyType.Unicode, JoinRecipients(document, EmailRecipientKind.Bcc));
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

    internal static MsgPropertyBuilder CreateRecipientProperties(EmailRecipient recipient) {
        var properties = new MsgPropertyBuilder(recipient.MapiProperties);
        int type = recipient.Kind == EmailRecipientKind.To ? 1 : recipient.Kind == EmailRecipientKind.Cc ? 2 :
            recipient.Kind == EmailRecipientKind.Bcc ? 3 : 0;
        properties.Set(0x0C15, MapiPropertyType.Integer32, type);
        properties.Set(0x3001, MapiPropertyType.Unicode, recipient.Address.DisplayName ?? recipient.Address.Address);
        properties.Set(0x3002, MapiPropertyType.Unicode, "SMTP");
        properties.Set(0x3003, MapiPropertyType.Unicode, recipient.Address.Address);
        properties.Set(0x39FE, MapiPropertyType.Unicode, recipient.Address.Address);
        return properties;
    }

    internal static MsgPropertyBuilder CreateAttachmentProperties(EmailAttachment attachment, int index, int method,
        IList<EmailDiagnostic> diagnostics, string location) {
        var properties = new MsgPropertyBuilder(attachment.MapiProperties);
        properties.Set(0x3705, MapiPropertyType.Integer32, method);
        properties.Set(0x0E21, MapiPropertyType.Integer32, index);
        properties.Set(0x3707, MapiPropertyType.Unicode, attachment.FileName);
        properties.Set(0x3704, MapiPropertyType.Unicode, attachment.FileName);
        properties.Set(0x3001, MapiPropertyType.Unicode, attachment.FileName);
        properties.Set(0x370E, MapiPropertyType.Unicode, attachment.ContentType);
        properties.Set(0x3712, MapiPropertyType.Unicode, attachment.ContentId);
        properties.Set(0x3713, MapiPropertyType.Unicode, attachment.ContentLocation);
        properties.Set(0x370B, MapiPropertyType.Integer32, attachment.IsInline ? 0 : -1);
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
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8232, MapiPropertyType.Unicode, item.RecurrencePattern);
            properties.SetNamed(MsgProjection.PsetidAppointment, 0x8216, MapiPropertyType.Binary, item.RecurrenceState);
        }
        if (document.Contact != null) {
            OutlookContact item = document.Contact;
            properties.Set(0x3A06, MapiPropertyType.Unicode, item.GivenName);
            properties.Set(0x3A11, MapiPropertyType.Unicode, item.Surname);
            properties.Set(0x3A16, MapiPropertyType.Unicode, item.CompanyName);
            properties.Set(0x3A17, MapiPropertyType.Unicode, item.JobTitle);
            properties.Set(0x3A08, MapiPropertyType.Unicode, item.BusinessPhone);
            properties.Set(0x3A09, MapiPropertyType.Unicode, item.HomePhone);
            properties.Set(0x3A1C, MapiPropertyType.Unicode, item.MobilePhone);
            properties.SetNamed(MsgProjection.PsetidAddress, 0x8005, MapiPropertyType.Unicode, item.FileAs);
            properties.SetNamed(MsgProjection.PsetidAddress, 0x8084, MapiPropertyType.Unicode, item.Email1Address);
        }
        if (document.Task != null) {
            OutlookTask item = document.Task;
            properties.SetNamed(MsgProjection.PsetidTask, 0x8104, MapiPropertyType.Time, item.Start);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8105, MapiPropertyType.Time, item.Due);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8101, MapiPropertyType.Integer32, item.Status);
            properties.SetNamed(MsgProjection.PsetidTask, 0x8102, MapiPropertyType.Floating64, item.PercentComplete);
            properties.SetNamed(MsgProjection.PsetidTask, 0x810F, MapiPropertyType.Boolean, item.IsComplete);
            properties.SetNamed(MsgProjection.PsetidTask, 0x811C, MapiPropertyType.Unicode, item.Owner);
        }
        if (document.Journal != null) {
            OutlookJournal item = document.Journal;
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8516, MapiPropertyType.Time, item.Start);
            properties.SetNamed(MsgProjection.PsetidCommon, 0x8517, MapiPropertyType.Time, item.End);
            properties.SetNamed(MsgProjection.PsetidLog, 0x8700, MapiPropertyType.Unicode, item.Type);
        }
        if (document.Note != null) {
            OutlookNote item = document.Note;
            properties.SetNamed(MsgProjection.PsetidNote, 0x8B00, MapiPropertyType.Integer32, item.Color);
            properties.SetNamed(MsgProjection.PsetidNote, 0x8B02, MapiPropertyType.Integer32, item.Width);
            properties.SetNamed(MsgProjection.PsetidNote, 0x8B03, MapiPropertyType.Integer32, item.Height);
        }
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

    private static string? JoinRecipients(EmailDocument document, EmailRecipientKind kind) {
        string[] values = document.Recipients.Where(recipient => recipient.Kind == kind)
            .Select(recipient => recipient.Address.ToString()).Where(value => value.Length > 0).ToArray();
        return values.Length == 0 ? null : string.Join("; ", values);
    }
}
