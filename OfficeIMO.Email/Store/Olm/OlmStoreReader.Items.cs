using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed partial class OlmStoreReader {
    private EmailDocument ProjectItem(XElement item, OutlookItemKind kind, string id,
        string folderId, string location) {
        var document = new EmailDocument {
            Format = EmailFileFormat.Unknown,
            OutlookItemKind = kind,
            MessageClass = GetMessageClass(kind)
        };
        document.Properties["EmailStore:Format"] = EmailStoreFormat.Olm.ToString();
        document.Properties["EmailStore:ItemId"] = id;
        document.Properties["EmailStore:FolderId"] = folderId;
        document.Properties["Olm:EntryPath"] = location.Substring(0, location.LastIndexOf('#'));
        document.Properties["Olm:ElementName"] = item.Name.LocalName;
        PreserveScalarProperties(document, item);
        ApplyCategories(document, item);

        switch (kind) {
            case OutlookItemKind.Message:
                ProjectMessage(document, item, location);
                break;
            case OutlookItemKind.Appointment:
                ProjectAppointment(document, item, location);
                break;
            case OutlookItemKind.Contact:
                ProjectContact(document, item);
                break;
            case OutlookItemKind.Task:
                ProjectTask(document, item);
                break;
            case OutlookItemKind.Note:
                ProjectNote(document, item);
                break;
        }
        return document;
    }

    private void ProjectMessage(EmailDocument document, XElement item, string location) {
        document.Subject = Value(item, "OPFMessageCopySubject");
        document.MessageId = Value(item, "OPFMessageCopyMessageID");
        document.Body.Text = Value(item, "OPFMessageCopyBody");
        document.Body.Html = Value(item, "OPFMessageCopyHTMLBody");
        document.Date = DateValue(item, "OPFMessageCopySentTime");
        document.ReceivedDate = DateValue(item, "OPFMessageCopyReceivedTime");
        document.MessageMetadata.ModifiedDate = DateValue(item, "OPFMessageCopyModDate");
        document.MessageMetadata.ConversationTopic = Value(item, "OPFMessageCopyThreadTopic");
        document.MessageMetadata.InReplyToId = Value(item, "OPFMessageCopyInReplyTo");
        document.MessageMetadata.InternetReferences = Value(item, "OPFMessageCopyReferences");
        document.MessageMetadata.IsRead = BooleanValue(item, "OPFMessageGetIsRead");
        document.MessageMetadata.Priority = PriorityValue(item, "OPFMessageGetPriority");
        document.MessageMetadata.IsDraft = location.IndexOf("/Drafts/", StringComparison.OrdinalIgnoreCase) >= 0;

        EmailAddress? from = FirstAddress(item, "OPFMessageCopyFromAddresses");
        EmailAddress? sender = FirstAddress(item, "OPFMessageCopySenderAddress");
        document.From = from ?? sender;
        document.Sender = sender;
        string? receivedRepresenting = Value(item, "OPFMessageCopyReceivedRepresentingName");
        if (!string.IsNullOrWhiteSpace(receivedRepresenting)) {
            document.ReceivedRepresenting = new EmailAddress(null, receivedRepresenting);
        }

        AddRecipients(document, item, "OPFMessageCopyToAddresses", EmailRecipientKind.To);
        AddRecipients(document, item, "OPFMessageCopyCCAddresses", EmailRecipientKind.Cc);
        AddRecipients(document, item, "OPFMessageCopyBCCAddresses", EmailRecipientKind.Bcc);
        AddRecipients(document, item, "OPFMessageCopyReplyToAddresses", EmailRecipientKind.ReplyTo);
        AddAttachments(document, item, "OPFMessageCopyAttachmentList", location);

        string? meetingData = Value(item, "OPFMessageCopyMeetingData");
        if (meetingData != null && meetingData.Trim().Length > 0 &&
            meetingData.EndsWith(".ics", StringComparison.OrdinalIgnoreCase)) {
            AddAttachment(document, meetingData, "meeting.ics", "text/calendar", null, location);
        }
    }

    private void ProjectAppointment(EmailDocument document, XElement item, string location) {
        document.Subject = Value(item, "OPFCalendarEventCopySummary");
        document.MessageId = Value(item, "OPFCalendarEventCopyUUID");
        document.Body.Text = Value(item, "OPFCalendarEventCopyDescriptionPlain");
        document.Body.Html = Value(item, "OPFCalendarEventCopyDescription");
        document.Date = DateValue(item, "OPFCalendarEventCopyStartTime");
        document.MessageMetadata.ModifiedDate = DateValue(item, "OPFCalendarEventCopyModDate");
        document.Appointment = new OutlookAppointment {
            Start = DateValue(item, "OPFCalendarEventCopyStartTime"),
            End = DateValue(item, "OPFCalendarEventCopyEndTime"),
            Location = Value(item, "OPFCalendarEventCopyLocation"),
            IsAllDay = BooleanValue(item, "OPFCalendarEventGetIsAllDayEvent"),
            BusyStatus = IntegerValue(item, "OPFCalendarEventCopyFreeBusyStatus"),
            ResponseStatus = IntegerValue(item, "OPFCalendarEventGetAcceptStatus"),
            IsRecurring = BooleanValue(item, "OPFCalendarEventIsRecurring"),
            ReminderIsSet = BooleanValue(item, "OPFCalendarEventGetHasReminder"),
            ReminderDeltaMinutes = IntegerValue(item, "OPFCalendarEventCopyReminderDelta"),
            ReminderTime = DateValue(item, "OPFCalendarEventCopyReminderTime"),
            TimeZoneDescription = Value(item, "OPFCalendarEventCopyStartTimeZone")
        };
        if (document.Appointment.Start.HasValue && document.Appointment.End.HasValue) {
            double minutes = (document.Appointment.End.Value - document.Appointment.Start.Value).TotalMinutes;
            if (minutes >= 0 && minutes <= int.MaxValue) document.Appointment.DurationMinutes = (int)minutes;
        }

        AddAppointmentRecipients(document, item);
        string? organizer = Value(item, "OPFCalendarEventCopyOrganizer");
        if (organizer != null) {
            document.From = organizer.IndexOf('@') >= 0
                ? new EmailAddress(organizer, rawValue: organizer)
                : new EmailAddress(null, organizer, organizer);
        }
        AddAttachments(document, item, "OPFCalendarEventCopyAttachmentList", location);
    }

    private static void ProjectContact(EmailDocument document, XElement item) {
        var contact = new OutlookContact {
            DisplayName = Value(item, "OPFContactCopyDisplayName"),
            Prefix = Value(item, "OPFContactCopyTitle"),
            GivenName = Value(item, "OPFContactCopyFirstName"),
            MiddleName = Value(item, "OPFContactCopyMiddleName"),
            Surname = Value(item, "OPFContactCopyLastName"),
            Generation = Value(item, "OPFContactCopySetNameSuffix"),
            CompanyName = Value(item, "OPFContactCopyBusinessCompany"),
            JobTitle = Value(item, "OPFContactCopyBusinessTitle"),
            Department = Value(item, "OPFContactCopyBusinessDepartment"),
            NickName = Value(item, "OPFContactCopyNickName"),
            SpouseName = Value(item, "OPFContactCopySpousesName"),
            OfficeLocation = Value(item, "OPFContactCopyBusinessOffice"),
            Birthday = DateValue(item, "OPFContactCopyBirthday"),
            WeddingAnniversary = DateValue(item, "OPFContactCopyAnniversary"),
            HasPicture = BooleanValue(item, "OPFContactCopyContactImage"),
            BusinessHomePage = Value(item, "OPFContactCopyBusinessHomePage"),
            PersonalHomePage = Value(item, "OPFContactCopyHomeWebPage"),
            InstantMessagingAddress = Value(item, "OPFContactCopyDefaultIMAddress")
        };
        document.Contact = contact;
        document.Subject = contact.DisplayName;
        document.Body.Text = Value(item, "OPFContactCopyNotesPlain");
        document.Body.Html = Value(item, "OPFContactCopyNotes");
        document.MessageMetadata.ModifiedDate = DateValue(item, "OPFContactCopyModDate");

        contact.BusinessAddress.Street = Value(item, "OPFContactCopyBusinessStreetAddress");
        contact.BusinessAddress.City = Value(item, "OPFContactCopyBusinessCity");
        contact.BusinessAddress.StateOrProvince = Value(item, "OPFContactCopyBusinessState");
        contact.BusinessAddress.PostalCode = Value(item, "OPFContactCopyBusinessZip");
        contact.BusinessAddress.Country = Value(item, "OPFContactCopyBusinessCountry");
        contact.HomeAddress.Street = Value(item, "OPFContactCopyHomeStreetAddress");
        contact.HomeAddress.City = Value(item, "OPFContactCopyHomeCity");
        contact.HomeAddress.StateOrProvince = Value(item, "OPFContactCopyHomeState");
        contact.HomeAddress.PostalCode = Value(item, "OPFContactCopyHomeZip");
        contact.HomeAddress.Country = Value(item, "OPFContactCopyHomeCountry");
        contact.OtherAddress.Street = Value(item, "OPFContactCopyOtherStreetAddress");
        contact.OtherAddress.City = Value(item, "OPFContactCopyOtherCity");
        contact.OtherAddress.StateOrProvince = Value(item, "OPFContactCopyOtherState");
        contact.OtherAddress.PostalCode = Value(item, "OPFContactCopyOtherZip");
        contact.OtherAddress.Country = Value(item, "OPFContactCopyOtherCountry");

        contact.Phones.Business = Value(item, "OPFContactCopyBusinessPhone");
        contact.Phones.Business2 = Value(item, "OPFContactCopyBusinessPhone2");
        contact.Phones.Home = Value(item, "OPFContactCopyHomePhone");
        contact.Phones.Home2 = Value(item, "OPFContactCopyHomePhone2");
        contact.Phones.Mobile = Value(item, "OPFContactCopyCellPhone");
        contact.Phones.Other = Value(item, "OPFContactCopyOtherPhone");
        contact.Phones.BusinessFax = Value(item, "OPFContactCopyBusinessFax");
        contact.Phones.HomeFax = Value(item, "OPFContactCopyHomeFax");
        contact.Phones.Other = contact.Phones.Other ?? Value(item, "OPFContactCopyRadioPhone");
        contact.Phones.Pager = Value(item, "OPFContactCopyPager");
        contact.Phones.Assistant = Value(item, "OPFContactCopyAssistantPhone");
        contact.Phones.CompanyMain = Value(item, "OPFContactMainPhone");

        IReadOnlyList<EmailAddress> addresses = Addresses(item, "OPFContactCopyEmailAddressList")
            .Concat(Addresses(item, "OPFContactCopyEmailAddressList1"))
            .Concat(Addresses(item, "OPFContactCopyEmailAddressList2"))
            .GroupBy(address => string.Concat(address.Address, "\u001f", address.DisplayName), StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .Take(3)
            .ToList();
        for (int index = 0; index < addresses.Count; index++) {
            OutlookContactEmailAddress target = index == 0 ? contact.Email1 : index == 1 ? contact.Email2 : contact.Email3;
            target.Address = addresses[index].Address;
            target.DisplayName = addresses[index].DisplayName;
            target.AddressType = addresses[index].AddressType;
        }
    }

    private static void ProjectTask(EmailDocument document, XElement item) {
        document.Subject = Value(item, "OPFTaskCopyName");
        document.Body.Text = Value(item, "OPFTaskCopyNotePlain");
        document.Body.Html = Value(item, "OPFTaskCopyNote");
        document.Date = DateValue(item, "OPFTaskCopyStartDateTime");
        document.MessageMetadata.ModifiedDate = DateValue(item, "OPFTaskCopyModDate");
        document.MessageMetadata.Priority = PriorityValue(item, "OPFTaskGetPriority");
        document.Task = new OutlookTask {
            Start = DateValue(item, "OPFTaskCopyStartDateTime"),
            Due = DateValue(item, "OPFTaskCopyDueDateTime"),
            CompletedAt = DateValue(item, "OPFTaskCopyCompletedDateTime"),
            ReminderTime = DateValue(item, "OPFTaskCopyReminderTime"),
            ReminderIsSet = DateValue(item, "OPFTaskCopyReminderTime").HasValue
        };
    }

    private static void ProjectNote(EmailDocument document, XElement item) {
        document.Subject = Value(item, "OPFNoteCopyTitle");
        document.Body.Text = Value(item, "OPFNoteCopyText");
        document.Date = DateValue(item, "OPFNoteCopyCreationDate");
        document.MessageMetadata.CreatedDate = DateValue(item, "OPFNoteCopyCreationDate");
        document.MessageMetadata.ModifiedDate = DateValue(item, "OPFNoteCopyModDate");
        document.Note = new OutlookNote();
    }

    private void AddAttachments(EmailDocument document, XElement item, string containerName, string location) {
        XElement? container = Child(item, containerName);
        if (container == null) return;
        foreach (XElement attachment in container.Descendants().Where(element =>
                     Attribute(element, "OPFAttachmentURL") != null)) {
            string? path = Attribute(attachment, "OPFAttachmentURL");
            string? name = Attribute(attachment, "OPFAttachmentName");
            string? type = Attribute(attachment, "OPFAttachmentContentType");
            string? contentId = Attribute(attachment, "OPFAttachmentContentID");
            AddAttachment(document, path, name, type, contentId, location);
        }
    }

    private void AddAttachment(EmailDocument document, string? sourcePath, string? fileName,
        string? contentType, string? contentId, string location) {
        if (document.Attachments.Count >= _options.MaxAttachmentsPerItem) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxAttachmentsPerItem),
                document.Attachments.Count + 1L, _options.MaxAttachmentsPerItem);
        }

        var attachment = new EmailAttachment {
            FileName = fileName,
            ContentType = contentType,
            ContentId = TrimContentId(contentId),
            ContentLocation = sourcePath,
            IsInline = !string.IsNullOrWhiteSpace(contentId),
            IsMimeRelated = !string.IsNullOrWhiteSpace(contentId)
        };
        document.Attachments.Add(attachment);
        if (sourcePath == null || sourcePath.Trim().Length == 0) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_OLM_ATTACHMENT_PATH_MISSING",
                "An attachment has no archive entry path.",
                EmailStoreDiagnosticSeverity.Warning,
                location));
            return;
        }
        if (!TryNormalizeArchivePath(sourcePath, out string normalized)) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_OLM_ATTACHMENT_PATH_UNSAFE",
                "An attachment path was rejected because it is not archive-root relative.",
                EmailStoreDiagnosticSeverity.Warning,
                location));
            return;
        }
        if (!_entries.TryGetValue(normalized, out ZipArchiveEntry? entry)) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_OLM_ATTACHMENT_MISSING",
                "The attachment entry referenced by the item is absent from the archive.",
                EmailStoreDiagnosticSeverity.Warning,
                normalized));
            return;
        }

        attachment.Length = entry.Length;
        if (attachment.Length > _options.MaxAttachmentBytes) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxAttachmentBytes),
                attachment.Length, _options.MaxAttachmentBytes);
        }
        _totalAttachmentBytes = AddBounded(_totalAttachmentBytes, attachment.Length,
            nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes), _options.MaxTotalAttachmentBytes);
        if (_options.RetainAttachmentContent) {
            byte[] content = ReadEntryBytes(entry);
            if (content.LongLength > attachment.Length) {
                _totalAttachmentBytes = AddBounded(
                    _totalAttachmentBytes, content.LongLength - attachment.Length,
                    nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes),
                    _options.MaxTotalAttachmentBytes);
            }
            attachment.Length = content.LongLength;
            attachment.Content = content;
        }
    }

    private byte[] ReadEntryBytes(ZipArchiveEntry entry) {
        long maximumBytes = Math.Min(_options.MaxAttachmentBytes, int.MaxValue);
        if (entry.Length > maximumBytes) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxAttachmentBytes),
                entry.Length, maximumBytes);
        }
        using (Stream stream = OpenDecodedEntry(entry, maximumBytes,
            nameof(EmailStoreReaderOptions.MaxAttachmentBytes)))
        using (var output = new MemoryStream(
            entry.Length > 0 ? checked((int)entry.Length) : 0)) {
            var buffer = new byte[81920];
            while (true) {
                _cancellationToken.ThrowIfCancellationRequested();
                int read = stream.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                output.Write(buffer, 0, read);
            }
            return output.ToArray();
        }
    }

    private static void AddRecipients(EmailDocument document, XElement item,
        string containerName, EmailRecipientKind kind) {
        foreach (EmailAddress address in Addresses(item, containerName)) {
            document.Recipients.Add(new EmailRecipient(kind, address));
        }
    }

    private static void AddAppointmentRecipients(EmailDocument document, XElement item) {
        XElement? container = Child(item, "OPFCalendarEventCopyAttendeeList");
        if (container == null) return;
        var all = new List<string>();
        var required = new List<string>();
        var optional = new List<string>();
        foreach (XElement attendee in container.Descendants().Where(element =>
                     string.Equals(element.Name.LocalName, "appointmentAttendee", StringComparison.OrdinalIgnoreCase))) {
            string? address = Attribute(attendee, "OPFCalendarAttendeeAddress");
            string? name = Attribute(attendee, "OPFCalendarAttendeeName");
            if (address == null && name == null) continue;
            int attendeeType = ParseInteger(Attribute(attendee, "OPFCalendarAttendeeType")) ?? 1;
            EmailRecipientKind kind = attendeeType == 2 ? EmailRecipientKind.Cc
                : attendeeType == 3 ? EmailRecipientKind.Resource
                : EmailRecipientKind.To;
            document.Recipients.Add(new EmailRecipient(kind, new EmailAddress(address, name)));
            string display = name ?? address ?? string.Empty;
            all.Add(display);
            if (kind == EmailRecipientKind.Cc) optional.Add(display);
            else if (kind == EmailRecipientKind.To) required.Add(display);
        }
        OutlookAppointment? appointment = document.Appointment;
        if (appointment == null) return;
        appointment.AllAttendees = all.Count == 0 ? null : string.Join("; ", all);
        appointment.RequiredAttendees = required.Count == 0 ? null : string.Join("; ", required);
        appointment.OptionalAttendees = optional.Count == 0 ? null : string.Join("; ", optional);
    }

    private static EmailAddress? FirstAddress(XElement item, string containerName) {
        return Addresses(item, containerName).FirstOrDefault();
    }

    private static IReadOnlyList<EmailAddress> Addresses(XElement item, string containerName) {
        XElement? container = Child(item, containerName);
        if (container == null) return Array.Empty<EmailAddress>();
        return container.Descendants().Where(element =>
                string.Equals(element.Name.LocalName, "emailAddress", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(element.Name.LocalName, "contactEmailAddress", StringComparison.OrdinalIgnoreCase))
            .Select(element => new EmailAddress(
                Attribute(element, "OPFContactEmailAddressAddress"),
                Attribute(element, "OPFContactEmailAddressName")) {
                AddressType = GetAddressType(element)
            })
            .Where(address => !string.IsNullOrWhiteSpace(address.Address) ||
                              !string.IsNullOrWhiteSpace(address.DisplayName))
            .ToList();
    }

    private static string? GetAddressType(XElement element) {
        string? type = Attribute(element, "OPFContactEmailAddressType");
        string? address = Attribute(element, "OPFContactEmailAddressAddress");
        return address != null && address.Trim().Length > 0 && address.IndexOf('@') >= 0 ? "SMTP" : type;
    }

    private static void ApplyCategories(EmailDocument document, XElement item) {
        foreach (XElement category in item.Descendants().Where(element =>
                     string.Equals(element.Name.LocalName, "category", StringComparison.OrdinalIgnoreCase))) {
            string? name = Attribute(category, "OPFCategoryCopyName");
            if (name != null && name.Trim().Length > 0 &&
                !document.MessageMetadata.Categories.Contains(name, StringComparer.OrdinalIgnoreCase)) {
                document.MessageMetadata.Categories.Add(name);
            }
        }
    }

    private static void PreserveScalarProperties(EmailDocument document, XElement item) {
        foreach (XElement element in item.Elements().Where(element => !element.HasElements)) {
            string name = element.Name.LocalName;
            if (IsProjectedLargeBody(name) || string.IsNullOrEmpty(element.Value)) continue;
            document.Properties[string.Concat("Olm:", name)] = element.Value;
        }
    }

    private static bool IsProjectedLargeBody(string name) {
        return name.EndsWith("CopyBody", StringComparison.OrdinalIgnoreCase) ||
               name.EndsWith("CopyHTMLBody", StringComparison.OrdinalIgnoreCase) ||
               name.EndsWith("CopyDescription", StringComparison.OrdinalIgnoreCase) ||
               name.EndsWith("CopyDescriptionPlain", StringComparison.OrdinalIgnoreCase) ||
               name.EndsWith("CopyNote", StringComparison.OrdinalIgnoreCase) ||
               name.EndsWith("CopyNotePlain", StringComparison.OrdinalIgnoreCase) ||
               name.EndsWith("CopyNotes", StringComparison.OrdinalIgnoreCase) ||
               name.EndsWith("CopyNotesPlain", StringComparison.OrdinalIgnoreCase) ||
               name.EndsWith("CopyText", StringComparison.OrdinalIgnoreCase);
    }

    private static string GetMessageClass(OutlookItemKind kind) {
        switch (kind) {
            case OutlookItemKind.Appointment: return "IPM.Appointment";
            case OutlookItemKind.Contact: return "IPM.Contact";
            case OutlookItemKind.Task: return "IPM.Task";
            case OutlookItemKind.Note: return "IPM.StickyNote";
            default: return "IPM.Note";
        }
    }

    private static XElement? Child(XElement item, string name) {
        return item.Elements().FirstOrDefault(element =>
            string.Equals(element.Name.LocalName, name, StringComparison.OrdinalIgnoreCase));
    }

    private static string? Value(XElement item, string name) {
        string? value = Child(item, name)?.Value;
        return string.IsNullOrEmpty(value) ? null : value;
    }

    private static string? Attribute(XElement item, string name) {
        XAttribute? attribute = item.Attributes().FirstOrDefault(candidate =>
            string.Equals(candidate.Name.LocalName, name, StringComparison.OrdinalIgnoreCase));
        return attribute == null || attribute.Value.Length == 0 ? null : attribute.Value;
    }

    private static DateTimeOffset? DateValue(XElement item, string name) {
        string? value = Value(item, name);
        if (value == null) return null;
        if (DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeUniversal,
                out DateTimeOffset result)) return result;
        return null;
    }

    private static int? IntegerValue(XElement item, string name) {
        return ParseInteger(Value(item, name));
    }

    private static int? ParseInteger(string? value) {
        if (value == null) return null;
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result)) return result;
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double floating) &&
            floating >= int.MinValue && floating <= int.MaxValue) return (int)floating;
        return null;
    }

    private static bool? BooleanValue(XElement item, string name) {
        string? value = Value(item, name);
        if (value == null) return null;
        if (bool.TryParse(value, out bool boolean)) return boolean;
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
            return Math.Abs(number) > double.Epsilon;
        }
        return null;
    }

    private static EmailMessagePriority? PriorityValue(XElement item, string name) {
        int? priority = IntegerValue(item, name);
        if (!priority.HasValue) return null;
        return priority.Value < 3 ? EmailMessagePriority.Urgent
            : priority.Value > 3 ? EmailMessagePriority.NonUrgent
            : EmailMessagePriority.Normal;
    }

    private static string? TrimContentId(string? value) {
        return value == null || value.Trim().Length == 0 ? null : value.Trim().Trim('<', '>');
    }
}
