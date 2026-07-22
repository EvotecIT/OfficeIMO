namespace OfficeIMO.Email;

internal static class MsgProjection {
    internal static readonly Guid PsetidAppointment = MapiPropertySets.Appointment;
    internal static readonly Guid PsetidTask = MapiPropertySets.Task;
    internal static readonly Guid PsetidAddress = MapiPropertySets.Address;
    internal static readonly Guid PsetidCommon = MapiPropertySets.Common;
    internal static readonly Guid PsetidLog = MapiPropertySets.Log;
    internal static readonly Guid PsetidNote = MapiPropertySets.Note;
    internal static readonly Guid PsPublicStrings = MapiPropertySets.PublicStrings;
    internal static readonly Guid PsInternetHeaders = MapiPropertySets.InternetHeaders;
    internal static readonly Guid PsetidCalendarAssistant = MapiPropertySets.CalendarAssistant;
    internal static readonly Guid PsetidSharing = MapiPropertySets.Sharing;
    internal static readonly Guid PsetidReactions = MapiPropertySets.Reactions;

    internal static void Apply(EmailDocument document, MsgParserState state, string location,
        MapiStringEncodingContext encoding) {
        MapiPropertyBag mapi = document.Mapi;
        IList<MapiProperty> properties = mapi.Properties;
        document.MessageClass = mapi.GetValueOrDefault(MapiKnownProperties.PidTag.MessageClass) ?? "IPM.Note";
        document.OutlookItemKind = Classify(document.MessageClass);
        document.Subject = mapi.GetValueOrDefault(MapiKnownProperties.PidTag.Subject) ??
            mapi.GetValueOrDefault(MapiKnownProperties.PidTag.NormalizedSubject);
        document.MessageId = TrimAngle(mapi.GetValueOrDefault(MapiKnownProperties.PidTag.InternetMessageId));
        document.Date = mapi.GetNullableValue(MapiKnownProperties.PidTag.ClientSubmitTime) ??
            mapi.GetNullableValue(MapiKnownProperties.PidTag.CreationTime);
        document.ReceivedDate = mapi.GetNullableValue(MapiKnownProperties.PidTag.MessageDeliveryTime);
        ApplyMessageMetadata(document, mapi);
        OutlookMessageSemanticsProjection.Apply(document, mapi, encoding.PrimaryCodePage,
            state.Diagnostics, location);
        document.Body.Text = mapi.GetValueOrDefault(MapiKnownProperties.PidTag.Body);
        document.Body.Html = GetHtml(mapi, encoding, state.Diagnostics, location);
        document.Body.Rtf = GetRtf(mapi, state, location);
        if (document.Body.Html == null && document.Body.Rtf != null) {
            document.Body.Html = MsgRtfBodyProjection.TryGetEncapsulatedHtml(
                document.Body.Rtf,
                state,
                string.Concat(location, "/rtf-html"));
        }

        document.From = MsgAddressProjection.ReadAddress(
            properties, MapiKnownProperties.PidTag.SentRepresentingName,
            MapiKnownProperties.PidTag.SentRepresentingSmtpAddress,
            MapiKnownProperties.PidTag.SentRepresentingEmailAddress,
            MapiKnownProperties.PidTag.SentRepresentingAddressType);
        document.Sender = MsgAddressProjection.ReadAddress(
            properties, MapiKnownProperties.PidTag.SenderName, MapiKnownProperties.PidTag.SenderSmtpAddress,
            MapiKnownProperties.PidTag.SenderEmailAddress, MapiKnownProperties.PidTag.SenderAddressType);
        document.ReceivedBy = MsgAddressProjection.ReadAddress(
            properties, MapiKnownProperties.PidTag.ReceivedByName, MapiKnownProperties.PidTag.ReceivedByEmailAddress,
            MapiKnownProperties.PidTag.ReceivedByEmailAddress, MapiKnownProperties.PidTag.ReceivedByAddressType);
        document.ReceivedRepresenting = MsgAddressProjection.ReadAddress(
            properties, MapiKnownProperties.PidTag.ReceivedRepresentingName,
            MapiKnownProperties.PidTag.ReceivedRepresentingEmailAddress,
            MapiKnownProperties.PidTag.ReceivedRepresentingEmailAddress,
            MapiKnownProperties.PidTag.ReceivedRepresentingAddressType);

        string? transportHeaders = mapi.GetValueOrDefault(MapiKnownProperties.PidTag.TransportMessageHeaders);
        if (!string.IsNullOrWhiteSpace(transportHeaders)) {
            byte[] bytes = Encoding.UTF8.GetBytes(string.Concat(transportHeaders, "\r\n\r\n"));
            var parsedHeaders = new List<EmailHeader>();
            MimeHeaderParser.Parse(bytes, 0, bytes.Length, state.Options, parsedHeaders, state.Diagnostics,
                string.Concat(location, "/transport-headers"));
            foreach (EmailHeader header in parsedHeaders) document.Headers.Add(header);
            MimeMessageMetadataProjection.ApplyReceiptDestinations(document, parsedHeaders);
            document.From = document.From ?? MimeAddressParser.ParseOne(
                MimeHeaderParser.GetRawValue(parsedHeaders, "From"), state.Diagnostics,
                string.Concat(location, "/transport-headers/From"));
            document.Sender = document.Sender ?? MimeAddressParser.ParseOne(
                MimeHeaderParser.GetRawValue(parsedHeaders, "Sender"), state.Diagnostics,
                string.Concat(location, "/transport-headers/Sender"));
            document.Subject = document.Subject ?? MimeHeaderParser.GetValue(parsedHeaders, "Subject");
            document.MessageId = document.MessageId ?? TrimAngle(MimeHeaderParser.GetValue(parsedHeaders, "Message-ID"));
        }

        foreach (MapiProperty property in properties) {
            string key = property.Name == null
                ? string.Concat("Mapi:0x", property.PropertyTag.ToString("X8", CultureInfo.InvariantCulture))
                : string.Concat("Mapi:", property.Name.ToString(), ":", property.PropertyType.ToString());
            document.Properties[key] = property.Value;
        }

        ApplyReplyTo(document, mapi);
        ApplyTyped(document);
    }

    private static void ApplyReplyTo(EmailDocument document, MapiPropertyBag properties) {
        string? replyTo = properties.GetValueOrDefault(MapiKnownProperties.PidTag.ReplyRecipientNames);
        if (string.IsNullOrWhiteSpace(replyTo)) return;

        foreach (EmailAddress address in MimeAddressParser.ParseMany(replyTo, allowSemicolonSeparator: true)) {
            if (document.Recipients.Any(recipient => recipient.Kind == EmailRecipientKind.ReplyTo &&
                string.Equals(recipient.Address.Address, address.Address, StringComparison.OrdinalIgnoreCase))) {
                continue;
            }
            document.Recipients.Add(new EmailRecipient(EmailRecipientKind.ReplyTo, address));
        }
    }

    internal static void ApplyTransportHeaderRecipients(EmailDocument document, MsgParserState state, string location) {
        AddTransportHeaderRecipients(document, state, location, "To", EmailRecipientKind.To);
        AddTransportHeaderRecipients(document, state, location, "Cc", EmailRecipientKind.Cc);
        AddTransportHeaderRecipients(document, state, location, "Bcc", EmailRecipientKind.Bcc);
        AddTransportHeaderRecipients(document, state, location, "Reply-To", EmailRecipientKind.ReplyTo);
    }

    private static void AddTransportHeaderRecipients(EmailDocument document, MsgParserState state, string location,
        string headerName, EmailRecipientKind kind) {
        if (document.Recipients.Any(recipient => recipient.Kind == kind)) return;
        foreach (string value in MimeHeaderParser.GetRawValues(document.Headers, headerName)) {
            foreach (EmailAddress address in MimeAddressParser.ParseMany(value, state.Diagnostics,
                string.Concat(location, "/transport-headers/", headerName))) {
                document.Recipients.Add(new EmailRecipient(kind, address));
            }
        }
    }

    private static void ApplyMessageMetadata(EmailDocument document, MapiPropertyBag properties) {
        EmailMessageMetadata metadata = document.MessageMetadata;
        metadata.SubjectPrefix = properties.GetValueOrDefault(MapiKnownProperties.PidTag.SubjectPrefix);
        metadata.NormalizedSubject = properties.GetValueOrDefault(MapiKnownProperties.PidTag.NormalizedSubject);
        metadata.ConversationTopic = properties.GetValueOrDefault(MapiKnownProperties.PidTag.ConversationTopic);
        metadata.ConversationIndex = properties.GetValueOrDefault(MapiKnownProperties.PidTag.ConversationIndex);
        metadata.InternetReferences = properties.GetValueOrDefault(MapiKnownProperties.PidTag.InternetReferences);
        metadata.InReplyToId = properties.GetValueOrDefault(MapiKnownProperties.PidTag.InReplyToId);
        int? importance = properties.GetNullableValue(MapiKnownProperties.PidTag.Importance);
        if (importance.HasValue && Enum.IsDefined(typeof(EmailMessageImportance), importance.Value)) {
            metadata.Importance = (EmailMessageImportance)importance.Value;
        }
        int? priority = properties.GetNullableValue(MapiKnownProperties.PidTag.Priority);
        if (priority.HasValue && Enum.IsDefined(typeof(EmailMessagePriority), priority.Value)) {
            metadata.Priority = (EmailMessagePriority)priority.Value;
        }
        metadata.IconIndex = properties.GetNullableValue(MapiKnownProperties.PidTag.IconIndex);
        int flags = properties.GetValueOrDefault(MapiKnownProperties.PidTag.MessageFlags);
        metadata.IsDraft = (flags & 0x0008) != 0;
        metadata.IsRead = (flags & 0x0001) != 0;
        metadata.ReadReceiptRequested = properties.GetValueOrDefault(MapiKnownProperties.PidTag.ReadReceiptRequested);
        metadata.DeliveryReceiptRequested = properties.GetValueOrDefault(
            MapiKnownProperties.PidTag.OriginatorDeliveryReportRequested);
        metadata.Sensitivity = properties.GetNullableValue(MapiKnownProperties.PidTag.Sensitivity);
        metadata.OriginalSensitivity = properties.GetNullableValue(MapiKnownProperties.PidTag.OriginalSensitivity);
        metadata.CreatedDate = properties.GetNullableValue(MapiKnownProperties.PidTag.CreationTime);
        metadata.ModifiedDate = properties.GetNullableValue(MapiKnownProperties.PidTag.LastModificationTime);
        metadata.LastModifierName = properties.GetValueOrDefault(MapiKnownProperties.PidTag.LastModifierName);
        metadata.LocaleId = properties.GetNullableValue(MapiKnownProperties.PidTag.MessageLocaleId);
        metadata.DeclaredSize = properties.GetNullableValue(MapiKnownProperties.PidTag.MessageSize);
        metadata.ConversationId = properties.GetValueOrDefault(MapiKnownProperties.PidTag.ConversationId);
        metadata.EditorFormat = properties.GetNullableValue(MapiKnownProperties.PidTag.MessageEditorFormat);
        metadata.ReactionsSummary = properties.GetValueOrDefault(MapiKnownProperties.PidName.ReactionsSummary);
        metadata.OwnerReactionHistory = properties.GetValueOrDefault(MapiKnownProperties.PidName.OwnerReactionHistory);
        metadata.OwnerReactionType = properties.GetValueOrDefault(MapiKnownProperties.PidName.OwnerReactionType);
        metadata.OwnerReactionTime = properties.GetNullableValue(MapiKnownProperties.PidName.OwnerReactionTime);
        metadata.ReactionsCount = properties.GetNullableValue(MapiKnownProperties.PidName.ReactionsCount);
        MapiProperty? categories = properties.Find(MapiKnownProperties.PidName.Keywords) ??
            properties.Find(MapiKnownProperties.PidLid.LegacyKeywords);
        AddStrings(metadata.Categories, categories?.Value);
    }

    internal static void ApplyTyped(EmailDocument document) {
        MapiPropertyBag properties = document.Mapi;
        switch (document.OutlookItemKind) {
            case OutlookItemKind.Appointment:
                document.Appointment = CreateAppointment(properties, document.OutlookCodePage ?? 1252);
                document.MeetingCommunication = CreateMeetingCommunication(document.MessageClass, properties);
                break;
            case OutlookItemKind.Contact: document.Contact = CreateContact(properties); break;
            case OutlookItemKind.DistributionList:
                document.DistributionList = OutlookDistributionList.Project(properties);
                break;
            case OutlookItemKind.Task:
                document.Task = CreateTask(properties);
                document.TaskCommunication = CreateTaskCommunication(document.MessageClass);
                break;
            case OutlookItemKind.Journal: document.Journal = CreateJournal(properties); break;
            case OutlookItemKind.Note: document.Note = CreateNote(properties); break;
        }
    }

    internal static void ApplyAttachmentSemantics(EmailDocument document) {
        OutlookTaskCommunication? communication = document.TaskCommunication;
        if (communication == null || communication.Kind == OutlookTaskCommunicationKind.None) return;
        EmailAttachment? payload = document.Attachments.FirstOrDefault();
        communication.PayloadAttachment = payload;
        communication.EmbeddedTask = payload?.EmbeddedDocument;
    }

    internal static OutlookItemKind Classify(string? messageClass) {
        if (messageClass == null) return OutlookItemKind.Unknown;
        if (messageClass.StartsWith("IPM.Appointment", StringComparison.OrdinalIgnoreCase) ||
            messageClass.StartsWith("IPM.Schedule.Meeting", StringComparison.OrdinalIgnoreCase)) return OutlookItemKind.Appointment;
        if (messageClass.StartsWith("IPM.DistList", StringComparison.OrdinalIgnoreCase)) return OutlookItemKind.DistributionList;
        if (messageClass.StartsWith("IPM.Contact", StringComparison.OrdinalIgnoreCase)) return OutlookItemKind.Contact;
        if (messageClass.StartsWith("IPM.Task", StringComparison.OrdinalIgnoreCase)) return OutlookItemKind.Task;
        if (messageClass.StartsWith("IPM.Activity", StringComparison.OrdinalIgnoreCase)) return OutlookItemKind.Journal;
        if (messageClass.StartsWith("IPM.StickyNote", StringComparison.OrdinalIgnoreCase)) return OutlookItemKind.Note;
        return OutlookItemKind.Message;
    }

    private static OutlookAppointment CreateAppointment(MapiPropertyBag properties, int string8CodePage) {
        byte[]? recurrenceState = properties.GetValueOrDefault(MapiKnownProperties.PidLid.AppointmentRecur);
        byte[]? legacyTimeZone = properties.GetValueOrDefault(MapiKnownProperties.PidLid.TimeZoneStruct);
        byte[]? startTimeZone = properties.GetValueOrDefault(
            MapiKnownProperties.PidLid.AppointmentTimeZoneDefinitionStartDisplay);
        byte[]? endTimeZone = properties.GetValueOrDefault(
            MapiKnownProperties.PidLid.AppointmentTimeZoneDefinitionEndDisplay);
        byte[]? recurrenceTimeZone = properties.GetValueOrDefault(
            MapiKnownProperties.PidLid.AppointmentTimeZoneDefinitionRecur);
        var appointment = new OutlookAppointment {
            GlobalObjectId = properties.GetValueOrDefault(MapiKnownProperties.PidLid.GlobalObjectId),
            CleanGlobalObjectId = properties.GetValueOrDefault(MapiKnownProperties.PidLid.CleanGlobalObjectId),
            Start = properties.GetNullableValue(MapiKnownProperties.PidLid.AppointmentStartWhole) ??
                properties.GetNullableValue(MapiKnownProperties.PidLid.CommonStart),
            End = properties.GetNullableValue(MapiKnownProperties.PidLid.AppointmentEndWhole) ??
                properties.GetNullableValue(MapiKnownProperties.PidLid.CommonEnd),
            Location = properties.GetValueOrDefault(MapiKnownProperties.PidLid.Location),
            IsAllDay = properties.GetNullableValue(MapiKnownProperties.PidLid.AppointmentSubType),
            BusyStatus = properties.GetNullableValue(MapiKnownProperties.PidLid.BusyStatus),
            MeetingStatus = properties.GetNullableValue(MapiKnownProperties.PidLid.AppointmentStateFlags),
            ResponseStatus = properties.GetNullableValue(MapiKnownProperties.PidLid.ResponseStatus),
            Sequence = properties.GetNullableValue(MapiKnownProperties.PidLid.AppointmentSequence),
            DurationMinutes = properties.GetNullableValue(MapiKnownProperties.PidLid.AppointmentDuration),
            AllAttendees = properties.GetValueOrDefault(MapiKnownProperties.PidLid.AllAttendeesString),
            RequiredAttendees = properties.GetValueOrDefault(MapiKnownProperties.PidLid.ToAttendeesString),
            OptionalAttendees = properties.GetValueOrDefault(MapiKnownProperties.PidLid.CcAttendeesString),
            NotAllowPropose = properties.GetNullableValue(MapiKnownProperties.PidLid.AppointmentNotAllowPropose),
            RecurrenceType = properties.GetNullableValue(MapiKnownProperties.PidLid.RecurrenceType),
            RecurrencePattern = properties.GetValueOrDefault(MapiKnownProperties.PidLid.RecurrencePattern),
            RecurrenceState = recurrenceState,
            Recurrence = recurrenceState == null ? null :
                OutlookRecurrenceBinary.DecodeAppointment(recurrenceState, string8CodePage),
            IsRecurring = properties.GetNullableValue(MapiKnownProperties.PidLid.Recurring),
            ClientIntentFlags = properties.GetNullableValue(MapiKnownProperties.PidLid.ClientIntent),
            TimeZoneStructure = legacyTimeZone,
            LegacyTimeZone = legacyTimeZone == null ? null : OutlookTimeZoneBinary.DecodeStructure(legacyTimeZone),
            TimeZoneDescription = properties.GetValueOrDefault(MapiKnownProperties.PidLid.TimeZoneDescription),
            StartTimeZoneDefinition = startTimeZone,
            StartTimeZone = startTimeZone == null ? null : OutlookTimeZoneBinary.DecodeDefinition(startTimeZone),
            EndTimeZoneDefinition = endTimeZone,
            EndTimeZone = endTimeZone == null ? null : OutlookTimeZoneBinary.DecodeDefinition(endTimeZone),
            RecurrenceTimeZoneDefinition = recurrenceTimeZone,
            RecurrenceTimeZone = recurrenceTimeZone == null ? null :
                OutlookTimeZoneBinary.DecodeDefinition(recurrenceTimeZone)
        };
        if (appointment.Recurrence != null)
            appointment.Recurrence.TimeZoneId = appointment.RecurrenceTimeZone?.KeyName ??
                appointment.StartTimeZone?.KeyName ?? appointment.TimeZoneDescription;
        OutlookMessageSemanticsProjection.ApplyReminder(appointment.Reminder, properties);
        return appointment;
    }

    private static OutlookMeetingCommunication? CreateMeetingCommunication(string? messageClass,
        MapiPropertyBag properties) {
        OutlookMeetingCommunicationKind kind = ClassifyMeetingCommunication(messageClass);
        if (kind == OutlookMeetingCommunicationKind.None) return null;
        return new OutlookMeetingCommunication {
            Kind = kind,
            RequestTypeValue = properties.GetNullableValue(MapiKnownProperties.PidLid.MeetingType),
            IntendedBusyStatus = properties.GetNullableValue(MapiKnownProperties.PidLid.IntendedBusyStatus),
            OwnerCriticalChange = properties.GetNullableValue(MapiKnownProperties.PidLid.OwnerCriticalChange),
            AttendeeCriticalChange = properties.GetNullableValue(MapiKnownProperties.PidLid.AttendeeCriticalChange),
            IsSilent = properties.GetNullableValue(MapiKnownProperties.PidLid.IsSilent),
            IsCounterProposal = properties.GetNullableValue(MapiKnownProperties.PidLid.AppointmentCounterProposal),
            ProposedStart = properties.GetNullableValue(MapiKnownProperties.PidLid.AppointmentProposedStartWhole),
            ProposedEnd = properties.GetNullableValue(MapiKnownProperties.PidLid.AppointmentProposedEndWhole),
            ProposedDurationMinutes = properties.GetNullableValue(MapiKnownProperties.PidLid.AppointmentProposedDuration),
            ReplyAt = properties.GetNullableValue(MapiKnownProperties.PidLid.AppointmentReplyTime),
            ReplyName = properties.GetValueOrDefault(MapiKnownProperties.PidLid.AppointmentReplyName)
        };
    }

    private static OutlookMeetingCommunicationKind ClassifyMeetingCommunication(string? messageClass) {
        if (messageClass == null) return OutlookMeetingCommunicationKind.None;
        if (messageClass.StartsWith("IPM.Schedule.Meeting.Request", StringComparison.OrdinalIgnoreCase))
            return OutlookMeetingCommunicationKind.RequestOrUpdate;
        if (messageClass.StartsWith("IPM.Schedule.Meeting.Canceled", StringComparison.OrdinalIgnoreCase))
            return OutlookMeetingCommunicationKind.Cancellation;
        if (messageClass.StartsWith("IPM.Schedule.Meeting.Resp.Pos", StringComparison.OrdinalIgnoreCase))
            return OutlookMeetingCommunicationKind.ResponseAccepted;
        if (messageClass.StartsWith("IPM.Schedule.Meeting.Resp.Tent", StringComparison.OrdinalIgnoreCase))
            return OutlookMeetingCommunicationKind.ResponseTentative;
        if (messageClass.StartsWith("IPM.Schedule.Meeting.Resp.Neg", StringComparison.OrdinalIgnoreCase))
            return OutlookMeetingCommunicationKind.ResponseDeclined;
        if (messageClass.StartsWith("IPM.Schedule.Meeting.Forward.Notification", StringComparison.OrdinalIgnoreCase))
            return OutlookMeetingCommunicationKind.ForwardNotification;
        return OutlookMeetingCommunicationKind.None;
    }

    private static OutlookTaskCommunication? CreateTaskCommunication(string? messageClass) {
        OutlookTaskCommunicationKind kind;
        if (messageClass?.StartsWith("IPM.TaskRequest.Accept", StringComparison.OrdinalIgnoreCase) == true)
            kind = OutlookTaskCommunicationKind.Accept;
        else if (messageClass?.StartsWith("IPM.TaskRequest.Decline", StringComparison.OrdinalIgnoreCase) == true)
            kind = OutlookTaskCommunicationKind.Decline;
        else if (messageClass?.StartsWith("IPM.TaskRequest.Update", StringComparison.OrdinalIgnoreCase) == true)
            kind = OutlookTaskCommunicationKind.Update;
        else if (messageClass?.StartsWith("IPM.TaskRequest", StringComparison.OrdinalIgnoreCase) == true)
            kind = OutlookTaskCommunicationKind.Request;
        else return null;
        return new OutlookTaskCommunication { Kind = kind };
    }

    private static OutlookContact CreateContact(MapiPropertyBag properties) {
        var contact = new OutlookContact {
            DisplayName = properties.GetValueOrDefault(MapiKnownProperties.PidTag.DisplayName),
            Prefix = properties.GetValueOrDefault(MapiKnownProperties.PidTag.DisplayNamePrefix),
            Initials = properties.GetValueOrDefault(MapiKnownProperties.PidTag.Initials),
            GivenName = properties.GetValueOrDefault(MapiKnownProperties.PidTag.GivenName),
            MiddleName = properties.GetValueOrDefault(MapiKnownProperties.PidTag.MiddleName),
            Surname = properties.GetValueOrDefault(MapiKnownProperties.PidTag.Surname),
            Generation = properties.GetValueOrDefault(MapiKnownProperties.PidTag.Generation),
            CompanyName = properties.GetValueOrDefault(MapiKnownProperties.PidTag.CompanyName),
            JobTitle = properties.GetValueOrDefault(MapiKnownProperties.PidTag.Title),
            Department = properties.GetValueOrDefault(MapiKnownProperties.PidTag.DepartmentName),
            FileAs = properties.GetValueOrDefault(MapiKnownProperties.PidLid.FileUnder),
            NickName = properties.GetValueOrDefault(MapiKnownProperties.PidTag.Nickname),
            ManagerName = properties.GetValueOrDefault(MapiKnownProperties.PidTag.ManagerName),
            AssistantName = properties.GetValueOrDefault(MapiKnownProperties.PidTag.Assistant),
            SpouseName = properties.GetValueOrDefault(MapiKnownProperties.PidTag.SpouseName),
            Profession = properties.GetValueOrDefault(MapiKnownProperties.PidTag.Profession),
            Language = properties.GetValueOrDefault(MapiKnownProperties.PidTag.Language),
            Location = properties.GetValueOrDefault(MapiKnownProperties.PidTag.MailUserLocation),
            OfficeLocation = properties.GetValueOrDefault(MapiKnownProperties.PidTag.OfficeLocation),
            Birthday = properties.GetNullableValue(MapiKnownProperties.PidTag.Birthday) ??
                properties.GetNullableValue(MapiKnownProperties.PidLid.BirthdayLocal),
            WeddingAnniversary = properties.GetNullableValue(MapiKnownProperties.PidTag.WeddingAnniversary) ??
                properties.GetNullableValue(MapiKnownProperties.PidLid.WeddingAnniversaryLocal),
            IsPrivate = properties.GetNullableValue(MapiKnownProperties.PidLid.Private),
            HasPicture = properties.GetNullableValue(MapiKnownProperties.PidLid.HasPicture),
            InstantMessagingAddress = properties.GetValueOrDefault(MapiKnownProperties.PidLid.InstantMessagingAddress),
            BusinessHomePage = properties.GetValueOrDefault(MapiKnownProperties.PidTag.BusinessHomePage),
            PersonalHomePage = properties.GetValueOrDefault(MapiKnownProperties.PidTag.PersonalHomePage),
            Html = properties.GetValueOrDefault(MapiKnownProperties.PidLid.ContactHtml)
        };
        AddStrings(contact.Children, properties.Find(MapiKnownProperties.PidTag.ChildrensNames)?.Value);

        PopulateAddress(contact.BusinessAddress, properties, null,
            MapiKnownProperties.PidTag.StreetAddress, MapiKnownProperties.PidTag.Locality,
            MapiKnownProperties.PidTag.StateOrProvince, MapiKnownProperties.PidTag.PostalCode,
            MapiKnownProperties.PidTag.Country, MapiKnownProperties.PidTag.PostOfficeBox);
        PopulateAddress(contact.HomeAddress, properties, MapiKnownProperties.PidLid.HomeAddress,
            MapiKnownProperties.PidTag.HomeAddressStreet, MapiKnownProperties.PidTag.HomeAddressCity,
            MapiKnownProperties.PidTag.HomeAddressStateOrProvince, MapiKnownProperties.PidTag.HomeAddressPostalCode,
            MapiKnownProperties.PidTag.HomeAddressCountry, MapiKnownProperties.PidTag.HomeAddressPostOfficeBox);
        PopulateAddress(contact.OtherAddress, properties, MapiKnownProperties.PidLid.OtherAddress,
            MapiKnownProperties.PidTag.OtherAddressStreet, MapiKnownProperties.PidTag.OtherAddressCity,
            MapiKnownProperties.PidTag.OtherAddressStateOrProvince, MapiKnownProperties.PidTag.OtherAddressPostalCode,
            MapiKnownProperties.PidTag.OtherAddressCountry, MapiKnownProperties.PidTag.OtherAddressPostOfficeBox);
        contact.WorkAddress.Formatted = properties.GetValueOrDefault(MapiKnownProperties.PidLid.WorkAddress);
        contact.WorkAddress.Street = properties.GetValueOrDefault(MapiKnownProperties.PidLid.WorkAddressStreet);
        contact.WorkAddress.City = properties.GetValueOrDefault(MapiKnownProperties.PidLid.WorkAddressCity);
        contact.WorkAddress.StateOrProvince = properties.GetValueOrDefault(MapiKnownProperties.PidLid.WorkAddressState);
        contact.WorkAddress.PostalCode = properties.GetValueOrDefault(MapiKnownProperties.PidLid.WorkAddressPostalCode);
        contact.WorkAddress.Country = properties.GetValueOrDefault(MapiKnownProperties.PidLid.WorkAddressCountry);
        contact.WorkAddress.PostOfficeBox = properties.GetValueOrDefault(MapiKnownProperties.PidLid.WorkAddressPostOfficeBox);
        contact.WorkAddress.CountryCode = properties.GetValueOrDefault(MapiKnownProperties.PidLid.WorkAddressCountryCode);
        contact.BusinessAddress.Formatted = contact.WorkAddress.Formatted ??
            properties.GetValueOrDefault(MapiKnownProperties.PidTag.PostalAddress);

        contact.Phones.Business = properties.GetValueOrDefault(MapiKnownProperties.PidTag.BusinessTelephoneNumber);
        contact.Phones.Business2 = properties.GetValueOrDefault(MapiKnownProperties.PidTag.Business2TelephoneNumber);
        contact.Phones.Home = properties.GetValueOrDefault(MapiKnownProperties.PidTag.HomeTelephoneNumber);
        contact.Phones.Home2 = properties.GetValueOrDefault(MapiKnownProperties.PidTag.Home2TelephoneNumber);
        contact.Phones.Mobile = properties.GetValueOrDefault(MapiKnownProperties.PidTag.MobileTelephoneNumber);
        contact.Phones.Other = properties.GetValueOrDefault(MapiKnownProperties.PidTag.OtherTelephoneNumber);
        contact.Phones.Primary = properties.GetValueOrDefault(MapiKnownProperties.PidTag.PrimaryTelephoneNumber);
        contact.Phones.BusinessFax = properties.GetValueOrDefault(MapiKnownProperties.PidTag.BusinessFaxNumber);
        contact.Phones.HomeFax = properties.GetValueOrDefault(MapiKnownProperties.PidTag.HomeFaxNumber);
        contact.Phones.PrimaryFax = properties.GetValueOrDefault(MapiKnownProperties.PidTag.PrimaryFaxNumber);
        contact.Phones.Assistant = properties.GetValueOrDefault(MapiKnownProperties.PidTag.AssistantTelephoneNumber);
        contact.Phones.CompanyMain = properties.GetValueOrDefault(MapiKnownProperties.PidTag.CompanyMainPhoneNumber);
        contact.Phones.Car = properties.GetValueOrDefault(MapiKnownProperties.PidTag.CarTelephoneNumber);
        contact.Phones.Radio = properties.GetValueOrDefault(MapiKnownProperties.PidTag.RadioTelephoneNumber);
        contact.Phones.Pager = properties.GetValueOrDefault(MapiKnownProperties.PidTag.PagerTelephoneNumber);
        contact.Phones.Callback = properties.GetValueOrDefault(MapiKnownProperties.PidTag.CallbackTelephoneNumber);
        contact.Phones.Telex = properties.GetValueOrDefault(MapiKnownProperties.PidTag.TelexNumber);
        contact.Phones.TextTelephone = properties.GetValueOrDefault(MapiKnownProperties.PidTag.TtyTddPhoneNumber);
        contact.Phones.Isdn = properties.GetValueOrDefault(MapiKnownProperties.PidTag.IsdnNumber);

        PopulateEmail(contact.Email1, properties, MapiKnownProperties.PidLid.Email1DisplayName,
            MapiKnownProperties.PidLid.Email1AddressType, MapiKnownProperties.PidLid.Email1EmailAddress,
            MapiKnownProperties.PidLid.Email1OriginalDisplayName, MapiKnownProperties.PidLid.Email1OriginalEntryId);
        PopulateEmail(contact.Email2, properties, MapiKnownProperties.PidLid.Email2DisplayName,
            MapiKnownProperties.PidLid.Email2AddressType, MapiKnownProperties.PidLid.Email2EmailAddress,
            MapiKnownProperties.PidLid.Email2OriginalDisplayName, MapiKnownProperties.PidLid.Email2OriginalEntryId);
        PopulateEmail(contact.Email3, properties, MapiKnownProperties.PidLid.Email3DisplayName,
            MapiKnownProperties.PidLid.Email3AddressType, MapiKnownProperties.PidLid.Email3EmailAddress,
            MapiKnownProperties.PidLid.Email3OriginalDisplayName, MapiKnownProperties.PidLid.Email3OriginalEntryId);
        return contact;
    }

    private static OutlookTask CreateTask(MapiPropertyBag properties) {
        byte[]? recurrenceState = properties.GetValueOrDefault(MapiKnownProperties.PidLid.TaskRecurrence);
        var task = new OutlookTask {
            Start = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskStartDate),
            Due = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskDueDate),
            Status = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskStatus),
            PercentComplete = properties.GetNullableValue(MapiKnownProperties.PidLid.PercentComplete),
            IsComplete = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskComplete),
            Owner = properties.GetValueOrDefault(MapiKnownProperties.PidLid.TaskOwner),
            ActualEffort = ToMinutes(properties.GetNullableValue(MapiKnownProperties.PidLid.TaskActualEffort)),
            EstimatedEffort = ToMinutes(properties.GetNullableValue(MapiKnownProperties.PidLid.TaskEstimatedEffort)),
            SendUpdates = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskUpdates),
            SendStatusOnComplete = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskStatusOnComplete),
            Ownership = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskOwnership),
            AcceptanceState = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskAcceptanceState),
            Version = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskVersion),
            State = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskState),
            Assigner = properties.GetValueOrDefault(MapiKnownProperties.PidLid.TaskAssigner),
            IsTeamTask = properties.GetNullableValue(MapiKnownProperties.PidLid.TeamTask),
            Ordinal = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskOrdinal),
            IsRecurring = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskFRecurring),
            RecurrenceState = recurrenceState,
            Recurrence = recurrenceState == null ? null : OutlookRecurrenceBinary.DecodeTask(recurrenceState),
            CommonStart = properties.GetNullableValue(MapiKnownProperties.PidLid.CommonStart),
            CommonEnd = properties.GetNullableValue(MapiKnownProperties.PidLid.CommonEnd),
            Mode = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskMode),
            IsAccepted = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskAccepted),
            History = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskHistory),
            LastUpdate = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskLastUpdate),
            LastUser = properties.GetValueOrDefault(MapiKnownProperties.PidLid.TaskLastUser),
            LastDelegate = properties.GetValueOrDefault(MapiKnownProperties.PidLid.TaskLastDelegate),
            GlobalId = ToGuid(properties.GetValueOrDefault(MapiKnownProperties.PidLid.TaskGlobalId)),
            ToDoOrdinalDate = properties.GetNullableValue(MapiKnownProperties.PidLid.ToDoOrdinalDate),
            ToDoSubOrdinal = properties.GetValueOrDefault(MapiKnownProperties.PidLid.ToDoSubOrdinal),
            BillingInformation = properties.GetValueOrDefault(MapiKnownProperties.PidLid.Billing),
            Mileage = properties.GetValueOrDefault(MapiKnownProperties.PidLid.Mileage),
            CompletedAt = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskDateCompleted)
        };
        OutlookMessageSemanticsProjection.ApplyReminder(task.Reminder, properties);
        AddStrings(task.Contacts, properties.Find(MapiKnownProperties.PidLid.Contacts)?.Value);
        AddStrings(task.Companies, properties.Find(MapiKnownProperties.PidLid.Companies)?.Value);
        return task;
    }

    private static Guid? ToGuid(byte[]? value) => value?.Length == 16 ? new Guid(value) : (Guid?)null;

    private static OutlookJournal CreateJournal(MapiPropertyBag properties) {
        return new OutlookJournal {
            Start = properties.GetNullableValue(MapiKnownProperties.PidLid.CommonStart),
            End = properties.GetNullableValue(MapiKnownProperties.PidLid.CommonEnd) ??
                properties.GetNullableValue(MapiKnownProperties.PidLid.LogEnd),
            DurationMinutes = properties.GetNullableValue(MapiKnownProperties.PidLid.LogDuration),
            Type = properties.GetValueOrDefault(MapiKnownProperties.PidLid.LogType),
            TypeDescription = properties.GetValueOrDefault(MapiKnownProperties.PidLid.LogTypeDesc),
            Flags = properties.GetNullableValue(MapiKnownProperties.PidLid.LogFlags),
            DocumentPrinted = properties.GetNullableValue(MapiKnownProperties.PidLid.LogDocumentPrinted),
            DocumentSaved = properties.GetNullableValue(MapiKnownProperties.PidLid.LogDocumentSaved),
            DocumentRouted = properties.GetNullableValue(MapiKnownProperties.PidLid.LogDocumentRouted),
            DocumentPosted = properties.GetNullableValue(MapiKnownProperties.PidLid.LogDocumentPosted)
        };
    }

    private static OutlookNote CreateNote(MapiPropertyBag properties) {
        return new OutlookNote {
            Color = properties.GetNullableValue(MapiKnownProperties.PidLid.NoteColor),
            Width = properties.GetNullableValue(MapiKnownProperties.PidLid.NoteWidth),
            Height = properties.GetNullableValue(MapiKnownProperties.PidLid.NoteHeight),
            X = properties.GetNullableValue(MapiKnownProperties.PidLid.NoteX),
            Y = properties.GetNullableValue(MapiKnownProperties.PidLid.NoteY)
        };
    }

    private static void PopulateAddress(OutlookPostalAddress address, MapiPropertyBag properties,
        MapiPropertyKey<string>? formattedKey, MapiPropertyKey<string> streetKey, MapiPropertyKey<string> cityKey,
        MapiPropertyKey<string> stateKey, MapiPropertyKey<string> postalKey, MapiPropertyKey<string> countryKey,
        MapiPropertyKey<string> postOfficeBoxKey) {
        if (formattedKey != null) address.Formatted = properties.GetValueOrDefault(formattedKey);
        address.Street = properties.GetValueOrDefault(streetKey);
        address.City = properties.GetValueOrDefault(cityKey);
        address.StateOrProvince = properties.GetValueOrDefault(stateKey);
        address.PostalCode = properties.GetValueOrDefault(postalKey);
        address.Country = properties.GetValueOrDefault(countryKey);
        address.PostOfficeBox = properties.GetValueOrDefault(postOfficeBoxKey);
    }

    private static void PopulateEmail(OutlookContactEmailAddress email, MapiPropertyBag properties,
        MapiPropertyKey<string> displayKey, MapiPropertyKey<string> addressTypeKey,
        MapiPropertyKey<string> addressKey, MapiPropertyKey<string> originalDisplayKey,
        MapiPropertyKey<byte[]> entryKey) {
        email.DisplayName = properties.GetValueOrDefault(displayKey);
        email.AddressType = properties.GetValueOrDefault(addressTypeKey);
        email.Address = properties.GetValueOrDefault(addressKey);
        email.OriginalDisplayName = properties.GetValueOrDefault(originalDisplayKey);
        email.OriginalEntryId = properties.GetValueOrDefault(entryKey);
    }

    private static void AddDelimited(IList<string> target, string? value) {
        if (string.IsNullOrWhiteSpace(value)) return;
        foreach (string item in value!.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)) {
            string clean = item.Trim();
            if (clean.Length > 0) target.Add(clean);
        }
    }

    private static void AddStrings(IList<string> target, object? value) {
        if (value is string scalar) {
            AddDelimited(target, scalar);
        } else if (value is object[] values) {
            foreach (string item in values.OfType<string>()) target.Add(item);
        }
    }

    private static TimeSpan? ToMinutes(int? minutes) => minutes.HasValue ? TimeSpan.FromMinutes(minutes.Value) : null;

    private static string? GetHtml(MapiPropertyBag properties, MapiStringEncodingContext encoding,
        IList<EmailDiagnostic> diagnostics, string location) {
        MapiProperty? property = properties.Find(MapiKnownProperties.PidTag.Html);
        if (property?.Value is string text) return text;
        if (property?.Value is byte[] bytes) {
            return encoding.Decode(bytes, diagnostics, location).TrimEnd('\0');
        }
        return null;
    }

    private static string? GetRtf(MapiPropertyBag properties, MsgParserState state, string location) {
        MapiProperty? property = properties.Find(MapiKnownProperties.PidTag.RtfCompressed);
        if (!(property?.Value is byte[] compressed)) return null;
        if (!MapiCompressedRtfCodec.TryDecompress(compressed, state.RemainingDecodedPropertyBytes,
            state.Diagnostics, string.Concat(location, "/rtf"), state.CancellationToken, out byte[] rtfBytes)) return null;
        state.CountDecodedBytes(rtfBytes.Length);
        char[] characters = new char[rtfBytes.Length];
        for (int index = 0; index < rtfBytes.Length; index++) characters[index] = (char)rtfBytes[index];
        return new string(characters);
    }

    private static string? TrimAngle(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return value;
        return value!.Trim().Trim('<', '>');
    }
}
