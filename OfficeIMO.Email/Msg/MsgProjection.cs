namespace OfficeIMO.Email;

internal static class MsgProjection {
    internal static readonly Guid PsetidAppointment = new Guid("00062002-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidTask = new Guid("00062003-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidAddress = new Guid("00062004-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidCommon = new Guid("00062008-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidLog = new Guid("0006200A-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidNote = new Guid("0006200E-0000-0000-C000-000000000046");
    internal static readonly Guid PsPublicStrings = new Guid("00020329-0000-0000-C000-000000000046");
    internal static readonly Guid PsInternetHeaders = new Guid("00020386-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidCalendarAssistant = new Guid("11000E07-B51B-40D6-AF21-CAA85EDAB1D0");
    internal static readonly Guid PsetidSharing = new Guid("00062041-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidReactions = new Guid("41F28F13-83F4-4114-A584-EEDB5A6B0BFF");

    internal static void Apply(EmailDocument document, MsgParserState state, string location,
        MapiStringEncodingContext encoding) {
        IList<MapiProperty> properties = document.MapiProperties;
        document.MessageClass = GetString(properties, 0x001A) ?? "IPM.Note";
        document.OutlookItemKind = Classify(document.MessageClass);
        document.Subject = GetString(properties, 0x0037) ?? GetString(properties, 0x0E1D);
        document.MessageId = TrimAngle(GetString(properties, 0x1035));
        document.Date = GetDate(properties, 0x0039) ?? GetDate(properties, 0x3007);
        document.ReceivedDate = GetDate(properties, 0x0E06);
        ApplyMessageMetadata(document, properties);
        document.Body.Text = GetString(properties, 0x1000);
        document.Body.Html = GetHtml(properties, encoding, state.Diagnostics, location);
        document.Body.Rtf = GetRtf(properties, state, location);
        if (document.Body.Html == null && document.Body.Rtf != null) {
            document.Body.Html = MsgRtfBodyProjection.TryGetEncapsulatedHtml(
                document.Body.Rtf,
                state,
                string.Concat(location, "/rtf-html"));
        }

        document.From = MsgAddressProjection.ReadAddress(
            properties, 0x0042, 0x5D02, 0x0065, 0x0064);
        document.Sender = MsgAddressProjection.ReadAddress(
            properties, 0x0C1A, 0x5D01, 0x0C1F, 0x0C1E);
        document.ReceivedBy = MsgAddressProjection.ReadAddress(
            properties, 0x0040, 0x0076, 0x0076, 0x0075);
        document.ReceivedRepresenting = MsgAddressProjection.ReadAddress(
            properties, 0x0044, 0x0078, 0x0078, 0x0077);

        string? transportHeaders = GetString(properties, 0x007D);
        if (!string.IsNullOrWhiteSpace(transportHeaders)) {
            byte[] bytes = Encoding.UTF8.GetBytes(string.Concat(transportHeaders, "\r\n\r\n"));
            var parsedHeaders = new List<EmailHeader>();
            MimeHeaderParser.Parse(bytes, 0, bytes.Length, state.Options, parsedHeaders, state.Diagnostics,
                string.Concat(location, "/transport-headers"));
            foreach (EmailHeader header in parsedHeaders) document.Headers.Add(header);
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

        ApplyReplyTo(document, properties);
        ApplyTyped(document);
    }

    private static void ApplyReplyTo(EmailDocument document, IEnumerable<MapiProperty> properties) {
        string? replyTo = GetString(properties, 0x0050);
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

    private static void ApplyMessageMetadata(EmailDocument document, IEnumerable<MapiProperty> properties) {
        EmailMessageMetadata metadata = document.MessageMetadata;
        metadata.SubjectPrefix = GetString(properties, 0x003D);
        metadata.NormalizedSubject = GetString(properties, 0x0E1D);
        metadata.ConversationTopic = GetString(properties, 0x0070);
        metadata.ConversationIndex = properties.FirstOrDefault(property => property.PropertyId == 0x0071)?.Value as byte[];
        metadata.InternetReferences = GetString(properties, 0x1039);
        metadata.InReplyToId = GetString(properties, 0x1042);
        int? importance = GetInt(properties, 0x0017);
        if (importance.HasValue && Enum.IsDefined(typeof(EmailMessageImportance), importance.Value)) {
            metadata.Importance = (EmailMessageImportance)importance.Value;
        }
        int? priority = GetInt(properties, 0x0026);
        if (priority.HasValue && Enum.IsDefined(typeof(EmailMessagePriority), priority.Value)) {
            metadata.Priority = (EmailMessagePriority)priority.Value;
        }
        metadata.IconIndex = GetInt(properties, 0x1080);
        int flags = GetInt(properties, 0x0E07) ?? 0;
        metadata.IsDraft = (flags & 0x0008) != 0;
        metadata.IsRead = (flags & 0x0001) != 0;
        metadata.ReadReceiptRequested = GetBool(properties, 0x0029) ?? false;
        metadata.DeliveryReceiptRequested = GetBool(properties, 0x0023) ?? false;
        metadata.Sensitivity = GetInt(properties, 0x0036);
        metadata.OriginalSensitivity = GetInt(properties, 0x002E);
        metadata.CreatedDate = GetDate(properties, 0x3007);
        metadata.ModifiedDate = GetDate(properties, 0x3008);
        metadata.LastModifierName = GetString(properties, 0x3FFA);
        metadata.LocaleId = GetInt(properties, 0x3FF1);
        metadata.DeclaredSize = GetInt(properties, 0x0E08);
        metadata.ConversationId = properties.FirstOrDefault(property => property.PropertyId == 0x3013)?.Value as byte[];
        metadata.EditorFormat = GetInt(properties, 0x5909);
        metadata.ReactionsSummary = GetNamedByName(properties, "ReactionsSummary")?.Value as byte[];
        metadata.OwnerReactionHistory = GetNamedByName(properties, "OwnerReactionHistory")?.Value as byte[];
        metadata.OwnerReactionType = GetNamedByName(properties, "OwnerReactionType")?.Value as string;
        metadata.OwnerReactionTime = ConvertDate(GetNamedByName(properties, "OwnerReactionTime")?.Value);
        metadata.ReactionsCount = ConvertInt(GetNamedByName(properties, "ReactionsCount")?.Value);
        foreach (string category in GetNamedStrings(properties, PsPublicStrings, "Keywords", 0x9000)) {
            metadata.Categories.Add(category);
        }
    }

    internal static void ApplyTyped(EmailDocument document) {
        IList<MapiProperty> properties = document.MapiProperties;
        switch (document.OutlookItemKind) {
            case OutlookItemKind.Appointment: document.Appointment = CreateAppointment(properties); break;
            case OutlookItemKind.Contact: document.Contact = CreateContact(properties); break;
            case OutlookItemKind.Task: document.Task = CreateTask(properties); break;
            case OutlookItemKind.Journal: document.Journal = CreateJournal(properties); break;
            case OutlookItemKind.Note: document.Note = CreateNote(properties); break;
        }
    }

    internal static OutlookItemKind Classify(string? messageClass) {
        if (messageClass == null) return OutlookItemKind.Unknown;
        if (messageClass.StartsWith("IPM.Appointment", StringComparison.OrdinalIgnoreCase) ||
            messageClass.StartsWith("IPM.Schedule.Meeting", StringComparison.OrdinalIgnoreCase)) return OutlookItemKind.Appointment;
        if (messageClass.StartsWith("IPM.Contact", StringComparison.OrdinalIgnoreCase) ||
            messageClass.StartsWith("IPM.DistList", StringComparison.OrdinalIgnoreCase)) return OutlookItemKind.Contact;
        if (messageClass.StartsWith("IPM.Task", StringComparison.OrdinalIgnoreCase)) return OutlookItemKind.Task;
        if (messageClass.StartsWith("IPM.Activity", StringComparison.OrdinalIgnoreCase)) return OutlookItemKind.Journal;
        if (messageClass.StartsWith("IPM.StickyNote", StringComparison.OrdinalIgnoreCase)) return OutlookItemKind.Note;
        return OutlookItemKind.Message;
    }

    internal static string? GetString(IEnumerable<MapiProperty> properties, ushort id) {
        return properties.FirstOrDefault(property => property.PropertyId == id && property.Value is string)?.Value as string;
    }

    internal static int? GetInt(IEnumerable<MapiProperty> properties, ushort id) {
        return ConvertInt(properties.FirstOrDefault(property => property.PropertyId == id)?.Value);
    }

    internal static DateTimeOffset? GetDate(IEnumerable<MapiProperty> properties, ushort id) {
        return ConvertDate(properties.FirstOrDefault(property => property.PropertyId == id)?.Value);
    }

    internal static bool? GetBool(IEnumerable<MapiProperty> properties, ushort id) {
        return ConvertBool(properties.FirstOrDefault(property => property.PropertyId == id)?.Value);
    }

    internal static MapiProperty? GetNamed(IEnumerable<MapiProperty> properties, Guid set, uint localId) {
        return properties.FirstOrDefault(property => property.Name?.PropertySet == set && property.Name.LocalId == localId);
    }

    private static MapiProperty? GetNamedByName(IEnumerable<MapiProperty> properties, string name) {
        return properties.FirstOrDefault(property =>
            string.Equals(property.Name?.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    private static IEnumerable<string> GetNamedStrings(IEnumerable<MapiProperty> properties, Guid set,
        string name, uint legacyLocalId) {
        MapiProperty? property = properties.FirstOrDefault(item => item.Name?.PropertySet == set &&
            (string.Equals(item.Name.Name, name, StringComparison.OrdinalIgnoreCase) || item.Name.LocalId == legacyLocalId));
        if (property?.Value is string scalar) return new[] { scalar };
        if (property?.Value is object[] values) return values.OfType<string>();
        return Array.Empty<string>();
    }

    private static OutlookAppointment CreateAppointment(IEnumerable<MapiProperty> properties) {
        return new OutlookAppointment {
            Start = GetNamedDate(properties, PsetidAppointment, 0x820D) ?? GetNamedDate(properties, PsetidCommon, 0x8516),
            End = GetNamedDate(properties, PsetidAppointment, 0x820E) ?? GetNamedDate(properties, PsetidCommon, 0x8517),
            Location = GetNamedString(properties, PsetidAppointment, 0x8208),
            IsAllDay = GetNamedBool(properties, PsetidAppointment, 0x8215),
            BusyStatus = GetNamedInt(properties, PsetidAppointment, 0x8205),
            MeetingStatus = GetNamedInt(properties, PsetidAppointment, 0x8217),
            ResponseStatus = GetNamedInt(properties, PsetidAppointment, 0x8218),
            Sequence = GetNamedInt(properties, PsetidAppointment, 0x8201),
            DurationMinutes = GetNamedInt(properties, PsetidAppointment, 0x8213),
            AllAttendees = GetNamedString(properties, PsetidAppointment, 0x8238),
            RequiredAttendees = GetNamedString(properties, PsetidAppointment, 0x823B),
            OptionalAttendees = GetNamedString(properties, PsetidAppointment, 0x823C),
            NotAllowPropose = GetNamedBool(properties, PsetidAppointment, 0x825A),
            RecurrenceType = GetNamedInt(properties, PsetidAppointment, 0x8231),
            RecurrencePattern = GetNamedString(properties, PsetidAppointment, 0x8232),
            RecurrenceState = GetNamedBinary(properties, PsetidAppointment, 0x8216),
            IsRecurring = GetNamedBool(properties, PsetidAppointment, 0x8223),
            ClientIntentFlags = GetNamedInt(properties, PsetidCalendarAssistant, 0x0015),
            ReminderDeltaMinutes = GetNamedInt(properties, PsetidCommon, 0x8501),
            ReminderTime = GetNamedDate(properties, PsetidCommon, 0x8502),
            ReminderIsSet = GetNamedBool(properties, PsetidCommon, 0x8503),
            ReminderSignalTime = GetNamedDate(properties, PsetidCommon, 0x8560),
            TimeZoneStructure = GetNamedBinary(properties, PsetidAppointment, 0x8233),
            TimeZoneDescription = GetNamedString(properties, PsetidAppointment, 0x8234),
            StartTimeZoneDefinition = GetNamedBinary(properties, PsetidAppointment, 0x825E),
            EndTimeZoneDefinition = GetNamedBinary(properties, PsetidAppointment, 0x825F),
            RecurrenceTimeZoneDefinition = GetNamedBinary(properties, PsetidAppointment, 0x8260)
        };
    }

    private static OutlookContact CreateContact(IEnumerable<MapiProperty> properties) {
        var contact = new OutlookContact {
            DisplayName = GetString(properties, 0x3001),
            Prefix = GetString(properties, 0x3A45),
            Initials = GetString(properties, 0x3A0A),
            GivenName = GetString(properties, 0x3A06),
            MiddleName = GetString(properties, 0x3A44),
            Surname = GetString(properties, 0x3A11),
            Generation = GetString(properties, 0x3A05),
            CompanyName = GetString(properties, 0x3A16),
            JobTitle = GetString(properties, 0x3A17),
            Department = GetString(properties, 0x3A18),
            FileAs = GetNamedString(properties, PsetidAddress, 0x8005),
            NickName = GetString(properties, 0x3A4F),
            ManagerName = GetString(properties, 0x3A4E),
            AssistantName = GetString(properties, 0x3A30),
            SpouseName = GetString(properties, 0x3A48),
            Profession = GetString(properties, 0x3A46),
            Language = GetString(properties, 0x3A0C),
            Location = GetString(properties, 0x3A0D),
            OfficeLocation = GetString(properties, 0x3A19),
            Birthday = GetDate(properties, 0x3A42) ?? GetNamedDate(properties, PsetidAddress, 0x80DE),
            WeddingAnniversary = GetDate(properties, 0x3A41) ?? GetNamedDate(properties, PsetidAddress, 0x80DF),
            IsPrivate = GetNamedBool(properties, PsetidCommon, 0x8506) ?? GetNamedBool(properties, PsetidSharing, 0x8506),
            HasPicture = GetNamedBool(properties, PsetidAddress, 0x8015),
            InstantMessagingAddress = GetNamedString(properties, PsetidAddress, 0x8062),
            BusinessHomePage = GetString(properties, 0x3A51),
            PersonalHomePage = GetString(properties, 0x3A50),
            Html = GetNamedString(properties, PsetidAddress, 0x802B)
        };
        AddDelimited(contact.Children, GetString(properties, 0x3A58));

        PopulateAddress(contact.BusinessAddress, properties, null, 0x3A29, 0x3A27, 0x3A28, 0x3A2A, 0x3A26, 0x3A2B);
        PopulateAddress(contact.HomeAddress, properties, 0x801A, 0x3A5D, 0x3A59, 0x3A5C, 0x3A5B, 0x3A5A, 0x3A5E);
        PopulateAddress(contact.OtherAddress, properties, 0x801C, 0x3A63, 0x3A5F, 0x3A62, 0x3A61, 0x3A60, 0x3A64);
        contact.WorkAddress.Formatted = GetNamedString(properties, PsetidAddress, 0x801B);
        contact.WorkAddress.Street = GetNamedString(properties, PsetidAddress, 0x8045);
        contact.WorkAddress.City = GetNamedString(properties, PsetidAddress, 0x8046);
        contact.WorkAddress.StateOrProvince = GetNamedString(properties, PsetidAddress, 0x8047);
        contact.WorkAddress.PostalCode = GetNamedString(properties, PsetidAddress, 0x8048);
        contact.WorkAddress.Country = GetNamedString(properties, PsetidAddress, 0x8049);
        contact.WorkAddress.PostOfficeBox = GetNamedString(properties, PsetidAddress, 0x804A);
        contact.WorkAddress.CountryCode = GetNamedString(properties, PsetidAddress, 0x80DB);
        contact.BusinessAddress.Formatted = contact.WorkAddress.Formatted ?? GetString(properties, 0x3A15);

        contact.Phones.Business = GetString(properties, 0x3A08);
        contact.Phones.Business2 = GetString(properties, 0x3A1B);
        contact.Phones.Home = GetString(properties, 0x3A09);
        contact.Phones.Home2 = GetString(properties, 0x3A2F);
        contact.Phones.Mobile = GetString(properties, 0x3A1C);
        contact.Phones.Other = GetString(properties, 0x3A1F);
        contact.Phones.Primary = GetString(properties, 0x3A1A);
        contact.Phones.BusinessFax = GetString(properties, 0x3A24);
        contact.Phones.HomeFax = GetString(properties, 0x3A25);
        contact.Phones.PrimaryFax = GetString(properties, 0x3A23);
        contact.Phones.Assistant = GetString(properties, 0x3A2E);
        contact.Phones.CompanyMain = GetString(properties, 0x3A57);
        contact.Phones.Car = GetString(properties, 0x3A1E);
        contact.Phones.Radio = GetString(properties, 0x3A1D);
        contact.Phones.Pager = GetString(properties, 0x3A21);
        contact.Phones.Callback = GetString(properties, 0x3A02);
        contact.Phones.Telex = GetString(properties, 0x3A2C);
        contact.Phones.TextTelephone = GetString(properties, 0x3A4B);
        contact.Phones.Isdn = GetString(properties, 0x3A2D);

        PopulateEmail(contact.Email1, properties, 0x8080, 0x8082, 0x8083, 0x8084, 0x8085);
        PopulateEmail(contact.Email2, properties, 0x8090, 0x8092, 0x8093, 0x8094, 0x8095);
        PopulateEmail(contact.Email3, properties, 0x80A0, 0x80A2, 0x80A3, 0x80A4, 0x80A5);
        return contact;
    }

    private static OutlookTask CreateTask(IEnumerable<MapiProperty> properties) {
        var task = new OutlookTask {
            Start = GetNamedDate(properties, PsetidTask, 0x8104),
            Due = GetNamedDate(properties, PsetidTask, 0x8105),
            Status = GetNamedInt(properties, PsetidTask, 0x8101),
            PercentComplete = ConvertDouble(GetNamed(properties, PsetidTask, 0x8102)?.Value),
            IsComplete = ConvertBool(GetNamed(properties, PsetidTask, 0x811C)?.Value),
            Owner = GetNamedString(properties, PsetidTask, 0x811F),
            ActualEffort = ToMinutes(GetNamedInt(properties, PsetidTask, 0x8110)),
            EstimatedEffort = ToMinutes(GetNamedInt(properties, PsetidTask, 0x8111)),
            SendUpdates = GetNamedBool(properties, PsetidTask, 0x811B),
            SendStatusOnComplete = GetNamedBool(properties, PsetidTask, 0x8119),
            Ownership = GetNamedInt(properties, PsetidTask, 0x8129),
            AcceptanceState = GetNamedInt(properties, PsetidTask, 0x812A),
            Version = GetNamedInt(properties, PsetidTask, 0x8112),
            State = GetNamedInt(properties, PsetidTask, 0x8113),
            Assigner = GetNamedString(properties, PsetidTask, 0x8121),
            IsTeamTask = GetNamedBool(properties, PsetidTask, 0x8103),
            Ordinal = GetNamedInt(properties, PsetidTask, 0x8123),
            IsRecurring = GetNamedBool(properties, PsetidAppointment, 0x8223),
            ReminderDeltaMinutes = GetNamedInt(properties, PsetidCommon, 0x8501),
            ReminderTime = GetNamedDate(properties, PsetidCommon, 0x8502),
            ReminderIsSet = GetNamedBool(properties, PsetidCommon, 0x8503),
            ReminderSignalTime = GetNamedDate(properties, PsetidCommon, 0x8560),
            CommonStart = GetNamedDate(properties, PsetidCommon, 0x8516),
            CommonEnd = GetNamedDate(properties, PsetidCommon, 0x8517),
            Mode = GetNamedInt(properties, PsetidCommon, 0x8518),
            ToDoOrdinalDate = GetNamedDate(properties, PsetidCommon, 0x85A0),
            ToDoSubOrdinal = GetNamedString(properties, PsetidCommon, 0x85A1),
            BillingInformation = GetNamedString(properties, PsetidCommon, 0x8535),
            Mileage = GetNamedString(properties, PsetidCommon, 0x8534),
            CompletedAt = GetDate(properties, 0x1091)
        };
        AddStrings(task.Contacts, GetNamed(properties, PsetidCommon, 0x853A)?.Value);
        AddStrings(task.Companies, GetNamed(properties, PsetidCommon, 0x8539)?.Value);
        return task;
    }

    private static OutlookJournal CreateJournal(IEnumerable<MapiProperty> properties) {
        return new OutlookJournal {
            Start = GetNamedDate(properties, PsetidCommon, 0x8516),
            End = GetNamedDate(properties, PsetidCommon, 0x8517) ?? GetNamedDate(properties, PsetidLog, 0x8708),
            DurationMinutes = GetNamedInt(properties, PsetidLog, 0x8707),
            Type = GetNamedString(properties, PsetidLog, 0x8700),
            TypeDescription = GetNamedString(properties, PsetidLog, 0x8712),
            Flags = GetNamedInt(properties, PsetidLog, 0x870C),
            DocumentPrinted = GetNamedBool(properties, PsetidLog, 0x870E),
            DocumentSaved = GetNamedBool(properties, PsetidLog, 0x870F),
            DocumentRouted = GetNamedBool(properties, PsetidLog, 0x8710),
            DocumentPosted = GetNamedBool(properties, PsetidLog, 0x8711)
        };
    }

    private static OutlookNote CreateNote(IEnumerable<MapiProperty> properties) {
        return new OutlookNote {
            Color = GetNamedInt(properties, PsetidNote, 0x8B00),
            Width = GetNamedInt(properties, PsetidNote, 0x8B02),
            Height = GetNamedInt(properties, PsetidNote, 0x8B03),
            X = GetNamedInt(properties, PsetidNote, 0x8B04),
            Y = GetNamedInt(properties, PsetidNote, 0x8B05)
        };
    }

    private static void PopulateAddress(OutlookPostalAddress address, IEnumerable<MapiProperty> properties,
        uint? formattedNamedId, ushort streetId, ushort cityId, ushort stateId, ushort postalId, ushort countryId,
        ushort postOfficeBoxId) {
        if (formattedNamedId.HasValue) address.Formatted = GetNamedString(properties, PsetidAddress, formattedNamedId.Value);
        address.Street = GetString(properties, streetId);
        address.City = GetString(properties, cityId);
        address.StateOrProvince = GetString(properties, stateId);
        address.PostalCode = GetString(properties, postalId);
        address.Country = GetString(properties, countryId);
        address.PostOfficeBox = GetString(properties, postOfficeBoxId);
    }

    private static void PopulateEmail(OutlookContactEmailAddress email, IEnumerable<MapiProperty> properties,
        uint displayId, uint addressTypeId, uint addressId, uint originalDisplayId, uint entryId) {
        email.DisplayName = GetNamedString(properties, PsetidAddress, displayId);
        email.AddressType = GetNamedString(properties, PsetidAddress, addressTypeId);
        email.Address = GetNamedString(properties, PsetidAddress, addressId);
        email.OriginalDisplayName = GetNamedString(properties, PsetidAddress, originalDisplayId);
        email.OriginalEntryId = GetNamedBinary(properties, PsetidAddress, entryId);
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

    private static string? GetHtml(IEnumerable<MapiProperty> properties, MapiStringEncodingContext encoding,
        IList<EmailDiagnostic> diagnostics, string location) {
        MapiProperty? property = properties.FirstOrDefault(item => item.PropertyId == 0x1013);
        if (property?.Value is string text) return text;
        if (property?.Value is byte[] bytes) {
            return encoding.Decode(bytes, diagnostics, location).TrimEnd('\0');
        }
        return null;
    }

    private static string? GetRtf(IEnumerable<MapiProperty> properties, MsgParserState state, string location) {
        MapiProperty? property = properties.FirstOrDefault(item => item.PropertyId == 0x1009);
        if (!(property?.Value is byte[] compressed)) return null;
        if (!MapiCompressedRtfCodec.TryDecompress(compressed, state.Options.MaxDecodedPropertyBytes,
            state.Diagnostics, string.Concat(location, "/rtf"), state.CancellationToken, out byte[] rtfBytes)) return null;
        state.CountDecodedBytes(rtfBytes.Length);
        char[] characters = new char[rtfBytes.Length];
        for (int index = 0; index < rtfBytes.Length; index++) characters[index] = (char)rtfBytes[index];
        return new string(characters);
    }

    private static string? GetNamedString(IEnumerable<MapiProperty> properties, Guid set, uint localId) {
        return GetNamed(properties, set, localId)?.Value as string;
    }

    private static byte[]? GetNamedBinary(IEnumerable<MapiProperty> properties, Guid set, uint localId) {
        return GetNamed(properties, set, localId)?.Value as byte[];
    }

    private static int? GetNamedInt(IEnumerable<MapiProperty> properties, Guid set, uint localId) {
        return ConvertInt(GetNamed(properties, set, localId)?.Value);
    }

    private static bool? GetNamedBool(IEnumerable<MapiProperty> properties, Guid set, uint localId) {
        return ConvertBool(GetNamed(properties, set, localId)?.Value);
    }

    private static DateTimeOffset? GetNamedDate(IEnumerable<MapiProperty> properties, Guid set, uint localId) {
        return ConvertDate(GetNamed(properties, set, localId)?.Value);
    }

    private static int? ConvertInt(object? value) {
        if (value is int intValue) return intValue;
        if (value is short shortValue) return shortValue;
        if (value is long longValue && longValue >= int.MinValue && longValue <= int.MaxValue) return (int)longValue;
        return null;
    }

    private static double? ConvertDouble(object? value) {
        if (value is double doubleValue) return doubleValue;
        if (value is float floatValue) return floatValue;
        return null;
    }

    private static bool? ConvertBool(object? value) => value is bool boolean ? boolean : (bool?)null;

    private static DateTimeOffset? ConvertDate(object? value) {
        if (value is DateTimeOffset offset) return offset;
        if (value is DateTime date) return new DateTimeOffset(date);
        return null;
    }

    private static string? TrimAngle(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return value;
        return value!.Trim().Trim('<', '>');
    }
}
