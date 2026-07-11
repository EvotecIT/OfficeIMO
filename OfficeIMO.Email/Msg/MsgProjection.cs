namespace OfficeIMO.Email;

internal static class MsgProjection {
    internal static readonly Guid PsetidAppointment = new Guid("00062002-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidTask = new Guid("00062003-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidAddress = new Guid("00062004-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidCommon = new Guid("00062008-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidLog = new Guid("0006200A-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidNote = new Guid("0006200E-0000-0000-C000-000000000046");
    internal static readonly Guid PsPublicStrings = new Guid("00020329-0000-0000-C000-000000000046");

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

        string? transportHeaders = GetString(properties, 0x007D);
        if (!string.IsNullOrWhiteSpace(transportHeaders)) {
            byte[] bytes = Encoding.UTF8.GetBytes(string.Concat(transportHeaders, "\r\n\r\n"));
            var parsedHeaders = new List<EmailHeader>();
            MimeHeaderParser.Parse(bytes, 0, bytes.Length, state.Options, parsedHeaders, state.Diagnostics,
                string.Concat(location, "/transport-headers"));
            foreach (EmailHeader header in parsedHeaders) document.Headers.Add(header);
            document.From = document.From ?? MimeAddressParser.ParseOne(MimeHeaderParser.GetValue(parsedHeaders, "From"));
            document.Sender = document.Sender ?? MimeAddressParser.ParseOne(MimeHeaderParser.GetValue(parsedHeaders, "Sender"));
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
        metadata.CreatedDate = GetDate(properties, 0x3007);
        metadata.ModifiedDate = GetDate(properties, 0x3008);
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
            RecurrencePattern = GetNamedString(properties, PsetidAppointment, 0x8232),
            RecurrenceState = GetNamed(properties, PsetidAppointment, 0x8216)?.Value as byte[]
        };
    }

    private static OutlookContact CreateContact(IEnumerable<MapiProperty> properties) {
        return new OutlookContact {
            GivenName = GetString(properties, 0x3A06),
            Surname = GetString(properties, 0x3A11),
            CompanyName = GetString(properties, 0x3A16),
            JobTitle = GetString(properties, 0x3A17),
            BusinessPhone = GetString(properties, 0x3A08),
            HomePhone = GetString(properties, 0x3A09),
            MobilePhone = GetString(properties, 0x3A1C),
            FileAs = GetNamedString(properties, PsetidAddress, 0x8005),
            Email1Address = GetNamedString(properties, PsetidAddress, 0x8084)
        };
    }

    private static OutlookTask CreateTask(IEnumerable<MapiProperty> properties) {
        return new OutlookTask {
            Start = GetNamedDate(properties, PsetidTask, 0x8104),
            Due = GetNamedDate(properties, PsetidTask, 0x8105),
            Status = GetNamedInt(properties, PsetidTask, 0x8101),
            PercentComplete = ConvertDouble(GetNamed(properties, PsetidTask, 0x8102)?.Value),
            IsComplete = ConvertBool(GetNamed(properties, PsetidTask, 0x810F)?.Value),
            Owner = GetNamedString(properties, PsetidTask, 0x811C)
        };
    }

    private static OutlookJournal CreateJournal(IEnumerable<MapiProperty> properties) {
        return new OutlookJournal {
            Start = GetNamedDate(properties, PsetidCommon, 0x8516),
            End = GetNamedDate(properties, PsetidCommon, 0x8517),
            Type = GetNamedString(properties, PsetidLog, 0x8700)
        };
    }

    private static OutlookNote CreateNote(IEnumerable<MapiProperty> properties) {
        return new OutlookNote {
            Color = GetNamedInt(properties, PsetidNote, 0x8B00),
            Width = GetNamedInt(properties, PsetidNote, 0x8B02),
            Height = GetNamedInt(properties, PsetidNote, 0x8B03)
        };
    }

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
