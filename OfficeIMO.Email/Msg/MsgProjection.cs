namespace OfficeIMO.Email;

internal static class MsgProjection {
    internal static readonly Guid PsetidAppointment = new Guid("00062002-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidTask = new Guid("00062003-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidAddress = new Guid("00062004-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidCommon = new Guid("00062008-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidLog = new Guid("0006200A-0000-0000-C000-000000000046");
    internal static readonly Guid PsetidNote = new Guid("0006200E-0000-0000-C000-000000000046");

    internal static void Apply(EmailDocument document, MsgParserState state, string location) {
        IList<MapiProperty> properties = document.MapiProperties;
        document.MessageClass = GetString(properties, 0x001A) ?? "IPM.Note";
        document.OutlookItemKind = Classify(document.MessageClass);
        document.Subject = GetString(properties, 0x0037) ?? GetString(properties, 0x0E1D);
        document.MessageId = TrimAngle(GetString(properties, 0x1035));
        document.Date = GetDate(properties, 0x0039) ?? GetDate(properties, 0x3007);
        document.ReceivedDate = GetDate(properties, 0x0E06);
        document.Body.Text = GetString(properties, 0x1000);
        document.Body.Html = GetHtml(properties, state.Diagnostics, location);
        document.Body.Rtf = GetRtf(properties, state, location);

        string? fromName = GetString(properties, 0x0042);
        string? fromAddress = GetString(properties, 0x5D02) ?? GetString(properties, 0x0065);
        if (fromName != null || fromAddress != null) document.From = new EmailAddress(fromAddress, fromName);
        string? senderName = GetString(properties, 0x0C1A);
        string? senderAddress = GetString(properties, 0x5D01) ?? GetString(properties, 0x0C1F);
        if (senderName != null || senderAddress != null) document.Sender = new EmailAddress(senderAddress, senderName);

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

        ApplyTyped(document);
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

    internal static MapiProperty? GetNamed(IEnumerable<MapiProperty> properties, Guid set, uint localId) {
        return properties.FirstOrDefault(property => property.Name?.PropertySet == set && property.Name.LocalId == localId);
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

    private static string? GetHtml(IEnumerable<MapiProperty> properties, IList<EmailDiagnostic> diagnostics, string location) {
        MapiProperty? property = properties.FirstOrDefault(item => item.PropertyId == 0x1013);
        if (property?.Value is string text) return text;
        if (property?.Value is byte[] bytes) {
            int codePage = GetInt(properties, 0x3FDE) ?? GetInt(properties, 0x3FFD) ?? 65001;
            string charset = codePage == 65001 ? "utf-8" : codePage == 20127 ? "us-ascii" :
                codePage == 28591 ? "iso-8859-1" : string.Concat("windows-", codePage.ToString(CultureInfo.InvariantCulture));
            return MimeTextCodec.DecodeText(bytes, charset, diagnostics, location).TrimEnd('\0');
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
