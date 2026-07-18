using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class OutlookItemProjectionTests {
    [Fact]
    public void ReadsMsgKitNamedPropertyLayoutWithoutMappingDiagnostics() {
        var sender = new MsgKit.Sender("sender@example.com", "Sender");
        using var contact = new MsgKit.Contact(sender, "MsgKit contact") {
            GivenName = "Ada",
            SurName = "Lovelace",
            FileUnder = "Lovelace, Ada",
            Email1 = new MsgKit.Address("ada@example.com", "Ada Lovelace")
        };
        using MemoryStream stream = new MemoryStream();
        contact.Save(stream);

        EmailReadResult result = new EmailDocumentReader().Read(stream.ToArray());

        Assert.Equal("Ada", result.Document.Contact!.GivenName);
        Assert.Equal("Lovelace, Ada", result.Document.Contact.FileAs);
        Assert.Equal("ada@example.com", result.Document.Contact.Email1.Address);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MSG_NAMEID_GUID_INVALID" ||
            diagnostic.Code == "EMAIL_MSG_NAMEID_STRING_INVALID");
    }

    [Fact]
    public void RoundTripsAppointmentNamedPropertiesAndMsgReaderProjection() {
        DateTimeOffset start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero);
        var appointment = new OutlookAppointment {
            Start = start,
            End = start.AddHours(2),
            Location = "Room 1",
            IsAllDay = false,
            BusyStatus = 2,
            MeetingStatus = 1,
            ResponseStatus = 3,
            Sequence = 7,
            DurationMinutes = 120,
            AllAttendees = "Required; Optional",
            RequiredAttendees = "Required",
            OptionalAttendees = "Optional",
            NotAllowPropose = true,
            RecurrenceType = 2,
            RecurrencePattern = "weekly",
            RecurrenceState = new byte[] { 1, 3, 5 },
            IsRecurring = true,
            ClientIntentFlags = 32 | 512,
            ReminderIsSet = true,
            ReminderDeltaMinutes = 15,
            ReminderTime = start.AddMinutes(-15),
            ReminderSignalTime = start.AddMinutes(-15),
            TimeZoneDescription = "UTC",
            TimeZoneStructure = new byte[] { 2, 4 },
            StartTimeZoneDefinition = new byte[] { 6 },
            EndTimeZoneDefinition = new byte[] { 7 },
            RecurrenceTimeZoneDefinition = new byte[] { 8 }
        };
        EmailDocument source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            Subject = "Planning",
            Appointment = appointment
        };

        byte[] bytes = Write(source);
        EmailDocument result = Read(bytes);

        Assert.Equal(OutlookItemKind.Appointment, result.OutlookItemKind);
        Assert.Equal("IPM.Appointment", result.MessageClass);
        Assert.Equal(start, result.Appointment!.Start);
        Assert.Equal(start.AddHours(2), result.Appointment.End);
        Assert.Equal("Room 1", result.Appointment.Location);
        Assert.Equal(120, result.Appointment.DurationMinutes);
        Assert.Equal("Required", result.Appointment.RequiredAttendees);
        Assert.True(result.Appointment.NotAllowPropose);
        Assert.Equal(2, result.Appointment.RecurrenceType);
        Assert.Equal("weekly", result.Appointment.RecurrencePattern);
        Assert.Equal(new byte[] { 1, 3, 5 }, result.Appointment.RecurrenceState);
        Assert.Equal(32 | 512, result.Appointment.ClientIntentFlags);
        Assert.Equal("UTC", result.Appointment.TimeZoneDescription);

        using var oracle = OpenOracle(bytes);
        Assert.Equal("Room 1", oracle.Appointment!.Location);
        Assert.Equal(start, oracle.Appointment.Start);
        Assert.Equal(start.AddHours(2), oracle.Appointment.End);
        Assert.Equal("Required", oracle.Appointment.ToAttendees);
        Assert.True(oracle.Appointment.NotAllowPropose);
        Assert.Equal("weekly", oracle.Appointment.RecurrencePattern);
    }

    [Fact]
    public void RoundTripsCompleteContactProjectionAndMsgReaderProjection() {
        DateTimeOffset birthday = new DateTimeOffset(1815, 12, 10, 0, 0, 0, TimeSpan.Zero);
        var contact = new OutlookContact {
            DisplayName = "Ada Lovelace",
            Prefix = "Countess",
            Initials = "AAL",
            GivenName = "Ada",
            MiddleName = "Augusta",
            Surname = "Lovelace",
            Generation = "I",
            CompanyName = "Analytical",
            JobTitle = "Programmer",
            Department = "Research",
            FileAs = "Lovelace, Ada",
            NickName = "Ada",
            ManagerName = "Charles Babbage",
            AssistantName = "Assistant",
            SpouseName = "William",
            Profession = "Mathematician",
            Language = "English",
            Location = "London",
            OfficeLocation = "Engine room",
            Birthday = birthday,
            WeddingAnniversary = birthday.AddYears(20),
            IsPrivate = true,
            HasPicture = false,
            InstantMessagingAddress = "ada@example.com",
            BusinessHomePage = "https://example.com/business",
            PersonalHomePage = "https://example.com/ada",
            Html = "<p>Ada</p>"
        };
        contact.Children.Add("Byron");
        contact.Children.Add("Anne");
        contact.BusinessAddress.Street = "1 Engine Way";
        contact.BusinessAddress.City = "London";
        contact.BusinessAddress.StateOrProvince = "London";
        contact.BusinessAddress.PostalCode = "SW1";
        contact.BusinessAddress.Country = "United Kingdom";
        contact.BusinessAddress.PostOfficeBox = "42";
        contact.WorkAddress.Formatted = "1 Engine Way\nSW1 London\nUnited Kingdom";
        contact.WorkAddress.CountryCode = "GB";
        contact.HomeAddress.Street = "Home Street";
        contact.OtherAddress.City = "Paris";
        contact.Phones.Business = "100";
        contact.Phones.Business2 = "101";
        contact.Phones.Home = "200";
        contact.Phones.Mobile = "300";
        contact.Phones.BusinessFax = "400";
        contact.Phones.Assistant = "500";
        contact.Email1.Address = "ada@example.com";
        contact.Email1.DisplayName = "Ada Lovelace";
        contact.Email1.AddressType = "SMTP";
        contact.Email2.Address = "/O=EXAMPLE/CN=ADA";
        contact.Email2.AddressType = "EX";
        contact.Email3.Address = "ada@example.net";

        EmailDocument source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Contact,
            Subject = "Ada Lovelace",
            Contact = contact
        };
        byte[] bytes = Write(source);
        OutlookContact result = Read(bytes).Contact!;

        Assert.Equal("Ada", result.GivenName);
        Assert.Equal("Augusta", result.MiddleName);
        Assert.Equal("Research", result.Department);
        Assert.Equal(new[] { "Byron", "Anne" }, result.Children);
        Assert.Equal("1 Engine Way", result.BusinessAddress.Street);
        Assert.Equal("GB", result.WorkAddress.CountryCode);
        Assert.Equal("Home Street", result.HomeAddress.Street);
        Assert.Equal("Paris", result.OtherAddress.City);
        Assert.Equal("101", result.Phones.Business2);
        Assert.Equal("400", result.Phones.BusinessFax);
        Assert.Equal("ada@example.com", result.Email1.Address);
        Assert.Equal("/O=EXAMPLE/CN=ADA", result.Email2.Address);
        Assert.Equal("ada@example.net", result.Email3.Address);
        Assert.Equal("<p>Ada</p>", result.Html);

        using var oracle = OpenOracle(bytes);
        Assert.Equal("Ada Lovelace", oracle.Contact!.DisplayName);
        Assert.Equal("ada@example.com", oracle.Contact.Email1EmailAddress);
        Assert.Equal("/O=EXAMPLE/CN=ADA", oracle.Contact.Email2EmailAddress);
        Assert.Equal("1 Engine Way", oracle.Contact.BusinessAddressStreet);
        Assert.Equal("100", oracle.Contact.BusinessTelephoneNumber);
        Assert.Equal("300", oracle.Contact.CellularTelephoneNumber);
        Assert.Equal(birthday, oracle.Contact.Birthday);
    }

    [Fact]
    public void RoundTripsCompleteTaskProjectionAndMsgReaderProjection() {
        DateTimeOffset start = new DateTimeOffset(2026, 9, 2, 8, 0, 0, TimeSpan.Zero);
        var task = new OutlookTask {
            Start = start,
            Due = start.AddDays(2),
            Status = 1,
            PercentComplete = 0.5,
            IsComplete = false,
            Owner = "Ada",
            EstimatedEffort = TimeSpan.FromMinutes(90),
            ActualEffort = TimeSpan.FromMinutes(30),
            SendUpdates = true,
            SendStatusOnComplete = true,
            Ownership = 2,
            AcceptanceState = 1,
            Version = 4,
            State = 1,
            Assigner = "Charles",
            IsTeamTask = true,
            Ordinal = 3,
            IsRecurring = true,
            ReminderDeltaMinutes = 10,
            ReminderIsSet = true,
            ReminderTime = start.AddMinutes(-10),
            ReminderSignalTime = start.AddMinutes(-10),
            CommonStart = start,
            CommonEnd = start.AddDays(2),
            Mode = 1,
            ToDoOrdinalDate = start,
            ToDoSubOrdinal = "A",
            BillingInformation = "BILL-1",
            Mileage = "12 km",
            CompletedAt = start.AddDays(1)
        };
        task.Contacts.Add("Grace");
        task.Companies.Add("Analytical");
        EmailDocument source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Task,
            Subject = "Calculate",
            Task = task
        };
        byte[] bytes = Write(source);
        OutlookTask result = Read(bytes).Task!;

        Assert.Equal(start.AddDays(2), result.Due);
        Assert.Equal(0.5, result.PercentComplete);
        Assert.False(result.IsComplete);
        Assert.Equal("Ada", result.Owner);
        Assert.Equal(TimeSpan.FromMinutes(90), result.EstimatedEffort);
        Assert.Equal(TimeSpan.FromMinutes(30), result.ActualEffort);
        Assert.True(result.SendUpdates);
        Assert.Equal("Charles", result.Assigner);
        Assert.Equal(new[] { "Grace" }, result.Contacts);
        Assert.Equal(new[] { "Analytical" }, result.Companies);
        Assert.Equal("BILL-1", result.BillingInformation);
        Assert.Equal(start.AddDays(1), result.CompletedAt);

        using var oracle = OpenOracle(bytes);
        Assert.Equal(global::MsgReader.Outlook.MessageType.Task, oracle.Type);
        Assert.Equal(global::MsgReader.Outlook.TaskStatus.InProgress, oracle.Task!.Status);
        Assert.Equal("Ada", oracle.Task.Owner);
        Assert.False(oracle.Task.Complete);
        Assert.Equal("BILL-1", oracle.Task.BillingInformation);
        Assert.Equal("12 km", oracle.Task.Mileage);
    }

    [Fact]
    public void RoundTripsJournalAndNoteProjections() {
        DateTimeOffset now = new DateTimeOffset(2026, 9, 2, 8, 0, 0, TimeSpan.Zero);
        EmailDocument journal = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Journal,
            Journal = new OutlookJournal {
                Start = now,
                End = now.AddMinutes(30),
                DurationMinutes = 30,
                Type = "Phone call",
                TypeDescription = "Customer call",
                Flags = 3,
                DocumentPrinted = true,
                DocumentSaved = true,
                DocumentRouted = false,
                DocumentPosted = true
            }
        };
        byte[] journalBytes = Write(journal);
        OutlookJournal journalResult = Read(journalBytes).Journal!;
        Assert.Equal("Phone call", journalResult.Type);
        Assert.Equal("Customer call", journalResult.TypeDescription);
        Assert.Equal(30, journalResult.DurationMinutes);
        Assert.True(journalResult.DocumentPrinted);
        using (var oracle = OpenOracle(journalBytes)) {
            Assert.Equal(global::MsgReader.Outlook.MessageType.Journal, oracle.Type);
            Assert.Equal("Phone call", oracle.Log!.Type);
            Assert.Equal("Customer call", oracle.Log.TypeDescription);
            Assert.Equal(30, oracle.Log.Duration);
            Assert.True(oracle.Log.DocumentPrinted);
        }

        EmailDocument note = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Note,
            Note = new OutlookNote { Color = 3, Width = 500, Height = 300, X = 40, Y = 60 }
        };
        OutlookNote noteResult = Read(Write(note)).Note!;
        Assert.Equal(3, noteResult.Color);
        Assert.Equal(500, noteResult.Width);
        Assert.Equal(300, noteResult.Height);
        Assert.Equal(40, noteResult.X);
        Assert.Equal(60, noteResult.Y);
    }

    private static byte[] Write(EmailDocument source) =>
        new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);

    private static EmailDocument Read(byte[] bytes) {
        EmailReadResult result = new EmailDocumentReader().Read(bytes);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
        return result.Document;
    }

    private static global::MsgReader.Outlook.Storage.Message OpenOracle(byte[] bytes) =>
        new global::MsgReader.Outlook.Storage.Message(new MemoryStream(bytes), FileAccess.Read, true);
}
