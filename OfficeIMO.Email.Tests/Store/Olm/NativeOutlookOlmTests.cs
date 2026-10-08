using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;
using OfficeIMO.Email.Tests;

namespace OfficeIMO.Email.Store.Tests.Olm;

public sealed class NativeOutlookOlmTests {
    [Fact]
    public void NativeMailRetainsUnicodeHtmlCidAndExactAttachmentBytesThroughEml() {
        EmailStoreReadResult result = Read(File.ReadAllBytes(FixturePath));
        EmailDocument document = Assert.Single(Items(result), item => item.OutlookItemKind == OutlookItemKind.Message);
        Assert.Equal("OfficeIMO validation — Zażółć 日本語", document.Subject);
        Assert.Equal("synthetic@example.test", document.From!.Address);
        Assert.Contains("Zażółć", document.Body.Html!);
        Assert.Equal(new[] {
            "3d5b3a5d3cbfe3b002f15d1b73331b9ab4f528831c9631cc09e41e032d8ceb2f",
            "a31bc9e0d54f843f4a3c088c88f7542e3ca553a0d9f2ec5956bb1e3ae23621d5"
        }, document.Attachments.Select(attachment => Hash(attachment.Content!)).OrderBy(hash => hash));
        Assert.Contains(document.Attachments, attachment => attachment.IsInline && attachment.ContentId != null);
        byte[] eml = new EmailDocumentWriter(new EmailWriterOptions(EmailConversionLossPolicy.Warn)).ToBytes(document, EmailFileFormat.Eml);
        EmailDocument reopened = new EmailDocumentReader().Read(eml).Document;
        Assert.Equal(document.Body.Html, reopened.Body.Html);
        Assert.Equal(document.Attachments.Select(attachment => Hash(attachment.Content!)).OrderBy(hash => hash),
            reopened.Attachments.Select(attachment => Hash(attachment.Content!)).OrderBy(hash => hash));
    }

    [Fact]
    public void NativeWeeklyCalendarKeepsTenMinuteAlarmAndLocalClockAcrossDstThroughMsgAndEml() {
        EmailDocument document = Assert.Single(Items(Read(File.ReadAllBytes(FixturePath))),
            item => item.OutlookItemKind == OutlookItemKind.Appointment);
        OutlookAppointment appointment = document.Appointment!;
        Assert.Equal(10, appointment.ReminderDeltaMinutes);
        Assert.Equal(new DateTimeOffset(2026, 10, 18, 7, 0, 0, TimeSpan.Zero), appointment.Start);
        Assert.True(appointment.ReminderIsSet);
        Assert.Equal(appointment.Start!.Value.AddMinutes(-10), appointment.ReminderTime);
        AssertNativeRecurrence(appointment);

        var writer = new EmailDocumentWriter(new EmailWriterOptions(EmailConversionLossPolicy.Warn));
        byte[] msg = writer.ToBytes(document, EmailFileFormat.OutlookMsg);
        EmailDocument fromMsg = new EmailDocumentReader().Read(msg).Document;
        Assert.Equal(10, fromMsg.Appointment!.ReminderDeltaMinutes);
        AssertNativeRecurrence(fromMsg.Appointment);
        byte[] eml = writer.ToBytes(fromMsg, EmailFileFormat.Eml);
        EmailReadResult fromEml = new EmailDocumentReader().Read(eml);
        Assert.Equal(10, fromEml.Document.Appointment!.ReminderDeltaMinutes);
        AssertNativeRecurrence(fromEml.Document.Appointment);
        Assert.DoesNotContain(fromEml.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_ICALENDAR_TIMEZONE_PROJECTION_UNSUPPORTED");
    }

    [Fact]
    public void NativeContactPhotoFlagWithoutPayloadReportsExplicitSourceAndConversionLoss() {
        EmailStoreReadResult result = Read(File.ReadAllBytes(FixturePath));
        Assert.Equal(3, Items(result).Count());
        EmailDocument document = Assert.Single(Items(result), item => item.OutlookItemKind == OutlookItemKind.Contact);
        Assert.Equal("OfficeIMO native Zażółć 日本語 Validation", document.Contact!.DisplayName);
        Assert.Equal("officeimo-native-contact@example.test", document.Contact.Email1.Address);
        Assert.True(document.Contact.HasPicture);
        Assert.Empty(document.Attachments);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_CONTACT_PICTURE_UNAVAILABLE");
        using var output = new MemoryStream();
        EmailWriteResult conversion = new EmailDocumentWriter(new EmailWriterOptions(EmailConversionLossPolicy.Warn)).Write(document, output, EmailFileFormat.Eml);
        Assert.Contains(conversion.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_OLM_PROJECTION_INCOMPLETE" &&
            diagnostic.Message.IndexOf("picture", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    [Theory]
    [InlineData("reminder", "EMAIL_STORE_OLM_REMINDER_UNSUPPORTED")]
    [InlineData("pattern", "EMAIL_STORE_OLM_RECURRENCE_UNSUPPORTED")]
    [InlineData("interval", "EMAIL_STORE_OLM_RECURRENCE_UNSUPPORTED")]
    [InlineData("exception", "EMAIL_STORE_OLM_RECURRENCE_UNSUPPORTED")]
    [InlineData("timezone", "EMAIL_STORE_OLM_TIMEZONE_UNSUPPORTED")]
    public void UnrepresentableNativeCalendarFieldsRemainRawAndProduceLossDiagnostics(string variant, string code) {
        byte[] archive = MutateCalendar(item => {
            XElement recurrence = item.Element("OPFCalendarEventCopyRecurrence")!;
            switch (variant) {
                case "reminder": item.Element("OPFCalendarEventCopyReminderDelta")!.Value = "601"; break;
                case "pattern": recurrence.Element("OPFRecurrencePattern")!.Element("OPFRecurrencePatternType")!.Value = "OPFRecurrencePatternMonthly"; break;
                case "interval": recurrence.Element("OPFRecurrencePattern")!.Element("OPFRecurrencePatternInterval")!.Value = "1.5E0"; break;
                case "exception": recurrence.Add(new XElement("OPFRecurrenceExceptions", new XElement("exception", "modified occurrence"))); break;
                case "timezone": item.Element("OPFCalendarEventGetStartTimeZoneICSData")!.Value =
                    item.Element("OPFCalendarEventGetStartTimeZoneICSData")!.Value.Replace("BYHOUR=3", "BYHOUR=4"); break;
            }
        });
        EmailStoreReadResult result = Read(archive);
        EmailDocument document = Assert.Single(Items(result), item => item.OutlookItemKind == OutlookItemKind.Appointment);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == code);
        Assert.NotEmpty(Assert.IsType<string[]>(document.Properties["Olm:StructuredProperties"]));
        Assert.NotEmpty(Assert.IsType<string[]>(document.Properties["Olm:ProjectionIssues"]));
        if (variant == "reminder") Assert.Null(document.Appointment!.ReminderDeltaMinutes);
        else Assert.Null(document.Appointment!.Recurrence);
        using var output = new MemoryStream();
        EmailWriteResult conversion = new EmailDocumentWriter(new EmailWriterOptions(EmailConversionLossPolicy.Warn)).Write(document, output, EmailFileFormat.Eml);
        Assert.Contains(conversion.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_OLM_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData(false, EmailConversionLossPolicy.Block)]
    [InlineData(false, EmailConversionLossPolicy.Warn)]
    [InlineData(false, EmailConversionLossPolicy.Allow)]
    [InlineData(true, EmailConversionLossPolicy.Block)]
    [InlineData(true, EmailConversionLossPolicy.Warn)]
    [InlineData(true, EmailConversionLossPolicy.Allow)]
    public void SourceSemanticLossHonorsConversionPolicyBeforeWriting(bool contact, EmailConversionLossPolicy policy) {
        byte[] bytes = contact ? File.ReadAllBytes(FixturePath) : MutateCalendar(item =>
            item.Element("OPFCalendarEventCopyReminderDelta")!.Value = "601");
        EmailDocument document = Assert.Single(Items(Read(bytes)), item => item.OutlookItemKind ==
            (contact ? OutlookItemKind.Contact : OutlookItemKind.Appointment));
        using var output = new MemoryStream();
        EmailWriteResult result = new EmailDocumentWriter(new EmailWriterOptions(policy)).Write(document, output);
        EmailDiagnostic loss = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_OLM_PROJECTION_INCOMPLETE");
        Assert.Equal(policy == EmailConversionLossPolicy.Block ? EmailDiagnosticSeverity.Error :
            policy == EmailConversionLossPolicy.Warn ? EmailDiagnosticSeverity.Warning : EmailDiagnosticSeverity.Information, loss.Severity);
        Assert.Equal(policy == EmailConversionLossPolicy.Block, result.HasErrors);
        if (policy == EmailConversionLossPolicy.Block) Assert.Equal(0, output.Length);
        else Assert.True(output.Length > 0);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DstOverlapStartOrEndThatResolvesToAnotherInstantRemainsUnprojected(bool overlapEnd) {
        byte[] bytes = MutateCalendar(item => {
            item.Element("OPFCalendarEventCopyStartTime")!.Value = overlapEnd ? "2026-10-25T00:30:00Z" : "2026-10-25T01:30:00Z";
            item.Element("OPFCalendarEventCopyEndTime")!.Value = overlapEnd ? "2026-10-25T01:45:00Z" : "2026-10-25T02:30:00Z";
            XElement recurrence = item.Element("OPFCalendarEventCopyRecurrence")!;
            recurrence.Element("OPFRecurrenceCopyStartDate")!.Value = "2026-10-24T22:00:00Z";
            recurrence.Element("OPFRecurrenceCopyEndDate")!.Value = "2026-11-14T23:00:00Z";
        });
        EmailStoreReadResult result = Read(bytes);
        EmailDocument document = Assert.Single(Items(result), item => item.OutlookItemKind == OutlookItemKind.Appointment);
        Assert.Null(document.Appointment!.Recurrence);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_RECURRENCE_UNSUPPORTED" &&
            diagnostic.Message.IndexOf("ambiguous", StringComparison.OrdinalIgnoreCase) >= 0);
        using var output = new MemoryStream();
        Assert.True(new EmailDocumentWriter().Write(document, output).HasErrors);
        Assert.Equal(0, output.Length);
    }

    private static void AssertNativeRecurrence(OutlookAppointment appointment) {
        OutlookRecurrence recurrence = Assert.IsType<OutlookRecurrence>(appointment.Recurrence);
        Assert.Equal(new DateTime(2026, 10, 18, 9, 0, 0), recurrence.Start);
        Assert.Equal(TimeSpan.FromHours(1), recurrence.Duration);
        Assert.Equal(OutlookRecurrenceDays.Sunday, recurrence.DaysOfWeek);
        Assert.Equal(OutlookRecurrenceRangeKind.OccurrenceCount, recurrence.RangeKind);
        Assert.Equal(4, recurrence.OccurrenceCount);
        OutlookTimeZoneDefinition zone = Assert.IsType<OutlookTimeZoneDefinition>(appointment.RecurrenceTimeZone);
        Assert.Equal("Central European Standard Time", zone.KeyName);
        OutlookRecurrenceExpansionResult expanded = OutlookRecurrenceExpander.Expand(recurrence);
        Assert.False(expanded.Truncated);
        Assert.Equal(new[] { 18, 25, 1, 8 }, expanded.Occurrences.Select(value => value.Start.Day));
        Assert.All(expanded.Occurrences, value => Assert.Equal(9, value.Start.Hour));
        Assert.Equal(new[] { 7, 8, 8, 8 }, expanded.Occurrences.Select(value => value.ResolveStart(zone).UtcDateTime.Hour));
    }

    private static IEnumerable<EmailDocument> Items(EmailStoreReadResult result) =>
        result.Store.Folders.SelectMany(folder => folder.Items).Select(item => item.Document);

    private static EmailStoreReadResult Read(byte[] bytes) {
        using var source = new MemoryStream(bytes, writable: false);
        return new EmailStoreReader().Read(source, "native-outlook.olm");
    }

    private static byte[] MutateCalendar(Action<XElement> mutate) {
        using var input = ZipFile.OpenRead(FixturePath);
        using var bytes = new MemoryStream();
        using (var output = new ZipArchive(bytes, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in input.Entries) {
                using Stream source = entry.Open();
                using Stream destination = output.CreateEntry(entry.FullName).Open();
                if (entry.FullName.EndsWith("/Calendar.xml", StringComparison.Ordinal)) {
                    XDocument xml = XDocument.Load(source);
                    mutate(xml.Root!.Elements().Single());
                    xml.Save(destination);
                } else source.CopyTo(destination);
            }
        }
        return bytes.ToArray();
    }

    private static string FixturePath => Path.Combine(EmailTestRepository.FindRoot(), "OfficeIMO.Email.Tests", "Corpora",
        "native-apple-20261003", "outlook", "native-outlook.olm");
    private static string Hash(byte[] bytes) {
        using SHA256 algorithm = SHA256.Create();
        return string.Concat(algorithm.ComputeHash(bytes).Select(value => value.ToString("x2")));
    }
}
