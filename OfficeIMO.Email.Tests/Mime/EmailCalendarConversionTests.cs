using MimeKit;
using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailCalendarConversionTests {
    [Fact]
    public void ConvertsAppointmentThroughStandardsBasedEmlWithoutLosingCoreSemantics() {
        DateTimeOffset start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.FromHours(2));
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            MessageClass = "IPM.Schedule.Meeting.Request",
            Subject = "Planning",
            MessageId = "planning@example.com",
            From = new EmailAddress("organizer@example.com", "Organizer"),
            Appointment = new OutlookAppointment {
                Start = start,
                End = start.AddHours(2),
                Location = "Room 1",
                BusyStatus = 2,
                Sequence = 3,
                ReminderIsSet = true,
                ReminderDeltaMinutes = 20,
                ReminderTime = start.AddMinutes(-20),
                ReminderSignalTime = start.AddMinutes(-20)
            }
        };
        source.Body.Text = "Agenda";
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("attendee@example.com", "Attendee")));

        byte[] eml = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml);
        EmailReadResult read = new EmailDocumentReader().Read(eml);
        using var oracleStream = new MemoryStream(eml);
        MimeMessage oracle = MimeMessage.Load(oracleStream);

        Assert.Contains(oracle.BodyParts.OfType<MimePart>(), part => part.ContentType.MimeType == "text/calendar");
        Assert.Equal(OutlookItemKind.Appointment, read.Document.OutlookItemKind);
        Assert.Equal("Planning", read.Document.Subject);
        Assert.Equal(start.UtcDateTime, read.Document.Appointment!.Start!.Value.UtcDateTime);
        Assert.Equal(start.AddHours(2).UtcDateTime, read.Document.Appointment.End!.Value.UtcDateTime);
        Assert.Equal("Room 1", read.Document.Appointment.Location);
        Assert.Equal(3, read.Document.Appointment.Sequence);
        Assert.True(read.Document.Appointment.ReminderIsSet);
        Assert.Equal(20, read.Document.Appointment.ReminderDeltaMinutes);
        Assert.Equal(start.AddMinutes(-20).UtcDateTime,
            read.Document.Appointment.ReminderSignalTime!.Value.UtcDateTime);
        Assert.Equal("attendee@example.com", Assert.Single(read.Document.Recipients).Address.Address);
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void ConvertsTaskThroughVtodo() {
        DateTimeOffset start = new DateTimeOffset(2026, 9, 3, 8, 0, 0, TimeSpan.Zero);
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Task,
            MessageClass = "IPM.Task",
            Subject = "Ship release",
            Task = new OutlookTask {
                Start = start,
                Due = start.AddDays(2),
                PercentComplete = 0.5,
                Status = 1,
                Owner = "Release Manager",
                EstimatedEffort = TimeSpan.FromHours(6),
                ActualEffort = TimeSpan.FromHours(2),
                ReminderIsSet = true,
                ReminderDeltaMinutes = 30,
                BillingInformation = "Internal",
                Mileage = "12 km"
            }
        };
        source.Task.Contacts.Add("Ada Lovelace");
        source.Task.Companies.Add("Analytical Engines");

        EmailDocument result = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml)).Document;

        Assert.Equal(OutlookItemKind.Task, result.OutlookItemKind);
        Assert.Equal("Ship release", result.Subject);
        Assert.Equal(start, result.Task!.Start);
        Assert.Equal(start.AddDays(2), result.Task.Due);
        Assert.Equal(0.5, result.Task.PercentComplete);
        Assert.Equal(1, result.Task.Status);
        Assert.Equal("Release Manager", result.Task.Owner);
        Assert.Equal(TimeSpan.FromHours(6), result.Task.EstimatedEffort);
        Assert.Equal(TimeSpan.FromHours(2), result.Task.ActualEffort);
        Assert.True(result.Task.ReminderIsSet);
        Assert.Equal(30, result.Task.ReminderDeltaMinutes);
        Assert.Equal("Internal", result.Task.BillingInformation);
        Assert.Equal("12 km", result.Task.Mileage);
        Assert.Equal("Ada Lovelace", Assert.Single(result.Task.Contacts));
        Assert.Equal("Analytical Engines", Assert.Single(result.Task.Companies));
    }

    [Fact]
    public void BlocksOpaqueOutlookRecurrenceInsteadOfDroppingIt() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            Appointment = new OutlookAppointment {
                Start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero),
                End = new DateTimeOffset(2026, 8, 1, 11, 0, 0, TimeSpan.Zero),
                RecurrenceState = new byte[] { 1, 2, 3 }
            }
        };

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(source, EmailFileFormat.Eml);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_OPAQUE_RECURRENCE");
    }

    [Fact]
    public void BlocksAttendeeDisplayTextWhenNoCalendarAddressExists() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            Appointment = new OutlookAppointment {
                Start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero),
                RequiredAttendees = "Person without an address"
            }
        };

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(source, EmailFileFormat.Eml);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_ATTENDEE_ADDRESS_REQUIRED");
    }

    [Fact]
    public void KeepsRecurringIcalendarInEmlButBlocksIncompleteStoreProjection() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Subject: Recurring\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:r@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nDTEND:20260801T110000Z\r\n" +
            "RRULE:FREQ=WEEKLY;COUNT=4\r\nSUMMARY:Recurring\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;
        var writer = new EmailDocumentWriter();

        EmailConversionReport emlReport = writer.AnalyzeConversion(document, EmailFileFormat.Eml);
        EmailConversionReport msgReport = writer.AnalyzeConversion(document, EmailFileFormat.OutlookMsg);

        Assert.True(emlReport.CanWrite);
        Assert.NotEmpty(writer.ToBytes(document, EmailFileFormat.Eml));
        Assert.False(msgReport.CanWrite);
        Assert.Contains(msgReport.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }
}
