using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class OutlookItemProjectionTests {
    [Fact]
    public void RoundTripsAppointmentNamedProperties() {
        DateTimeOffset start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero);
        EmailDocument source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            Subject = "Planning",
            Appointment = new OutlookAppointment {
                Start = start,
                End = start.AddHours(2),
                Location = "Room 1",
                IsAllDay = false,
                BusyStatus = 2,
                MeetingStatus = 1,
                ResponseStatus = 3,
                RecurrencePattern = "weekly",
                RecurrenceState = new byte[] { 1, 3, 5 }
            }
        };

        EmailDocument result = RoundTrip(source);

        Assert.Equal(OutlookItemKind.Appointment, result.OutlookItemKind);
        Assert.Equal("IPM.Appointment", result.MessageClass);
        Assert.Equal(source.Appointment.Start, result.Appointment!.Start);
        Assert.Equal(source.Appointment.End, result.Appointment.End);
        Assert.Equal("Room 1", result.Appointment.Location);
        Assert.Equal("weekly", result.Appointment.RecurrencePattern);
        Assert.Equal(new byte[] { 1, 3, 5 }, result.Appointment.RecurrenceState);
    }

    [Fact]
    public void RoundTripsContactTaskJournalAndNoteProjections() {
        EmailDocument contact = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Contact,
            Contact = new OutlookContact {
                GivenName = "Ada", Surname = "Lovelace", CompanyName = "Analytical",
                JobTitle = "Programmer", BusinessPhone = "1", HomePhone = "2", MobilePhone = "3",
                FileAs = "Lovelace, Ada", Email1Address = "ada@example.com"
            }
        };
        EmailDocument contactResult = RoundTrip(contact);
        Assert.Equal("Ada", contactResult.Contact!.GivenName);
        Assert.Equal("ada@example.com", contactResult.Contact.Email1Address);

        DateTimeOffset now = new DateTimeOffset(2026, 9, 2, 8, 0, 0, TimeSpan.Zero);
        EmailDocument task = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Task,
            Task = new OutlookTask { Start = now, Due = now.AddDays(2), Status = 1, PercentComplete = 0.5, IsComplete = false, Owner = "Ada" }
        };
        EmailDocument taskResult = RoundTrip(task);
        Assert.Equal(now.AddDays(2), taskResult.Task!.Due);
        Assert.Equal(0.5, taskResult.Task.PercentComplete);

        EmailDocument journal = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Journal,
            Journal = new OutlookJournal { Start = now, End = now.AddMinutes(30), Type = "Phone call" }
        };
        Assert.Equal("Phone call", RoundTrip(journal).Journal!.Type);

        EmailDocument note = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Note,
            Note = new OutlookNote { Color = 3, Width = 500, Height = 300 }
        };
        OutlookNote noteResult = RoundTrip(note).Note!;
        Assert.Equal(3, noteResult.Color);
        Assert.Equal(500, noteResult.Width);
        Assert.Equal(300, noteResult.Height);
    }

    private static EmailDocument RoundTrip(EmailDocument source) {
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);
        EmailReadResult result = new EmailDocumentReader().Read(bytes);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
        return result.Document;
    }
}
