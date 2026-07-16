using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class MsgKitExampleCompatibilityTests {
    [Fact]
    public void ProcessesMsgKitGeneratedMailAppointmentContactAndTaskAcrossEveryFormat() {
        MsgKitArtifact[] artifacts = {
            CreateMail(), CreateAppointment(), CreateContact(), CreateTask()
        };
        var reader = new EmailDocumentReader();
        var writer = new EmailDocumentWriter(new EmailWriterOptions(
            conversionLossPolicy: EmailConversionLossPolicy.Warn));

        foreach (MsgKitArtifact artifact in artifacts) {
            EmailReadResult source = reader.Read(artifact.Bytes);
            Assert.DoesNotContain(source.Diagnostics,
                diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
            Assert.Equal(artifact.Kind, source.Document.OutlookItemKind);
            Assert.Equal(artifact.Subject, source.Document.Subject);

            foreach (EmailFileFormat format in new[] {
                EmailFileFormat.OutlookMsg, EmailFileFormat.Eml, EmailFileFormat.Tnef
            }) {
                byte[] rewritten = writer.ToBytes(source.Document, format);
                EmailReadResult reopened = reader.Read(rewritten);
                Assert.DoesNotContain(reopened.Diagnostics,
                    diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
                Assert.Equal(format, reopened.Document.Format);
                Assert.Equal(artifact.Kind, reopened.Document.OutlookItemKind);
                Assert.Equal(artifact.Subject, reopened.Document.Subject);
                if (format == EmailFileFormat.OutlookMsg) {
                    using var oracle = new global::MsgReader.Outlook.Storage.Message(
                        new MemoryStream(rewritten), FileAccess.Read, true);
                    Assert.Equal(artifact.Subject, oracle.Subject);
                } else if (format == EmailFileFormat.Eml) {
                    using var stream = new MemoryStream(rewritten);
                    using MimeKit.MimeMessage oracle = MimeKit.MimeMessage.Load(stream);
                    Assert.Equal(artifact.Subject, oracle.Subject);
                }
            }
        }
    }

    private static MsgKitArtifact CreateMail() {
        using var message = new MsgKit.Email(new MsgKit.Sender("sender@example.com", "Sender"),
            "MsgKit generated mail", draft: true, readReceipt: true) {
            BodyText = "MsgKit plain body",
            BodyHtml = "<html><body><b>MsgKit HTML body</b></body></html>",
            SentOn = new DateTime(2026, 7, 15, 10, 0, 0, DateTimeKind.Utc)
        };
        message.Recipients.AddTo("to@example.com", "To Person");
        message.Recipients.AddCc("cc@example.com", "Cc Person");
        message.Attachments.Add(new MemoryStream(Encoding.UTF8.GetBytes("attachment")), "sample.txt");
        return Save(message, OutlookItemKind.Message, "MsgKit generated mail");
    }

    private static MsgKitArtifact CreateAppointment() {
        using var appointment = new MsgKit.Appointment(new MsgKit.Sender("sender@example.com", "Sender"),
            new MsgKit.Representing("organizer@example.com", "Organizer"), "MsgKit generated appointment") {
            Location = "Neverland",
            MeetingStart = new DateTime(2026, 8, 1, 9, 0, 0, DateTimeKind.Utc),
            MeetingEnd = new DateTime(2026, 8, 1, 10, 30, 0, DateTimeKind.Utc),
            BodyText = "Appointment body"
        };
        appointment.Recipients.AddTo("attendee@example.com", "Attendee");
        return Save(appointment, OutlookItemKind.Appointment, "MsgKit generated appointment");
    }

    private static MsgKitArtifact CreateContact() {
        using var contact = new MsgKit.Contact(new MsgKit.Sender("sender@example.com", "Sender"),
            "MsgKit generated contact") {
            FileUnder = "Lovelace, Ada",
            GivenName = "Ada",
            SurName = "Lovelace",
            DepartmentName = "Research",
            Title = "Mathematician",
            Email1 = new MsgKit.Address("ada@example.com", "Ada Lovelace"),
            MobileTelephoneNumber = "+44 7000 000000",
            Business = new MsgKit.ContactBusiness {
                City = "London",
                Country = "United Kingdom",
                Street = "1 Engine Way"
            }
        };
        return Save(contact, OutlookItemKind.Contact, "MsgKit generated contact");
    }

    private static MsgKitArtifact CreateTask() {
        using var task = new MsgKit.Task(new MsgKit.Sender("sender@example.com", "Sender"),
            new MsgKit.Representing("owner@example.com", "Owner"), "MsgKit generated task") {
            Status = MsgKit.Enums.TaskStatus.InProgress,
            PercentageComplete = 0.25,
            StartDate = new DateTime(2026, 9, 1),
            DueDate = new DateTime(2026, 9, 3),
            ReminderDelta = 30,
            ReminderTime = new DateTime(2026, 9, 1, 8, 30, 0, DateTimeKind.Utc),
            BodyText = "Task body"
        };
        task.Attachments.Add(new MemoryStream(Encoding.UTF8.GetBytes("task attachment")), "task.txt");
        return Save(task, OutlookItemKind.Task, "MsgKit generated task");
    }

    private static MsgKitArtifact Save(MsgKit.Email item, OutlookItemKind kind, string subject) {
        using var stream = new MemoryStream();
        if (item is MsgKit.Appointment appointment) appointment.Save(stream);
        else if (item is MsgKit.Contact contact) contact.Save(stream);
        else if (item is MsgKit.Task task) task.Save(stream);
        else item.Save(stream);
        return new MsgKitArtifact(stream.ToArray(), kind, subject);
    }

    private sealed class MsgKitArtifact {
        internal MsgKitArtifact(byte[] bytes, OutlookItemKind kind, string subject) {
            Bytes = bytes;
            Kind = kind;
            Subject = subject;
        }
        internal byte[] Bytes { get; }
        internal OutlookItemKind Kind { get; }
        internal string Subject { get; }
    }
}
