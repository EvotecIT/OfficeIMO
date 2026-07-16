using OfficeIMO.Email;
using System.Globalization;
using System.Runtime.InteropServices;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class OutlookInteropTests {
    [Fact]
    public void ExchangesMailAppointmentContactAndTaskMsgFilesWithInstalledOutlookWhenEnabled() {
        if (!string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_OUTLOOK_INTEROP"), "1",
            StringComparison.Ordinal)) return;
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
        if (outlookType == null) return;

        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Email.Outlook." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        object? outlookObject = null;
        object? sessionObject = null;
        try {
            outlookObject = Activator.CreateInstance(outlookType);
            Assert.NotNull(outlookObject);
            dynamic outlook = outlookObject!;
            sessionObject = outlook.GetNamespace("MAPI");
            dynamic session = sessionObject!;

            ValidateOfficeImoFilesInOutlook(directory, session);
            ValidateOfficeImoStandardsFilesInOutlook(directory, session);
            ValidateOutlookFilesInOfficeImo(directory, outlook);
        } finally {
            ReleaseComObject(sessionObject);
            ReleaseComObject(outlookObject);
            try { Directory.Delete(directory, recursive: true); } catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static void ValidateOfficeImoStandardsFilesInOutlook(string directory, dynamic session) {
        DateTimeOffset start = new DateTimeOffset(2026, 10, 4, 9, 0, 0, TimeSpan.Zero);
        var appointment = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            MessageClass = "IPM.Appointment",
            Subject = "OfficeIMO Outlook iCalendar",
            Appointment = new OutlookAppointment {
                Start = start,
                End = start.AddHours(1),
                ReminderIsSet = true,
                ReminderDeltaMinutes = 15
            }
        };
        EmailDocument projectedAppointment = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(appointment, EmailFileFormat.Eml)).Document;
        EmailAttachment calendar = Assert.Single(projectedAppointment.Attachments,
            attachment => string.Equals(attachment.ContentType, "text/calendar", StringComparison.OrdinalIgnoreCase));
        string calendarPath = Path.Combine(directory, "officeimo.ics");
        File.WriteAllBytes(calendarPath, Assert.IsType<byte[]>(calendar.Content));

        var contact = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Contact,
            MessageClass = "IPM.Contact",
            Subject = "OfficeIMO Outlook vCard",
            Contact = new OutlookContact {
                DisplayName = "OfficeIMO Outlook vCard",
                GivenName = "OfficeIMO",
                Surname = "vCard"
            }
        };
        contact.Contact.Email1.Address = "officeimo.vcard@example.com";
        EmailDocument projectedContact = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(contact, EmailFileFormat.Eml)).Document;
        EmailAttachment vcard = Assert.Single(projectedContact.Attachments,
            attachment => string.Equals(attachment.ContentType, "text/vcard", StringComparison.OrdinalIgnoreCase));
        string vcardPath = Path.Combine(directory, "officeimo.vcf");
        File.WriteAllBytes(vcardPath, Assert.IsType<byte[]>(vcard.Content));

        object? appointmentObject = null;
        object? contactObject = null;
        try {
            appointmentObject = session.OpenSharedItem(calendarPath);
            dynamic outlookAppointment = appointmentObject!;
            Assert.Equal(appointment.Subject, (string)outlookAppointment.Subject);
            Assert.True((bool)outlookAppointment.ReminderSet);

            contactObject = session.OpenSharedItem(vcardPath);
            dynamic outlookContact = contactObject!;
            Assert.Equal(contact.Contact.DisplayName, (string)outlookContact.FullName);
            Assert.Equal(contact.Contact.Email1.Address, (string)outlookContact.Email1Address);
        } finally {
            ReleaseComObject(contactObject);
            ReleaseComObject(appointmentObject);
        }
    }

    private static void ValidateOfficeImoFilesInOutlook(string directory, dynamic session) {
        DateTimeOffset start = new DateTimeOffset(2026, 10, 4, 9, 0, 0, TimeSpan.Zero);
        var documents = new[] {
            new EmailDocument {
                Format = EmailFileFormat.OutlookMsg,
                OutlookItemKind = OutlookItemKind.Message,
                Subject = "OfficeIMO Outlook mail"
            },
            new EmailDocument {
                Format = EmailFileFormat.OutlookMsg,
                OutlookItemKind = OutlookItemKind.Appointment,
                Subject = "OfficeIMO Outlook appointment",
                Appointment = new OutlookAppointment { Start = start, End = start.AddHours(1), Location = "Room 42" }
            },
            new EmailDocument {
                Format = EmailFileFormat.OutlookMsg,
                OutlookItemKind = OutlookItemKind.Contact,
                Subject = "OfficeIMO Outlook contact",
                Contact = new OutlookContact { DisplayName = "OfficeIMO Contact", GivenName = "OfficeIMO", Surname = "Contact" }
            },
            new EmailDocument {
                Format = EmailFileFormat.OutlookMsg,
                OutlookItemKind = OutlookItemKind.Task,
                Subject = "OfficeIMO Outlook task",
                Task = new OutlookTask { Start = start, Due = start.AddDays(1), PercentComplete = 0.25 }
            }
        };

        for (int index = 0; index < documents.Length; index++) {
            string path = Path.Combine(directory, "officeimo-" + index.ToString(CultureInfo.InvariantCulture) + ".msg");
            File.WriteAllBytes(path, new EmailDocumentWriter().ToBytes(documents[index], EmailFileFormat.OutlookMsg));
            object? itemObject = null;
            try {
                itemObject = session.OpenSharedItem(path);
                dynamic item = itemObject!;
                Assert.Equal(documents[index].Subject, (string)item.Subject);
                if (documents[index].OutlookItemKind == OutlookItemKind.Appointment) {
                    Assert.Equal("Room 42", (string)item.Location);
                }
            } finally {
                ReleaseComObject(itemObject);
            }
        }
    }

    private static void ValidateOutlookFilesInOfficeImo(string directory, dynamic outlook) {
        string[] subjects = {
            "Outlook OfficeIMO mail", "Outlook OfficeIMO appointment",
            "Outlook OfficeIMO contact", "Outlook OfficeIMO task"
        };
        OutlookItemKind[] kinds = {
            OutlookItemKind.Message, OutlookItemKind.Appointment, OutlookItemKind.Contact, OutlookItemKind.Task
        };
        int[] outlookKinds = { 0, 1, 2, 3 };
        for (int index = 0; index < outlookKinds.Length; index++) {
            object? itemObject = null;
            string path = Path.Combine(directory, "outlook-" + index.ToString(CultureInfo.InvariantCulture) + ".msg");
            try {
                itemObject = outlook.CreateItem(outlookKinds[index]);
                dynamic item = itemObject!;
                item.Subject = subjects[index];
                if (kinds[index] == OutlookItemKind.Appointment) {
                    item.Start = new DateTime(2026, 11, 5, 10, 0, 0, DateTimeKind.Local);
                    item.End = new DateTime(2026, 11, 5, 11, 30, 0, DateTimeKind.Local);
                    item.Location = "Outlook Room";
                } else if (kinds[index] == OutlookItemKind.Contact) {
                    item.FullName = "Outlook Contact";
                    item.Email1Address = "outlook.contact@example.com";
                } else if (kinds[index] == OutlookItemKind.Task) {
                    item.StartDate = new DateTime(2026, 11, 5);
                    item.DueDate = new DateTime(2026, 11, 7);
                    item.PercentComplete = 50;
                } else {
                    item.Body = "Created by Outlook for OfficeIMO validation";
                }
                item.SaveAs(path, 9);
                item.Close(1);
            } finally {
                ReleaseComObject(itemObject);
            }

            EmailReadResult read = new EmailDocumentReader().Read(path);
            Assert.Equal(kinds[index], read.Document.OutlookItemKind);
            Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
            if (kinds[index] == OutlookItemKind.Appointment) Assert.Equal("Outlook Room", read.Document.Appointment!.Location);
            if (kinds[index] == OutlookItemKind.Contact) Assert.Equal("outlook.contact@example.com", read.Document.Contact!.Email1.Address);
        }
    }

    private static void ReleaseComObject(object? value) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) && value != null && Marshal.IsComObject(value)) {
            Marshal.FinalReleaseComObject(value);
        }
    }
}
