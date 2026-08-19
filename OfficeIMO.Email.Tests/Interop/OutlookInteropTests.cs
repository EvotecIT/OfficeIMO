using OfficeIMO.Email;
using System.Globalization;
using System.Runtime.InteropServices;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class OutlookInteropTests {
    private const string NamedInteropProperty =
        "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/OfficeIMOInterop";

    [EmailArtifactOutlookInteropFact]
    public void ExchangesMailAppointmentContactAndTaskMsgFilesWithInstalledOutlookWhenEnabled() {
#pragma warning disable CA1416
        Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
#pragma warning restore CA1416
        Assert.NotNull(outlookType);

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
        string linkedPath = Path.Combine(directory, "outlook-linked.txt");
        File.WriteAllText(linkedPath, "Outlook independently produced attachment", Encoding.UTF8);
        string embeddedPath = CreateOutlookEmbeddedMessage(directory, outlook);
        for (int index = 0; index < outlookKinds.Length; index++) {
            object? itemObject = null;
            string path = Path.Combine(directory, "outlook-" + index.ToString(CultureInfo.InvariantCulture) + ".msg");
            try {
                itemObject = outlook.CreateItem(outlookKinds[index]);
                dynamic item = itemObject!;
                item.Subject = subjects[index];
                item.PropertyAccessor.SetProperty(NamedInteropProperty, "Outlook named evidence " + index);
                if (kinds[index] == OutlookItemKind.Appointment) {
                    item.Start = new DateTime(2026, 11, 5, 10, 0, 0, DateTimeKind.Local);
                    item.End = new DateTime(2026, 11, 5, 11, 30, 0, DateTimeKind.Local);
                    item.Location = "Outlook Room";
                    object? recurrenceObject = null;
                    try {
                        recurrenceObject = item.GetRecurrencePattern();
                        dynamic recurrence = recurrenceObject!;
                        recurrence.RecurrenceType = 1; // olRecursWeekly
                        recurrence.Interval = 1;
                        recurrence.DayOfWeekMask = 32; // Thursday
                        recurrence.PatternStartDate = new DateTime(2026, 11, 5);
                        recurrence.NoEndDate = false;
                        recurrence.Occurrences = 4;
                    } finally {
                        ReleaseComObject(recurrenceObject);
                    }
                } else if (kinds[index] == OutlookItemKind.Contact) {
                    item.FullName = "Outlook Contact";
                    item.Email1Address = "outlook.contact@example.com";
                } else if (kinds[index] == OutlookItemKind.Task) {
                    item.StartDate = new DateTime(2026, 11, 5);
                    item.DueDate = new DateTime(2026, 11, 7);
                    item.PercentComplete = 50;
                } else {
                    item.Body = "Created by Outlook for OfficeIMO validation";
                    object? recipientObject = null;
                    object? attachmentsObject = null;
                    object? byValueObject = null;
                    object? byReferenceObject = null;
                    object? embeddedObject = null;
                    try {
                        recipientObject = item.Recipients.Add("officeimo.smtp@example.com");
                        dynamic recipient = recipientObject!;
                        recipient.Type = 1;
                        attachmentsObject = item.Attachments;
                        dynamic attachments = attachmentsObject!;
                        byValueObject = attachments.Add(linkedPath, 1, Type.Missing, "Outlook by-value evidence");
                        byReferenceObject = attachments.Add(linkedPath, 4, Type.Missing,
                            "Outlook by-reference evidence");
                        embeddedObject = attachments.Add(embeddedPath, 5, Type.Missing, "Outlook embedded evidence");
                    } finally {
                        ReleaseComObject(embeddedObject);
                        ReleaseComObject(byReferenceObject);
                        ReleaseComObject(byValueObject);
                        ReleaseComObject(attachmentsObject);
                        ReleaseComObject(recipientObject);
                    }
                }
                item.SaveAs(path, 9);
                item.Close(1);
            } finally {
                ReleaseComObject(itemObject);
            }

            using EmailReadResult read = new EmailDocumentReader().Read(path);
            Assert.Equal(kinds[index], read.Document.OutlookItemKind);
            Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
            Assert.Contains(read.Document.MapiProperties, property =>
                string.Equals(property.Name?.Name, "OfficeIMOInterop", StringComparison.OrdinalIgnoreCase) &&
                Equals(property.Value, "Outlook named evidence " + index));
            if (kinds[index] == OutlookItemKind.Appointment) Assert.Equal("Outlook Room", read.Document.Appointment!.Location);
            if (kinds[index] == OutlookItemKind.Appointment) {
                Assert.True(read.Document.Appointment!.IsRecurring);
                Assert.NotNull(read.Document.Appointment.RecurrenceState);
                Assert.True(read.Document.Appointment.StartTimeZoneDefinition != null ||
                    read.Document.Appointment.RecurrenceTimeZoneDefinition != null ||
                    read.Document.Appointment.TimeZoneStructure != null);
            }
            if (kinds[index] == OutlookItemKind.Contact) Assert.Equal("outlook.contact@example.com", read.Document.Contact!.Email1.Address);
            if (kinds[index] == OutlookItemKind.Message) {
                Assert.Contains(read.Document.MapiProperties, property => property.Name == null &&
                    MapiKnownProperties.Find(property) == null);
                Assert.Contains(read.Document.Recipients, recipient =>
                    string.Equals(recipient.Address?.AddressType, "SMTP", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(recipient.Address?.Address, "officeimo.smtp@example.com", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(read.Document.Attachments, attachment => attachment.MapiAttachMethod == 1 &&
                    attachment.FileName == "outlook-linked.txt");
                Assert.Contains(read.Document.Attachments, attachment =>
                    attachment.FileName == "outlook-linked.txt" &&
                    (attachment.MapiAttachMethod == 1 || attachment.MapiAttachMethod == 4));
                Assert.Contains(read.Document.Attachments, attachment => attachment.MapiAttachMethod == 5 &&
                    attachment.EmbeddedDocument?.Subject == "Outlook embedded evidence");
            }

            ValidateOfficeImoSemanticReopen(directory, index, read.Document, outlook);
        }
    }

    private static string CreateOutlookEmbeddedMessage(string directory, dynamic outlook) {
        string path = Path.Combine(directory, "outlook-embedded.msg");
        object? itemObject = null;
        try {
            itemObject = outlook.CreateItem(0);
            dynamic item = itemObject!;
            item.Subject = "Outlook embedded evidence";
            item.Body = "Nested Outlook message body";
            item.SaveAs(path, 9);
            item.Close(1);
            return path;
        } finally {
            ReleaseComObject(itemObject);
        }
    }

    private static void ValidateOfficeImoSemanticReopen(string directory, int index,
        EmailDocument document, dynamic outlook) {
        string path = Path.Combine(directory,
            "officeimo-reopen-" + index.ToString(CultureInfo.InvariantCulture) + ".msg");
        File.WriteAllBytes(path, new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg));
        using (EmailReadResult officeImoReopen = new EmailDocumentReader().Read(path)) {
            MapiProperty[] unknown = document.MapiProperties.Where(property => property.Name == null &&
                MapiKnownProperties.Find(property) == null).ToArray();
            foreach (MapiProperty property in unknown) {
                Assert.Contains(officeImoReopen.Document.MapiProperties, candidate =>
                    candidate.PropertyTag == property.PropertyTag);
            }
        }
        object? reopenedObject = null;
        try {
            reopenedObject = outlook.Session.OpenSharedItem(path);
            dynamic reopened = reopenedObject!;
            Assert.Equal(document.Subject, Convert.ToString(reopened.Subject));
            Assert.Equal("Outlook named evidence " + index,
                Convert.ToString(reopened.PropertyAccessor.GetProperty(NamedInteropProperty)));
            if (document.OutlookItemKind == OutlookItemKind.Appointment) {
                Assert.True((bool)reopened.IsRecurring);
                Assert.Equal("Outlook Room", Convert.ToString(reopened.Location));
            }
            if (document.OutlookItemKind == OutlookItemKind.Message) {
                Assert.Equal(document.Attachments.Count, (int)reopened.Attachments.Count);
            }
            reopened.Close(1);
        } finally {
            ReleaseComObject(reopenedObject);
        }
    }

    private static void ReleaseComObject(object? value) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) && value != null && Marshal.IsComObject(value)) {
            Marshal.FinalReleaseComObject(value);
        }
    }
}
