using MimeKit;
using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailConversionMatrixTests {
    [Theory]
    [InlineData(OutlookItemKind.Message)]
    [InlineData(OutlookItemKind.Appointment)]
    [InlineData(OutlookItemKind.Contact)]
    [InlineData(OutlookItemKind.Task)]
    public void ConvertsSupportedItemsAcrossEmlMsgTnefAndBack(OutlookItemKind kind) {
        EmailDocument source = CreateDocument(kind);
        var writer = new EmailDocumentWriter();
        var reader = new EmailDocumentReader();
        EmailDocument current = source;

        foreach (EmailFileFormat target in new[] {
            EmailFileFormat.Eml, EmailFileFormat.OutlookMsg, EmailFileFormat.Tnef, EmailFileFormat.Eml
        }) {
            EmailConversionReport report = writer.AnalyzeConversion(current, target);
            Assert.True(report.CanWrite, string.Join(Environment.NewLine,
                report.Diagnostics.Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message)));
            byte[] artifact = writer.ToBytes(current, target, out EmailWriteResult write);
            Assert.DoesNotContain(write.Diagnostics,
                diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
            EmailReadResult read = reader.Read(artifact);
            Assert.DoesNotContain(read.Diagnostics,
                diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
            Assert.Equal(target, read.Document.Format);
            current = read.Document;

            if (target == EmailFileFormat.Eml) {
                using var stream = new MemoryStream(artifact);
                Assert.NotNull(MimeMessage.Load(stream).Body);
            } else if (target == EmailFileFormat.OutlookMsg) {
                using var oracle = new global::MsgReader.Outlook.Storage.Message(
                    new MemoryStream(artifact), FileAccess.Read, true);
                Assert.False(string.IsNullOrWhiteSpace(oracle.Subject));
            }
        }

        Assert.Equal(kind, current.OutlookItemKind);
        Assert.Equal(source.Subject, current.Subject);
        if (kind == OutlookItemKind.Message) {
            Assert.Equal("Matrix body", current.Body.Text);
            Assert.Equal("matrix.txt", Assert.Single(current.Attachments).FileName);
        } else if (kind == OutlookItemKind.Appointment) {
            Assert.Equal("Matrix Room", current.Appointment!.Location);
            Assert.Equal(source.Appointment!.Start!.Value.UtcDateTime,
                current.Appointment.Start!.Value.UtcDateTime);
        } else if (kind == OutlookItemKind.Contact) {
            Assert.Equal("matrix.contact@example.com", current.Contact!.Email1.Address);
        } else {
            Assert.Equal(0.25, current.Task!.PercentComplete);
        }
    }

    [Fact]
    public void UnchangedProtectedContentCannotCrossFormatsWithoutExplicitLossAcceptance() {
        byte[] source = Encoding.ASCII.GetBytes(
            "Subject: Signed\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: multipart/signed; protocol=\"application/pkcs7-signature\"; boundary=\"s\"\r\n\r\n" +
            "--s\r\nContent-Type: text/plain\r\n\r\nbody\r\n" +
            "--s\r\nContent-Type: application/pkcs7-signature\r\n\r\nsignature\r\n--s--\r\n");
        EmailDocument protectedDocument = new EmailDocumentReader().Read(source).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            protectedDocument, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_PROTECTED_CONTENT_REWRITE");
    }

    private static EmailDocument CreateDocument(OutlookItemKind kind) {
        DateTimeOffset start = new DateTimeOffset(2026, 12, 1, 9, 0, 0, TimeSpan.Zero);
        var document = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = kind,
            Subject = "Matrix " + kind,
            MessageId = "matrix-" + kind.ToString().ToLowerInvariant() + "@example.com",
            Date = start,
            From = new EmailAddress("sender@example.com", "Sender")
        };
        if (kind == OutlookItemKind.Message) {
            document.Body.Text = "Matrix body";
            document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
                new EmailAddress("recipient@example.com", "Recipient")));
            document.Attachments.Add(new EmailAttachment {
                FileName = "matrix.txt",
                ContentType = "text/plain",
                Content = Encoding.UTF8.GetBytes("attachment"),
                Length = 10
            });
        } else if (kind == OutlookItemKind.Appointment) {
            document.MessageClass = "IPM.Appointment";
            document.Appointment = new OutlookAppointment {
                Start = start,
                End = start.AddHours(1),
                Location = "Matrix Room",
                BusyStatus = 2
            };
        } else if (kind == OutlookItemKind.Contact) {
            document.MessageClass = "IPM.Contact";
            document.Contact = new OutlookContact {
                DisplayName = "Matrix Contact",
                GivenName = "Matrix",
                Surname = "Contact",
                CompanyName = "Evotec"
            };
            document.Contact.Email1.Address = "matrix.contact@example.com";
        } else {
            document.MessageClass = "IPM.Task";
            document.Task = new OutlookTask {
                Start = start,
                Due = start.AddDays(1),
                Status = 1,
                PercentComplete = 0.25
            };
        }
        return document;
    }
}
