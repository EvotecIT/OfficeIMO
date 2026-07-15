using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailSemanticProjectionLossTests {
    [Fact]
    public void PreservesMissingMimeCalendarMethodWhenReusingUnchangedContent() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:methodless-rewrite@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "ATTENDEE:mailto:reader@example.com\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        string rewritten = Encoding.ASCII.GetString(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Eml));

        Assert.Contains("Content-Type: text/calendar", rewritten, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("method=", rewritten, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("SUMMARY;LANGUAGE=fr:Reunion")]
    [InlineData("DESCRIPTION;ALTREP=\"cid:description\":Notes")]
    public void BlocksParameterizedCalendarTextBeforeStoreConversion(string property) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:parameterized-text@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            property + "\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");

        AssertStoreProjectionBlocked(eml);
    }

    [Fact]
    public void BlocksCalendarRelationshipsBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:related@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "RELATED-TO;RELTYPE=PARENT:parent@example.com\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");

        AssertStoreProjectionBlocked(eml);
    }

    [Theory]
    [InlineData("FN;LANGUAGE=fr:Jean Dupont")]
    [InlineData("NOTE;ALTID=1:Notes")]
    public void BlocksParameterizedVcardTextBeforeStoreConversion(string property) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            property + "\r\nEND:VCARD\r\n");

        AssertStoreProjectionBlocked(eml);
    }

    [Theory]
    [InlineData("AGENT:BEGIN:VCARD\\nFN:Assistant\\nEND:VCARD")]
    [InlineData("SORT-STRING:Lovelace, Ada")]
    [InlineData("MAILER:OfficeIMO.Email")]
    public void BlocksUnprojectedVcardDelegationAndSortFieldsBeforeStoreConversion(string property) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\n" + property + "\r\nEND:VCARD\r\n");

        AssertStoreProjectionBlocked(eml);
    }

    [Theory]
    [InlineData("METHOD:PUBLISH\r\n", "ORGANIZER:mailto:alice@example.com?subject=Meeting")]
    [InlineData("METHOD:REQUEST\r\n", "ATTENDEE:mailto:alice@example.com?subject=Meeting")]
    public void BlocksCalendarMailtoUrisWithHeaderFieldsBeforeStoreConversion(
        string method, string addressProperty) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            method + "BEGIN:VEVENT\r\nUID:mailto-header@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            addressProperty + "\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");

        AssertStoreProjectionBlocked(eml);
    }

    [Theory]
    [InlineData("")]
    [InlineData("VERSION:3.0\r\nVERSION:3.0\r\n")]
    public void RequiresExactlyOneVcardVersionBeforeStoreConversion(string versions) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\n" + versions +
            "FN:Ada Lovelace\r\nEND:VCARD\r\n");

        AssertStoreProjectionBlocked(eml);
    }

    [Theory]
    [InlineData("")]
    [InlineData("SUMMARY:Calendar subject\r\n")]
    public void BlocksTransportSubjectPromotionOrReplacementBeforeStoreConversion(string summary) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Subject: Transport subject\r\nContent-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nMETHOD:PUBLISH\r\nBEGIN:VEVENT\r\n" +
            "UID:subject-provenance@example.com\r\nDTSTART:20260801T100000Z\r\n" + summary +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");

        AssertStoreProjectionBlocked(eml);
    }

    [Fact]
    public void BlocksTransportFromPromotionIntoCalendarOrganizer() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "From: Owner <owner@example.com>\r\nContent-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nMETHOD:PUBLISH\r\nBEGIN:VEVENT\r\n" +
            "UID:organizer-provenance@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");

        AssertStoreProjectionBlocked(eml);
    }

    [Fact]
    public void BlocksUnknownCalendarExtensionsBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "METHOD:PUBLISH\r\nBEGIN:VEVENT\r\nUID:vendor-extension@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nX-APPLE-STRUCTURED-LOCATION:geo:52.2,21.0\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");

        AssertStoreProjectionBlocked(eml);
    }

    [Fact]
    public void BlocksDuplicateCalendarAttendeesBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "METHOD:REQUEST\r\nBEGIN:VEVENT\r\nUID:duplicate-attendee@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nATTENDEE;ROLE=REQ-PARTICIPANT:mailto:alice@example.com\r\n" +
            "ATTENDEE;ROLE=OPT-PARTICIPANT:mailto:alice@example.com\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");

        AssertStoreProjectionBlocked(eml);
    }

    [Theory]
    [InlineData("X-SOCIALPROFILE:https://social.example/ada")]
    [InlineData("item1.X-ABLABEL:Social")]
    public void BlocksUnknownVcardExtensionsBeforeStoreConversion(string extension) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\n" + extension + "\r\nEND:VCARD\r\n");

        AssertStoreProjectionBlocked(eml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BlocksNonSmtpCalendarRecipientsAndOmitsInvalidMailtoValues(bool task) {
        var document = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = task ? OutlookItemKind.Task : OutlookItemKind.Appointment,
            Subject = "Portable attendee",
            Task = task ? new OutlookTask() : null,
            Appointment = task ? null : new OutlookAppointment {
                Start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero)
            }
        };
        document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("/o=Example/ou=Exchange/cn=Recipients/cn=Reader", "Reader") { AddressType = "EX" }));

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.Eml);
        byte[] warned = new EmailDocumentWriter(new EmailWriterOptions(EmailConversionLossPolicy.Warn))
            .ToBytes(document, EmailFileFormat.Eml);
        EmailAttachment calendar = Assert.Single(new EmailDocumentReader().Read(warned).Document.Attachments,
            attachment => string.Equals(attachment.ContentType, "text/calendar", StringComparison.OrdinalIgnoreCase));
        string calendarText = Encoding.UTF8.GetString(Assert.IsType<byte[]>(calendar.Content));

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_ATTENDEE_ADDRESS_REQUIRED");
        Assert.DoesNotContain("/o=Example", calendarText, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BlocksNonSmtpCalendarOrganizersAndOmitsInvalidMailtoValues(bool task) {
        const string organizerAddress = "owner@example.com";
        var document = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = task ? OutlookItemKind.Task : OutlookItemKind.Appointment,
            Subject = "Portable organizer",
            From = new EmailAddress(organizerAddress, "Owner") { AddressType = "EX" },
            Task = task ? new OutlookTask { Owner = organizerAddress } : null,
            Appointment = task ? null : new OutlookAppointment {
                Start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero)
            }
        };

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.Eml);
        byte[] warned = new EmailDocumentWriter(new EmailWriterOptions(EmailConversionLossPolicy.Warn))
            .ToBytes(document, EmailFileFormat.Eml);
        EmailAttachment calendar = Assert.Single(new EmailDocumentReader().Read(warned).Document.Attachments,
            attachment => string.Equals(attachment.ContentType, "text/calendar", StringComparison.OrdinalIgnoreCase));
        string calendarText = Encoding.UTF8.GetString(Assert.IsType<byte[]>(calendar.Content));

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_ORGANIZER_ADDRESS_REQUIRED");
        Assert.DoesNotContain("\r\nORGANIZER", calendarText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void BlocksMimeBodyPromotionIntoCalendarDescription() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "MIME-Version: 1.0\r\nContent-Type: multipart/alternative; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nWrapper text\r\n" +
            "--x\r\nContent-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:no-description@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n--x--\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.Equal("Wrapper text", document.Body.Text!.Trim());
        AssertProjectionBlocked(report);
    }

    [Fact]
    public void BlocksMimeBodyPromotionIntoVcardNote() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "MIME-Version: 1.0\r\nContent-Type: multipart/alternative; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nWrapper text\r\n" +
            "--x\r\nContent-Type: text/vcard; charset=utf-8\r\n\r\n" +
            "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Ada Lovelace\r\nEND:VCARD\r\n--x--\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.Equal("Wrapper text", document.Body.Text!.Trim());
        AssertProjectionBlocked(report);
    }

    [Fact]
    public void BlocksDistributionListConversionToIndividualVcard() {
        var document = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Contact,
            MessageClass = "IPM.DistList",
            Contact = new OutlookContact { DisplayName = "Engineering" }
        };

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.Eml);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_VCARD_DISTRIBUTION_LIST_UNSUPPORTED");
    }

    private static void AssertStoreProjectionBlocked(byte[] eml) {
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;
        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);
        AssertProjectionBlocked(report);
    }

    private static void AssertProjectionBlocked(EmailConversionReport report) {
        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }
}
