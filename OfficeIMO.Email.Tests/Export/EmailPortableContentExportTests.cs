using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailPortableContentExportTests {
    [Fact]
    public void ExportsTypedAppointmentTaskAndContactAsStandaloneArtifacts() {
        var appointment = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg, OutlookItemKind = OutlookItemKind.Appointment,
            Subject = "Design; review", Appointment = new OutlookAppointment {
                Start = new DateTimeOffset(2026, 10, 1, 9, 0, 0, TimeSpan.FromHours(2)),
                End = new DateTimeOffset(2026, 10, 1, 10, 0, 0, TimeSpan.FromHours(2)), Location = "Room 1"
            }
        };
        var calendar = EmailPortableContentExport.ToCalendar(appointment);
        Assert.False(calendar.RetainedSemanticSource);
        ContentLineComponent item = Assert.Single(IcsDocument.Load(calendar.ToBytes()).GetComponents("VEVENT"));
        Assert.Equal("Design\\; review", item.GetFirstProperty("SUMMARY")!.Value);
        Assert.Equal("20261001T070000Z", item.GetFirstProperty("DTSTART")!.Value);
        Assert.Empty(calendar.Diagnostics);

        var task = new EmailDocument { OutlookItemKind = OutlookItemKind.Task, Subject = "Prepare notes", Task = new OutlookTask() };
        Assert.Single(EmailPortableContentExport.ToCalendar(task).Document.GetComponents("VTODO"));
        var contact = Contact("Ada", "ada@example.com");
        contact.Contact!.Phones.Mobile = "+44 123";
        var card = EmailPortableContentExport.ToVCard(contact);
        ContentLineComponent exported = Assert.Single(VCardDocument.Load(card.ToBytes()).Cards);
        Assert.Equal("Ada", exported.GetVCardText("FN"));
        Assert.Equal("ada@example.com", exported.GetFirstProperty("EMAIL")!.Value);
        Assert.Equal("+44 123", exported.GetFirstProperty("TEL")!.Value);
        byte[] independent = card.ToBytes(); independent[0] = 0;
        Assert.Equal((byte)'B', card.ToBytes()[0]);
        Assert.Equal("Ada", contact.Contact.DisplayName);
    }

    [Theory]
    [InlineData("text/calendar", "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:test\r\nBEGIN:VEVENT\r\nUID:1\r\nDTSTART:20261001T090000Z\r\nX-PRIVATE-EXT:keep\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n", true)]
    [InlineData("text/vcard", "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Ada\r\nX-PRIVATE-EXT:keep\r\nEND:VCARD\r\n", false)]
    public void RetainsImportedExtensionsAndBlocksRegenerationAfterMutation(string type, string body, bool isCalendar) {
        byte[] eml = Encoding.UTF8.GetBytes("MIME-Version: 1.0\r\nContent-Type: " + type + "; charset=utf-8\r\n\r\n" + body);
        using var read = new EmailDocumentReader().Read(eml);
        EmailDocument document = read.Document;
        byte[] semanticBytes = (byte[])Assert.Single(document.Attachments).Content!.Clone();
        if (isCalendar) {
            var export = EmailPortableContentExport.ToCalendar(document);
            Assert.True(export.RetainedSemanticSource);
            Assert.Equal("keep", Assert.Single(export.Document.GetComponents("VEVENT")).GetFirstProperty("X-PRIVATE-EXT")!.Value);
        } else {
            var export = EmailPortableContentExport.ToVCard(document);
            Assert.True(export.RetainedSemanticSource);
            Assert.Equal("keep", Assert.Single(export.Document.Cards).GetFirstProperty("X-PRIVATE-EXT")!.Value);
        }
        document.Subject = "changed";
        InvalidOperationException stopped = isCalendar
            ? Assert.Throws<InvalidOperationException>(() => EmailPortableContentExport.ToCalendar(document))
            : Assert.Throws<InvalidOperationException>(() => EmailPortableContentExport.ToVCard(document));
        Assert.Contains("EMAIL_MIME_SEMANTIC_CONTENT_CHANGED", stopped.Message);
        var warn = new EmailPortableContentExportOptions(EmailConversionLossPolicy.Warn);
        IReadOnlyList<EmailDiagnostic> diagnostics = isCalendar
            ? EmailPortableContentExport.ToCalendar(document, warn).Diagnostics
            : EmailPortableContentExport.ToVCard(document, warn).Diagnostics;
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "EMAIL_MIME_SEMANTIC_CONTENT_CHANGED");
        Assert.Equal(semanticBytes, Assert.Single(document.Attachments).Content);
    }

    [Fact]
    public void ReportsOpaqueStateAndRejectsWrongKindsOversizedExportsAndCancellation() {
        var contact = Contact("Ada", "/o=directory/cn=Ada");
        contact.Contact!.Email1.AddressType = "EX";
        Assert.Contains("EMAIL_VCARD_OPAQUE_CONTACT_IDENTITY", Assert.Throws<InvalidOperationException>(
            () => EmailPortableContentExport.ToVCard(contact)).Message);
        Assert.Contains(EmailPortableContentExport.ToVCard(contact,
            new EmailPortableContentExportOptions(EmailConversionLossPolicy.Warn)).Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_VCARD_OPAQUE_CONTACT_IDENTITY");
        Assert.Throws<ArgumentException>(() => EmailPortableContentExport.ToCalendar(contact));
        var large = Contact("Ada", "ada@example.com"); large.Body.Text = new string('x', 10000);
        Assert.Throws<EmailLimitExceededException>(() => EmailPortableContentExport.ToVCard(large,
            new EmailPortableContentExportOptions(maxOutputBytes: 300)));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => EmailPortableContentExport.ToVCard(large, cancellationToken: cancellation.Token));
    }

    internal static EmailDocument Contact(string name, string address) {
        var contact = new OutlookContact { DisplayName = name };
        contact.Email1.Address = address;
        return new EmailDocument { OutlookItemKind = OutlookItemKind.Contact, Contact = contact };
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ExportsEveryUnchangedSemanticPartAndBoundsTheirCombinedSourceBytes(bool calendar) {
        string Part(string id) => calendar
            ? "Content-Type: text/calendar\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:test\r\nBEGIN:VEVENT\r\nUID:" + id + "\r\nDTSTART:20261001T090000Z\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n"
            : "Content-Type: text/vcard\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\nFN:" + id + "\r\nEND:VCARD\r\n";
        string eml = "MIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=parts\r\n\r\n--parts\r\n" + Part("a") + "--parts\r\n" + Part("b") + "--parts--\r\n";
        using var read = new EmailDocumentReader().Read(Encoding.UTF8.GetBytes(eml));
        if (calendar) Assert.Equal(2, EmailPortableContentExport.ToCalendar(read.Document).Document.Calendars.Count);
        else Assert.Equal(2, EmailPortableContentExport.ToVCard(read.Document).Document.Cards.Count);
        long firstBytes = read.Document.Attachments[0].Content!.LongLength;
        var small = new EmailPortableContentExportOptions(maxSourceBytes: firstBytes);
        if (calendar) Assert.Throws<EmailLimitExceededException>(() => EmailPortableContentExport.ToCalendar(read.Document, small));
        else Assert.Throws<EmailLimitExceededException>(() => EmailPortableContentExport.ToVCard(read.Document, small));
    }

    [Fact]
    public void Utf8ExportNormalizesUnencodedLegacyCharsetWithoutChangingQuotedPrintableBytes() {
        string card = "BEGIN:VCARD\r\nVERSION:2.1\r\nFN;CHARSET=iso-8859-1:José\r\nNOTE;CHARSET=iso-8859-1;ENCODING=QUOTED-PRINTABLE:Jos=E9\r\nEND:VCARD\r\n";
        byte[] source = Encoding.GetEncoding(28591).GetBytes("MIME-Version: 1.0\r\nContent-Type: text/vcard; charset=iso-8859-1\r\n\r\n" + card);
        using var read = new EmailDocumentReader().Read(source);
        var exported = EmailPortableContentExport.ToVCard(read.Document);
        ContentLineComponent result = VCardDocument.Load(exported.ToBytes()).Cards[0];
        Assert.Equal("utf-8", Assert.Single(result.GetFirstProperty("FN")!.GetParameter("CHARSET")!.Values));
        Assert.Equal("José", result.GetVCardText("FN"));
        // Independent property-aware decoding of actual output octets, rather than another Unicode model round trip.
        string wireLine = Encoding.GetEncoding(28591).GetString(exported.ToBytes()).Split(new[] { "\r\n" }, StringSplitOptions.None)
            .Single(line => line.StartsWith("FN;", StringComparison.Ordinal));
        byte[] valueBytes = Encoding.GetEncoding(28591).GetBytes(wireLine.Substring(wireLine.IndexOf(':') + 1));
        Assert.Equal("José", Encoding.GetEncoding(result.GetFirstProperty("FN")!.GetParameter("CHARSET")!.Values[0]).GetString(valueBytes));
        Assert.Equal("iso-8859-1", Assert.Single(result.GetFirstProperty("NOTE")!.GetParameter("CHARSET")!.Values));
        Assert.Equal("Jos=E9", result.GetFirstProperty("NOTE")!.Value);
        Assert.Contains(exported.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_VCARD_CHARSET_NORMALIZED");
        Assert.Equal(Encoding.GetEncoding(28591).GetBytes(card), Assert.Single(read.Document.Attachments).Content);
    }

    [Fact]
    public void RetainedCalendarKeepsEffectiveMimeMethodPerPart() {
        string Part(string uid, string method) => "Content-Type: text/calendar; method=" + method + "\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:test\r\nBEGIN:VEVENT\r\nUID:" + uid + "\r\nDTSTART:20261001T090000Z\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n";
        string eml = "MIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=parts\r\n\r\n--parts\r\n" + Part("a", "CANCEL") + "--parts\r\n" + Part("b", "REQUEST") + "--parts--\r\n";
        using var read = new EmailDocumentReader().Read(Encoding.UTF8.GetBytes(eml));
        var export = EmailPortableContentExport.ToCalendar(read.Document);
        Assert.Equal(new[] { "CANCEL", "REQUEST" }, export.Document.Calendars.Select(calendar => calendar.GetFirstProperty("METHOD")?.Value));
    }

    [Fact]
    public void SourceByteBudgetDoesNotRejectLargerUtf8Transcoding() {
        string card = "BEGIN:VCARD\r\nVERSION:2.1\r\nFN:José\r\nEND:VCARD\r\n";
        byte[] payload = Encoding.GetEncoding(28591).GetBytes(card);
        byte[] eml = Encoding.GetEncoding(28591).GetBytes("MIME-Version: 1.0\r\nContent-Type: text/vcard; charset=iso-8859-1\r\n\r\n" + card);
        using var read = new EmailDocumentReader().Read(eml);
        var export = EmailPortableContentExport.ToVCard(read.Document,
            new EmailPortableContentExportOptions(maxSourceBytes: payload.Length, maxOutputBytes: payload.Length * 2));
        Assert.Equal("José", export.Document.Cards[0].GetVCardText("FN"));
        Assert.True(export.ToBytes().Length > payload.Length);
    }
}
