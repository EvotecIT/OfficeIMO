using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailMimeReaderTests {
    [Fact]
    public void ReadsMultipartAlternativesAndAttachmentsWithStandardEncodings() {
        const string eml = "From: =?utf-8?B?Sm9zw6kgU2VuZGVy?= <sender@example.com>\r\n" +
            "To: Alice <alice@example.com>, bob@example.com\r\n" +
            "Cc: Team <team@example.com>\r\n" +
            "Subject: =?utf-8?B?UsOpc3Vtw6k=?=\r\n" +
            "X-Trace: first\r\n" +
            "X-Trace: second\r\n" +
            "Message-ID: <sample@example.com>\r\n" +
            "Date: Fri, 10 Jul 2026 12:30:00 +0200\r\n" +
            "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=outer\r\n\r\n" +
            "preamble\r\n--outer\r\n" +
            "Content-Type: multipart/alternative; boundary=inner\r\n\r\n" +
            "--inner\r\nContent-Type: text/plain; charset=utf-8\r\n" +
            "Content-Transfer-Encoding: quoted-printable\r\n\r\nHello =C5=BC=C3=B3=C5=82=C4=87\r\n" +
            "--inner\r\nContent-Type: text/html; charset=utf-8\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nPHA+SGVsbG88L3A+\r\n" +
            "--inner--\r\n" +
            "--outer\r\nContent-Type: application/octet-stream\r\n" +
            "Content-Disposition: attachment; filename*=utf-8''c%C3%A9.txt\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nAQIDBA==\r\n" +
            "--outer--\r\nepilogue\r\n";

        EmailReadResult result = new EmailDocumentReader().Read(Encoding.UTF8.GetBytes(eml));

        Assert.Equal(EmailFileFormat.Eml, result.Document.Format);
        Assert.Equal("Résumé", result.Document.Subject);
        Assert.Equal("José Sender", result.Document.From!.DisplayName);
        Assert.Equal("sender@example.com", result.Document.From.Address);
        Assert.Equal(3, result.Document.Recipients.Count);
        Assert.Equal("Hello żółć", result.Document.Body.Text!.Trim());
        Assert.Equal("<p>Hello</p>", result.Document.Body.Html!.Trim());
        Assert.Equal(2, result.Document.Headers.Count(header => header.Name == "X-Trace"));
        EmailAttachment attachment = Assert.Single(result.Document.Attachments);
        Assert.Equal("cé.txt", attachment.FileName);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, attachment.Content);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void PreservesNamedMultipartEntitiesAsSingleAttachments() {
        const string eml = "Subject: multipart attachment\r\n" +
            "Content-Type: multipart/mixed; boundary=outer\r\n\r\n" +
            "--outer\r\nContent-Type: text/plain\r\n\r\nmessage body\r\n" +
            "--outer\r\nContent-Type: multipart/report; boundary=report; name=delivery-report.mime\r\n" +
            "Content-Disposition: attachment; filename=delivery-report.mime\r\n\r\n" +
            "--report\r\nContent-Type: text/plain\r\n\r\ninner report\r\n" +
            "--report\r\nContent-Type: text/plain; name=details.txt\r\n" +
            "Content-Disposition: attachment; filename=details.txt\r\n\r\ndetails\r\n" +
            "--report--\r\n--outer--\r\n";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("message body", document.Body.Text!.Trim());
        EmailAttachment attachment = Assert.Single(document.Attachments);
        Assert.Equal("delivery-report.mime", attachment.FileName);
        Assert.Equal("multipart/report", attachment.ContentType);
        Assert.Contains("inner report", Encoding.ASCII.GetString(Assert.IsType<byte[]>(attachment.Content)),
            StringComparison.Ordinal);

        byte[] rewritten = new EmailDocumentWriter().WriteToBytes(document, EmailFileFormat.Eml,
            out EmailWriteResult writeResult);
        EmailAttachment rewrittenAttachment = Assert.Single(new EmailDocumentReader().Read(rewritten).Document.Attachments);
        Assert.Contains(writeResult.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_MULTIPART_ATTACHMENT_WRITTEN_OPAQUE");
        Assert.Equal("application/octet-stream", rewrittenAttachment.ContentType);
        Assert.Equal(attachment.Content, rewrittenAttachment.Content);
    }

    [Fact]
    public void ReadsEmbeddedMessageAsStructuredAttachment() {
        const string eml = "Subject: Parent\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain\r\n\r\nParent body\r\n" +
            "--x\r\nContent-Type: message/rfc822; name=child.eml\r\n" +
            "Content-Disposition: attachment; filename=child.eml\r\n\r\n" +
            "From: child@example.com\r\nSubject: Child\r\nContent-Type: text/plain\r\n\r\nChild body\r\n" +
            "--x--\r\n";

        EmailReadResult result = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml));

        EmailAttachment attachment = Assert.Single(result.Document.Attachments);
        Assert.Equal("child.eml", attachment.FileName);
        Assert.NotNull(attachment.EmbeddedDocument);
        Assert.Equal("Child", attachment.EmbeddedDocument!.Subject);
        Assert.Equal("Child body", attachment.EmbeddedDocument.Body.Text!.Trim());
    }

    [Fact]
    public void ReportsRecoverableMalformedContent() {
        const string eml = "Subject: malformed\r\nContent-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: application/octet-stream\r\nContent-Transfer-Encoding: base64\r\n\r\n%%%\r\n";

        EmailReadResult result = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml));

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MIME_BASE64_INVALID");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MIME_BOUNDARY_NOT_CLOSED");
        Assert.Single(result.Document.Attachments);
    }

    [Fact]
    public void DecodesLegacyCodePagesWithoutPriorMsgInitialization() {
        string encoded = Convert.ToBase64String(new byte[] { 0x5A, 0x61, 0xBF, 0xF3, 0xB3, 0xE6 });
        string eml = string.Concat("Subject: =?windows-1250?B?", encoded, "?=\r\n\r\nbody");

        EmailReadResult result = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml));

        Assert.Equal("Zażółć", result.Document.Subject);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MIME_CHARSET_UNSUPPORTED");
    }

    [Fact]
    public void IgnoresWhitespaceBetweenAdjacentEncodedWords() {
        string first = Convert.ToBase64String(Encoding.UTF8.GetBytes("Zaż"));
        string second = Convert.ToBase64String(Encoding.UTF8.GetBytes("ółć"));
        string eml = string.Concat("Subject: =?utf-8?B?", first, "?=  \t=?utf-8?B?", second, "?=\r\n\r\nbody");

        EmailReadResult result = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml));

        Assert.Equal("Zażółć", result.Document.Subject);
    }

    [Fact]
    public void ReadsRawUtf8HeadersAndUsesTopmostReceivedTimestamp() {
        const string eml = "From: José <jose@example.com>\r\n" +
            "Subject: Café\r\n" +
            "Received: from final.example; Fri, 10 Jul 2026 12:30:00 +0000\r\n" +
            "Received: from origin.example; Fri, 10 Jul 2026 10:00:00 +0000\r\n\r\nbody";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.UTF8.GetBytes(eml)).Document;

        Assert.Equal("Café", document.Subject);
        Assert.Equal("José", document.From!.DisplayName);
        Assert.Equal(new DateTimeOffset(2026, 7, 10, 12, 30, 0, TimeSpan.Zero), document.ReceivedDate);
    }

    [Fact]
    public void DecodesRfc2231ContinuationAfterJoiningEncodedSegments() {
        const string eml = "Subject: continuation\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain\r\n\r\nbody\r\n" +
            "--x\r\nContent-Type: application/octet-stream\r\n" +
            "Content-Disposition: attachment; filename*0*=utf-8''price-%E2%82; filename*1*=%AC.txt\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nAQ==\r\n--x--\r\n";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("price-€.txt", Assert.Single(document.Attachments).FileName);
    }

    [Fact]
    public void AcceptsUtf8BomWithoutPollutingTheFirstHeaderName() {
        byte[] message = Encoding.UTF8.GetPreamble()
            .Concat(Encoding.UTF8.GetBytes("Subject: BOM message\r\nFrom: sender@example.com\r\n\r\nbody"))
            .ToArray();

        EmailReadResult result = new EmailDocumentReader().Read(message);

        Assert.Equal(EmailFileFormat.Eml, result.Document.Format);
        Assert.Equal("BOM message", result.Document.Subject);
        Assert.Contains(result.Document.Headers, header => header.Name == "Subject");
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_FORMAT_UNKNOWN" ||
            diagnostic.Code == "EMAIL_MIME_HEADER_MALFORMED");
    }

    [Fact]
    public void SplitsRawAddressListsBeforeDecodingDisplayNames() {
        const string eml = "To: =?utf-8?B?RG9lLCBKb2hu?= <john@example.com>, Jane <jane@example.com>\r\n\r\nbody";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal(2, document.Recipients.Count);
        Assert.Equal("Doe, John", document.Recipients[0].Address.DisplayName);
        Assert.Equal("john@example.com", document.Recipients[0].Address.Address);
        Assert.Equal("jane@example.com", document.Recipients[1].Address.Address);
    }

    [Fact]
    public void KeepsLiteralCommasInsideQEncodedDisplayNames() {
        const string eml = "To: =?utf-8?Q?Doe,_John?= <john@example.com>, Jane <jane@example.com>\r\n\r\nbody";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal(2, document.Recipients.Count);
        Assert.Equal("Doe, John", document.Recipients[0].Address.DisplayName);
        Assert.Equal("john@example.com", document.Recipients[0].Address.Address);
        Assert.Equal("jane@example.com", document.Recipients[1].Address.Address);
    }

    [Fact]
    public void PreservesCidTextPartsAsInlineAttachments() {
        const string eml = "Subject: related\r\nContent-Type: multipart/related; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/html; charset=utf-8\r\n\r\n<p>body</p>\r\n" +
            "--x\r\nContent-Type: text/plain; charset=utf-8\r\nContent-ID: <caption>\r\n\r\ninline caption\r\n" +
            "--x--\r\n";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("<p>body</p>", document.Body.Html!.Trim());
        Assert.Null(document.Body.Text);
        EmailAttachment attachment = Assert.Single(document.Attachments);
        Assert.True(attachment.IsInline);
        Assert.Equal("caption", attachment.ContentId);
        Assert.Equal("inline caption", Encoding.UTF8.GetString(Assert.IsType<byte[]>(attachment.Content)).Trim());
    }
}
