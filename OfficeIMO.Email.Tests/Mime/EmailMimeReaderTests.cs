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
        Assert.Equal("report", attachment.ContentTypeParameters["boundary"]);
        Assert.Contains("inner report", Encoding.ASCII.GetString(Assert.IsType<byte[]>(attachment.Content)),
            StringComparison.Ordinal);

        byte[] rewritten = new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Eml,
            out EmailWriteResult writeResult);
        EmailAttachment rewrittenAttachment = Assert.Single(new EmailDocumentReader().Read(rewritten).Document.Attachments);
        Assert.DoesNotContain(writeResult.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_MULTIPART_ATTACHMENT_WRITTEN_OPAQUE");
        Assert.Equal("multipart/report", rewrittenAttachment.ContentType);
        Assert.Equal("report", rewrittenAttachment.ContentTypeParameters["boundary"]);
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
    public void PrefersExtendedFilenameWhenThePlainFallbackAppearsLater() {
        const string eml = "Subject: extended filename\r\n" +
            "Content-Type: application/octet-stream\r\n" +
            "Content-Disposition: attachment; filename*=utf-8''caf%C3%A9.txt; filename=cafe.txt\r\n\r\n" +
            "content";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("café.txt", Assert.Single(document.Attachments).FileName);
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
    public void ParsesEncodedDisplayNamesThatDecodeToAngleBrackets() {
        const string eml = "To: =?utf-8?Q?Team_=3CEU=3E?= <team@example.com>\r\n\r\nbody";

        EmailRecipient recipient = Assert.Single(
            new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document.Recipients);

        Assert.Equal("team@example.com", recipient.Address.Address);
        Assert.Equal("Team <EU>", recipient.Address.DisplayName);
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

    [Fact]
    public void TreatsCidTextPartsOutsideMultipartRelatedAsBodies() {
        const string eml = "Subject: alternative\r\nContent-Type: multipart/alternative; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain; charset=utf-8\r\nContent-ID: <plain>\r\n\r\nplain body\r\n" +
            "--x\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <html>\r\n\r\n<p>html body</p>\r\n" +
            "--x--\r\n";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("plain body", document.Body.Text!.Trim());
        Assert.Equal("<p>html body</p>", document.Body.Html!.Trim());
        Assert.Empty(document.Attachments);
    }

    [Fact]
    public void TreatsInlineTextWithoutAttachmentIdentityAsBody() {
        const string eml = "Subject: inline body\r\nContent-Type: text/plain; charset=utf-8\r\n" +
            "Content-Disposition: inline\r\n\r\nbody text";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("body text", document.Body.Text);
        Assert.Empty(document.Attachments);
    }

    [Fact]
    public void RetainsAdditionalInlineTextPartsAfterSelectingThePrimaryBody() {
        const string eml = "Subject: multiple inline bodies\r\n" +
            "Content-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nprimary body\r\n" +
            "--x\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nadditional text\r\n" +
            "--x--\r\n";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("primary body", document.Body.Text!.Trim());
        EmailAttachment additional = Assert.Single(document.Attachments);
        Assert.True(additional.IsInline);
        Assert.Equal("text/plain", additional.ContentType);
        Assert.Equal("additional text", Encoding.UTF8.GetString(Assert.IsType<byte[]>(additional.Content)).Trim());

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document)).Document;
        Assert.Equal("primary body", roundTrip.Body.Text!.Trim());
        Assert.Equal("additional text", Encoding.UTF8.GetString(
            Assert.IsType<byte[]>(Assert.Single(roundTrip.Attachments).Content)).Trim());
    }

    [Fact]
    public void OmitsOrdinaryCalendarAndVcardAttachmentContentWhenRequested() {
        const string eml = "MIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/calendar; name=invite.ics\r\n" +
            "Content-Disposition: attachment; filename=invite.ics\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nEND:VCALENDAR\r\n" +
            "--x\r\nContent-Type: text/vcard; name=person.vcf\r\n" +
            "Content-Disposition: attachment; filename=person.vcf\r\n\r\n" +
            "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Ada Lovelace\r\nEND:VCARD\r\n--x--\r\n";

        EmailDocument document = new EmailDocumentReader(new EmailReaderOptions(includeAttachmentContent: false))
            .Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal(OutlookItemKind.Message, document.OutlookItemKind);
        Assert.Equal(2, document.Attachments.Count);
        Assert.All(document.Attachments, attachment => {
            Assert.Null(attachment.Content);
            Assert.False(attachment.IsProjectedSemanticContent);
            Assert.True(attachment.Length > 0);
        });
    }

    [Fact]
    public void PreservesEmptyAddressGroupsThroughStoreTransportHeaders() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "To: undisclosed-recipients:;\r\nSubject: private list\r\n\r\nbody");
        EmailDocument source = new EmailDocumentReader().Read(eml).Document;

        EmailDocument stored = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg)).Document;
        string regenerated = Encoding.UTF8.GetString(
            new EmailDocumentWriter().ToBytes(stored, EmailFileFormat.Eml));

        Assert.Empty(stored.Recipients);
        Assert.Contains(stored.Headers, header => header.Name == "To" &&
            (header.RawValue ?? header.Value).Contains("undisclosed-recipients:;", StringComparison.Ordinal));
        Assert.Contains("To: undisclosed-recipients:;", regenerated, StringComparison.Ordinal);
    }

    [Fact]
    public void ParsesAddressGroupsAndIgnoresEmptyGroups() {
        const string eml = "To: undisclosed-recipients:; (no recipients), Team: Alice <alice@example.com>, " +
            "Bob <bob@example.com>;\r\n\r\nbody";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal(2, document.Recipients.Count);
        Assert.Equal("alice@example.com", document.Recipients[0].Address.Address);
        Assert.Equal("bob@example.com", document.Recipients[1].Address.Address);
    }

    [Fact]
    public void RecoversBrokenMultipartAndBase64AsWarnings() {
        const string eml = "Subject: recovery\r\nContent-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: application/octet-stream\r\nContent-Transfer-Encoding: base64\r\n\r\n" +
            "not valid base64!\r\n";

        EmailReadResult result = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml));

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MIME_BASE64_INVALID" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Warning);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MIME_BOUNDARY_NOT_CLOSED" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Warning);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void DefaultsMultipartDigestChildrenToEmbeddedMessages() {
        const string eml = "Subject: digest\r\nContent-Type: multipart/digest; boundary=d\r\n\r\n" +
            "--d\r\n\r\nFrom: child@example.com\r\nSubject: Digest child\r\n\r\nchild body\r\n" +
            "--d--\r\n";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        EmailAttachment attachment = Assert.Single(document.Attachments);
        Assert.Equal("message/rfc822", attachment.ContentType);
        Assert.Equal("Digest child", attachment.EmbeddedDocument!.Subject);
        Assert.Equal("child body", attachment.EmbeddedDocument.Body.Text!.Trim());
    }

    [Fact]
    public void RoundTripsSemanticAttachmentContentTypeParameters() {
        const string eml = "Subject: invite\r\nContent-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain\r\n\r\nbody\r\n" +
            "--x\r\nContent-Type: text/calendar; method=REQUEST; charset=utf-8; name=invite.ics\r\n" +
            "Content-Disposition: attachment; filename=invite.ics\r\n\r\nBEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n" +
            "--x--\r\n";

        EmailDocument parsed = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;
        EmailAttachment attachment = Assert.Single(parsed.Attachments);
        byte[] rewritten = new EmailDocumentWriter().ToBytes(parsed);
        EmailAttachment roundTrip = Assert.Single(new EmailDocumentReader().Read(rewritten).Document.Attachments);

        Assert.Equal("REQUEST", attachment.ContentTypeParameters["method"]);
        Assert.Equal("utf-8", attachment.ContentTypeParameters["charset"]);
        Assert.False(attachment.ContentTypeParameters.ContainsKey("name"));
        Assert.Equal(attachment.ContentTypeParameters, roundTrip.ContentTypeParameters);
    }

    [Fact]
    public void PreservesCidMultipartResourcesAsInlineAttachments() {
        const string eml = "Subject: related multipart\r\nContent-Type: multipart/related; boundary=outer\r\n\r\n" +
            "--outer\r\nContent-Type: text/html\r\n\r\n<p>body</p>\r\n" +
            "--outer\r\nContent-Type: multipart/alternative; boundary=resource\r\nContent-ID: <resource>\r\n\r\n" +
            "--resource\r\nContent-Type: text/plain\r\n\r\nresource text\r\n" +
            "--resource\r\nContent-Type: text/html\r\n\r\n<p>resource</p>\r\n" +
            "--resource--\r\n--outer--\r\n";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("<p>body</p>", document.Body.Html!.Trim());
        EmailAttachment attachment = Assert.Single(document.Attachments);
        Assert.Equal("multipart/alternative", attachment.ContentType);
        Assert.Equal("resource", attachment.ContentId);
        Assert.True(attachment.IsInline);
        Assert.Contains("resource text", Encoding.ASCII.GetString(Assert.IsType<byte[]>(attachment.Content)),
            StringComparison.Ordinal);
    }

    [Fact]
    public void UsesMultipartRelatedStartCidAsTheMessageBody() {
        const string eml = "Subject: related start\r\n" +
            "Content-Type: multipart/related; boundary=outer; start=\"<root>\"\r\n\r\n" +
            "--outer\r\nContent-Type: image/png\r\nContent-ID: <logo>\r\n\r\npng\r\n" +
            "--outer\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <root>\r\n\r\n<p>root body</p>\r\n" +
            "--outer--\r\n";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("<p>root body</p>", document.Body.Html!.Trim());
        EmailAttachment attachment = Assert.Single(document.Attachments);
        Assert.Equal("logo", attachment.ContentId);
    }

    [Fact]
    public void UsesFirstRelatedPartAsBodyWhenStartIsAbsent() {
        const string eml = "Subject: default related root\r\nContent-Type: multipart/related; boundary=outer\r\n\r\n" +
            "--outer\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <root>\r\n\r\n<p>default root</p>\r\n" +
            "--outer\r\nContent-Type: image/png\r\nContent-ID: <logo>\r\n\r\npng\r\n" +
            "--outer--\r\n";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("<p>default root</p>", document.Body.Html!.Trim());
        Assert.Equal("logo", Assert.Single(document.Attachments).ContentId);
    }

    [Fact]
    public void KeepsRelatedTextWithContentLocationAsAnInlineAttachment() {
        const string eml = "Subject: related text resource\r\nContent-Type: multipart/related; boundary=outer\r\n\r\n" +
            "--outer\r\nContent-Type: text/html; charset=utf-8\r\n\r\n<a href=\"caption.txt\">root</a>\r\n" +
            "--outer\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Location: caption.txt\r\n\r\n" +
            "inline caption\r\n--outer--\r\n";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("<a href=\"caption.txt\">root</a>", document.Body.Html!.Trim());
        Assert.Null(document.Body.Text);
        EmailAttachment attachment = Assert.Single(document.Attachments);
        Assert.Equal("caption.txt", attachment.ContentLocation);
        Assert.True(attachment.IsInline);
        Assert.Equal("inline caption", Encoding.UTF8.GetString(Assert.IsType<byte[]>(attachment.Content)).Trim());

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document)).Document;
        EmailAttachment roundTripAttachment = Assert.Single(roundTrip.Attachments);
        Assert.Equal("caption.txt", roundTripAttachment.ContentLocation);
        Assert.True(roundTripAttachment.IsInline);
    }

    [Fact]
    public void DecodesFormatFlowedTextWithDelSpAndSpaceStuffing() {
        const string eml = "Content-Type: text/plain; charset=utf-8; format=flowed; delsp=yes\r\n\r\n" +
            "This is flow \r\ned.\r\n From space-stuffed\r\n>quoted \r\n>continuation\r\n-- \r\nsignature";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("This is flowed.\r\nFrom space-stuffed\r\n>quotedcontinuation\r\n-- \r\nsignature",
            document.Body.Text);
    }

    [Fact]
    public void KeepsFlowedJoinSpaceWhenDelSpIsNotEnabled() {
        const string eml = "Content-Type: text/plain; charset=utf-8; format=flowed\r\n\r\n" +
            "This is flow \r\ned.";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("This is flow ed.", document.Body.Text);
    }

    [Fact]
    public void RemovesFinalFlowedSpaceMarkerWhenDelSpIsEnabled() {
        const string eml = "Content-Type: text/plain; charset=utf-8; format=flowed; delsp=yes\r\n\r\n" +
            "https://example.com/ ";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(eml)).Document;

        Assert.Equal("https://example.com/", document.Body.Text);
    }

    [Fact]
    public void RecoversCharsetlessInvalidUtf8AsWindows1252() {
        byte[] headers = Encoding.ASCII.GetBytes("Content-Type: text/plain\r\n\r\nPrice: ");
        byte[] eml = new byte[headers.Length + 1];
        Buffer.BlockCopy(headers, 0, eml, 0, headers.Length);
        eml[eml.Length - 1] = 0xA3;

        EmailReadResult result = new EmailDocumentReader().Read(eml);

        Assert.Equal("Price: £", result.Document.Body.Text);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MIME_CHARSET_GUESSED" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Warning);
    }

    [Fact]
    public void KeepsCharsetlessValidUtf8WithoutGuessing() {
        const string body = "Zażółć gęślą jaźń";
        byte[] headers = Encoding.ASCII.GetBytes("Content-Type: text/plain\r\n\r\n");
        byte[] bodyBytes = Encoding.UTF8.GetBytes(body);
        byte[] eml = new byte[headers.Length + bodyBytes.Length];
        Buffer.BlockCopy(headers, 0, eml, 0, headers.Length);
        Buffer.BlockCopy(bodyBytes, 0, eml, headers.Length, bodyBytes.Length);

        EmailReadResult result = new EmailDocumentReader().Read(eml);

        Assert.Equal(body, result.Document.Body.Text);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MIME_CHARSET_GUESSED");
    }
}
