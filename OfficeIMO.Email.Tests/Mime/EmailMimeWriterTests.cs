using OfficeIMO.Email;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailMimeWriterTests {
    [Fact]
    public void DeterministicWriterRoundTripsEnvelopeBodiesAndAttachments() {
        EmailDocument document = new EmailDocument {
            Format = EmailFileFormat.Eml,
            Subject = "Zażółć",
            From = new EmailAddress("sender@example.com", "José"),
            MessageId = "stable@example.com",
            Date = new DateTimeOffset(2026, 7, 10, 15, 0, 0, TimeSpan.FromHours(2))
        };
        document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("to@example.com", "Receiver")));
        document.Body.Text = "plain";
        document.Body.Html = "<strong>html</strong>";
        document.Headers.Add(new EmailHeader("X-Correlation", "abc"));
        document.Attachments.Add(new EmailAttachment {
            FileName = "dane-ą.bin",
            ContentType = "application/octet-stream",
            Content = new byte[] { 9, 8, 7 },
            Length = 3
        });

        EmailDocumentWriter writer = new EmailDocumentWriter();
        byte[] first = writer.WriteToBytes(document);
        byte[] second = writer.WriteToBytes(document);
        EmailReadResult parsed = new EmailDocumentReader().Read(first);

        Assert.Equal(first, second);
        Assert.Equal(document.Subject, parsed.Document.Subject);
        Assert.Equal(document.Body.Text, parsed.Document.Body.Text);
        Assert.Equal(document.Body.Html, parsed.Document.Body.Html);
        Assert.Equal(document.Attachments[0].FileName, parsed.Document.Attachments[0].FileName);
        Assert.Equal(document.Attachments[0].Content, parsed.Document.Attachments[0].Content);
        Assert.Equal("abc", parsed.Document.Headers.Single(header => header.Name == "X-Correlation").Value);
    }

    [Fact]
    public void PreservedSourceWritingIsExplicitAndVerbatim() {
        byte[] source = Encoding.ASCII.GetBytes("Subject: raw\n\nbody\n");
        EmailReaderOptions readerOptions = new EmailReaderOptions(preserveRawSource: true);
        EmailDocument document = new EmailDocumentReader(readerOptions).Read(source).Document;
        EmailDocumentWriter writer = new EmailDocumentWriter(new EmailWriterOptions(usePreservedRawSource: true));

        byte[] result = writer.WriteToBytes(document);

        Assert.Equal(source, result);
    }

    [Fact]
    public void WritesAndReadsEmbeddedMessages() {
        EmailDocument child = new EmailDocument { Format = EmailFileFormat.Eml, Subject = "Child" };
        child.Body.Text = "inside";
        EmailDocument parent = new EmailDocument { Format = EmailFileFormat.Eml, Subject = "Parent" };
        parent.Body.Text = "outside";
        parent.Attachments.Add(new EmailAttachment {
            FileName = "child.eml",
            ContentType = "message/rfc822",
            EmbeddedDocument = child
        });

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(parent);
        EmailDocument parsed = new EmailDocumentReader().Read(bytes).Document;

        Assert.Equal("Child", Assert.Single(parsed.Attachments).EmbeddedDocument!.Subject);
    }

    [Fact]
    public void AppliesConfiguredBase64LineLengthToTextBodies() {
        var document = new EmailDocument { Subject = "line length" };
        document.Body.Text = new string('x', 120);

        byte[] bytes = new EmailDocumentWriter(new EmailWriterOptions(base64LineLength: 20))
            .WriteToBytes(document);
        string[] lines = Encoding.ASCII.GetString(bytes).Split(new[] { "\r\n" }, StringSplitOptions.None);
        int bodyStart = Array.IndexOf(lines, string.Empty) + 1;
        string[] payloadLines = lines.Skip(bodyStart).Where(line => line.Length > 0).ToArray();

        Assert.True(payloadLines.Length > 1);
        Assert.All(payloadLines, line => Assert.InRange(line.Length, 1, 20));
    }

    [Fact]
    public void QuotesAddressSpecialsAndPreventsMessageIdHeaderInjection() {
        var document = new EmailDocument {
            From = new EmailAddress("john@example.com", "Doe, John"),
            MessageId = "safe@example.com\r\nX-Injected: yes"
        };
        document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("jane@example.com", "Doe, Jane")));
        document.Attachments.Add(new EmailAttachment {
            FileName = "logo.png",
            ContentType = "image/png",
            ContentId = "logo>\r\nX-Attachment-Injected: yes",
            Content = new byte[] { 1 },
            Length = 1,
            IsInline = true
        });
        document.Body.Text = "body";

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(document);
        string eml = Encoding.ASCII.GetString(bytes);
        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes).Document;

        Assert.Contains("From: \"Doe, John\" <john@example.com>\r\n", eml, StringComparison.Ordinal);
        Assert.Contains("To: \"Doe, Jane\" <jane@example.com>\r\n", eml, StringComparison.Ordinal);
        Assert.DoesNotContain("\r\nX-Injected:", eml, StringComparison.Ordinal);
        Assert.DoesNotContain("\r\nX-Attachment-Injected:", eml, StringComparison.Ordinal);
        Assert.Equal("Doe, John", roundTrip.From!.DisplayName);
        Assert.Equal("Doe, Jane", Assert.Single(roundTrip.Recipients).Address.DisplayName);
    }

    [Fact]
    public void FoldsLongInternationalHeadersIntoCompliantEncodedWords() {
        string subject = string.Concat(Enumerable.Repeat("日本語の長い件名", 12));
        var document = new EmailDocument { Subject = subject };
        document.Body.Text = "body";

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(document);
        string eml = Encoding.ASCII.GetString(bytes);
        MatchCollection words = Regex.Matches(eml, @"=\?utf-8\?B\?[^?]*\?=",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        EmailReadResult result = new EmailDocumentReader().Read(bytes);

        Assert.True(words.Count > 1);
        Assert.All(words.Cast<Match>(), word => Assert.InRange(word.Value.Length, 1, 75));
        Assert.Contains("\r\n =?utf-8?B?", eml, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(subject, result.Document.Subject);
    }

    [Fact]
    public void WritesRtfOnlyBodyAsAPreservedMimeAlternative() {
        const string rtf = "{\\rtf1\\ansi RTF-only body \\'e9\\par}";
        var document = new EmailDocument { Subject = "RTF body" };
        document.Body.Rtf = rtf;

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(
            document, EmailFileFormat.Eml, out EmailWriteResult writeResult);
        EmailReadResult readResult = new EmailDocumentReader().Read(bytes);

        Assert.Contains("Content-Type: text/rtf; charset=utf-8\r\n", Encoding.ASCII.GetString(bytes),
            StringComparison.Ordinal);
        Assert.False(writeResult.HasErrors);
        Assert.Equal(rtf, readResult.Document.Body.Rtf);
    }

    [Fact]
    public void WritesDateHeadersWithRfcNumericTimeZones() {
        var document = new EmailDocument {
            Date = new DateTimeOffset(2026, 7, 10, 15, 0, 0, TimeSpan.FromMinutes(150))
        };
        document.Body.Text = "body";

        string eml = Encoding.ASCII.GetString(new EmailDocumentWriter().WriteToBytes(document));

        Assert.Contains("Date: Fri, 10 Jul 2026 15:00:00 +0230\r\n", eml, StringComparison.Ordinal);
        Assert.DoesNotContain("+02:30", eml, StringComparison.Ordinal);
    }

    [Fact]
    public void RoundTripsQuotedPairsInDisplayNames() {
        const string displayName = "John \\\"JD\\\" \\\\ Smith";
        var document = new EmailDocument {
            From = new EmailAddress("john@example.com", displayName)
        };
        document.Body.Text = "body";

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().WriteToBytes(document)).Document;

        Assert.Equal(displayName, roundTrip.From!.DisplayName);
    }
}
