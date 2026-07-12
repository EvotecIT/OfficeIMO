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
    public void PreservedSourceWritingRegeneratesAfterModelEdits() {
        byte[] source = Encoding.ASCII.GetBytes("Subject: original\r\n\r\noriginal body\r\n");
        EmailDocument document = new EmailDocumentReader(new EmailReaderOptions(preserveRawSource: true))
            .Read(source).Document;
        document.Subject = "edited";
        document.Body.Text = "edited body";
        document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("recipient@example.com")));
        document.Attachments.Add(new EmailAttachment {
            FileName = "edited.txt",
            ContentType = "text/plain",
            Content = Encoding.UTF8.GetBytes("attachment edit"),
            Length = Encoding.UTF8.GetByteCount("attachment edit")
        });
        var writer = new EmailDocumentWriter(new EmailWriterOptions(usePreservedRawSource: true));

        byte[] result = writer.WriteToBytes(document, EmailFileFormat.Eml, out EmailWriteResult writeResult);
        EmailDocument roundTrip = new EmailDocumentReader().Read(result).Document;

        Assert.False(writeResult.UsedPreservedSource);
        Assert.Contains(writeResult.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_RAW_SOURCE_SKIPPED_MODEL_CHANGED");
        Assert.Equal("edited", roundTrip.Subject);
        Assert.Equal("edited body", roundTrip.Body.Text!.Trim());
        Assert.Equal("recipient@example.com", Assert.Single(roundTrip.Recipients).Address.Address);
        Assert.Equal("edited.txt", Assert.Single(roundTrip.Attachments).FileName);
    }

    [Fact]
    public void PreservesUtf8AddressSpecsForInternationalizedEmail() {
        var document = new EmailDocument {
            Format = EmailFileFormat.Eml,
            From = new EmailAddress("josé@example.com")
        };
        document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("δοκιμή@example.com")));
        document.Body.Text = "body";

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(document);
        string eml = Encoding.UTF8.GetString(bytes);
        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes).Document;

        Assert.Contains("From: josé@example.com", eml, StringComparison.Ordinal);
        Assert.Contains("To: δοκιμή@example.com", eml, StringComparison.Ordinal);
        Assert.Equal(document.From.Address, roundTrip.From!.Address);
        Assert.Equal(document.Recipients[0].Address.Address, Assert.Single(roundTrip.Recipients).Address.Address);
    }

    [Fact]
    public void WritesAndReadsEmbeddedMessages() {
        EmailDocument child = new EmailDocument { Format = EmailFileFormat.Eml, Subject = "Child" };
        child.Body.Text = "inside";
        EmailDocument parent = new EmailDocument { Format = EmailFileFormat.Eml, Subject = "Parent" };
        parent.Body.Text = "outside";
        parent.Attachments.Add(new EmailAttachment {
            FileName = "child.eml",
            ContentType = "application/octet-stream",
            EmbeddedDocument = child
        });

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(parent);
        EmailDocument parsed = new EmailDocumentReader().Read(bytes).Document;

        Assert.Contains("Content-Type: message/rfc822;", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
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
    public void FoldsLongAsciiHeadersWithoutChangingTheirValues() {
        string subject = new string('a', 1200);
        var document = new EmailDocument { Subject = subject };
        document.Body.Text = "body";

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(document);
        string eml = Encoding.ASCII.GetString(bytes);
        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes).Document;

        Assert.All(eml.Split(new[] { "\r\n" }, StringSplitOptions.None),
            line => Assert.InRange(line.Length, 0, 998));
        Assert.Contains("Subject: =?utf-8?B?", eml, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(subject, roundTrip.Subject);
    }

    [Fact]
    public void PreservesRetainedStructuredHeaderSyntaxWithoutEncodedWords() {
        const string received = "from relay.example.test by destination.example.test with ESMTPS id 1234567890 for <reader@example.test>; Fri, 10 Jul 2026 12:30:00 +0000";
        const string signature = "v=1; a=rsa-sha256; d=example.test; s=mail; h=from:to:subject:date:message-id; bh=YWJjZGVmZ2hpamtsbW5vcA==; b=cXdlcnR5dWlvcGFzZGZnaGprbA==";
        var document = new EmailDocument { Subject = "structured headers" };
        document.Body.Text = "body";
        document.Headers.Add(new EmailHeader("Received", received, received));
        document.Headers.Add(new EmailHeader("DKIM-Signature", signature, signature));

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(document);
        string eml = Encoding.ASCII.GetString(bytes);
        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes).Document;

        Assert.DoesNotContain("Received: =?utf-8?", eml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("DKIM-Signature: =?utf-8?", eml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("\r\n ", eml, StringComparison.Ordinal);
        Assert.Equal(received, roundTrip.Headers.Single(header => header.Name == "Received").RawValue);
        Assert.Equal(signature, roundTrip.Headers.Single(header => header.Name == "DKIM-Signature").RawValue);
    }

    [Fact]
    public void FoldsLongRecipientListsWithoutChangingRecipients() {
        var document = new EmailDocument { Subject = "recipient folding" };
        document.Body.Text = "body";
        for (int index = 0; index < 24; index++) {
            document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
                new EmailAddress(string.Concat("recipient", index, "@example.com"),
                    string.Concat(new string('A', 90), index))));
        }

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(document);
        string eml = Encoding.ASCII.GetString(bytes);
        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes).Document;

        Assert.All(eml.Split(new[] { "\r\n" }, StringSplitOptions.None),
            line => Assert.InRange(line.Length, 0, 998));
        Assert.Contains("\r\n =?utf-8?B?", eml, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(document.Recipients.Count, roundTrip.Recipients.Count);
        Assert.Equal(document.Recipients.Select(recipient => recipient.Address.DisplayName),
            roundTrip.Recipients.Select(recipient => recipient.Address.DisplayName));
    }

    [Fact]
    public void FoldsLongUnicodeFileNamesIntoRfc2231Continuations() {
        string fileName = string.Concat(Enumerable.Repeat("資料-zażółć-", 40), "report.bin");
        var document = new EmailDocument { Subject = "filename continuations" };
        document.Body.Text = "body";
        document.Attachments.Add(new EmailAttachment {
            FileName = fileName,
            ContentType = "application/octet-stream",
            Content = new byte[] { 1, 2, 3 },
            Length = 3
        });

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(document);
        string eml = Encoding.ASCII.GetString(bytes);
        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes).Document;

        Assert.Contains("filename*0*=utf-8''", eml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("filename*1*=", eml, StringComparison.OrdinalIgnoreCase);
        Assert.All(eml.Split(new[] { "\r\n" }, StringSplitOptions.None),
            line => Assert.InRange(line.Length, 0, 998));
        Assert.Equal(fileName, Assert.Single(roundTrip.Attachments).FileName);
    }

    [Fact]
    public void EmitsRetainedMapiThreadingMetadataWhenTransportHeadersAreAbsent() {
        var source = new EmailDocument { Subject = "thread metadata" };
        source.Body.Text = "body";
        source.MessageMetadata.InternetReferences = "<root@example.test> <parent@example.test>";
        source.MessageMetadata.InReplyToId = "<parent@example.test>";
        byte[] msg = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);
        EmailDocument retained = new EmailDocumentReader().Read(msg).Document;

        byte[] emlBytes = new EmailDocumentWriter().WriteToBytes(retained, EmailFileFormat.Eml);
        string eml = Encoding.ASCII.GetString(emlBytes);
        EmailDocument roundTrip = new EmailDocumentReader().Read(emlBytes).Document;

        Assert.Contains("References: <root@example.test> <parent@example.test>\r\n", eml,
            StringComparison.Ordinal);
        Assert.Contains("In-Reply-To: <parent@example.test>\r\n", eml, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Headers, header => header.Name == "References" &&
            header.Value == "<root@example.test> <parent@example.test>");
    }

    [Fact]
    public void GroupsReferencedCidResourcesInsideMultipartRelated() {
        var document = new EmailDocument { Subject = "related resources" };
        document.Body.Html = "<html><img src=\"cid:logo\"></html>";
        document.Attachments.Add(new EmailAttachment {
            FileName = "logo.png",
            ContentType = "image/png",
            ContentId = "logo",
            IsInline = true,
            Content = new byte[] { 1, 2 },
            Length = 2
        });
        document.Attachments.Add(new EmailAttachment {
            FileName = "notes.txt",
            ContentType = "text/plain",
            Content = Encoding.UTF8.GetBytes("notes"),
            Length = 5
        });

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(document);
        string eml = Encoding.ASCII.GetString(bytes);
        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes).Document;

        Assert.Contains("Content-Type: multipart/mixed", eml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Content-Type: multipart/related", eml, StringComparison.OrdinalIgnoreCase);
        Assert.True(eml.IndexOf("multipart/related", StringComparison.OrdinalIgnoreCase) <
            eml.IndexOf("Content-ID: <logo>", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(document.Body.Html, roundTrip.Body.Html);
        Assert.Contains(roundTrip.Attachments, attachment => attachment.ContentId == "logo" && attachment.IsInline);
        Assert.Contains(roundTrip.Attachments, attachment => attachment.FileName == "notes.txt");
    }

    [Fact]
    public void GroupsOnlyTheExactPercentDecodedCidResource() {
        var document = new EmailDocument { Subject = "exact related resource" };
        document.Body.Html = "<html><img src=\"cid:logo%40example.com\"></html>";
        document.Attachments.Add(new EmailAttachment {
            FileName = "exact.png",
            ContentType = "image/png",
            ContentId = "logo@example.com",
            IsInline = true,
            Content = new byte[] { 1 },
            Length = 1
        });
        document.Attachments.Add(new EmailAttachment {
            FileName = "prefix.png",
            ContentType = "image/png",
            ContentId = "logo",
            IsInline = true,
            Content = new byte[] { 2 },
            Length = 1
        });

        string eml = Encoding.ASCII.GetString(new EmailDocumentWriter().WriteToBytes(document));
        int relatedHeader = eml.IndexOf("multipart/related; boundary=\"", StringComparison.OrdinalIgnoreCase);
        Assert.True(relatedHeader >= 0);
        int boundaryStart = relatedHeader + "multipart/related; boundary=\"".Length;
        int boundaryEnd = eml.IndexOf('"', boundaryStart);
        string closingBoundary = string.Concat("--", eml.Substring(boundaryStart, boundaryEnd - boundaryStart), "--");
        int relatedEnd = eml.IndexOf(closingBoundary, boundaryEnd, StringComparison.Ordinal);
        int exactId = eml.IndexOf("Content-ID: <logo@example.com>", StringComparison.OrdinalIgnoreCase);
        int prefixId = eml.IndexOf("Content-ID: <logo>", StringComparison.OrdinalIgnoreCase);

        Assert.InRange(exactId, relatedHeader, relatedEnd);
        Assert.True(prefixId > relatedEnd);
    }

    [Fact]
    public void GroupsReferencedContentLocationResourcesInsideMultipartRelated() {
        var document = new EmailDocument { Subject = "location related resource" };
        document.Body.Html = "<picture><source srcset=\"image001.png 1x, image002.png 2x\"></picture>";
        document.Attachments.Add(new EmailAttachment {
            FileName = "image001.png",
            ContentType = "image/png",
            ContentLocation = "image001.png",
            IsInline = true,
            Content = new byte[] { 1 },
            Length = 1
        });
        document.Attachments.Add(new EmailAttachment {
            FileName = "unreferenced.png",
            ContentType = "image/png",
            ContentLocation = "image001.png.bak",
            IsInline = true,
            Content = new byte[] { 2 },
            Length = 1
        });

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(document);
        string eml = Encoding.ASCII.GetString(bytes);
        int relatedHeader = eml.IndexOf("multipart/related; boundary=\"", StringComparison.OrdinalIgnoreCase);
        Assert.True(relatedHeader >= 0);
        int boundaryStart = relatedHeader + "multipart/related; boundary=\"".Length;
        int boundaryEnd = eml.IndexOf('"', boundaryStart);
        string closingBoundary = string.Concat("--", eml.Substring(boundaryStart, boundaryEnd - boundaryStart), "--");
        int relatedEnd = eml.IndexOf(closingBoundary, boundaryEnd, StringComparison.Ordinal);
        int referenced = eml.IndexOf("Content-Location: image001.png\r\n", StringComparison.OrdinalIgnoreCase);
        int unreferenced = eml.IndexOf("Content-Location: image001.png.bak\r\n", StringComparison.OrdinalIgnoreCase);

        Assert.InRange(referenced, relatedHeader, relatedEnd);
        Assert.True(unreferenced > relatedEnd);
        EmailAttachment roundTrip = new EmailDocumentReader().Read(bytes).Document.Attachments
            .Single(attachment => attachment.ContentLocation == "image001.png");
        Assert.True(roundTrip.IsInline);
    }

    [Fact]
    public void WritesRtfOnlyBodyAsAPreservedMimeAlternative() {
        string rtf = string.Concat("{\\rtf1\\ansi RTF-only body ", (char)0xE9, "\\par}");
        var document = new EmailDocument { Subject = "RTF body" };
        document.Body.Rtf = rtf;

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(
            document, EmailFileFormat.Eml, out EmailWriteResult writeResult);
        EmailReadResult readResult = new EmailDocumentReader().Read(bytes);

        Assert.Contains("Content-Type: text/rtf; charset=iso-8859-1\r\n", Encoding.ASCII.GetString(bytes),
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
