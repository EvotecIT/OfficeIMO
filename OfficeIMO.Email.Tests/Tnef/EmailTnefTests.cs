using MimeKit.Tnef;
using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailTnefTests {
    [Fact]
    public void RoundTripsMessageMapiRecipientsAndAttachmentKinds() {
        DateTimeOffset start = new DateTimeOffset(2026, 10, 3, 9, 0, 0, TimeSpan.Zero);
        EmailDocument child = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "child" };
        child.Body.Text = "nested";
        EmailDocument source = new EmailDocument {
            Format = EmailFileFormat.Tnef,
            OutlookItemKind = OutlookItemKind.Appointment,
            Subject = "TNEF subject",
            MessageId = "tnef@example.com",
            Date = start,
            Appointment = new OutlookAppointment { Start = start, End = start.AddHours(1), Location = "Room" }
        };
        source.Body.Text = "TNEF body";
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("to@example.com", "To")));
        source.MapiProperties.Add(new MapiProperty(0x66AA, MapiPropertyType.MultipleUnicode, new object[] { "one", "two" }));
        source.TnefAttributes.Add(new TnefAttribute(TnefAttributeLevel.Message, 0x0006F001, new byte[] { 7, 8 }));
        source.Attachments.Add(new EmailAttachment {
            FileName = "data.bin", ContentType = "application/octet-stream", Content = new byte[] { 1, 2, 3 }, Length = 3
        });
        source.Attachments.Add(new EmailAttachment { FileName = "child.dat", EmbeddedDocument = child });
        var storage = new EmailAttachment { FileName = "ole.dat", MapiAttachMethod = 6 };
        storage.StructuredStorageStreams["Contents"] = new byte[] { 9, 8, 7 };
        source.Attachments.Add(storage);

        byte[] first = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);
        byte[] second = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);
        EmailReadResult result = new EmailDocumentReader().Read(first);

        Assert.Equal(first, second);
        Assert.Equal(EmailFileFormat.Tnef, result.Document.Format);
        Assert.Equal("TNEF subject", result.Document.Subject);
        Assert.Equal("TNEF body", result.Document.Body.Text);
        Assert.True(result.Document.Appointment != null,
            string.Join(Environment.NewLine, result.Diagnostics.Select(diagnostic => string.Concat(diagnostic.Code, ": ", diagnostic.Message))));
        Assert.Equal("Room", result.Document.Appointment!.Location);
        Assert.Equal("to@example.com", Assert.Single(result.Document.Recipients).Address.Address);
        Assert.Equal(new byte[] { 1, 2, 3 }, result.Document.Attachments[0].Content);
        Assert.Equal("child", result.Document.Attachments[1].EmbeddedDocument!.Subject);
        Assert.Equal(new byte[] { 9, 8, 7 }, result.Document.Attachments[2].StructuredStorageStreams["Contents"]);
        Assert.Contains(result.Document.TnefAttributes, attribute => attribute.Tag == 0x0006F001);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void OutputIsAcceptedByMimeKitTnefReaderOracle() {
        EmailDocument source = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "oracle" };
        source.Body.Text = "body";
        source.Attachments.Add(new EmailAttachment { FileName = "a.txt", Content = Encoding.UTF8.GetBytes("a"), Length = 1 });
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);

        using MemoryStream stream = new MemoryStream(bytes);
        using var reader = new global::MimeKit.Tnef.TnefReader(stream);
        int count = 0;
        var compliance = new List<string>();
        while (reader.ReadNextAttribute()) {
            count++;
            if (reader.ComplianceStatus != TnefComplianceStatus.Compliant) {
                compliance.Add(string.Concat(reader.AttributeTag.ToString(), ": ", reader.ComplianceStatus.ToString()));
                reader.ResetComplianceStatus();
            }
        }

        Assert.True(count >= 6);
        Assert.Equal(0x00010000, reader.TnefVersion);
        Assert.True(compliance.Count == 0, string.Join(Environment.NewLine, compliance));
    }

    [Fact]
    public void RoundTripsCompressedRtfThroughTnefMapiProperties() {
        const string rtf = "{\\rtf1\\ansi TNEF RTF body\\par}";
        var source = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "rtf" };
        source.Body.Rtf = rtf;

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);
        EmailReadResult result = new EmailDocumentReader().Read(bytes);

        Assert.Equal(rtf, result.Document.Body.Rtf);
        Assert.Contains(result.Document.MapiProperties, property => property.PropertyId == 0x1009);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void ReportsChecksumDamageAndEnforcesAttributeLimit() {
        EmailDocument source = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "checksum" };
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);
        bytes[bytes.Length - 1] ^= 0x01;

        EmailReadResult damaged = new EmailDocumentReader().Read(bytes);

        Assert.Contains(damaged.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_TNEF_CHECKSUM_MISMATCH");
        Assert.Throws<EmailLimitExceededException>(() => new EmailDocumentReader(
            new EmailReaderOptions(maxTnefAttributeCount: 1)).Read(bytes));
    }

    [Fact]
    public void TruncatedMapiPropertyRowsReturnDiagnostics() {
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(TnefConstants.Signature);
            writer.Write((ushort)1);
            writer.Write((byte)TnefAttributeLevel.Message);
            writer.Write(TnefConstants.MessageProperties);
            writer.Write(4U);
            writer.Write(1U);
            writer.Write((ushort)1);
        }

        EmailReadResult result = new EmailDocumentReader().Read(stream.ToArray());

        Assert.Equal(EmailFileFormat.Tnef, result.Document.Format);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_TNEF_MAPI_TRUNCATED" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }
}
