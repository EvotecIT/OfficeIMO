using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class MsgSerializationRegressionTests {
    [Fact]
    public void PreservesUnknownMessageFlagBitsWhileUpdatingManagedFlags() {
        const int unknownFlag = 0x20000000;
        var source = new EmailDocument { Format = EmailFileFormat.OutlookMsg, Subject = "Flags" };
        source.MapiProperties.Add(new MapiProperty(0x0E07, MapiPropertyType.Integer32, unknownFlag));
        source.MessageMetadata.IsDraft = true;

        EmailDocument result = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg)).Document;
        int flags = Assert.IsType<int>(result.MapiProperties.Single(property =>
            property.PropertyId == 0x0E07).Value);

        Assert.NotEqual(0, flags & unknownFlag);
        Assert.True(result.MessageMetadata.IsDraft);
    }

    [Fact]
    public void ClearingManagedValuesRemovesRetainedMsgProperties() {
        var source = new EmailDocument {
            Subject = "retained subject",
            MessageId = "retained@example.com"
        };
        source.Body.Text = "retained text";
        source.Body.Html = "<p>retained html</p>";
        source.MessageMetadata.OwnerReactionType = "like";
        source.MessageMetadata.Categories.Add("retained category");
        source.Headers.Add(new EmailHeader("X-Retained", "value"));
        byte[] first = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailDocument edited = new EmailDocumentReader().Read(first).Document;

        edited.Subject = null;
        edited.MessageId = null;
        edited.Body.Text = null;
        edited.Body.Html = null;
        edited.Body.Rtf = null;
        edited.MessageMetadata.OwnerReactionType = null;
        edited.MessageMetadata.Categories.Clear();
        edited.Headers.Clear();
        byte[] second = new EmailDocumentWriter().ToBytes(edited, EmailFileFormat.OutlookMsg);
        EmailDocument roundTrip = new EmailDocumentReader().Read(second).Document;

        Assert.Null(roundTrip.Subject);
        Assert.Null(roundTrip.MessageId);
        Assert.Null(roundTrip.Body.Text);
        Assert.Null(roundTrip.Body.Html);
        Assert.Null(roundTrip.Body.Rtf);
        Assert.Null(roundTrip.MessageMetadata.OwnerReactionType);
        Assert.Empty(roundTrip.MessageMetadata.Categories);
        Assert.Empty(roundTrip.Headers);
        Assert.DoesNotContain(roundTrip.MapiProperties, property =>
            property.Name == null && (property.PropertyId == 0x0037 || property.PropertyId == 0x003D ||
                property.PropertyId == 0x007D || property.PropertyId == 0x0E1D ||
                property.PropertyId == 0x1000 || property.PropertyId == 0x1013));
        Assert.DoesNotContain(roundTrip.MapiProperties, property =>
            property.Name?.PropertySet == MsgProjection.PsetidReactions &&
            string.Equals(property.Name.Name, "OwnerReactionType", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(roundTrip.MapiProperties, property =>
            property.Name?.PropertySet == MsgProjection.PsPublicStrings &&
            string.Equals(property.Name.Name, "Keywords", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ClearingAttachmentContentRemovesTheRetainedMsgPayload() {
        var source = new EmailDocument { Subject = "attachment" };
        source.Attachments.Add(new EmailAttachment {
            FileName = "payload.bin",
            Content = new byte[] { 1, 2, 3, 4 },
            Length = 4
        });
        EmailDocument edited = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg)).Document;
        EmailAttachment attachment = Assert.Single(edited.Attachments);
        attachment.Content = null;
        attachment.Length = 0;

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(edited, EmailFileFormat.OutlookMsg)).Document;

        EmailAttachment cleared = Assert.Single(roundTrip.Attachments);
        Assert.Null(cleared.Content);
        Assert.Equal(0, cleared.Length);
        Assert.DoesNotContain(cleared.MapiProperties, property => property.PropertyId == 0x3701);
    }

    [Fact]
    public void SanitizesTransportHeadersBeforeMsgSerialization() {
        var source = new EmailDocument { Subject = "headers" };
        source.Headers.Add(new EmailHeader(
            "X-Good\r\nBcc",
            "decoded",
            "ok\r\nBcc: injected@example.com"));

        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes).Document;
        MapiProperty transportHeaders = roundTrip.MapiProperties.Single(property => property.PropertyId == 0x007D);

        EmailHeader header = Assert.Single(roundTrip.Headers);
        Assert.Equal("X-GoodBcc", header.Name);
        Assert.Equal("ok Bcc: injected@example.com", header.Value);
        Assert.DoesNotContain("\r\nBcc:", Assert.IsType<string>(transportHeaders.Value), StringComparison.Ordinal);
        Assert.DoesNotContain(roundTrip.Headers, item =>
            string.Equals(item.Name, "Bcc", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void UsesTransportHeaderRecipientsWhenMsgRecipientStoragesAreMissing() {
        var source = new EmailDocument { Subject = "transport recipients" };
        source.Headers.Add(new EmailHeader("To", "\"Doe, John\" <john@example.com>"));
        source.Headers.Add(new EmailHeader("Cc", "Jane <jane@example.com>"));
        source.Headers.Add(new EmailHeader("Bcc", "Hidden <hidden@example.com>"));
        source.Headers.Add(new EmailHeader("Reply-To", "Replies <reply@example.com>"));

        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes).Document;

        Assert.Equal(4, roundTrip.Recipients.Count);
        Assert.Contains(roundTrip.Recipients, recipient => recipient.Kind == EmailRecipientKind.To &&
            recipient.Address.Address == "john@example.com" && recipient.Address.DisplayName == "Doe, John");
        Assert.Contains(roundTrip.Recipients, recipient => recipient.Kind == EmailRecipientKind.Cc &&
            recipient.Address.Address == "jane@example.com");
        Assert.Contains(roundTrip.Recipients, recipient => recipient.Kind == EmailRecipientKind.Bcc &&
            recipient.Address.Address == "hidden@example.com");
        Assert.Contains(roundTrip.Recipients, recipient => recipient.Kind == EmailRecipientKind.ReplyTo &&
            recipient.Address.Address == "reply@example.com");
    }

    [Fact]
    public void ReusesAggregateParserStateAcrossTnefAttachments() {
        byte[] tnef = CreateMinimalTnef();
        var source = new EmailDocument { Subject = "aggregate TNEF" };
        source.Attachments.Add(new EmailAttachment {
            FileName = "one.dat", ContentType = "application/ms-tnef", Content = tnef, Length = tnef.Length
        });
        source.Attachments.Add(new EmailAttachment {
            FileName = "two.dat", ContentType = "application/ms-tnef", Content = tnef, Length = tnef.Length
        });
        byte[] msg = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailDocumentReader(new EmailReaderOptions(maxTnefAttributeCount: 1)).Read(msg));

        Assert.Equal(nameof(EmailReaderOptions.MaxTnefAttributeCount), exception.LimitName);
        Assert.Equal(2, exception.ActualValue);
    }

    private static byte[] CreateMinimalTnef() {
        byte[] subject = Encoding.ASCII.GetBytes("nested\0");
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(TnefConstants.Signature);
            writer.Write((ushort)1);
            writer.Write((byte)TnefAttributeLevel.Message);
            writer.Write(TnefConstants.Subject);
            writer.Write(unchecked((uint)subject.Length));
            writer.Write(subject);
            writer.Write(unchecked((ushort)subject.Sum(value => value)));
        }
        return stream.ToArray();
    }
}
