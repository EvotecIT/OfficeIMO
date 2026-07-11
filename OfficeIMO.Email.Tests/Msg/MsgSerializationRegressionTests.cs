using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class MsgSerializationRegressionTests {
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
        byte[] first = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);
        EmailDocument edited = new EmailDocumentReader().Read(first).Document;

        edited.Subject = null;
        edited.MessageId = null;
        edited.Body.Text = null;
        edited.Body.Html = null;
        edited.Body.Rtf = null;
        edited.MessageMetadata.OwnerReactionType = null;
        edited.MessageMetadata.Categories.Clear();
        edited.Headers.Clear();
        byte[] second = new EmailDocumentWriter().WriteToBytes(edited, EmailFileFormat.OutlookMsg);
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
    public void SanitizesTransportHeadersBeforeMsgSerialization() {
        var source = new EmailDocument { Subject = "headers" };
        source.Headers.Add(new EmailHeader(
            "X-Good\r\nBcc",
            "decoded",
            "ok\r\nBcc: injected@example.com"));

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);
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
    public void ReusesAggregateParserStateAcrossTnefAttachments() {
        byte[] tnef = CreateMinimalTnef();
        var source = new EmailDocument { Subject = "aggregate TNEF" };
        source.Attachments.Add(new EmailAttachment {
            FileName = "one.dat", ContentType = "application/ms-tnef", Content = tnef, Length = tnef.Length
        });
        source.Attachments.Add(new EmailAttachment {
            FileName = "two.dat", ContentType = "application/ms-tnef", Content = tnef, Length = tnef.Length
        });
        byte[] msg = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

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
