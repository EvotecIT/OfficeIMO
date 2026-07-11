using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class TnefResourceLimitTests {
    [Fact]
    public void RejectsEmbeddedMessagePayloadBeyondTheAttachmentLimit() {
        var child = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "embedded" };
        child.Body.Text = new string('x', 2048);
        var parent = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "parent" };
        parent.Attachments.Add(new EmailAttachment { FileName = "child.dat", EmbeddedDocument = child });
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(parent, EmailFileFormat.Tnef);

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailDocumentReader(new EmailReaderOptions(maxAttachmentBytes: 512)).Read(bytes));

        Assert.Equal(nameof(EmailReaderOptions.MaxAttachmentBytes), exception.LimitName);
        Assert.True(exception.ActualValue > 512);
    }

    [Fact]
    public void RejectsAttachDataBeforeCopyingAnOversizedAttribute() {
        byte[] bytes = CreateTnefAttachment(new byte[1024]);

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailDocumentReader(new EmailReaderOptions(
                maxAttachmentBytes: 512,
                includeAttachmentContent: false)).Read(bytes));

        Assert.Equal(nameof(EmailReaderOptions.MaxAttachmentBytes), exception.LimitName);
        Assert.Equal(1024, exception.ActualValue);
    }

    [Fact]
    public void RejectsMapiValueBeforeCopyingBeyondTheDecodedPropertyLimit() {
        byte[] properties = TnefMapiCodec.WriteProperties(new[] {
            new MapiProperty(0x66AA, MapiPropertyType.Binary, new byte[1024])
        }, 1252, new List<EmailDiagnostic>(), "tnef/mapi");
        byte[] bytes = CreateTnef((TnefAttributeLevel.Message, TnefConstants.MessageProperties, properties));

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailDocumentReader(new EmailReaderOptions(maxDecodedPropertyBytes: 512)).Read(bytes));

        Assert.Equal(nameof(EmailReaderOptions.MaxDecodedPropertyBytes), exception.LimitName);
        Assert.Equal(1024, exception.ActualValue);
    }

    [Fact]
    public void PrefersMapiPayloadWhenAttachDataAndMapiContentDiffer() {
        byte[] mapiPayload = { 3, 4, 5 };
        byte[] properties = TnefMapiCodec.WriteProperties(new[] {
            new MapiProperty(0x3701, MapiPropertyType.Binary, mapiPayload),
            new MapiProperty(0x3705, MapiPropertyType.Integer32, 1)
        }, 1252, new List<EmailDiagnostic>(), "tnef/attachment/mapi");
        byte[] bytes = CreateTnef(
            (TnefAttributeLevel.Attachment, TnefConstants.AttachRendData, new byte[14]),
            (TnefAttributeLevel.Attachment, TnefConstants.AttachData, new byte[] { 1, 2 }),
            (TnefAttributeLevel.Attachment, TnefConstants.AttachmentProperties, properties));

        EmailAttachment attachment = Assert.Single(new EmailDocumentReader().Read(bytes).Document.Attachments);

        Assert.Equal(mapiPayload, attachment.Content);
        Assert.Equal(mapiPayload.Length, attachment.Length);
    }

    [Fact]
    public void RoundTripsLegacyCodePageAcrossTnefAttributesAndMapiBytes() {
        var source = new EmailDocument {
            Format = EmailFileFormat.Tnef,
            OutlookCodePage = 932,
            Subject = "日本語の件名"
        };
        source.Body.Html = "<p>日本語の本文</p>";

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef)).Document;

        Assert.Equal(932, roundTrip.OutlookCodePage);
        Assert.Equal(source.Subject, roundTrip.Subject);
        Assert.Equal(source.Body.Html, roundTrip.Body.Html);
    }

    private static byte[] CreateTnefAttachment(byte[] payload) {
        return CreateTnef(
            (TnefAttributeLevel.Attachment, TnefConstants.AttachRendData, new byte[14]),
            (TnefAttributeLevel.Attachment, TnefConstants.AttachData, payload));
    }

    private static byte[] CreateTnef(params (TnefAttributeLevel Level, uint Tag, byte[] Data)[] attributes) {
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(TnefConstants.Signature);
            writer.Write((ushort)1);
            foreach ((TnefAttributeLevel level, uint tag, byte[] data) in attributes) {
                writer.Write((byte)level);
                writer.Write(tag);
                writer.Write(unchecked((uint)data.Length));
                writer.Write(data);
                writer.Write(unchecked((ushort)data.Sum(value => value)));
            }
        }
        return stream.ToArray();
    }
}
