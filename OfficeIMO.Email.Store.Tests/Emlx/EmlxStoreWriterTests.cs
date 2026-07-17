using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmlxStoreWriterTests {
    [Fact]
    public void WriteThenReadPreservesMessageAndAppleMetadata() {
        var document = new EmailDocument {
            Subject = "EMLX write test",
            MessageId = "emlx-write@example.com",
            Date = new DateTimeOffset(2026, 7, 17, 8, 30, 0, TimeSpan.Zero),
            ReceivedDate = new DateTimeOffset(2026, 7, 17, 8, 31, 0, TimeSpan.Zero)
        };
        document.Body.Text = "EMLX body";
        document.MessageMetadata.IsRead = true;
        document.MessageMetadata.IsDraft = true;
        document.Properties["Emlx:Flag:Flagged"] = true;
        document.Properties["Emlx:Flag:PriorityLevel"] = 73;

        byte[] bytes = new EmailStoreEmlxWriter().ToBytes(document);
        using var stream = new MemoryStream(bytes);
        EmailStoreReadResult result = new EmailStoreReader().Read(stream, "written.emlx");
        EmailDocument loaded = result.Store.Folders.Single().Items.Single().Document;

        Assert.Equal(EmailStoreFormat.Emlx, result.Store.Format);
        Assert.Equal("EMLX write test", loaded.Subject);
        Assert.Equal("emlx-write@example.com", loaded.MessageId);
        Assert.Equal(document.ReceivedDate, loaded.ReceivedDate);
        Assert.True(loaded.MessageMetadata.IsRead);
        Assert.True(loaded.MessageMetadata.IsDraft);
        Assert.Equal(true, loaded.Properties["Emlx:Flag:Flagged"]);
        Assert.Equal(73, loaded.Properties["Emlx:Flag:PriorityLevel"]);
    }

    [Fact]
    public void RewriteUsesPreservedAttachmentCountWhenAttachmentsAreNotMaterialized() {
        var document = new EmailDocument { Subject = "Partial EMLX" };
        document.Properties["Emlx:Flag:AttachmentCount"] = 37;
        document.Properties["Emlx:IsPartial"] = true;

        byte[] bytes = new EmailStoreEmlxWriter().ToBytes(document);
        using var stream = new MemoryStream(bytes);
        EmailDocument loaded = new EmailStoreReader().Read(stream, "partial.emlx")
            .Store.Folders.Single().Items.Single().Document;

        Assert.Empty(loaded.Attachments);
        Assert.Equal(37, loaded.Properties["Emlx:Flag:AttachmentCount"]);
    }

    [Fact]
    public void RewriteUsesMaterializedAttachmentCountAfterDocumentEdits() {
        var document = new EmailDocument { Subject = "Edited EMLX" };
        document.Properties["Emlx:Flag:AttachmentCount"] = 37;
        document.Attachments.Add(new EmailAttachment {
            FileName = "current.bin",
            ContentType = "application/octet-stream",
            Content = new byte[] { 1, 2, 3 },
            Length = 3
        });

        byte[] bytes = new EmailStoreEmlxWriter().ToBytes(document);
        using var stream = new MemoryStream(bytes);
        EmailDocument loaded = new EmailStoreReader().Read(stream, "edited.emlx")
            .Store.Folders.Single().Items.Single().Document;

        Assert.Single(loaded.Attachments);
        Assert.Equal(1, loaded.Properties["Emlx:Flag:AttachmentCount"]);
    }

    [Fact]
    public void WriterEnforcesCompleteArtifactLimit() {
        var document = new EmailDocument { Subject = "Bounded EMLX" };
        document.Body.Text = new string('x', 2048);
        var writer = new EmailStoreEmlxWriter(new EmailStoreEmlxWriterOptions(
            includeMetadata: false, maxOutputBytes: 128));

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            writer.ToBytes(document));

        Assert.Equal(nameof(EmailStoreEmlxWriterOptions.MaxOutputBytes), exception.LimitName);
    }

    [Fact]
    public void WriterBoundsMetadataBeforeMaterializingTheMessage() {
        var document = new EmailDocument { Subject = new string('x', 2048) };
        var writer = new EmailStoreEmlxWriter(new EmailStoreEmlxWriterOptions(maxOutputBytes: 128));

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            writer.ToBytes(document));

        Assert.Equal(nameof(EmailStoreEmlxWriterOptions.MaxOutputBytes), exception.LimitName);
        Assert.True(exception.ActualValue > exception.MaximumValue);
    }

    [Fact]
    public void WriterNormalizesXmlForbiddenMetadataTextToInvalidDataException() {
        var document = new EmailDocument { Subject = "invalid\u0001subject" };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            new EmailStoreEmlxWriter().ToBytes(document));

        Assert.Contains("XML cannot represent", exception.Message, StringComparison.Ordinal);
    }
}
