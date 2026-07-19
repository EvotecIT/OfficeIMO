using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmailStoreItemContentAvailabilityTests {
    [Fact]
    public void MarksMissingHeaderOnlyOstContentUnavailable() {
        EmailDocument document = CreateDocumentWithMissingAttachment();
        document.MapiProperties.Add(new MapiProperty(
            0x8001, MapiPropertyType.Boolean, true,
            name: new MapiNamedProperty(
                EmailStoreItemContentAvailability.PsetidCommon, 0x8578)));
        EmailStoreItemReadParts requested = EmailStoreItemReadParts.Metadata |
            EmailStoreItemReadParts.Bodies |
            EmailStoreItemReadParts.AttachmentMetadata |
            EmailStoreItemReadParts.AttachmentContent;

        var item = new EmailStoreItem(
            "pst:00008004", "pst:00008022", document,
            loadedParts: requested, format: EmailStoreFormat.Ost);

        Assert.True(item.ContentAvailability.IsHeaderOnly);
        Assert.True(item.ContentAvailability.IsPotentiallyPartial);
        Assert.True(item.ContentAvailability.AvailableParts.HasFlag(EmailStoreItemReadParts.Metadata));
        Assert.True(item.ContentAvailability.AvailableParts.HasFlag(EmailStoreItemReadParts.AttachmentMetadata));
        Assert.True(item.ContentAvailability.UnavailableParts.HasFlag(EmailStoreItemReadParts.Bodies));
        Assert.True(item.ContentAvailability.UnavailableParts.HasFlag(EmailStoreItemReadParts.AttachmentContent));
        Assert.Equal(EmailStoreItemReadParts.None, item.ContentAvailability.IndeterminateParts);
    }

    [Fact]
    public void KeepsMissingOstContentIndeterminateWithoutAHeaderOnlySignal() {
        EmailDocument document = CreateDocumentWithMissingAttachment();
        document.MapiProperties.Add(new MapiProperty(
            0x0E17, MapiPropertyType.Integer32, 0x00001000));
        EmailStoreItemReadParts requested = EmailStoreItemReadParts.Metadata |
            EmailStoreItemReadParts.Bodies |
            EmailStoreItemReadParts.AttachmentMetadata |
            EmailStoreItemReadParts.AttachmentContent;

        var item = new EmailStoreItem(
            "pst:00008004", "pst:00008022", document,
            loadedParts: requested, format: EmailStoreFormat.Ost);

        Assert.Null(item.ContentAvailability.IsHeaderOnly);
        Assert.True(item.ContentAvailability.IsMarkedForDownload);
        Assert.True(item.ContentAvailability.IndeterminateParts.HasFlag(EmailStoreItemReadParts.Bodies));
        Assert.True(item.ContentAvailability.IndeterminateParts.HasFlag(
            EmailStoreItemReadParts.AttachmentContent));
        Assert.Equal(EmailStoreItemReadParts.None, item.ContentAvailability.UnavailableParts);
    }

    [Fact]
    public void TreatsAnEmptyBodyAsCompleteInANonCachedStore() {
        var item = new EmailStoreItem(
            "pst:00008004", "pst:00008022", new EmailDocument(),
            loadedParts: EmailStoreItemReadParts.Metadata | EmailStoreItemReadParts.Bodies,
            format: EmailStoreFormat.Pst);

        Assert.False(item.ContentAvailability.IsHeaderOnly);
        Assert.False(item.ContentAvailability.IsPotentiallyPartial);
        Assert.True(item.ContentAvailability.AvailableParts.HasFlag(EmailStoreItemReadParts.Bodies));
        Assert.Equal(EmailStoreItemReadParts.None, item.ContentAvailability.UnavailableParts);
        Assert.Equal(EmailStoreItemReadParts.None, item.ContentAvailability.IndeterminateParts);
    }

    private static EmailDocument CreateDocumentWithMissingAttachment() {
        var document = new EmailDocument();
        document.Attachments.Add(new EmailAttachment {
            FileName = "remote.bin",
            Length = 4096,
            MapiAttachMethod = 1
        });
        return document;
    }
}
