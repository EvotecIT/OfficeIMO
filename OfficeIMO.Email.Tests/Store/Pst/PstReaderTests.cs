using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstReaderTests {
    [Fact]
    public void DetectsPstAndOstFromTheNdbClientSignatureWithoutConsumingTheStream() {
        using var pst = new MemoryStream(PstTestFileBuilder.Create());
        using var ost = new MemoryStream(PstTestFileBuilder.Create(ost: true));
        pst.Position = 7;
        ost.Position = 9;

        Assert.Equal(EmailStoreFormat.Pst, EmailStoreReader.DetectFormat(pst, "wrong.ost"));
        Assert.Equal(EmailStoreFormat.Ost, EmailStoreReader.DetectFormat(ost, "wrong.pst"));
        Assert.Equal(7, pst.Position);
        Assert.Equal(9, ost.Position);
    }

    [Fact]
    public void ReadsUnicodePstHierarchyAndProjectsMessagesThroughOfficeImoEmail() {
        using var stream = new MemoryStream(PstTestFileBuilder.Create());

        EmailStoreReadResult result = new EmailStoreReader().Read(stream, "mailbox.pst");

        Assert.False(result.HasErrors);
        Assert.Equal(EmailStoreFormat.Pst, result.Store.Format);
        Assert.Equal("Test Store", result.Store.DisplayName);
        Assert.Equal(2, result.Store.Folders.Count);
        EmailStoreFolder root = Assert.Single(result.Store.RootFolders);
        Assert.Equal("Root", root.Name);
        Assert.Equal(EmailStoreSpecialFolderKind.Root, root.SpecialFolderKind);
        Assert.Equal(EmailStoreFolderClassificationSource.SourceIdentifier, root.ClassificationSource);
        EmailStoreFolder inbox = Assert.Single(result.Store.Folders, folder => folder.Name == "Inbox");
        Assert.Equal(root.Id, inbox.ParentId);
        Assert.Equal(EmailStoreSpecialFolderKind.Inbox, inbox.SpecialFolderKind);
        Assert.Equal(EmailStoreFolderClassificationSource.DisplayName, inbox.ClassificationSource);
        Assert.False(inbox.IsSearchFolder);
        EmailStoreItem message = Assert.Single(inbox.Items);
        Assert.Equal("Synthetic PST message", message.Document.Subject);
        Assert.Equal("Body from the PST property context", message.Document.Body.Text);
        Assert.Equal("IPM.Note", message.Document.MessageClass);
        Assert.Equal("Pst", message.Document.Properties["EmailStore:Format"]);
    }

    [Fact]
    public void ReadsAnsiPstHierarchyAndProjectsMessagesThroughOfficeImoEmail() {
        using var stream = new MemoryStream(PstTestFileBuilder.Create(ansi: true));

        EmailStoreReadResult result = new EmailStoreReader().Read(stream, "mailbox.pst");

        Assert.False(result.HasErrors);
        Assert.Equal("Test Store", result.Store.DisplayName);
        EmailStoreFolder inbox = Assert.Single(result.Store.Folders, folder => folder.Name == "Inbox");
        EmailStoreItem message = Assert.Single(inbox.Items);
        Assert.Equal("Synthetic PST message", message.Document.Subject);
        Assert.Equal("Body from the PST property context", message.Document.Body.Text);
    }

    [Fact]
    public void ReadsCyclicEncodedPstDataBlocks() {
        using var stream = new MemoryStream(PstTestFileBuilder.Create(cryptMethod: 2));

        EmailStoreReadResult result = new EmailStoreReader().Read(stream, "mailbox.pst");

        Assert.False(result.HasErrors);
        EmailStoreItem message = Assert.Single(
            Assert.Single(result.Store.Folders, folder => folder.Name == "Inbox").Items);
        Assert.Equal("Synthetic PST message", message.Document.Subject);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReadsModernFourKilobyteOstBlocks(bool compressed) {
        using var stream = new MemoryStream(PstTestFileBuilder.Create(
            ost: true, fourK: true, compressBlocks: compressed));

        EmailStoreReadResult result = new EmailStoreReader().Read(stream, "mailbox.ost");

        Assert.False(result.HasErrors);
        Assert.Equal(EmailStoreFormat.Ost, result.Store.Format);
        EmailStoreItem message = Assert.Single(
            Assert.Single(result.Store.Folders, folder => folder.Name == "Inbox").Items);
        Assert.Equal("Synthetic PST message", message.Document.Subject);
        Assert.Equal("Body from the PST property context", message.Document.Body.Text);
    }

    [Fact]
    public void ProjectsEmbeddedPstMessagesFromAttachmentSubnodes() {
        using var stream = new MemoryStream(PstTestFileBuilder.Create(includeEmbeddedMessage: true));

        EmailStoreReadResult result = new EmailStoreReader().Read(stream, "mailbox.pst");

        Assert.False(result.HasErrors);
        EmailStoreItem message = Assert.Single(
            Assert.Single(result.Store.Folders, folder => folder.Name == "Inbox").Items);
        EmailAttachment attachment = Assert.Single(message.Document.Attachments);
        Assert.Equal(5, attachment.MapiAttachMethod);
        Assert.Equal("forwarded.msg", attachment.FileName);
        Assert.Null(attachment.Content);
        Assert.True(attachment.Length > 0);
        Assert.NotNull(attachment.EmbeddedDocument);
        Assert.Equal("Embedded PST message", attachment.EmbeddedDocument!.Subject);
        Assert.Equal("Body from the embedded PST item", attachment.EmbeddedDocument.Body.Text);
    }

    [Fact]
    public void EmbeddedPstObjectAttachmentsHonorAttachmentByteLimit() {
        using var stream = new MemoryStream(PstTestFileBuilder.Create(
            includeEmbeddedMessage: true));
        var reader = new EmailStoreReader(new EmailStoreReaderOptions(
            maxAttachmentBytes: 16,
            maxTotalAttachmentBytes: 1024));

        Assert.Throws<EmailStoreLimitExceededException>(() =>
            reader.Read(stream, "mailbox.pst"));
    }

    [Fact]
    public void PreservesEmbeddedAttachmentMetadataWhenDepthLimitIsZero() {
        using var stream = new MemoryStream(PstTestFileBuilder.Create(includeEmbeddedMessage: true));
        var reader = new EmailStoreReader(new EmailStoreReaderOptions(maxNestedMessageDepth: 0));

        EmailStoreReadResult result = reader.Read(stream, "mailbox.pst");

        EmailAttachment attachment = Assert.Single(
            Assert.Single(result.Store.Folders, folder => folder.Name == "Inbox").Items.Single().Document.Attachments);
        Assert.Null(attachment.EmbeddedDocument);
        Assert.Contains(result.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_PST_EMBEDDED_DEPTH_LIMIT");
    }

    [Fact]
    public void SelectiveReadsSkipBodiesRecipientsAndAttachments() {
        using var stream = new MemoryStream(PstTestFileBuilder.Create(
            attachmentContent: Encoding.UTF8.GetBytes("private payload")));
        using EmailStoreSession session = EmailStoreSession.Open(stream, "mailbox.pst");
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());

        EmailStoreItem item = session.ReadItem(reference,
            new EmailStoreItemReadOptions(EmailStoreItemReadParts.Metadata));

        Assert.Equal(EmailStoreItemReadParts.Metadata, item.LoadedParts);
        Assert.Equal("Synthetic PST message", item.Document.Subject);
        Assert.Null(item.Document.Body.Text);
        Assert.Empty(item.Document.Recipients);
        Assert.Empty(item.Document.Attachments);
    }

    [Fact]
    public void PreservesFolderClassificationWhenASessionMaterializesTheStore() {
        using var stream = new MemoryStream(PstTestFileBuilder.Create());
        using EmailStoreSession session = EmailStoreSession.Open(stream, "mailbox.pst");

        EmailStoreReadResult result = session.ReadAll();

        EmailStoreFolder root = Assert.Single(result.Store.Folders,
            folder => folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Root);
        Assert.Equal(EmailStoreFolderClassificationSource.SourceIdentifier,
            root.ClassificationSource);
        EmailStoreFolder inbox = Assert.Single(result.Store.Folders,
            folder => folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Inbox);
        Assert.Equal(EmailStoreFolderClassificationSource.DisplayName,
            inbox.ClassificationSource);
    }

    [Fact]
    public void StreamsPstAttachmentContentWhileSessionIsOpen() {
        byte[] expected = Enumerable.Range(0, 60_000).Select(index => (byte)(index % 251)).ToArray();
        var options = new EmailStoreReaderOptions(retainAttachmentContent: false);
        using var stream = new MemoryStream(PstTestFileBuilder.Create(attachmentContent: expected));
        EmailAttachment attachment;
        using (EmailStoreSession session = EmailStoreSession.Open(stream, "mailbox.pst", options)) {
            EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
            EmailStoreItem item = session.ReadItem(reference, new EmailStoreItemReadOptions(
                EmailStoreItemReadParts.Metadata |
                EmailStoreItemReadParts.AttachmentMetadata |
                EmailStoreItemReadParts.AttachmentContent));

            attachment = Assert.Single(item.Document.Attachments);
            Assert.Null(attachment.Content);
            Assert.NotNull(attachment.ContentSource);
            Assert.Equal(expected.LongLength, attachment.Length);
            Assert.True(item.ContentAvailability.AvailableParts.HasFlag(
                EmailStoreItemReadParts.AttachmentContent));
            Assert.Equal(EmailStoreItemReadParts.None,
                item.ContentAvailability.UnavailableParts & EmailStoreItemReadParts.AttachmentContent);
            using Stream payload = attachment.OpenContentStream();
            using var copy = new MemoryStream();
            payload.CopyTo(copy);
            Assert.Equal(expected, copy.ToArray());
        }

        Assert.Throws<ObjectDisposedException>(() => attachment.OpenContentStream());
    }

    [Fact]
    public void StreamingAttachmentEnforcesActualAggregateBytesWhenDeclaredSizeUnderreports() {
        byte[] content = Enumerable.Range(0, 60_000).Select(index => (byte)(index % 251)).ToArray();
        var options = new EmailStoreReaderOptions(
            maxAttachmentBytes: 100_000,
            maxTotalAttachmentBytes: 50_000,
            retainAttachmentContent: false);
        using var stream = new MemoryStream(PstTestFileBuilder.Create(
            attachmentContent: content, attachmentDeclaredLength: 1));
        using EmailStoreSession session = EmailStoreSession.Open(stream, "mailbox.pst", options);
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
        EmailStoreItem item = session.ReadItem(reference, new EmailStoreItemReadOptions(
            EmailStoreItemReadParts.Metadata |
            EmailStoreItemReadParts.AttachmentMetadata |
            EmailStoreItemReadParts.AttachmentContent,
            preferStreamingAttachmentContent: true));
        EmailAttachment attachment = Assert.Single(item.Document.Attachments);

        using Stream payload = attachment.OpenContentStream();
        EmailStoreLimitExceededException exception = Assert.Throws<EmailStoreLimitExceededException>(
            () => payload.CopyTo(Stream.Null));

        Assert.Equal(nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes), exception.LimitName);
        Assert.True(exception.Actual > exception.Maximum);
    }

    [Fact]
    public void ReopeningStreamingAttachmentDoesNotChargeItsActualBytesTwice() {
        byte[] content = Enumerable.Range(0, 60_000).Select(index => (byte)(index % 251)).ToArray();
        var options = new EmailStoreReaderOptions(
            maxAttachmentBytes: 100_000,
            maxTotalAttachmentBytes: 70_000,
            retainAttachmentContent: false);
        using var stream = new MemoryStream(PstTestFileBuilder.Create(
            attachmentContent: content, attachmentDeclaredLength: 1));
        using EmailStoreSession session = EmailStoreSession.Open(stream, "mailbox.pst", options);
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
        EmailStoreItem item = session.ReadItem(reference, new EmailStoreItemReadOptions(
            EmailStoreItemReadParts.Metadata |
            EmailStoreItemReadParts.AttachmentMetadata |
            EmailStoreItemReadParts.AttachmentContent,
            preferStreamingAttachmentContent: true));
        EmailAttachment attachment = Assert.Single(item.Document.Attachments);

        for (int index = 0; index < 2; index++) {
            using Stream payload = attachment.OpenContentStream();
            using var copy = new MemoryStream();
            payload.CopyTo(copy);
            Assert.Equal(content, copy.ToArray());
        }
    }

    [Fact]
    public void RetainedAttachmentEnforcesActualAggregateBytesWhenDeclaredSizeUnderreports() {
        byte[] content = Enumerable.Range(0, 60_000).Select(index => (byte)(index % 251)).ToArray();
        var options = new EmailStoreReaderOptions(
            maxAttachmentBytes: 100_000,
            maxTotalAttachmentBytes: 50_000);
        using var stream = new MemoryStream(PstTestFileBuilder.Create(
            attachmentContent: content, attachmentDeclaredLength: 1));
        using EmailStoreSession session = EmailStoreSession.Open(stream, "mailbox.pst", options);
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());

        EmailStoreLimitExceededException exception = Assert.Throws<EmailStoreLimitExceededException>(
            () => session.ReadItem(reference));

        Assert.Equal(nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes), exception.LimitName);
        Assert.Equal(content.LongLength, exception.Actual);
    }

    [Fact]
    public void Materialized_reader_does_not_return_expired_attachment_sources() {
        var options = new EmailStoreReaderOptions(retainAttachmentContent: false);
        using var stream = new MemoryStream(PstTestFileBuilder.Create(
            attachmentContent: new byte[] { 1, 2, 3, 4 }));

        EmailStoreReadResult result = new EmailStoreReader(options).Read(stream, "mailbox.pst");
        EmailStoreItem item = Assert.Single(result.Store.Folders.SelectMany(folder => folder.Items));
        EmailAttachment attachment = Assert.Single(item.Document.Attachments);

        Assert.Null(attachment.Content);
        Assert.Null(attachment.ContentSource);
        Assert.False(item.LoadedParts.HasFlag(EmailStoreItemReadParts.AttachmentContent));
    }

    [Fact]
    public void EnforcesInputAndSeekabilityContracts() {
        byte[] bytes = PstTestFileBuilder.Create();
        var options = new EmailStoreReaderOptions(maxInputBytes: bytes.Length - 1L);
        using var stream = new MemoryStream(bytes);

        EmailStoreLimitExceededException exception = Assert.Throws<EmailStoreLimitExceededException>(
            () => new EmailStoreReader(options).Read(stream, "mailbox.pst"));

        Assert.Equal(nameof(EmailStoreReaderOptions.MaxInputBytes), exception.LimitName);
        using var nonSeekable = new NonSeekableStream(bytes);
        Assert.Throws<ArgumentException>(() => new EmailStoreReader().Read(nonSeekable, "mailbox.pst"));
    }

    private sealed class NonSeekableStream : MemoryStream {
        internal NonSeekableStream(byte[] bytes) : base(bytes) { }
        public override bool CanSeek => false;
    }
}
