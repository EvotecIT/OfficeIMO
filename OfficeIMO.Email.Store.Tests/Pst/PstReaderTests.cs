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
        EmailStoreFolder inbox = Assert.Single(result.Store.Folders, folder => folder.Name == "Inbox");
        Assert.Equal(root.Id, inbox.ParentId);
        EmailStoreMessage message = Assert.Single(inbox.Messages);
        Assert.Equal("Synthetic PST message", message.Document.Subject);
        Assert.Equal("Body from the PST property context", message.Document.Body.Text);
        Assert.Equal("IPM.Note", message.Document.MessageClass);
        Assert.Equal("Pst", message.Document.Properties["EmailStore:Format"]);
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
