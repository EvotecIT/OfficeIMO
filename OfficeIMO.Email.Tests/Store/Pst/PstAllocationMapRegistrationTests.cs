namespace OfficeIMO.Email.Store.Tests;

public sealed class PstAllocationMapRegistrationTests {
    private const long FirstAmapOffset = 0x4400;
    private const long FirstPmapOffset = 0x4600;
    private const long AmapInterval = 0x3E000;

    [Fact]
    public void RepeatedOrLowerHorizonsDoNotRewriteRegisteredPages() {
        WithWriterFile(file => {
            PstAllocationMapTestSupport.CountingFileStream counted = PstAllocationMapTestSupport.CountMapWrites(file);

            PstAllocationMapTestSupport.RegisterThrough(file, file.Length);
            PstAllocationMapTestSupport.RegisterThrough(file, FirstAmapOffset);
            PstAllocationMapTestSupport.RegisterThrough(file, file.Length);

            Assert.Equal(0, counted.WriteCalls);
        });
    }

    [Fact]
    public void GrowthRegistersEachNewAmapAndPmapPageOnce() {
        WithWriterFile(file => {
            PstAllocationMapTestSupport.CountingFileStream counted = PstAllocationMapTestSupport.CountMapWrites(file);
            long end = FirstAmapOffset + AmapInterval * 9 + 1;

            PstAllocationMapTestSupport.RegisterThrough(file, end);
            Assert.Equal(10, counted.WriteCalls); // Nine new AMaps and one new PMap.
            PstWriterAllocationMap map = PstAllocationMapTestSupport.Field<PstWriterAllocationMap>(file, "_allocationMap");
            for (int index = 0; index <= 9; index++) {
                Assert.Equal(0xFF, map.Read(index)[0]);
            }
            Assert.Equal(0xFF, map.Read(0)[1]);
            Assert.Equal(0xFF, map.Read(8)[1]);

            PstAllocationMapTestSupport.RegisterThrough(file, FirstAmapOffset);
            PstAllocationMapTestSupport.RegisterThrough(file, end);
            Assert.Equal(10, counted.WriteCalls);
        });
    }

    [Theory]
    [InlineData(FirstAmapOffset + AmapInterval)]
    [InlineData(FirstPmapOffset + AmapInterval * 8)]
    public void MapPageIsRegisteredOnlyWhenTheHorizonPassesItsOffset(long pageOffset) {
        WithWriterFile(file => {
            PstAllocationMapTestSupport.RegisterThrough(file, pageOffset);
            PstAllocationMapTestSupport.CountingFileStream counted = PstAllocationMapTestSupport.CountMapWrites(file);

            PstAllocationMapTestSupport.RegisterThrough(file, pageOffset);
            Assert.Equal(0, counted.WriteCalls);
            PstAllocationMapTestSupport.RegisterThrough(file, pageOffset + 1);
            Assert.Equal(1, counted.WriteCalls);
            PstAllocationMapTestSupport.RegisterThrough(file, pageOffset + 1);
            Assert.Equal(1, counted.WriteCalls);
        });
    }

    [Fact]
    public void ResumeReconstructsTheHorizonAndPreservesItemsAttachmentsAndApplicationState() {
        string directory = TemporaryDirectory();
        string path = Path.Combine(directory, "resumed.pst");
        string checkpoint = Path.Combine(directory, "writer.checkpoint");
        byte[] expectedState = System.Text.Encoding.UTF8.GetBytes("synthetic-ordinal=40");
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions(failOnDataLoss: true, checkpointPath: checkpoint,
                    checkpointIntervalItems: int.MaxValue))) {
                string folder = writer.AddFolder("Original");
                for (int index = 0; index < 40; index++) writer.AddItem(folder, CreateMessage(index));
                writer.Checkpoint(expectedState);
            }

            using (EmailStorePstWriter resumed = EmailStorePstWriter.Resume(checkpoint, out byte[]? state)) {
                Assert.Equal(expectedState, state);
                PstWriterFile file = PstAllocationMapTestSupport.WriterFile(resumed);
                PstAllocationMapTestSupport.CountingFileStream counted = PstAllocationMapTestSupport.CountMapWrites(file);
                PstAllocationMapTestSupport.RegisterThrough(file,
                    PstAllocationMapTestSupport.Field<long>(file, "_nextOffset"));
                Assert.Equal(0, counted.WriteCalls);

                resumed.AddItem(resumed.AddFolder("After resume"), CreateMessage(40));
                EmailStorePstWriteReport report = resumed.Complete();
                Assert.Equal(41, report.ItemCount);
                Assert.False(report.HasErrors);
                Assert.False(report.HasDataLoss);
            }

            using EmailStoreSession session = EmailStoreSession.Open(path,
                new EmailStoreReaderOptions(retainAttachmentContent: true));
            EmailDocument[] documents = session.EnumerateItems().Select(item => session.ReadItem(item).Document).ToArray();
            Assert.Equal(41, documents.Length);
            for (int index = 0; index < 41; index++) {
                EmailDocument document = Assert.Single(documents, item => item.Subject == "Synthetic " + index);
                Assert.Equal("Body " + index, document.Body.Text);
                EmailAttachment attachment = Assert.Single(document.Attachments);
                Assert.Equal("payload.bin", attachment.FileName);
                Assert.Equal(Enumerable.Repeat((byte)index, 64 * 1024).ToArray(), attachment.Content);
            }
            Assert.Contains(session.FolderCatalog.All, item => item.Name == "Original");
            Assert.Contains(session.FolderCatalog.All, item => item.Name == "After resume");
            Assert.False(File.Exists(checkpoint));
        } finally { DeleteDirectory(directory); }
    }

    private static EmailDocument CreateMessage(int index) {
        var document = new EmailDocument { Subject = "Synthetic " + index, MessageClass = "IPM.Note" };
        document.Body.Text = "Body " + index;
        document.Attachments.Add(new EmailAttachment {
            FileName = "payload.bin", Content = Enumerable.Repeat((byte)index, 64 * 1024).ToArray(), Length = 64 * 1024
        });
        return document;
    }

    private static void WithWriterFile(Action<PstWriterFile> action) {
        string directory = TemporaryDirectory();
        try {
            using var file = new PstWriterFile(Path.Combine(directory, "working.pst"));
            action(file);
        } finally { DeleteDirectory(directory); }
    }

    private static string TemporaryDirectory() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-map-registration-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }

    private static void DeleteDirectory(string directory) {
        try { if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
