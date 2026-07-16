namespace OfficeIMO.OneNote.Tests;

public sealed class FileHeaderTests {
    [Theory]
    [InlineData("testOneNote.one")]
    [InlineData("testOneNote2016.one")]
    [InlineData("testOneNoteEmbeddedWordDoc.one")]
    public void ReadsDesktopRevisionStoreHeaders(string fileName) {
        string path = FixturePath(fileName);

        OneNoteFileHeader header = OneNoteFileProbe.ReadHeader(path);

        Assert.Equal(OneNoteFileKind.Section, header.FileKind);
        Assert.Equal(OneNoteStorageFormat.RevisionStore, header.StorageFormat);
        Assert.NotEqual(Guid.Empty, header.FileId);
        Assert.NotEqual(Guid.Empty, header.FileVersionId);
        Assert.True(header.TransactionCount > 0);
        Assert.Equal((ulong)new FileInfo(path).Length, header.ExpectedFileLength);
        Assert.NotNull(header.RootFileNodeList);
        Assert.NotNull(header.TransactionLog);
    }

    [Theory]
    [InlineData("testOneNoteFromOffice365.one")]
    [InlineData("testOneNoteFromOffice365-2.one")]
    public void ReadsFileSynchronizationPackageHeaders(string fileName) {
        string path = FixturePath(fileName);

        OneNoteFileHeader header = OneNoteFileProbe.ReadHeader(path);

        Assert.Equal(OneNoteFileKind.Section, header.FileKind);
        Assert.Equal(OneNoteStorageFormat.FileSynchronizationPackage, header.StorageFormat);
        Assert.NotEqual(Guid.Empty, header.FileId);
        Assert.NotNull(header.StorageIndexId);
        Assert.NotEqual(Guid.Empty, header.StorageIndexId!.Identifier);
        Assert.NotNull(header.CellSchemaId);
    }

    [Fact]
    public async Task AsyncProbeRestoresSeekableStreamPosition() {
        byte[] bytes = File.ReadAllBytes(FixturePath("testOneNote.one"));
        using (var stream = new MemoryStream(bytes)) {
            stream.Position = 37;

            OneNoteFileHeader header = await OneNoteFileProbe.ReadHeaderAsync(stream);

            Assert.Equal(OneNoteStorageFormat.RevisionStore, header.StorageFormat);
            Assert.Equal(37, stream.Position);
        }
    }

    [Fact]
    public void RejectsInputOverConfiguredLimitBeforeParsing() {
        string path = FixturePath("testOneNote.one");
        var options = new OneNoteReaderOptions { MaxInputBytes = 1024 };

        IOException exception = Assert.Throws<IOException>(() => OneNoteFileProbe.ReadHeader(path, options));

        Assert.Contains("MaxInputBytes", exception.Message);
    }

    [Fact]
    public void RejectsUnknownFormatWithStableErrorCode() {
        using (var stream = new MemoryStream(new byte[1024])) {
            OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() => OneNoteFileProbe.ReadHeader(stream));

            Assert.Equal("ONENOTE_UNKNOWN_FILE_FORMAT", exception.Code);
            Assert.Equal(48, exception.Offset);
        }
    }

    [Fact]
    public void RejectsTruncatedDesktopHeader() {
        byte[] bytes = File.ReadAllBytes(FixturePath("testOneNote.one"));
        Array.Resize(ref bytes, 512);
        using (var stream = new MemoryStream(bytes)) {
            OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() => OneNoteFileProbe.ReadHeader(stream));

            Assert.Equal("ONENOTE_TRUNCATED_STRUCTURE", exception.Code);
        }
    }

    [Theory]
    [InlineData("testOneNote.one")]
    [InlineData("testOneNote2016.one")]
    [InlineData("testOneNoteEmbeddedWordDoc.one")]
    public void ReadsValidatedRootFileNodeList(string fileName) {
        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(FixturePath(fileName));

        Assert.NotEmpty(store.RootFileNodeList.Fragments);
        Assert.NotEmpty(store.RootFileNodeList.Nodes);
        Assert.Equal(0U, store.RootFileNodeList.Fragments[0].Sequence);
        Assert.Contains(store.RootFileNodeList.Nodes, node => node.Id == OneNoteFileNodeId.ObjectSpaceManifestRoot);
        Assert.Contains(store.RootFileNodeList.Nodes, node => node.Id == OneNoteFileNodeId.ObjectSpaceManifestListReference);
        Assert.All(store.RootFileNodeList.Nodes, node => Assert.True(node.Size >= 4));
        Assert.True(store.FileNodeLists.Count > 1);
        Assert.Contains(store.FileNodeLists.SelectMany(list => list.Nodes), node => node.Id == OneNoteFileNodeId.RevisionManifestListStart);
        Assert.Contains(store.FileNodeLists.SelectMany(list => list.Nodes), node => node.Id == OneNoteFileNodeId.ObjectGroupStart);
        Assert.NotEmpty(store.Revisions);
        Assert.NotEmpty(store.Objects);
        Assert.Contains(store.Objects, item => item.PropertySet != null);
    }

    [Fact]
    public void RevisionStoreReaderRestoresStreamPosition() {
        byte[] bytes = File.ReadAllBytes(FixturePath("testOneNote.one"));
        using (var stream = new MemoryStream(bytes)) {
            stream.Position = 113;

            OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(stream);

            Assert.NotEmpty(store.RootFileNodeList.Nodes);
            Assert.Equal(113, stream.Position);
        }
    }

    [Fact]
    public void TransactionLogCountMatchesMaterializedRootList() {
        string path = FixturePath("testOneNote.one");
        using (var stream = File.OpenRead(path)) {
            var options = new OneNoteReaderOptions();
            OneNoteFileHeader header = OneNoteFileProbe.ReadHeader(stream, options);
            IReadOnlyDictionary<uint, int> counts = OneNoteTransactionLogReader.Read(stream, header, options);
            OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(stream, options);

            Assert.Equal(store.RootFileNodeList.Nodes.Count, counts[store.RootFileNodeList.Id]);
            Assert.All(store.FileNodeLists, list => Assert.Equal(list.Nodes.Count, counts[list.Id]));
        }
    }

    [Fact]
    public void OversizedTransactionLogFragmentFailsWithBoundedFormatError() {
        var header = new OneNoteFileHeader {
            TransactionLog = new OneNoteFileChunkReference(0, uint.MaxValue),
            TransactionCount = 1,
            ExpectedFileLength = uint.MaxValue
        };

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteTransactionLogReader.Read(
                new MemoryStream(),
                header,
                new OneNoteReaderOptions { MaxInputBytes = null }));

        Assert.Equal("ONENOTE_TRANSACTION_FRAGMENT_SIZE", exception.Code);
        Assert.Equal(0, exception.Offset);
    }

    [Fact]
    public void RootValidationAcceptsStructuralFragmentTerminator() {
        var nodes = new[] {
            Node(OneNoteFileNodeId.ObjectSpaceManifestListReference),
            Node(OneNoteFileNodeId.ObjectSpaceManifestRoot),
            Node(OneNoteFileNodeId.ChunkTerminator)
        };
        var root = new OneNoteFileNodeList(0x10, Array.Empty<OneNoteFileNodeListFragment>(), nodes);

        OneNoteRevisionStoreReader.ValidateRootList(root, OneNoteFileKind.Section);
    }

    [Fact]
    public void RevisionStoreReaderReconstructsPackageEncoding() {
        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(FixturePath("testOneNoteFromOffice365.one"));

        Assert.Equal(OneNoteStorageFormat.FileSynchronizationPackage, store.Header.StorageFormat);
        Assert.NotEmpty(store.Revisions);
        Assert.NotEmpty(store.Objects);
        Assert.Contains(store.Objects, item => item.PropertySet != null);
    }

    private static string FixturePath(string fileName) => Path.Combine(AppContext.BaseDirectory, "Fixtures", fileName);

    private static OneNoteFileNode Node(OneNoteFileNodeId id) =>
        new OneNoteFileNode((ushort)id, 4, 0, 0, OneNoteFileNodeBaseType.Inline, 0, null, Array.Empty<byte>());
}
