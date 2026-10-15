namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveInspectionLimitsTests {
    [Fact]
    public void RestoreFailurePropagatesWithoutTakingOwnershipOfTheCallerStream() {
        using var source = new ThrowOnRestoreStream(ArchiveFixtures.Msi());
        IOException error = Assert.Throws<IOException>(() => ArchiveInspector.Open(source));
        Assert.Equal("Position restoration failed.", error.Message);
        Assert.True(source.CanRead); Assert.Equal(1, source.RestoreAttempts);
        source.Position = 0; Assert.Equal(0xd0, source.ReadByte());
    }

    [Fact]
    public void InputLimitPrecedesAllocationAndReadAndPreservesCallerPosition() {
        using var source = new RejectReadStream(new byte[128]); source.Position = 9;
        Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveInspector.Open(source, new ArchiveInspectionOptions { MaxInputBytes = 64 }));
        Assert.Equal(9, source.Position); Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData("entries")]
    [InlineData("member")]
    [InlineData("expanded")]
    [InlineData("directory")]
    public void DeclaredCabContentAndDirectoryWorkAreBoundedBeforePayloadDecoding(string bound) {
        byte[] bytes = ArchiveFixtures.StoredCab(("first", new byte[] { 1, 2 }), ("second", new byte[] { 3, 4 }));
        var options = new ArchiveInspectionOptions();
        switch (bound) {
            case "entries": options.MaxEntryCount = 1; break;
            case "member": options.MaxEntryBytes = 1; break;
            case "expanded": options.MaxExpandedBytes = 3; break;
            case "directory": options.MaxDirectoryRecords = 3; break;
        }
        using var source = new MemoryStream(bytes);
        Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveInspector.Open(source, options));
    }

    [Theory]
    [InlineData("cabinets")]
    [InlineData("metadata")]
    [InlineData("rows")]
    [InlineData("strings")]
    [InlineData("files")]
    [InlineData("expansion")]
    [InlineData("depth")]
    public void MsiTablesCabinetsAndDeclaredContentObserveIndependentBounds(string bound) {
        using var source = new MemoryStream(ArchiveFixtures.Msi(twoCabinets: true));
        var options = new ArchiveInspectionOptions();
        switch (bound) {
            case "cabinets": options.MaxCabinetCount = 1; break;
            case "metadata": options.MaxMetadataBytes = 4; break;
            case "rows": options.MaxTableRows = 1; break;
            case "strings": options.MaxStringPoolEntries = 1; break;
            case "files": options.MaxEntryCount = 1; break;
            case "expansion": options.MaxExpandedBytes = 9; break;
            case "depth": options.MaxPathDepth = 2; break;
        }
        Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveInspector.Open(source, options));
    }

    [Fact]
    public void RepeatedReadsChargeTheEntireSharedFolderAndFailWithoutReturningAPrefix() {
        byte[] bytes = ArchiveFixtures.StoredCab(("small", new byte[] { 1 }), ("other", new byte[] { 2, 3, 4 }));
        using var source = new MemoryStream(bytes);
        using ArchiveInspection inspection = ArchiveInspector.Open(source, new ArchiveInspectionOptions { MaxExpandedBytes = 4 });
        using Stream first = inspection.OpenRead(0);
        Assert.Equal(new byte[] { 1 }, ArchiveInspectionTests.Bytes(first)); Assert.Equal(4, inspection.ReadBytes);
        Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.OpenRead(1));
        Assert.Equal(4, inspection.ReadBytes);
    }

    [Fact]
    public void WorkBudgetIsSharedAcrossMsiMemberReadsAndChargesFailedAttempts() {
        using var source = new MemoryStream(ArchiveFixtures.Msi(twoCabinets: true));
        using ArchiveInspection inspection = ArchiveInspector.Open(source,
            new ArchiveInspectionOptions { MaxDecompressionWorkBytes = 32768 });
        using Stream first = inspection.OpenRead(0);
        Assert.Equal(5, first.Length);
        Assert.Equal("MaxDecompressionWorkBytes", Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.OpenRead(1)).LimitName);
        Assert.Equal(10, inspection.ReadBytes);
    }

    [Fact]
    public void RawStreamReadsApplyCompleteMemberAndCumulativeBounds() {
        using var source = new MemoryStream(ArchiveFixtures.Msi());
        using ArchiveInspection inspection = ArchiveInspector.Open(source,
            new ArchiveInspectionOptions { MaxEntryBytes = 5, MaxExpandedBytes = 5 });
        ArchiveStorageStream cabinet = Assert.Single(inspection.StorageStreams, stream => stream.RawName == MsiPackageReader.EncodeStreamName("payload.cab", false));
        Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.OpenStorageStream(cabinet.Ordinal)); Assert.Equal(0, inspection.ReadBytes);
        using Stream member = inspection.OpenRead(0); Assert.Equal(5, member.Length);
        using var secondSource = new MemoryStream(ArchiveFixtures.Msi());
        using ArchiveInspection second = ArchiveInspector.Open(secondSource, new ArchiveInspectionOptions { MaxExpandedBytes = 5 });
        using Stream complete = second.OpenRead(0);
        ArchiveStorageStream raw = second.StorageStreams.First(stream => stream.Length > 0);
        Assert.Throws<OfficeInspectionResourceLimitException>(() => second.OpenStorageStream(raw.Ordinal)); Assert.Equal(5, second.ReadBytes);
    }

    [Theory]
    [InlineData("truncated")]
    [InlineData("range")]
    [InlineData("msi-size")]
    [InlineData("msi-order")]
    public void MalformedDirectoriesAndMsiMappingsCannotProduceAnInventory(string defect) {
        byte[] bytes;
        if (defect.StartsWith("msi", StringComparison.Ordinal)) {
            bytes = ArchiveFixtures.Msi(wrongSize: defect == "msi-size", twoCabinets: defect == "msi-order",
                editDatabase: (_, rows) => { if (defect == "msi-order") rows["File"][1][7] = 1; });
        } else {
            bytes = ArchiveFixtures.StoredCab("file", new byte[] { 1 });
            if (defect == "truncated") Array.Resize(ref bytes, bytes.Length - 1);
            else bytes[48] = 2; // A file offset beyond the declared one-byte folder.
        }
        using var source = new MemoryStream(bytes); source.Position = 2;
        Assert.Throws<InvalidDataException>(() => ArchiveInspector.Open(source)); Assert.Equal(2, source.Position);
    }

    [Fact]
    public void CancellationDuringSnapshotRestoresPositionAndReadCancellationDoesNotPoisonTheSession() {
        byte[] bytes = ArchiveFixtures.StoredCab("file", new byte[] { 1 });
        using var canceled = new CancellationTokenSource();
        using var canceling = new CancelAfterReadStream(bytes, canceled); canceling.Position = 9;
        Assert.Throws<OperationCanceledException>(() => ArchiveInspector.Open(canceling, cancellationToken: canceled.Token));
        Assert.Equal(9, canceling.Position); Assert.True(canceling.CanRead);
        using var source = new MemoryStream(bytes);
        using ArchiveInspection inspection = ArchiveInspector.Open(source);
        Assert.Throws<OperationCanceledException>(() => inspection.OpenRead(0, canceled.Token)); Assert.Equal(0, inspection.ReadBytes);
        using Stream member = inspection.OpenRead(0); Assert.Equal(new byte[] { 1 }, ArchiveInspectionTests.Bytes(member));
    }

    private sealed class RejectReadStream : MemoryStream {
        internal RejectReadStream(byte[] bytes) : base(bytes, writable: false) { }
        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Input length must be rejected before reading.");
    }
    private sealed class ThrowOnRestoreStream : MemoryStream {
        internal ThrowOnRestoreStream(byte[] bytes) : base(bytes, writable: false) { base.Position = 9; }
        internal int RestoreAttempts { get; private set; }
        public override long Position {
            get => base.Position;
            set {
                if (value == 9) { RestoreAttempts++; throw new IOException("Position restoration failed."); }
                base.Position = value;
            }
        }
    }
    private sealed class CancelAfterReadStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        internal CancelAfterReadStream(byte[] bytes, CancellationTokenSource cancellation) : base(bytes, writable: false) { _cancellation = cancellation; }
        public override int Read(byte[] buffer, int offset, int count) {
            int result = base.Read(buffer, offset, Math.Min(1, count)); _cancellation.Cancel(); return result;
        }
    }
}
