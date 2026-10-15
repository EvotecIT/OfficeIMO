namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveGroupedReadTests {
    [Fact]
    public void SelectedSolidFolderMembersFitOneDecodeBudgetAndKeepRawDuplicateIdentity() {
        byte[] cabinet = ArchiveFixtures.StoredCab(("../duplicate", Encoding.UTF8.GetBytes("first")),
            ("../duplicate", Encoding.UTF8.GetBytes("second")), ("unselected", new byte[] { 42, 43, 44 }));
        using var source = new MemoryStream(cabinet);
        var options = new ArchiveInspectionOptions { MaxExpandedBytes = 14, MaxDecompressionWorkBytes = 32768 };
        using ArchiveInspection inspection = ArchiveInspector.Open(source, options);
        using ArchiveReadBatch batch = inspection.ReadMembers(new[] { 1, 0 }, maxSelectedBytes: 11);
        Assert.Equal(new[] { 1, 0 }, batch.Members.Select(member => member.Entry.Ordinal));
        Assert.All(batch.Members, member => {
            Assert.Equal("../duplicate", member.Entry.RawName); Assert.False(member.Content.CanWrite);
            Assert.Equal(member.Entry.Length, member.Content.Length); Assert.Equal(0, member.Content.Position);
        });
        Assert.Equal(Encoding.UTF8.GetBytes("second"), ArchiveInspectionTests.Bytes(batch.Members[0].Content));
        Assert.Equal(Encoding.UTF8.GetBytes("first"), ArchiveInspectionTests.Bytes(batch.Members[1].Content));
        Assert.Equal(11, batch.SelectedBytes); Assert.Equal(14, batch.DecodedFolderBytes);
        Assert.Equal(32768, batch.DecompressionWorkBytes); Assert.Equal(14, inspection.ReadBytes);
        Assert.Equal(32768, inspection.DecompressionWorkBytes);
        // The original single-member API still consumes the folder for every request.
        using ArchiveInspection separate = ArchiveInspector.Open(source, options);
        using Stream first = separate.OpenRead(1);
        Assert.Throws<OfficeInspectionResourceLimitException>(() => separate.OpenRead(0));
    }

    [Fact]
    public void RequestedOrderSurvivesGroupingAndCabinetIdentitySeparatesMsiFolders() {
        using var source = new MemoryStream(ArchiveFixtures.Msi(twoCabinets: true, editDatabase: (_, rows) =>
            rows["File"][1][2] = "Template.adml"));
        using ArchiveInspection inspection = ArchiveInspector.Open(source,
            new ArchiveInspectionOptions { MaxExpandedBytes = 10, MaxDecompressionWorkBytes = 65536 });
        using ArchiveReadBatch batch = inspection.ReadMembers(new[] { 1, 0 }, 10);
        Assert.Equal(new[] { "PayloadTwo", "PayloadKey" }, batch.Members.Select(member => member.Entry.RawName));
        Assert.Equal(batch.Members[0].Entry.AuthoredPath, batch.Members[1].Entry.AuthoredPath);
        Assert.Equal(Encoding.UTF8.GetBytes("world"), ArchiveInspectionTests.Bytes(batch.Members[0].Content));
        Assert.Equal(Encoding.UTF8.GetBytes("hello"), ArchiveInspectionTests.Bytes(batch.Members[1].Content));
        Assert.Equal(10, batch.SelectedBytes); Assert.Equal(10, batch.DecodedFolderBytes);
        Assert.Equal(65536, batch.DecompressionWorkBytes);
    }

    [Fact]
    public void AnUnselectedUnsupportedFolderDoesNotConsumeDecodeWorkOrPreventSelectedReads() {
        using var source = new MemoryStream(ArchiveFixtures.Msi(twoCabinets: true,
            emptySecondPayload: true, emptySecondCompression: 2));
        using ArchiveInspection inspection = ArchiveInspector.Open(source,
            new ArchiveInspectionOptions { MaxExpandedBytes = 5, MaxDecompressionWorkBytes = 32768 });
        Assert.False(inspection.Entries[1].CanRead);
        using ArchiveReadBatch batch = inspection.ReadMembers(new[] { 0 }, 5);
        Assert.Equal(Encoding.UTF8.GetBytes("hello"), ArchiveInspectionTests.Bytes(Assert.Single(batch.Members).Content));
        Assert.Equal(5, batch.DecodedFolderBytes); Assert.Equal(32768, batch.DecompressionWorkBytes);
    }

    [Theory]
    [InlineData("makecab-lzx-testOneNote2016.cab", "testOneNote2016.one")]
    [InlineData("makecab-lzx-testOneNoteFromOffice365-2.cab", "testOneNoteFromOffice365-2.one")]
    public void GroupedLzxReadMatchesTheIndependentMakecabOriginal(string cabinet, string original) {
        string fixtures = Path.Combine(AppContext.BaseDirectory, "Fixtures");
        byte[] expected = File.ReadAllBytes(Path.Combine(fixtures, original));
        using ArchiveInspection inspection = ArchiveInspector.Open(Path.Combine(fixtures, cabinet));
        using ArchiveReadBatch batch = inspection.ReadMembers(new[] { 0 }, expected.LongLength);
        ArchiveReadMember member = Assert.Single(batch.Members);
        Assert.Equal(ArchiveCompression.Lzx, member.Entry.Compression);
        Assert.True(member.Entry.FolderChecksumFrameCount > 0);
        Assert.Equal(expected, ArchiveInspectionTests.Bytes(member.Content));
        Assert.Equal(expected.LongLength, batch.SelectedBytes);
        Assert.Equal(member.Entry.FolderExpandedBytes, batch.DecodedFolderBytes);
    }

    [Theory]
    [InlineData("selected")]
    [InlineData("duplicate")]
    [InlineData("ordinal")]
    [InlineData("canceled")]
    [InlineData("unsupported")]
    public void InvalidOrUnsupportedSelectionsAreRejectedBeforeAnyFolderWork(string refusal) {
        byte[] bytes = refusal == "unsupported"
            ? ArchiveFixtures.Msi(twoCabinets: true, emptySecondPayload: true, emptySecondCompression: 2)
            : ArchiveFixtures.StoredCab(("first", new byte[] { 1, 2 }), ("second", new byte[] { 3, 4 }));
        using var source = new MemoryStream(bytes);
        using ArchiveInspection inspection = ArchiveInspector.Open(source);
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        switch (refusal) {
            case "selected": Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.ReadMembers(new[] { 1, 0 }, 3)); break;
            case "duplicate": Assert.Throws<ArgumentException>(() => inspection.ReadMembers(new[] { 0, 0 }, 4)); break;
            case "ordinal": Assert.Throws<ArgumentOutOfRangeException>(() => inspection.ReadMembers(new[] { 0, -1 }, 4)); break;
            case "canceled": Assert.Throws<OperationCanceledException>(() => inspection.ReadMembers(new[] { 0 }, 4, canceled.Token)); break;
            case "unsupported": Assert.Throws<NotSupportedException>(() => inspection.ReadMembers(new[] { 0, 1 }, 5)); break;
        }
        Assert.Equal(0, inspection.ReadBytes); Assert.Equal(0, inspection.DecompressionWorkBytes);
    }

    [Fact]
    public void WholeBatchExpansionRefusalDoesNotConsumeMoreOfAnExistingSessionBudget() {
        using var source = new MemoryStream(ArchiveFixtures.StoredCab(("first", new byte[] { 1 }), ("second", new byte[] { 2 })));
        using ArchiveInspection inspection = ArchiveInspector.Open(source, new ArchiveInspectionOptions { MaxExpandedBytes = 2 });
        using Stream first = inspection.OpenRead(0);
        Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.ReadMembers(new[] { 0, 1 }, 2));
        Assert.Equal(2, inspection.ReadBytes); Assert.Equal(32768, inspection.DecompressionWorkBytes);
    }

    [Fact]
    public void FailedSecondCabinetKeepsAttemptedChargesAndReturnsNoPartialBatch() {
        using var source = new MemoryStream(ArchiveFixtures.Msi(twoCabinets: true));
        using ArchiveInspection inspection = ArchiveInspector.Open(source,
            new ArchiveInspectionOptions { MaxExpandedBytes = 10, MaxDecompressionWorkBytes = 32768 });
        Assert.Equal("MaxDecompressionWorkBytes", Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.ReadMembers(new[] { 0, 1 }, 10)).LimitName);
        Assert.Equal(10, inspection.ReadBytes); Assert.Equal(32768, inspection.DecompressionWorkBytes);
    }

    [Fact]
    public void CorruptFolderChecksumReturnsNoBatchAndAccountsForAttemptedWork() {
        byte[] cabinet = ArchiveFixtures.StoredCab(("first", new byte[] { 1 }), ("second", new byte[] { 2 }));
        cabinet[BitConverter.ToInt32(cabinet, 36)] = 1; // Wrong nonzero checksum; inventory alone does not verify it.
        using var source = new MemoryStream(cabinet);
        using ArchiveInspection inspection = ArchiveInspector.Open(source);
        Assert.StartsWith("CAB_CHECKSUM:", Assert.Throws<InvalidDataException>(() => inspection.ReadMembers(new[] { 0, 1 }, 2)).Message);
        Assert.Equal(2, inspection.ReadBytes); Assert.Equal(32768, inspection.DecompressionWorkBytes);
    }

    [Fact]
    public void BatchStreamsUseSnapshotOptionsAndSurviveInspectionButCloseWithTheirBatch() {
        var options = new ArchiveInspectionOptions { MaxExpandedBytes = 11, MaxDecompressionWorkBytes = 32768 };
        byte[] cabinet = ArchiveFixtures.StoredCab(("first", Encoding.UTF8.GetBytes("first")), ("second", Encoding.UTF8.GetBytes("second")));
        using var source = new MemoryStream(cabinet, writable: true);
        using var openCancellation = new CancellationTokenSource();
        ArchiveInspection inspection = ArchiveInspector.Open(source, options, openCancellation.Token);
        Array.Clear(cabinet, 0, cabinet.Length); source.Dispose(); openCancellation.Cancel();
        options.MaxEntryBytes = 1; options.MaxExpandedBytes = 1; options.MaxDecompressionWorkBytes = 1;
        var ordinals = new[] { 1, 0 };
        ArchiveReadBatch batch = inspection.ReadMembers(ordinals, 11);
        ordinals[0] = 0;
        ArchiveReadMember first = batch.Members[0]; ArchiveReadMember second = batch.Members[1];
        inspection.Dispose();
        Assert.Equal(1, first.Entry.Ordinal);
        Assert.Equal(Encoding.UTF8.GetBytes("second"), ArchiveInspectionTests.Bytes(first.Content));
        first.Content.Dispose();
        Assert.Equal(Encoding.UTF8.GetBytes("first"), ArchiveInspectionTests.Bytes(second.Content));
        batch.Dispose(); batch.Dispose();
        Assert.False(second.Content.CanRead); Assert.Empty(batch.Members);
        Assert.Equal("first", second.Entry.RawName);
        Assert.Throws<ObjectDisposedException>(() => inspection.ReadMembers(new[] { 0 }, 11));
    }

    [Fact]
    public void EmptySelectionsAndZeroLengthMembersRequireNoPayloadOrFrameBudget() {
        using var source = new MemoryStream(ArchiveFixtures.EmptyStoredCab("empty"));
        using ArchiveInspection inspection = ArchiveInspector.Open(source,
            new ArchiveInspectionOptions { MaxExpandedBytes = 1, MaxDecompressionWorkBytes = 1 });
        using ArchiveReadBatch empty = inspection.ReadMembers(Array.Empty<int>(), 0);
        Assert.Empty(empty.Members); Assert.Equal(0, empty.SelectedBytes); Assert.Equal(0, empty.DecodedFolderBytes);
        using ArchiveReadBatch selected = inspection.ReadMembers(new[] { 0 }, 0);
        Assert.Equal(0, Assert.Single(selected.Members).Content.Length);
        Assert.Equal(0, selected.DecodedFolderBytes); Assert.Equal(0, selected.DecompressionWorkBytes);
        Assert.Equal(0, inspection.ReadBytes);
    }
}
