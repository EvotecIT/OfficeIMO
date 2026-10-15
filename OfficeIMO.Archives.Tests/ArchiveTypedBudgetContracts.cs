using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveTypedBudgetContracts {
    [Theory]
    [InlineData("MaxEntryCount", 1, "CAB_ENTRY_COUNT_LIMIT")]
    [InlineData("MaxEntryBytes", 1, "CAB_ENTRY_LIMIT")]
    [InlineData("MaxExpandedBytes", 3, "CAB_EXPANDED_LIMIT")]
    [InlineData("MaxDirectoryRecords", 3, "CAB_DIRECTORY_LIMIT")]
    public void PassiveCabBoundsExposeTheirEffectivePublicLimit(string name, long maximum, string? code) {
        var options = new ArchiveInspectionOptions();
        typeof(ArchiveInspectionOptions).GetProperty(name)!.SetValue(options,
            name is "MaxEntryCount" or "MaxDirectoryRecords" ? (object)(int)maximum : maximum);
        using var source = new MemoryStream(ArchiveFixtures.StoredCab(("first", new byte[] { 1, 2 }), ("second", new byte[] { 3, 4 })));
        AssertBudget(() => ArchiveInspector.Open(source, options), name, maximum, code);
        Assert.Equal(0, source.Position); Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData("MaxCabinetCount", 1)]
    [InlineData("MaxMetadataBytes", 4)]
    [InlineData("MaxTableRows", 22)]
    [InlineData("MaxStringPoolEntries", 1)]
    [InlineData("MaxEntryCount", 1)]
    [InlineData("MaxExpandedBytes", 9)]
    [InlineData("MaxPathDepth", 2)]
    [InlineData("MaxPathLength", 10)]
    public void PassiveMsiBoundsAreDistinctFromMalformedInput(string name, long maximum) {
        var options = new ArchiveInspectionOptions();
        typeof(ArchiveInspectionOptions).GetProperty(name)!.SetValue(options,
            name is "MaxMetadataBytes" or "MaxExpandedBytes" ? (object)maximum : (int)maximum);
        using var source = new MemoryStream(ArchiveFixtures.Msi(twoCabinets: true));
        AssertBudget(() => ArchiveInspector.Open(source, options), name, maximum);
    }

    [Fact]
    public void MsiDirectoryQuotaRetainsTheCanonicalSlotLimitAndMalformedChainsStayMalformed() {
        using var source = new MemoryStream(ArchiveFixtures.Msi());
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            ArchiveInspector.Open(source, new ArchiveInspectionOptions { MaxDirectoryRecords = 3 }));
        Assert.Equal("archives", error.Owner); Assert.Equal("MaxDirectoryRecords", error.LimitName); Assert.Equal(3, error.Maximum);
        var native = Assert.IsType<OfficeIMO.Core.Internal.OfficeCompoundStreamLimitExceededException>(error.InnerException);
        Assert.Equal("MaxDirectoryEntries", native.LimitName); Assert.Equal(3, native.MaximumValue);
        using var cycle = new MemoryStream(ArchiveFixtures.Msi(directoryCycle: true));
        Assert.Throws<InvalidDataException>(() => ArchiveInspector.Open(cycle));
    }

    [Fact]
    public void NativeRemainingExpansionReportsZeroWithoutMisclassifyingMsiMappingCount() {
        using var fixture = new ArchiveFixtures();
        AssertBudget(() => ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi(twoCabinets: true)), fixture.Destination,
            new ArchiveExtractionOptions { MaxExpandedBytes = 5, IncludeFile = _ => false }), "MaxExpandedBytes", 0, "CAB_EXPANDED_LIMIT");
        byte[] unmapped = ArchiveFixtures.StoredCab(("a", new byte[] { 1 }), ("b", new byte[] { 2 }));
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(ArchiveFixtures.Msi(unmappedCabinet: unmapped)), fixture.Destination));
    }

    [Theory]
    [InlineData(64L, 64L)]
    [InlineData(long.MaxValue, (long)int.MaxValue)]
    public void InputCeilingPrecedesSnapshotAndReportsTheEffectiveManagedMaximum(long requested, long effective) {
        using var source = new DeclaredLengthStream(requested == 64 ? 65 : (long)int.MaxValue + 1);
        AssertBudget(() => ArchiveInspector.Open(source, new ArchiveInspectionOptions { MaxInputBytes = requested }),
            "MaxInputBytes", effective);
        Assert.Equal(0, source.Position); Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RepeatedAndGroupedMsiReadsPreserveNativeWorkIdentityAndAttemptedCharges(bool grouped) {
        using var source = new MemoryStream(ArchiveFixtures.Msi(twoCabinets: true));
        using ArchiveInspection inspection = ArchiveInspector.Open(source,
            new ArchiveInspectionOptions { MaxExpandedBytes = 10, MaxDecompressionWorkBytes = 32768 });
        if (grouped) AssertBudget(() => inspection.ReadMembers(new[] { 0, 1 }, 10), "MaxDecompressionWorkBytes", 32768, "CAB_WORK_LIMIT");
        else {
            using Stream first = inspection.OpenRead(0);
            AssertBudget(() => inspection.OpenRead(1), "MaxDecompressionWorkBytes", 32768, "CAB_WORK_LIMIT");
        }
        Assert.Equal(10, inspection.ReadBytes); Assert.Equal(32768, inspection.DecompressionWorkBytes);
    }

    [Fact]
    public void SelectedBatchAndRepeatedExpansionLimitsRemainSeparateAndDoNotChargeRejectedWork() {
        using var source = new MemoryStream(ArchiveFixtures.StoredCab(("first", new byte[] { 1, 2 }), ("second", new byte[] { 3, 4 })));
        using ArchiveInspection inspection = ArchiveInspector.Open(source, new ArchiveInspectionOptions { MaxExpandedBytes = 4 });
        AssertBudget(() => inspection.ReadMembers(new[] { 0, 1 }, 3), "maxSelectedBytes", 3);
        Assert.Equal(0, inspection.ReadBytes);
        using Stream first = inspection.OpenRead(0);
        AssertBudget(() => inspection.ReadMembers(new[] { 0, 1 }, 4), "MaxExpandedBytes", 0);
        AssertBudget(() => inspection.OpenRead(1), "MaxExpandedBytes", 4);
        Assert.Equal(4, inspection.ReadBytes); Assert.Equal(32768, inspection.DecompressionWorkBytes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WorkFailureRollsBackExtractionAndPreservesAnEmptyDestination(bool msi) {
        using var fixture = new ArchiveFixtures();
        Directory.CreateDirectory(fixture.Destination);
        byte[] package = msi ? ArchiveFixtures.Msi(twoCabinets: true) : ArchiveFixtures.StoredCab("file", new byte[] { 1 });
        long work = msi ? 32768 : 32767;
        AssertBudget(() => ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination,
            new ArchiveExtractionOptions { MaxDecompressionWorkBytes = work }), "MaxDecompressionWorkBytes", work, "CAB_WORK_LIMIT");
        Assert.True(Directory.Exists(fixture.Destination)); Assert.Empty(Directory.EnumerateFileSystemEntries(fixture.Destination));
        Assert.DoesNotContain(Directory.EnumerateDirectories(fixture.Root), path => Path.GetFileName(path).StartsWith(".officeimo-archive-", StringComparison.Ordinal));
    }

    [Fact]
    public void NegativeMemberAndCorruptHeaderRemainMalformedAndUnsupportedCompressionRemainsUnsupported() {
        using var msi = new MemoryStream(ArchiveFixtures.Msi(editDatabase: (_, rows) => rows["File"][0][3] = -1));
        Assert.Throws<InvalidDataException>(() => ArchiveInspector.Open(msi));
        byte[] bytes = ArchiveFixtures.StoredCab("file", new byte[] { 1 }); bytes[4] = 1;
        using var corrupt = new MemoryStream(bytes);
        Assert.Contains("CAB_HEADER", Assert.Throws<InvalidDataException>(() => ArchiveInspector.Open(corrupt)).Message);
        bytes[4] = 0; bytes[42] = 2;
        using var unsupported = new MemoryStream(bytes);
        using ArchiveInspection inspection = ArchiveInspector.Open(unsupported);
        Assert.Throws<NotSupportedException>(() => inspection.OpenRead(0));
    }

    private static void AssertBudget(Action action, string name, long maximum, string? code = null) {
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(action);
        Assert.Equal("archives", error.Owner); Assert.Equal(name, error.LimitName); Assert.Equal(maximum, error.Maximum);
        if (code != null) Assert.Equal(code, Assert.IsType<OfficeCabinetException>(error.InnerException).Code);
    }

    private sealed class DeclaredLengthStream : MemoryStream {
        private readonly long _length;
        internal DeclaredLengthStream(long length) => _length = length;
        public override long Length => _length;
        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Limit must precede read.");
    }
}
