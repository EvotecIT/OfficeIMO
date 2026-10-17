namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveOperationBudgetTests {
    [Fact]
    public void GroupedZipAdmissionRefusesTheWholeSelectionBeforeAnyPayloadDecode() {
        var budget = new ArchiveOperationBudget(new() { MaxSelectedBytes = 4 });
        using var inspection = ArchiveInspector.Open(new MemoryStream(ArchiveZipInspectionTests.Zip(("a", "abc"), ("b", "def"))),
            new() { OperationBudget = budget });
        Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.ReadMembers(new[] { 0, 1 }, 6));
        Assert.Equal(0, inspection.ReadBytes);
        Assert.Equal(0, budget.Usage.SelectedBytes);
        Assert.Equal(0, budget.Usage.ExpandedBytes);
    }
    [Fact]
    public void InputAndEntryAdmissionAreCumulativeAcrossMixedContainersAndDepth() {
        byte[] zip = ArchiveZipInspectionTests.Zip(("a.txt", "abc"));
        byte[] cab = ArchiveFixtures.StoredCab("b.txt", Encoding.UTF8.GetBytes("def"));
        var budget = new ArchiveOperationBudget(new() { MaxInputBytes = zip.Length + cab.Length, MaxEntryCount = 2, MaxDepth = 1 });
        using var first = ArchiveInspector.Open(new MemoryStream(zip), new() { OperationBudget = budget });
        using var second = ArchiveInspector.Open(new MemoryStream(cab), new() { OperationBudget = budget.CreateNested() });
        Assert.Equal(zip.Length + cab.Length, budget.Usage.InputBytes);
        Assert.Equal(2, budget.Usage.EntryCount);
        Assert.Equal(1, budget.Usage.MaximumDepth);
        var error = Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveInspector.Open(new MemoryStream(zip),
            new() { OperationBudget = budget.CreateNested().CreateNested() }));
        Assert.Equal("Operation.MaxDepth", error.LimitName);
        Assert.Equal(1, error.Maximum);
        Assert.Equal(2, budget.Usage.EntryCount);
    }

    [Fact]
    public void SolidCabinetGapsAndUnselectedMembersConsumeTheSharedExpansionBudget() {
        byte[] cab = ArchiveFixtures.StoredCab(("one", new byte[3]), ("two", new byte[5]));
        var budget = new ArchiveOperationBudget(new() { MaxExpandedBytes = 8 });
        using var inspection = ArchiveInspector.Open(new MemoryStream(cab), new() { OperationBudget = budget });
        using var selected = inspection.ReadMembers(new[] { 0 }, 3);
        Assert.Equal(8, budget.Usage.ExpandedBytes);
        Assert.Equal(3, budget.Usage.SelectedBytes);
        Assert.Equal(32768, budget.Usage.DecompressionWorkBytes);
        Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.OpenRead(1));
        Assert.Equal(8, budget.Usage.ExpandedBytes);
    }

    [Fact]
    public void FailedZipReadsKeepTheirChargesAndVerificationDoesNotSwallowQuotas() {
        byte[] zip = ArchiveZipInspectionTests.Zip(("bad", "abc"));
        zip[ArchiveZipInspectionTests.Payload(zip)] ^= 1;
        var budget = new ArchiveOperationBudget(new() { MaxExpandedBytes = 3 });
        using var inspection = ArchiveInspector.Open(new MemoryStream(zip), new() { OperationBudget = budget });
        Assert.Equal(ArchiveVerificationStatus.Invalid, Assert.Single(inspection.Verify().Members).Status);
        Assert.Equal(3, budget.Usage.ExpandedBytes);
        Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.Verify());
        Assert.Equal(3, budget.Usage.ExpandedBytes);
    }

    [Fact]
    public void MemoryReadAndExtractionUseTheSameBudgetAndRollbackOnFailure() {
        byte[] zip = ArchiveZipInspectionTests.Zip(("a.txt", "abc"));
        var budget = new ArchiveOperationBudget(new() { MaxSelectedBytes = 3 });
        Assert.Single(ArchiveReader.ReadZip(zip, new() { OperationBudget = budget }));
        using var fixture = new ArchiveFixtures();
        var error = Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveExtractor.Extract(
            fixture.Package(zip), fixture.Destination, new() { OperationBudget = budget }));
        Assert.Equal("Operation.MaxSelectedBytes", error.LimitName);
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
        Assert.Equal(2, budget.Usage.EntryCount);
    }
}
