namespace OfficeIMO.Archives.Tests {
    public sealed class ArchiveReadPlanTests {
        [Fact]
        public void ZipPlanningIsDetachedAndDoesNotReserveAnExhaustedReadBudget() {
            var budget = new ArchiveOperationBudget(new() { MaxSelectedBytes = 0, MaxExpandedBytes = 0, MaxDecompressionWorkBytes = 0 });
            byte[] input = ArchiveZipInspectionTests.Zip(("same", "first"), ("same", "second"));
            using var inspection = ArchiveInspector.Open(new MemoryStream(input), new() { OperationBudget = budget });
            ArchiveResourceUsage before = budget.Usage;
            ArchiveReadPlan plan = inspection.PlanRead(new[] { 1, 0 });
            Assert.Equal(new[] { 1, 0 }, plan.Entries.Select(e => e.Ordinal));
            Assert.Equal(11, plan.SelectedBytes); Assert.Equal(11, plan.DecodedBytes);
            Assert.Equal(inspection.Entries.Sum(e => e.CompressedLength!.Value), plan.CompressedPayloadBytes);
            Assert.Equal(plan.CompressedPayloadBytes + 22, plan.DecompressionWorkBytes);
            Assert.Equal(0, plan.CabinetFolderCount); Assert.Equal(0, plan.CabinetFrameCount);
            Assert.Equal(before.InputBytes, budget.Usage.InputBytes); Assert.Equal(before.EntryCount, budget.Usage.EntryCount);
            Assert.Equal(0, budget.Usage.SelectedBytes); Assert.Equal(0, budget.Usage.ExpandedBytes); Assert.Equal(0, budget.Usage.DecompressionWorkBytes);
            Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.ReadMembers(new[] { 1, 0 }, 11));
            inspection.Dispose();
            Assert.Equal("same", plan.Entries[0].RawName); Assert.Equal(11, plan.DecodedBytes);
            Assert.Throws<ObjectDisposedException>(() => inspection.PlanRead(new[] { 0 }));
        }

        [Fact]
        public void MsiPlanCountsEachSelectedCabinetAndFolderOnce() {
            using var inspection = ArchiveInspector.Open(new MemoryStream(ArchiveFixtures.Msi(twoCabinets: true)));
            ArchiveReadPlan plan = inspection.PlanRead(new[] { 1, 0 });
            Assert.Equal(2, plan.CabinetFolderCount); Assert.Equal(2, plan.CabinetFrameCount);
            Assert.Equal(10, plan.SelectedBytes); Assert.Equal(10, plan.DecodedBytes); Assert.Equal(10, plan.CompressedPayloadBytes);
            using var batch = inspection.ReadMembers(new[] { 1, 0 }, plan.SelectedBytes);
            Assert.Equal(plan.DecompressionWorkBytes, batch.DecompressionWorkBytes);
            Assert.Equal(plan.DecodedBytes, batch.DecodedFolderBytes);
        }

        [Fact]
        public void PlanningValidatesSelectionWithoutDecodingOrChargingIt() {
            using var inspection = ArchiveInspector.Open(new MemoryStream(ArchiveFixtures.StoredCab(("a", new byte[] { 1 }), ("b", new byte[] { 2 }))));
            Assert.Throws<ArgumentNullException>(() => inspection.PlanRead(null!));
            Assert.Throws<ArgumentException>(() => inspection.PlanRead(new[] { 0, 0 }));
            Assert.Throws<ArgumentOutOfRangeException>(() => inspection.PlanRead(new[] { -1 }));
            Assert.Throws<OperationCanceledException>(() => inspection.PlanRead(new[] { 0 }, new CancellationToken(true)));
            Assert.Empty(inspection.PlanRead(Array.Empty<int>()).Entries);
            Assert.Equal(0, inspection.ReadBytes); Assert.Equal(0, inspection.ResourceUsage.SelectedBytes);
            using var external = ArchiveInspector.Open(new MemoryStream(ArchiveFixtures.Msi(externalCab: true)));
            Assert.Throws<NotSupportedException>(() => external.PlanRead(new[] { 0 }));
        }
    }
}
