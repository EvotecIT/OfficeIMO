namespace OfficeIMO.Archives.Tests {
    public sealed class ArchiveCabinetCopyTests {
        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(4099)]
        public void VerifiedCopiesKeepCrossFrameRangesEmptyMembersAndFolderCharges(int compression) {
            byte[] expected = Enumerable.Range(0, 65536).Select(i => (byte)(i % 251)).ToArray();
            byte[][] raw = { expected.Take(32768).ToArray(), expected.Skip(32768).ToArray() };
            byte[][] frames = raw.Select((b, i) => compression == 0 ? b : compression == 1
                ? ArchiveFixtures.MsZipFrame(b) : ArchiveFixtures.LzxRawFrame(b, i == 0)).ToArray();
            byte[] input = ArchiveFixtures.FramedCab((ushort)compression, frames, new[] { 32768, 32768 },
                ("cross", 32760, 33), ("empty", 65536, 0));
            using var inspection = ArchiveInspector.Open(new MemoryStream(input));
            var plan = inspection.PlanRead(new[] { 0 });
            using var destination = new MemoryStream(); destination.WriteByte(42);
            inspection.CopyTo(0, destination);
            Assert.Equal(new byte[] { 42 }.Concat(expected.Skip(32760).Take(33)), destination.ToArray());
            Assert.Equal(plan.DecodedBytes, inspection.ReadBytes);
            Assert.Equal(plan.DecompressionWorkBytes, inspection.DecompressionWorkBytes);
            Assert.Equal(33, inspection.ResourceUsage.SelectedBytes);
            inspection.CopyTo(1, destination);
            Assert.Equal(34, destination.Length);
            Assert.Equal(2 * plan.DecodedBytes, inspection.ReadBytes);
            Assert.True(destination.CanWrite);
        }

        [Theory]
        [InlineData(ArchiveFormat.Zip)]
        [InlineData(ArchiveFormat.Cab)]
        [InlineData(ArchiveFormat.Msi)]
        public void SpoolLimitIsDetachedAndRejectsBeforeReadChargesOrDestinationWrites(ArchiveFormat format) {
            byte[] input = format == ArchiveFormat.Zip ? ArchiveZipInspectionTests.Zip(("member", "complete"))
                : format == ArchiveFormat.Cab ? ArchiveFixtures.StoredCab("member", Encoding.UTF8.GetBytes("complete"))
                : ArchiveFixtures.Msi();
            var options = new ArchiveInspectionOptions { MaxTemporaryBytes = 0 };
            using var inspection = ArchiveInspector.Open(new MemoryStream(input), options);
            options.MaxTemporaryBytes = long.MaxValue;
            using var destination = new MemoryStream(); destination.WriteByte(42);
            var error = Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.CopyTo(0, destination));
            Assert.Equal("MaxTemporaryBytes", error.LimitName); Assert.Equal(0, error.Maximum);
            Assert.Equal(new byte[] { 42 }, destination.ToArray());
            Assert.Equal(0, inspection.ReadBytes); Assert.Equal(0, inspection.DecompressionWorkBytes);
            Assert.Equal(0, inspection.ResourceUsage.SelectedBytes);
        }

        [Fact]
        public void MsiCopyMatchesItsVerifiedMemberAndChargesOneFolder() {
            byte[] input = ArchiveFixtures.Msi();
            byte[] expected;
            using (var baseline = ArchiveInspector.Open(new MemoryStream(input))) {
                using var member = baseline.OpenRead(0); using var bytes = new MemoryStream();
                member.CopyTo(bytes); expected = bytes.ToArray();
            }
            using var inspection = ArchiveInspector.Open(new MemoryStream(input));
            var plan = inspection.PlanRead(new[] { 0 });
            using var destination = new MemoryStream();
            inspection.CopyTo(0, destination);
            Assert.Equal(expected, destination.ToArray());
            Assert.Equal(plan.DecodedBytes, inspection.ReadBytes);
            Assert.Equal(plan.DecompressionWorkBytes, inspection.DecompressionWorkBytes);
            Assert.Equal(expected.Length, inspection.ResourceUsage.SelectedBytes);
        }

        [Fact]
        public void EmptyMemberCanBeVerifiedWithZeroSpoolAllowance() {
            using var inspection = ArchiveInspector.Open(new MemoryStream(ArchiveFixtures.EmptyStoredCab("empty")),
                new() { MaxTemporaryBytes = 0 });
            using var destination = new MemoryStream();
            inspection.CopyTo(0, destination);
            Assert.Equal(0, destination.Length);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void DestinationFailureOrCancellationKeepsDestinationOwnedAndAllowsAnotherRead(bool cancel) {
            byte[] expected = Enumerable.Repeat((byte)7, 100000).ToArray();
            byte[][] frames = Enumerable.Range(0, 4).Select(i => expected.Skip(i * 32768).Take(32768).ToArray()).ToArray();
            using var inspection = ArchiveInspector.Open(new MemoryStream(ArchiveFixtures.FramedCab(0, frames,
                frames.Select(f => f.Length).ToArray(), ("member", 0, expected.Length))));
            using var cancellation = new CancellationTokenSource();
            using var destination = new InterruptedDestination(cancel ? cancellation : null);
            if (cancel) Assert.Throws<OperationCanceledException>(() => inspection.CopyTo(0, destination, cancellation.Token));
            else Assert.Throws<IOException>(() => inspection.CopyTo(0, destination));
            Assert.True(destination.CanWrite);
            Assert.Equal(expected.Take(65536), destination.ToArray());
            using var recovered = new MemoryStream(); inspection.CopyTo(0, recovered);
            Assert.Equal(expected, recovered.ToArray());
        }

        [Fact]
        public void NegativeSpoolAllowanceFailsAtOpen() {
            Assert.Throws<ArgumentOutOfRangeException>(() => ArchiveInspector.Open(new MemoryStream(
                ArchiveFixtures.EmptyStoredCab("empty")), new() { MaxTemporaryBytes = -1 }));
        }

        private sealed class InterruptedDestination(CancellationTokenSource? cancellation) : MemoryStream {
            public override void Write(byte[] buffer, int offset, int count) {
                base.Write(buffer, offset, count);
                if (cancellation != null) cancellation.Cancel();
                else throw new IOException("Synthetic destination failure");
            }
        }
    }
}
