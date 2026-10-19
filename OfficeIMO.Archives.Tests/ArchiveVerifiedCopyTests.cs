namespace OfficeIMO.Archives.Tests {

    public sealed class ArchiveVerifiedCopyTests {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void CopyPublishesOnlyCompleteVerifiedMemberContentAndLeavesDestinationOpen(bool corrupt) {
            byte[] zip = ArchiveZipInspectionTests.Zip(("member.txt", new string('x', 180000)));
            if (corrupt) zip[ArchiveZipInspectionTests.Payload(zip)] ^= 1;
            using var inspection = ArchiveInspector.Open(new MemoryStream(zip));
            using var destination = new MemoryStream();
            destination.WriteByte(42);
            if (corrupt) {
                Assert.Throws<InvalidDataException>(() => inspection.CopyTo(0, destination));
                Assert.Equal(new byte[] { 42 }, destination.ToArray());
            } else {
                inspection.CopyTo(0, destination);
                Assert.Equal(180001, destination.Length);
                Assert.Equal(42, destination.ToArray()[0]);
                Assert.All(destination.ToArray().Skip(1), b => Assert.Equal((byte)'x', b));
            }
            Assert.True(destination.CanWrite);
        }

        [Fact]
        public void PreCanceledCopiesDoNotChargeBudgetsOrWriteBytes() {
            using var inspection = ArchiveInspector.Open(new MemoryStream(ArchiveZipInspectionTests.Zip(("member", "content"))));
            using var destination = new MemoryStream();
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => inspection.CopyTo(0, destination, cancellation.Token));
            Assert.Equal(0, inspection.ReadBytes);
            Assert.Equal(0, destination.Length);
        }
    }
}
