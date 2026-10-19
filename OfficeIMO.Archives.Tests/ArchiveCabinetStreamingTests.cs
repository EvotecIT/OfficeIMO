namespace OfficeIMO.Archives.Tests {
    public sealed class ArchiveCabinetStreamingTests {
        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(4099)]
        public void SelectedRangesCrossFramesOverlapAndKeepRequestOrder(int compression) {
            byte[] expected = Enumerable.Range(0, 2 * 32768 + 137).Select(i => (byte)(i % 251)).ToArray();
            byte[][] raw = { expected.Take(32768).ToArray(), expected.Skip(32768).Take(32768).ToArray(), expected.Skip(65536).ToArray() };
            byte[][] frames = raw.Select((bytes, i) => compression == 0 ? bytes : compression == 1
                ? ArchiveFixtures.MsZipFrame(bytes) : ArchiveFixtures.LzxRawFrame(bytes, i == 0)).ToArray();
            byte[] input = ArchiveFixtures.FramedCab((ushort)compression, frames, raw.Select(f => f.Length).ToArray(),
                ("tail", 65530, 20), ("cross", 32760, 33), ("overlap", 32760, 33), ("empty", (uint)expected.Length, 0), ("excluded", 0, 32760));
            using var source = new MemoryStream(input);
            using var inspection = ArchiveInspector.Open(source);
            int[] ordinals = { 2, 0, 3, 1 };
            ArchiveReadPlan plan = inspection.PlanRead(ordinals);
            Assert.Equal(86, plan.SelectedBytes); Assert.Equal(expected.Length, plan.DecodedBytes);
            Assert.Equal(frames.Sum(f => f.Length), plan.CompressedPayloadBytes);
            Assert.Equal(1, plan.CabinetFolderCount); Assert.Equal(3, plan.CabinetFrameCount);
            Assert.Equal(0, inspection.ReadBytes); Assert.Equal(0, inspection.ResourceUsage.SelectedBytes);
            using var batch = inspection.ReadMembers(ordinals, plan.SelectedBytes);
            Assert.Equal(plan.DecodedBytes, batch.DecodedFolderBytes);
            Assert.Equal(plan.DecompressionWorkBytes, batch.DecompressionWorkBytes);
            Assert.Equal(ordinals, batch.Members.Select(m => m.Entry.Ordinal));
            Assert.Equal(expected.Skip(32760).Take(33), Bytes(batch.Members[0].Content));
            Assert.Equal(expected.Skip(65530).Take(20), Bytes(batch.Members[1].Content));
            Assert.Empty(Bytes(batch.Members[2].Content));
            Assert.Equal(expected.Skip(32760).Take(33), Bytes(batch.Members[3].Content));
            using var fixture = new ArchiveFixtures();
            var extracted = ArchiveExtractor.Extract(fixture.Package(input), fixture.Destination,
                new() { IncludeFile = name => name == "cross" || name == "tail" });
            Assert.Equal(2, extracted.Files.Count);
            Assert.Equal(expected.Skip(32760).Take(33), File.ReadAllBytes(Path.Combine(fixture.Destination, "cross")));
            Assert.Equal(expected.Skip(65530).Take(20), File.ReadAllBytes(Path.Combine(fixture.Destination, "tail")));
        }

        [Fact]
        public void CircularLzxHistoryRetainsUntranslatedBytesAndUsesAbsoluteE8Positions() {
            byte[] lead = Enumerable.Repeat((byte)'A', 32768).ToArray();
            lead[64] = 0xe8; BitConverter.GetBytes(500).CopyTo(lead, 65);
            byte[] first = ArchiveFixtures.LzxRawFrame(lead, true, translationSize: 1000000, distance: 32768);
            byte[] second = ArchiveFixtures.LzxE8HistoryFrame();
            byte[] input = ArchiveFixtures.FramedCab((15 << 8) | 3, new[] { first, second }, new[] { 32768, 85 },
                ("lead", 0, 32768), ("tail", 32768, 85));
            using var inspection = ArchiveInspector.Open(new MemoryStream(input));
            using var batch = inspection.ReadMembers(new[] { 1, 0 }, 32853);
            byte[] tail = Bytes(batch.Members[0].Content);
            Assert.All(tail.Take(64), value => Assert.Equal((byte)'A', value));
            Assert.Equal(0xe8, tail[64]); Assert.Equal(500 - 32832, BitConverter.ToInt32(tail, 65));
            byte[] decodedLead = Bytes(batch.Members[1].Content);
            Assert.Equal(500 - 64, BitConverter.ToInt32(decodedLead, 65));
        }

        [Fact]
        public void CorruptUnselectedTailExposesNoSelectedMemberOrPublishedDirectory() {
            byte[] first = ArchiveFixtures.MsZipFrame(new byte[] { 1, 2, 3, 4 });
            byte[] unfinished = { 0x43, 0x4b, 0, 1, 0, 0xfe, 0xff, 5 }; // no final Deflate block
            byte[] input = ArchiveFixtures.FramedCab(1, new[] { first, unfinished }, new[] { 4, 1 }, ("first", 0, 4));
            using var inspection = ArchiveInspector.Open(new MemoryStream(input));
            Assert.Equal(5, inspection.PlanRead(new[] { 0 }).DecodedBytes);
            Assert.Throws<InvalidDataException>(() => inspection.OpenRead(0));
            Assert.Throws<InvalidDataException>(() => inspection.ReadMembers(new[] { 0 }, 4));
            using var destination = new MemoryStream(); destination.WriteByte(42);
            Assert.Throws<InvalidDataException>(() => inspection.CopyTo(0, destination));
            Assert.Equal(new byte[] { 42 }, destination.ToArray());
            using var fixture = new ArchiveFixtures();
            Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(input), fixture.Destination));
            Assert.False(Directory.Exists(fixture.Destination));
        }

        [Fact]
        public void LzxFrameSessionHonorsCancellationBetweenFrames() {
            using var cancellation = new CancellationTokenSource();
            var decoder = new OfficeLzxDecoder(65536, 15, cancellation.Token);
            byte[] frame = ArchiveFixtures.LzxRawFrame(new byte[32768], true);
            decoder.DecodeFrame(frame, 0, frame.Length, 32768);
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => decoder.DecodeFrame(frame, 0, frame.Length, 32768));
            Assert.Throws<OperationCanceledException>(() => decoder.Complete());
        }

        [Fact]
        public void ChmResetGroupsCanEndInsideABlockWhileCabinetFoldersRequireItsEnd() {
            byte[] frame = ArchiveFixtures.LzxRawFrame(Encoding.ASCII.GetBytes("ABCD"), true);
            // Advertise an eight-byte block, with this four-byte CHM reset group ending partway through it.
            frame[2] = 0x80; frame[3] = 0;
            byte[] decoded = OfficeLzxDecoder.Decompress(new[] { frame }, new[] { 4 }, 15, 4, requireCompleteBlocks: false);
            Assert.Equal("ABCD", Encoding.ASCII.GetString(decoded));
            Assert.Throws<OfficeLzxException>(() => OfficeLzxDecoder.Decompress(new[] { frame }, new[] { 4 }, 15, 4));
        }

        [Theory]
        [InlineData(15)]
        [InlineData(16)]
        [InlineData(17)]
        [InlineData(18)]
        [InlineData(19)]
        [InlineData(20)]
        [InlineData(21)]
        public void NativeMakecabE8FixturesRetainTheirOracleBytesForEveryWindow(int windowBits) {
            string file = windowBits == 21 ? "makecab-lzx-e8.cab" : "makecab-lzx" + windowBits + "-e8.cab";
            byte[] pattern = Encoding.ASCII.GetBytes("OfficeIMO-LZX-E8-independent-oracle-");
            byte[] expected = Enumerable.Range(0, 4096).Select(i => pattern[i % pattern.Length]).ToArray();
            foreach (var call in new[] { (64, 500), (1024, -50), (2048, 1000000) }) {
                expected[call.Item1] = 0xe8; BitConverter.GetBytes(call.Item2).CopyTo(expected, call.Item1 + 1);
            }
            using var inspection = ArchiveInspector.Open(Path.Combine(AppContext.BaseDirectory, "Fixtures", file));
            using var member = inspection.OpenRead(0);
            Assert.Equal(expected, Bytes(member));
        }

        private static byte[] Bytes(Stream content) {
            using var output = new MemoryStream(); content.CopyTo(output); return output.ToArray();
        }
    }
}
