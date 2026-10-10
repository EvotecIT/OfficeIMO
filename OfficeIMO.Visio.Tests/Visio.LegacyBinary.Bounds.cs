using System.Threading;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioLegacyBinaryBoundsTests {
        [Theory]
        [InlineData((byte)0)]
        [InlineData((byte)255)]
        public void CompressionAcceptsATrailingFlagWithoutConsumingTheDecodedBudget(byte trailingFlag) {
            byte[] compound = CompressedStream(new byte[] { 255, 65, 66, 67, 68, 69, 70, 71, 72, trailingFlag });
            VisioBinaryContainer.Node root = new VisioBinaryContainer(compound,
                new() { MaxDecompressedBytes = 8 }, CancellationToken.None).ReadRoot();
            Assert.Equal("ABCDEFGH", Encoding.ASCII.GetString(root.Data));
            Assert.Throws<InvalidDataException>(() => new VisioBinaryContainer(compound,
                new() { MaxDecompressedBytes = 7 }, CancellationToken.None).ReadRoot());
        }

        [Fact]
        public void CompressionRejectsAnIncompleteBackReference() {
            byte[] compound = CompressedStream(new byte[] { 255, 65, 66, 67, 68, 69, 70, 71, 72, 0, 238 });
            InvalidDataException error = Assert.Throws<InvalidDataException>(() =>
                new VisioBinaryContainer(compound, new(), CancellationToken.None).ReadRoot());
            Assert.Contains("match is truncated", error.Message);
        }

        [Fact]
        public void SharedPointerFanOutChargesTheLogicalRecordBudget() {
            byte[] compound = PointerGraph((0x14, Enumerable.Repeat(1, 20).ToArray()),
                (0x1d, Enumerable.Repeat(2, 20).ToArray()), (0x99, null));
            using var input = new MemoryStream(compound);
            var error = Assert.Throws<InvalidDataException>(() => VisioDocument.LoadLegacyBinary(input,
                options: new() { Limits = new() { MaxRecords = 64 } }));
            Assert.Contains("budget", error.Message);
            Assert.True(input.CanRead);
            Assert.Equal(0, input.Position);
        }

        [Fact]
        public void SharedSubtreesRespectDepthAtEveryReference() {
            // The first route reaches the leaf at depth3; the later alias reaches it at depth4.
            byte[] compound = PointerGraph((0x14, new[] { 1, 4 }), (0x1d, new[] { 2 }),
                (0x1d, new[] { 3 }), (0x99, null), (0x1d, new[] { 1 }));
            var root = new VisioBinaryContainer(compound, new() { MaxDepth = 4 }, CancellationToken.None).ReadRoot();
            Assert.Equal(2, root.Children.Count);
            using var input = new MemoryStream(compound);
            var error = Assert.Throws<InvalidDataException>(() => VisioDocument.LoadLegacyBinary(input,
                options: new() { MaxDepth = 3 }));
            Assert.Contains("depth", error.Message);
        }

        // Minimal standard CFB fixture with one regular-sector stream. Pointer-table graphs
        // are synthetic resource-boundary inputs, separate from the independent producer corpus.
        private static byte[] PointerGraph(params (uint Type, int[]? Children)[] nodes) {
            var native = new byte[4096];
            Encoding.ASCII.GetBytes("Visio (TM) Drawing\r\n\0").CopyTo(native, 0);
            native[26] = 11;
            var offsets = new int[nodes.Length];
            var lengths = new int[nodes.Length];
            int position = 64;
            for (int i = 0; i < nodes.Length; i++) {
                offsets[i] = position;
                lengths[i] = nodes[i].Children == null ? 4 : 16 + nodes[i].Children!.Length * 18;
                position += lengths[i];
            }
            Assert.True(position <= native.Length);
            void Pointer(int offset, int target) {
                var node = nodes[target];
                U32(native, offset, node.Type);
                U32(native, offset + 8, (uint)offsets[target]);
                U32(native, offset + 12, (uint)lengths[target]);
                U16(native, offset + 16, node.Children == null ? (ushort)0 : (ushort)0x50);
            }
            Pointer(36, 0);
            for (int i = 0; i < nodes.Length; i++) {
                int[]? children = nodes[i].Children;
                if (children == null) continue;
                U32(native, offsets[i], 8); // Table starts four bytes into the node.
                U32(native, offsets[i] + 8, (uint)children.Length);
                for (int j = 0; j < children.Length; j++) Pointer(offsets[i] + 16 + j * 18, children[j]);
            }
            return CompoundStream(native);
        }

        private static byte[] CompressedStream(byte[] encoded) {
            byte[] native = new byte[4096];
            Encoding.ASCII.GetBytes("Visio (TM) Drawing\r\n\0").CopyTo(native, 0);
            native[26] = 11;
            U32(native, 36, 0x99);
            U32(native, 44, 64);
            U32(native, 48, (uint)encoded.Length);
            U16(native, 52, 2);
            encoded.CopyTo(native, 64);
            return CompoundStream(native);
        }

        private static byte[] CompoundStream(byte[] native) {
            var cfb = new byte[11 * 512]; // Header, eight data sectors, directory, FAT.
            new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }.CopyTo(cfb, 0);
            U16(cfb, 24, 0x3e); U16(cfb, 26, 3); U16(cfb, 28, 0xfffe);
            U16(cfb, 30, 9); U16(cfb, 32, 6);
            U32(cfb, 44, 1); U32(cfb, 48, 8); U32(cfb, 56, 4096);
            U32(cfb, 60, 0xfffffffe); U32(cfb, 68, 0xfffffffe);
            for (int i = 0; i < 109; i++) U32(cfb, 76 + i * 4, i == 0 ? 9U : uint.MaxValue);
            native.CopyTo(cfb, 512);
            void Entry(int offset, string name, byte type, uint start, uint size, uint child = uint.MaxValue) {
                byte[] title = Encoding.Unicode.GetBytes(name + "\0"); title.CopyTo(cfb, offset);
                U16(cfb, offset + 64, (ushort)title.Length); cfb[offset + 66] = type; cfb[offset + 67] = 1;
                U32(cfb, offset + 68, uint.MaxValue); U32(cfb, offset + 72, uint.MaxValue);
                U32(cfb, offset + 76, child); U32(cfb, offset + 116, start); U32(cfb, offset + 120, size);
            }
            Entry(9 * 512, "Root Entry", 5, 0xfffffffe, 0, 1);
            Entry(9 * 512 + 128, "VisioDocument", 2, 0, 4096);
            for (int i = 0; i < 128; i++) U32(cfb, 10 * 512 + i * 4,
                i < 7 ? (uint)(i + 1) : i == 7 || i == 8 ? 0xfffffffe : i == 9 ? 0xfffffffd : uint.MaxValue);
            return cfb;
        }

        private static void U16(byte[] data, int offset, ushort value) {
            data[offset] = (byte)value; data[offset + 1] = (byte)(value >> 8);
        }
        private static void U32(byte[] data, int offset, uint value) {
            for (int index = 0; index < 4; index++) data[offset + index] = (byte)(value >> (index * 8));
        }
    }
}
