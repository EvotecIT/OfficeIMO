using System.IO.Compression;

namespace OfficeIMO.Archives.Tests {
    internal sealed partial class ArchiveFixtures {
        internal static byte[] FramedCab(ushort compression, byte[][] frames, int[] sizes,
            params (string Name, uint Offset, int Length)[] members) {
            byte[][] names = members.Select(m => Encoding.ASCII.GetBytes(m.Name + "\0")).ToArray();
            int start = 44 + names.Sum(n => n.Length) + members.Length * 16;
            using var stream = new MemoryStream();
            using var writer = new BinaryWriter(stream);
            writer.Write(Encoding.ASCII.GetBytes("MSCF")); writer.Write(0);
            writer.Write(start + frames.Sum(f => 8 + f.Length)); writer.Write(0); writer.Write(44); writer.Write(0);
            writer.Write((byte)3); writer.Write((byte)1); writer.Write((ushort)1); writer.Write((ushort)members.Length);
            writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)0);
            writer.Write(start); writer.Write((ushort)frames.Length); writer.Write(compression);
            for (int i = 0; i < members.Length; i++) {
                writer.Write(members[i].Length); writer.Write(members[i].Offset);
                writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)0);
                writer.Write(names[i]);
            }
            for (int i = 0; i < frames.Length; i++) {
                writer.Write(0); writer.Write((ushort)frames[i].Length); writer.Write((ushort)sizes[i]); writer.Write(frames[i]);
            }
            return stream.ToArray();
        }

        internal static byte[] MsZipFrame(byte[] bytes) {
            using var stream = new MemoryStream();
            stream.WriteByte(0x43); stream.WriteByte(0x4b);
            using (var deflate = new DeflateStream(stream, CompressionLevel.Optimal, leaveOpen: true))
                deflate.Write(bytes, 0, bytes.Length);
            return stream.ToArray();
        }

        internal static byte[] LzxRawFrame(byte[] bytes, bool first, int translationSize = 0, uint distance = 1) {
            var bits = new List<int>();
            if (first) {
                LzxBits(bits, translationSize == 0 ? 0 : 1, 1);
                if (translationSize != 0) LzxBits(bits, translationSize, 32);
            }
            LzxBits(bits, 3, 3); LzxBits(bits, bytes.Length, 24);
            using var stream = new MemoryStream();
            using var writer = new BinaryWriter(stream);
            writer.Write(LzxWords(bits)); writer.Write(distance); writer.Write(1u); writer.Write(1u); writer.Write(bytes);
            if ((bytes.Length & 1) != 0) writer.Write((byte)0);
            return stream.ToArray();
        }

        internal static byte[] LzxE8HistoryFrame() {
            var bits = new List<int>();
            LzxBits(bits, 1, 3); LzxBits(bits, 85, 24);
            Tree(0, 256); Tree(256, 240); Tree(0, 249, empty: true);
            for (int i = 0; i < 64; i++) LzxBits(bits, 0, 1); // literal A
            LzxBits(bits, 1, 1); // length 5, repeated offset R0 = 32768
            for (int i = 0; i < 16; i++) LzxBits(bits, 0, 1); // keep the E8 call outside the untranslated frame tail
            return LzxWords(bits);

            void Tree(int start, int count, bool empty = false) {
                for (int symbol = 0; symbol < 20; symbol++) LzxBits(bits, symbol == 0 || symbol == 16 ? 1 : 0, 4);
                for (int i = start; i < start + count; i++) LzxBits(bits, !empty && (i == 65 || i == 259) ? 1 : 0, 1);
            }
        }

        private static void LzxBits(List<int> bits, int value, int count) {
            for (int i = count - 1; i >= 0; i--) bits.Add((value >> i) & 1);
        }
        private static byte[] LzxWords(List<int> bits) {
            while ((bits.Count & 15) != 0) bits.Add(0);
            var output = new byte[bits.Count / 8];
            for (int i = 0; i < bits.Count; i += 16) {
                ushort value = 0;
                for (int j = 0; j < 16; j++) value = (ushort)((value << 1) | bits[i + j]);
                BitConverter.GetBytes(value).CopyTo(output, i / 8);
            }
            return output;
        }
    }
}
