using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    internal sealed class PacketMap {
        internal int[] Offsets = Array.Empty<int>();
        internal int[] Lengths = Array.Empty<int>();
        internal int DataOffset, StreamEnd;
        internal int Profile, Level;
    }

    internal static PacketMap ReadPackets(byte[] bytes, int offset, int length, FrameHeader frame, CancellationToken token) {
        int end = checked(offset + length);
        var bits = new Bits(bytes, frame.HeaderEnd, end - frame.HeaderEnd, token);
        int tileCount = checked(frame.TileWidths.Length * frame.TileHeights.Length);
        int packetCount = checked(tileCount * (frame.Frequency ? 4 - frame.Primary.Bands : 1));
        if (packetCount > 1_000_000 || (long)packetCount * sizeof(int) * 3 + bytes.Length > 256 * 1024 * 1024)
            throw new FormatException("JPEG-XR tile index exceeds the managed limit.");
        var map = new PacketMap { Offsets = new int[packetCount], Lengths = new int[packetCount], StreamEnd = end };
        if (frame.IndexTable) {
            if (bits.Read(16) != 1) throw new FormatException("JPEG-XR tile index signature is invalid.");
            for (int i = 0; i < packetCount; i++) {
                token.ThrowIfCancellationRequested();
                ulong value = ReadVariableWord(bits);
                if (value > int.MaxValue) throw new FormatException("JPEG-XR tile index offset exceeds the managed range.");
                map.Offsets[i] = (int)value;
            }
        } else if (packetCount != 1) {
            throw new FormatException("JPEG-XR multiple packets require a tile index.");
        }
        ulong subsequent = ReadVariableWord(bits);
        if (subsequent > (ulong)(end - bits.ByteOffset)) throw new FormatException("JPEG-XR profile segment exceeds the codestream.");
        if (subsequent != 0) {
            int remaining = (int)subsequent;
            bool last;
            do {
                if (remaining < 4) throw new FormatException("JPEG-XR profile level list is truncated.");
                map.Profile = (int)bits.Read(8); map.Level = (int)bits.Read(8);
                bits.Read(15); last = bits.Flag(); remaining -= 4;
            } while (!last);
            bits.SkipAligned(remaining);
        }
        map.DataOffset = bits.ByteOffset;
        if (packetCount == 1 && map.Offsets[0] != 0)
            throw new FormatException("JPEG-XR single packet index must have zero offset.");
        for (int i = 0; i < packetCount; i++) {
            token.ThrowIfCancellationRequested();
            if (map.Offsets[i] > end - map.DataOffset - 4)
                throw new FormatException("JPEG-XR tile packet offset is outside its codestream.");
            map.Offsets[i] += map.DataOffset;
            int p = map.Offsets[i];
            // T.832 8.7.10.2 requires decoders to ignore ARBITRARY_BYTE.
            // Packet bands follow index-table order; that byte is not a type tag.
            if (bytes[p] != 0 || bytes[p + 1] != 0 || bytes[p + 2] != 1)
                throw new FormatException("JPEG-XR tile packet signature is invalid.");
        }
        // Packets may be stored in an order different from their index entries.
        // Sort an index vector once to bound each reader without quadratic scans.
        var ordered = new int[packetCount];
        for (int i = 0; i < ordered.Length; i++) ordered[i] = i;
        Array.Sort(ordered, (a, b) => map.Offsets[a].CompareTo(map.Offsets[b]));
        token.ThrowIfCancellationRequested();
        for (int i = 0; i < ordered.Length; i++) {
            if ((i & 1023) == 0) token.ThrowIfCancellationRequested();
            int index = ordered[i], next = i + 1 == ordered.Length ? end : map.Offsets[ordered[i + 1]];
            map.Lengths[index] = next - map.Offsets[index];
            if (map.Lengths[index] < 4) throw new FormatException("JPEG-XR tile packet ranges overlap.");
        }
        return map;
    }

    private static ulong ReadVariableWord(Bits bits) {
        uint first = bits.Read(8);
        if (first < 0xFB) return first * 256UL + bits.Read(8);
        if (first == 0xFB) return bits.Read(32);
        if (first == 0xFC) return ((ulong)bits.Read(32) << 32) | bits.Read(32);
        return 0; // T.832 8.2.4 escape-mode values FD/FE/FF.
    }
}
