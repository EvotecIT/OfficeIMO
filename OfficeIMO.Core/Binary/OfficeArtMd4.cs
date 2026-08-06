using System;

namespace OfficeIMO.Drawing.Binary;

/// <summary>Computes the MD4 digest required by OfficeArt BLIP identifiers.</summary>
internal static class OfficeArtMd4 {
    internal static byte[] Compute(byte[] source) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        long bitLength = checked((long)source.Length * 8L);
        int paddingLength = 64 - ((source.Length + 9) & 63);
        if (paddingLength == 64) paddingLength = 0;
        var padded = new byte[checked(source.Length + 1 + paddingLength + 8)];
        Buffer.BlockCopy(source, 0, padded, 0, source.Length);
        padded[source.Length] = 0x80;
        WriteUInt64(padded, padded.Length - 8, unchecked((ulong)bitLength));

        uint a = 0x67452301U;
        uint b = 0xEFCDAB89U;
        uint c = 0x98BADCFEU;
        uint d = 0x10325476U;
        var words = new uint[16];
        var state = new uint[4];
        for (int blockOffset = 0; blockOffset < padded.Length;
             blockOffset += 64) {
            for (int index = 0; index < words.Length; index++) {
                words[index] = ReadUInt32(padded,
                    blockOffset + index * 4);
            }
            state[0] = a;
            state[1] = b;
            state[2] = c;
            state[3] = d;
            ApplyRound(state, words,
                new[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15 },
                new[] { 3, 7, 11, 19 }, 0U, round: 1);
            ApplyRound(state, words,
                new[] { 0, 4, 8, 12, 1, 5, 9, 13, 2, 6, 10, 14, 3, 7, 11, 15 },
                new[] { 3, 5, 9, 13 }, 0x5A827999U, round: 2);
            ApplyRound(state, words,
                new[] { 0, 8, 4, 12, 2, 10, 6, 14, 1, 9, 5, 13, 3, 11, 7, 15 },
                new[] { 3, 9, 11, 15 }, 0x6ED9EBA1U, round: 3);
            a = unchecked(a + state[0]);
            b = unchecked(b + state[1]);
            c = unchecked(c + state[2]);
            d = unchecked(d + state[3]);
        }

        var digest = new byte[16];
        WriteUInt32(digest, 0, a);
        WriteUInt32(digest, 4, b);
        WriteUInt32(digest, 8, c);
        WriteUInt32(digest, 12, d);
        return digest;
    }

    private static void ApplyRound(uint[] state, uint[] words,
        int[] order, int[] shifts, uint constant, int round) {
        for (int step = 0; step < 16; step++) {
            int target = (4 - (step & 3)) & 3;
            uint x = state[(target + 1) & 3];
            uint y = state[(target + 2) & 3];
            uint z = state[(target + 3) & 3];
            uint value = round == 1
                ? x & y | ~x & z
                : round == 2
                    ? x & y | x & z | y & z
                    : x ^ y ^ z;
            state[target] = RotateLeft(unchecked(state[target] + value
                + words[order[step]] + constant), shifts[step & 3]);
        }
    }

    private static uint RotateLeft(uint value, int count) =>
        value << count | value >> 32 - count;

    private static uint ReadUInt32(byte[] source, int offset) =>
        unchecked((uint)(source[offset]
            | source[offset + 1] << 8
            | source[offset + 2] << 16
            | source[offset + 3] << 24));

    private static void WriteUInt32(byte[] target, int offset, uint value) {
        target[offset] = unchecked((byte)value);
        target[offset + 1] = unchecked((byte)(value >> 8));
        target[offset + 2] = unchecked((byte)(value >> 16));
        target[offset + 3] = unchecked((byte)(value >> 24));
    }

    private static void WriteUInt64(byte[] target, int offset, ulong value) {
        for (int index = 0; index < 8; index++) {
            target[offset + index] = unchecked((byte)(value >> index * 8));
        }
    }
}
