namespace OfficeIMO.Archives.Rar5;

/// <summary>Unkeyed BLAKE2s and BLAKE2sp as defined in RFC 7693 and the BLAKE2 paper.</summary>
internal static class Blake2s {
    private const int BlockSize = 64;
    private const int DigestSize = 32;
    private const int Lanes = 8;
    private static readonly uint[] InitialWords = {
        0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19
    };
    private static readonly byte[] Schedule = {
        0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15,
        14, 10, 4, 8, 9, 15, 13, 6, 1, 12, 0, 2, 11, 7, 5, 3,
        11, 8, 12, 0, 5, 2, 15, 13, 10, 14, 3, 6, 7, 1, 9, 4,
        7, 9, 3, 1, 13, 12, 11, 14, 2, 6, 5, 10, 4, 0, 15, 8,
        9, 0, 5, 7, 2, 4, 10, 15, 14, 1, 11, 12, 6, 8, 3, 13,
        2, 12, 6, 10, 0, 11, 8, 3, 4, 13, 7, 5, 15, 14, 1, 9,
        12, 5, 1, 15, 14, 13, 4, 10, 0, 7, 6, 3, 9, 2, 8, 11,
        13, 11, 7, 14, 12, 1, 3, 9, 5, 0, 15, 4, 8, 6, 2, 10,
        6, 15, 14, 9, 11, 3, 0, 8, 12, 2, 13, 7, 1, 4, 10, 5,
        10, 2, 8, 4, 7, 6, 1, 5, 15, 11, 9, 14, 3, 12, 13, 0
    };

    internal static byte[] Digest(byte[] input, CancellationToken token = default) {
        var digest = new byte[DigestSize];
        HashNode(input, input.Length, digest, 0, 0, BlockSize, 0x01010020, 0, 0, false, token);
        return digest;
    }

    internal static byte[] TreeDigest(byte[] input, int offset, int length, CancellationToken token) {
        var digest = new byte[DigestSize];
        var leaves = new byte[Lanes * DigestSize];
        const uint treeParameter = 0x02080020; // Digest length 32, fanout 8, depth 2.
        for (int lane = 0; lane < Lanes; lane++) {
            HashNode(input, offset + length, leaves, lane * DigestSize, offset + Math.Min(length, lane * BlockSize), Lanes * BlockSize,
                treeParameter, (uint)lane, 0x20000000, lane == Lanes - 1, token);
        }
        HashNode(leaves, leaves.Length, digest, 0, 0, BlockSize, treeParameter, 0, 0x20010000, true, token);
        return digest;
    }

    private static void HashNode(byte[] input, int end, byte[] digest, int digestOffset, int start, int stride,
        uint mode, uint offset, uint depth, bool lastNode, CancellationToken token) {
        var state = (uint[])InitialWords.Clone();
        state[0] ^= mode;
        state[2] ^= offset;
        state[3] ^= depth;
        var finalBlock = new byte[BlockSize];
        var words = new uint[16];
        var work = new uint[16];
        long consumed = 0;
        int position = start;
        int remaining = end - start;
        while (remaining > stride) {
            if ((consumed & 0xffff) == 0) token.ThrowIfCancellationRequested();
            consumed += BlockSize;
            Compress(state, input, position, (ulong)consumed, false, false, words, work);
            position += stride;
            remaining -= stride;
        }

        token.ThrowIfCancellationRequested();
        int size = Math.Min(remaining, BlockSize);
        Buffer.BlockCopy(input, position, finalBlock, 0, size);
        Compress(state, finalBlock, 0, (ulong)(consumed + size), true, lastNode, words, work);
        for (int i = 0; i < state.Length; i++) {
            for (int j = 0; j < sizeof(uint); j++) digest[digestOffset + i * sizeof(uint) + j] = (byte)(state[i] >> (8 * j));
        }
    }

    private static void Compress(uint[] state, byte[] block, int offset, ulong count, bool lastBlock, bool lastNode,
        uint[] words, uint[] work) {
        for (int i = 0; i < words.Length; i++) words[i] = Rar5HeaderReader.UInt32At(block, offset + i * sizeof(uint));
        Array.Copy(state, work, state.Length);
        Array.Copy(InitialWords, 0, work, 8, InitialWords.Length);
        work[12] ^= (uint)count;
        work[13] ^= (uint)(count >> 32);
        if (lastBlock) work[14] = ~work[14];
        if (lastNode) work[15] = ~work[15];

        for (int round = 0; round < 10; round++) {
            int order = round * 16;
            for (int i = 0; i < 4; i++) {
                Mix(work, i, i + 4, i + 8, i + 12, words[Schedule[order + 2 * i]], words[Schedule[order + 2 * i + 1]]);
            }
            for (int i = 0; i < 4; i++) {
                Mix(work, i, 4 + ((i + 1) & 3), 8 + ((i + 2) & 3), 12 + ((i + 3) & 3),
                    words[Schedule[order + 8 + 2 * i]], words[Schedule[order + 9 + 2 * i]]);
            }
        }

        for (int i = 0; i < state.Length; i++) state[i] ^= work[i] ^ work[i + 8];
    }

    private static void Mix(uint[] v, int a, int b, int c, int d, uint first, uint second) {
        v[a] = unchecked(v[a] + v[b] + first);
        v[d] = RotateRight(v[d] ^ v[a], 16);
        v[c] = unchecked(v[c] + v[d]);
        v[b] = RotateRight(v[b] ^ v[c], 12);
        v[a] = unchecked(v[a] + v[b] + second);
        v[d] = RotateRight(v[d] ^ v[a], 8);
        v[c] = unchecked(v[c] + v[d]);
        v[b] = RotateRight(v[b] ^ v[c], 7);
    }

    private static uint RotateRight(uint word, int count) => (word >> count) | (word << (32 - count));
}
