using System.Security.Cryptography;
using OfficeIMO.Security;

namespace OfficeIMO.Pdf;

/// <summary>
/// AES-GCM decryption for crypt filter /AESV4 (ISO/TS 32003): each string or stream is a 12-byte nonce, the
/// ciphertext and a 16-byte authentication tag, under the 256-bit file encryption key with no additional data.
/// </summary>
/// <remarks>
/// .NET 8 and later use the platform <c>AesGcm</c>. .NET Framework, .NET Standard 2.0 hosts, platforms without
/// <c>AesGcm</c> (such as browser WebAssembly) and callers that supply an <see cref="IOfficeAesCryptographyProvider"/>
/// use the managed GCM construction of NIST SP 800-38D over the AES block cipher. Both verify the tag before any
/// plaintext is returned.
/// </remarks>
internal static class PdfAesGcm {
    internal const int NonceLength = 12;
    internal const int TagLength = 16;
    private const int ChunkBlocks = 4096;

    internal static byte[] Decrypt(byte[] key, byte[] data, IOfficeAesCryptographyProvider? provider, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (data.Length < NonceLength + TagLength) {
            throw new PdfEncryptedContentIntegrityException("An AES-GCM encrypted string or stream is shorter than its nonce and authentication tag.");
        }

        var nonce = new byte[NonceLength];
        Buffer.BlockCopy(data, 0, nonce, 0, NonceLength);
        var ciphertext = new byte[data.Length - NonceLength - TagLength];
        Buffer.BlockCopy(data, NonceLength, ciphertext, 0, ciphertext.Length);
        var tag = new byte[TagLength];
        Buffer.BlockCopy(data, data.Length - TagLength, tag, 0, TagLength);

#if NET8_0_OR_GREATER
        if (provider is null && AesGcm.IsSupported) {
            var plaintext = new byte[ciphertext.Length];
            try {
                using var gcm = new AesGcm(key, TagLength);
                gcm.Decrypt(nonce, ciphertext, tag, plaintext);
            } catch (CryptographicException) {
                throw TagMismatch();
            }

            cancellationToken.ThrowIfCancellationRequested();
            return plaintext;
        }
#endif
        return DecryptManaged(key, nonce, ciphertext, tag, provider, cancellationToken);
    }

    private static PdfEncryptedContentIntegrityException TagMismatch() =>
        new("AES-GCM authentication failed: encrypted PDF content was modified or does not belong to this document.");

    private static byte[] DecryptManaged(byte[] key, byte[] nonce, byte[] ciphertext, byte[] tag,
        IOfficeAesCryptographyProvider? provider, CancellationToken cancellationToken) {
        using var cipher = new BlockCipher(key, provider);
        byte[] hashKey = cipher.Encrypt(new byte[16]);
        ulong hHigh = ReadUInt64(hashKey, 0);
        ulong hLow = ReadUInt64(hashKey, 8);

        // GHASH over the ciphertext (no additional authenticated data) and the length block.
        ulong yHigh = 0;
        ulong yLow = 0;
        var block = new byte[16];
        for (int offset = 0; offset < ciphertext.Length; offset += 16) {
            if ((offset & 0xFFFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(16, ciphertext.Length - offset);
            Array.Clear(block, 0, 16);
            Buffer.BlockCopy(ciphertext, offset, block, 0, count);
            yHigh ^= ReadUInt64(block, 0);
            yLow ^= ReadUInt64(block, 8);
            Multiply(ref yHigh, ref yLow, hHigh, hLow);
        }

        yLow ^= (ulong)ciphertext.Length * 8UL;
        Multiply(ref yHigh, ref yLow, hHigh, hLow);

        var counter = new byte[16];
        Buffer.BlockCopy(nonce, 0, counter, 0, NonceLength);
        counter[15] = 1;
        byte[] tagMask = cipher.Encrypt(counter);
        int difference = 0;
        for (int index = 0; index < 8; index++) {
            difference |= tag[index] ^ tagMask[index] ^ (byte)(yHigh >> (56 - 8 * index));
            difference |= tag[8 + index] ^ tagMask[8 + index] ^ (byte)(yLow >> (56 - 8 * index));
        }

        if (difference != 0) throw TagMismatch();

        // CTR decryption from inc32(J0).
        var plaintext = new byte[ciphertext.Length];
        uint counterValue = 2;
        var counterBlocks = new byte[Math.Min(ChunkBlocks, (ciphertext.Length + 15) / 16) * 16];
        for (int offset = 0; offset < ciphertext.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int blocks = Math.Min(ChunkBlocks, (ciphertext.Length - offset + 15) / 16);
            for (int index = 0; index < blocks; index++) {
                int position = index * 16;
                Buffer.BlockCopy(nonce, 0, counterBlocks, position, NonceLength);
                counterBlocks[position + 12] = (byte)(counterValue >> 24);
                counterBlocks[position + 13] = (byte)(counterValue >> 16);
                counterBlocks[position + 14] = (byte)(counterValue >> 8);
                counterBlocks[position + 15] = (byte)counterValue;
                counterValue++;
            }

            byte[] keyStream = cipher.EncryptBlocks(counterBlocks, blocks * 16);
            int bytes = Math.Min(blocks * 16, ciphertext.Length - offset);
            for (int index = 0; index < bytes; index++) {
                plaintext[offset + index] = (byte)(ciphertext[offset + index] ^ keyStream[index]);
            }

            offset += bytes;
        }

        return plaintext;
    }

    // Multiplication in GF(2^128) with the GCM bit order (NIST SP 800-38D, algorithm 1).
    private static void Multiply(ref ulong xHigh, ref ulong xLow, ulong yHigh, ulong yLow) {
        ulong zHigh = 0;
        ulong zLow = 0;
        ulong vHigh = yHigh;
        ulong vLow = yLow;
        for (int bit = 0; bit < 128; bit++) {
            ulong word = bit < 64 ? xHigh : xLow;
            ulong mask = 0UL - ((word >> (63 - (bit & 63))) & 1UL);
            zHigh ^= vHigh & mask;
            zLow ^= vLow & mask;
            ulong reduce = 0UL - (vLow & 1UL);
            vLow = (vLow >> 1) | (vHigh << 63);
            vHigh = (vHigh >> 1) ^ (0xE100000000000000UL & reduce);
        }

        xHigh = zHigh;
        xLow = zLow;
    }

    private static ulong ReadUInt64(byte[] buffer, int offset) {
        ulong value = 0;
        for (int index = 0; index < 8; index++) value = (value << 8) | buffer[offset + index];
        return value;
    }

    /// <summary>AES block encryption (ECB of independent blocks) through the platform or the caller's CBC provider.</summary>
    private sealed class BlockCipher : IDisposable {
        private readonly byte[] _key;
        private readonly IOfficeAesCryptographyProvider? _provider;
        private readonly Aes? _aes;
        private readonly ICryptoTransform? _transform;

        internal BlockCipher(byte[] key, IOfficeAesCryptographyProvider? provider) {
            _key = key;
            _provider = provider;
            if (provider is not null) return;
            try {
                _aes = Aes.Create();
                _aes.Mode = CipherMode.ECB;
                _aes.Padding = PaddingMode.None;
                _aes.Key = key;
                _transform = _aes.CreateEncryptor();
            } catch (PlatformNotSupportedException exception) {
                _aes?.Dispose();
                throw new PlatformNotSupportedException(
                    "This host does not provide AES. Supply an IOfficeAesCryptographyProvider, such as OfficeManagedAesCryptographyProvider from OfficeIMO.Core.",
                    exception);
            }
        }

        internal byte[] Encrypt(byte[] block) => EncryptBlocks(block, 16);

        internal byte[] EncryptBlocks(byte[] blocks, int length) {
            var output = new byte[length];
            if (_transform is not null) {
                _transform.TransformBlock(blocks, 0, length, output, 0);
                return output;
            }

            // A one-block CBC encryption with a zero IV is the AES block function.
            var iv = new byte[16];
            var single = new byte[16];
            for (int offset = 0; offset < length; offset += 16) {
                Buffer.BlockCopy(blocks, offset, single, 0, 16);
                byte[] encrypted = PdfAesCryptography.EncryptNoPadding(_key, iv, single, _provider);
                Buffer.BlockCopy(encrypted, 0, output, offset, 16);
            }

            return output;
        }

        public void Dispose() {
            _transform?.Dispose();
            _aes?.Dispose();
        }
    }
}
