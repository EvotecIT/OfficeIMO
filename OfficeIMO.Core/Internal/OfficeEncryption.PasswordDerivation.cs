#nullable enable
using System;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
#if NET8_0_OR_GREATER
using System.Buffers.Binary;
#endif

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeEncryption {
        // The iterated password hash is shared only by the block-specific keys
        // within one encryption/decryption operation. Its caller clears it.
        private static byte[] DerivePasswordHash(string password, byte[] salt,
            int spinCount, string hashName, CancellationToken cancellationToken) {
            byte[]? passwordBytes = null;
            byte[]? initialInput = null;
            try {
                cancellationToken.ThrowIfCancellationRequested();
                passwordBytes = Encoding.Unicode.GetBytes(password);
                initialInput = Concat(salt, passwordBytes);
#if NET8_0_OR_GREATER
                using HashAlgorithm algorithm = CreateHash(hashName);
                return DerivePasswordHashModern(algorithm, initialInput,
                    spinCount, cancellationToken);
#else
                return DerivePasswordHashLegacy(initialInput, spinCount,
                    hashName, cancellationToken);
#endif
            } finally {
                Clear(passwordBytes, initialInput);
            }
        }

#if NET8_0_OR_GREATER
        private static byte[] DerivePasswordHashModern(HashAlgorithm algorithm,
            byte[] initialInput, int spinCount, CancellationToken cancellationToken) {
            // All supported digests fit in 64 bytes. Separate input/output
            // spans avoid depending on a hashing provider's overlap behavior.
            Span<byte> digest = stackalloc byte[64];
            Span<byte> iterationInput = stackalloc byte[sizeof(uint) + 64];
            int hashSize = algorithm.HashSize / 8;
            Span<byte> hash = digest.Slice(0, hashSize);
            Span<byte> input = iterationInput.Slice(0, sizeof(uint) + hashSize);
            try {
                HashPasswordInput(algorithm, initialInput, hash);
                for (uint iteration = 0; iteration < spinCount; iteration++) {
                    if ((iteration & 1023U) == 0U) {
                        cancellationToken.ThrowIfCancellationRequested();
                    }
                    BinaryPrimitives.WriteUInt32LittleEndian(input, iteration);
                    hash.CopyTo(input.Slice(sizeof(uint)));
                    HashPasswordInput(algorithm, input, hash);
                }
                cancellationToken.ThrowIfCancellationRequested();
                return hash.ToArray();
            } finally {
                CryptographicOperations.ZeroMemory(digest);
                CryptographicOperations.ZeroMemory(iterationInput);
            }
        }

        private static void HashPasswordInput(HashAlgorithm algorithm,
            ReadOnlySpan<byte> input, Span<byte> hash) {
            // TryComputeHash resets the operation-local provider for the next
            // iteration without retaining the digest in HashAlgorithm.Hash.
            if (!algorithm.TryComputeHash(input, hash, out int written)
                || written != hash.Length) {
                throw new CryptographicException("The password hash length is invalid.");
            }
        }
#else
        private static byte[] DerivePasswordHashLegacy(byte[] initialInput,
            int spinCount, string hashName, CancellationToken cancellationToken) {
            byte[]? hash = null;
            byte[]? iterationInput = null;
            try {
                // Framework ComputeHash retains a second digest in HashValue.
                // A fresh canonical provider clears it on every Dispose; reuse
                // would abandon earlier digests when HashValue is overwritten.
                hash = Hash(initialInput, hashName);
                iterationInput = new byte[sizeof(uint) + hash.Length];
                for (uint iteration = 0; iteration < spinCount; iteration++) {
                    if ((iteration & 1023U) == 0U) {
                        cancellationToken.ThrowIfCancellationRequested();
                    }
                    iterationInput[0] = (byte)(iteration & 0xff);
                    iterationInput[1] = (byte)((iteration >> 8) & 0xff);
                    iterationInput[2] = (byte)((iteration >> 16) & 0xff);
                    iterationInput[3] = (byte)((iteration >> 24) & 0xff);
                    Buffer.BlockCopy(hash, 0, iterationInput, sizeof(uint), hash.Length);
                    byte[] nextHash = Hash(iterationInput, hashName);
                    Array.Clear(hash, 0, hash.Length);
                    hash = nextHash;
                }
                cancellationToken.ThrowIfCancellationRequested();
                byte[] result = hash;
                hash = null;
                return result;
            } finally {
                Clear(hash, iterationInput);
            }
        }
#endif

        private static byte[] DeriveKeyFromPasswordHash(byte[] passwordHash,
            string hashName, int keyBits, byte[] blockKey,
            CancellationToken cancellationToken = default) {
            byte[]? finalInput = null;
            byte[]? finalHash = null;
            try {
                cancellationToken.ThrowIfCancellationRequested();
                finalInput = Concat(passwordHash, blockKey);
                finalHash = Hash(finalInput, hashName);
                byte[] result = new byte[keyBits / 8];
                int copy = Math.Min(finalHash.Length, result.Length);
                Buffer.BlockCopy(finalHash, 0, result, 0, copy);
                for (int index = copy; index < result.Length; index++) {
                    result[index] = 0x36;
                }
                return result;
            } finally {
                Clear(finalInput, finalHash);
            }
        }

        private static byte[] EncryptWithPasswordDerivedKey(byte[] data,
            byte[] passwordHash, byte[] salt, string hashName, int keyBits,
            byte[] blockKey) {
            byte[]? key = null;
            byte[]? iv = null;
            byte[]? padded = null;
            try {
                key = DeriveKeyFromPasswordHash(passwordHash, hashName, keyBits, blockKey);
                iv = GenerateIv(salt, null, hashName);
                padded = PadToBlock(data);
                return EncryptAes(padded, key, iv);
            } finally {
                Clear(key, iv, padded);
            }
        }
    }
}
