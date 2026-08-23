using System;
using System.Security.Cryptography;

namespace OfficeIMO.Security {
    /// <summary>
    /// Provides a portable managed AES-CBC implementation for hosts where the platform AES factory is unavailable.
    /// </summary>
    /// <remarks>
    /// The implementation supports 128-bit, 192-bit, and 256-bit AES keys with no padding or PKCS#7 padding.
    /// It is intended for compatibility hosts such as browser WebAssembly. Applications running on a platform with
    /// native AES support can leave the PDF provider unset and use the platform implementation.
    /// </remarks>
    public sealed class OfficeManagedAesCryptographyProvider : IOfficeAesCryptographyProvider {
        private const int BlockSize = 16;

        /// <summary>Gets the shared stateless provider instance.</summary>
        public static OfficeManagedAesCryptographyProvider Default { get; } = new OfficeManagedAesCryptographyProvider();

        /// <inheritdoc />
        public string Name => "OfficeIMO.Core/Managed AES";

        /// <inheritdoc />
        public byte[] EncryptCbc(byte[] key, byte[] iv, byte[] input, OfficeAesPadding padding) {
            ValidateArguments(key, iv, input, padding);
            if (padding == OfficeAesPadding.None && (input.Length & (BlockSize - 1)) != 0) {
                throw new ArgumentException("Unpadded AES-CBC input must contain complete 16-byte blocks.", nameof(input));
            }

            byte[] plaintext = ApplyPadding(input, padding);
            byte[] output = new byte[plaintext.Length];
            byte[] previous = (byte[])iv.Clone();
            byte[] block = new byte[BlockSize];
            bool completed = false;

            try {
                using (var cipher = new ManagedAesBlockCipher(key, forEncryption: true)) {
                    for (int offset = 0; offset < plaintext.Length; offset += BlockSize) {
                        for (int index = 0; index < BlockSize; index++) {
                            block[index] = (byte)(plaintext[offset + index] ^ previous[index]);
                        }

                        cipher.ProcessBlock(block, 0, output, offset);
                        Buffer.BlockCopy(output, offset, previous, 0, BlockSize);
                    }
                }

                completed = true;
                return output;
            } finally {
                if (!completed) {
                    Array.Clear(output, 0, output.Length);
                }
                Array.Clear(previous, 0, previous.Length);
                Array.Clear(block, 0, block.Length);
                if (!ReferenceEquals(plaintext, input)) {
                    Array.Clear(plaintext, 0, plaintext.Length);
                }
            }
        }

        /// <inheritdoc />
        public byte[] DecryptCbc(byte[] key, byte[] iv, byte[] input, OfficeAesPadding padding) {
            ValidateArguments(key, iv, input, padding);
            if ((input.Length & (BlockSize - 1)) != 0) {
                throw new ArgumentException("AES-CBC ciphertext must contain complete 16-byte blocks.", nameof(input));
            }
            if (input.Length == 0 && padding == OfficeAesPadding.Pkcs7) {
                throw new CryptographicException("PKCS#7 ciphertext must contain at least one AES block.");
            }

            byte[] plaintext = new byte[input.Length];
            byte[] previous = (byte[])iv.Clone();
            byte[] block = new byte[BlockSize];
            bool completed = false;

            try {
                using (var cipher = new ManagedAesBlockCipher(key, forEncryption: false)) {
                    for (int offset = 0; offset < input.Length; offset += BlockSize) {
                        cipher.ProcessBlock(input, offset, block, 0);
                        for (int index = 0; index < BlockSize; index++) {
                            plaintext[offset + index] = (byte)(block[index] ^ previous[index]);
                        }

                        Buffer.BlockCopy(input, offset, previous, 0, BlockSize);
                    }
                }

                byte[] result = padding == OfficeAesPadding.Pkcs7
                    ? RemovePadding(plaintext)
                    : plaintext;
                completed = true;
                return result;
            } finally {
                if (!completed) {
                    Array.Clear(plaintext, 0, plaintext.Length);
                }
                Array.Clear(previous, 0, previous.Length);
                Array.Clear(block, 0, block.Length);
            }
        }

        private static void ValidateArguments(byte[] key, byte[] iv, byte[] input, OfficeAesPadding padding) {
            if (key == null) {
                throw new ArgumentNullException(nameof(key));
            }
            if (key.Length != 16 && key.Length != 24 && key.Length != 32) {
                throw new ArgumentException("AES keys must contain 16, 24, or 32 bytes.", nameof(key));
            }
            if (iv == null) {
                throw new ArgumentNullException(nameof(iv));
            }
            if (iv.Length != BlockSize) {
                throw new ArgumentException("AES-CBC initialization vectors must contain 16 bytes.", nameof(iv));
            }
            if (input == null) {
                throw new ArgumentNullException(nameof(input));
            }
            if (padding != OfficeAesPadding.None && padding != OfficeAesPadding.Pkcs7) {
                throw new ArgumentOutOfRangeException(nameof(padding));
            }
        }

        private static byte[] ApplyPadding(byte[] input, OfficeAesPadding padding) {
            if (padding == OfficeAesPadding.None) {
                return input;
            }

            int paddingLength = BlockSize - (input.Length & (BlockSize - 1));
            byte[] padded = new byte[input.Length + paddingLength];
            Buffer.BlockCopy(input, 0, padded, 0, input.Length);
            for (int index = input.Length; index < padded.Length; index++) {
                padded[index] = (byte)paddingLength;
            }
            return padded;
        }

        private static byte[] RemovePadding(byte[] plaintext) {
            int paddingLength = plaintext[plaintext.Length - 1];
            int invalid = paddingLength < 1 || paddingLength > BlockSize ? 1 : 0;
            int mismatch = 0;

            for (int index = 1; index <= BlockSize; index++) {
                int mask = index <= paddingLength ? -1 : 0;
                mismatch |= (plaintext[plaintext.Length - index] ^ paddingLength) & mask;
            }

            if ((invalid | mismatch) != 0) {
                Array.Clear(plaintext, 0, plaintext.Length);
                throw new CryptographicException("Invalid PKCS#7 padding.");
            }

            byte[] result = new byte[plaintext.Length - paddingLength];
            Buffer.BlockCopy(plaintext, 0, result, 0, result.Length);
            Array.Clear(plaintext, 0, plaintext.Length);
            return result;
        }
    }
}
