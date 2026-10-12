#nullable enable
using System;
using System.Security.Cryptography;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeEncryption {
        private static byte[] GetIntegrityKey(AgileDescriptor descriptor, byte[] paddedKey) {
            int keyLength;
            // MS-OFFCRYPTO 2.3.4.14 specifies a saltSize-byte HMAC key. Some
            // producers use a hashSize-byte key instead. The encrypted length
            // distinguishes those layouts; neither permits arbitrary prefixes.
            int saltSize = descriptor.KeyDataSaltSize;
            if (saltSize > 0 && saltSize <= paddedKey.Length &&
                paddedKey.Length - saltSize < BlockSize) {
                keyLength = saltSize;
            } else {
                using HashAlgorithm algorithm = CreateHash(descriptor.KeyDataHashAlgorithm);
                keyLength = algorithm.HashSize / 8;
                if (RoundUp(keyLength, BlockSize) != paddedKey.Length) {
                    throw new CryptographicException("The encrypted integrity key has an unsupported length.");
                }
            }

            for (int index = keyLength; index < paddedKey.Length; index++) {
                if (paddedKey[index] != 0) {
                    throw new CryptographicException("The encrypted integrity key has invalid padding.");
                }
            }

            var key = new byte[keyLength];
            Buffer.BlockCopy(paddedKey, 0, key, 0, keyLength);
            return key;
        }
    }
}
