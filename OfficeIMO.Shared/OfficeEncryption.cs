#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Shared {
    /// <summary>
    /// Hash algorithms supported by Office Agile encryption.
    /// </summary>
    internal enum OfficeEncryptionHashAlgorithm {
        /// <summary>
        /// SHA-1, retained for compatibility with older Agile encryption descriptors.
        /// </summary>
        Sha1,
        /// <summary>
        /// SHA-256.
        /// </summary>
        Sha256,
        /// <summary>
        /// SHA-384.
        /// </summary>
        Sha384,
        /// <summary>
        /// SHA-512, the default for new encrypted packages.
        /// </summary>
        Sha512
    }

    /// <summary>
    /// Options used when encrypting Open XML Office packages.
    /// </summary>
    internal sealed class OfficeEncryptionOptions {
        /// <summary>
        /// Default modern Office Agile encryption options: AES-256, SHA-512, 100,000 spin count.
        /// </summary>
        public static OfficeEncryptionOptions Default { get; } = new OfficeEncryptionOptions();

        /// <summary>
        /// Number of password hash iterations. Office's modern default is 100,000.
        /// </summary>
        public int SpinCount { get; set; } = 100000;

        /// <summary>
        /// AES key size in bits. Supported values are 128, 192, and 256.
        /// </summary>
        public int KeyBits { get; set; } = 256;

        /// <summary>
        /// Hash algorithm used for key derivation and integrity.
        /// </summary>
        public OfficeEncryptionHashAlgorithm HashAlgorithm { get; set; } = OfficeEncryptionHashAlgorithm.Sha512;

        internal void Validate() {
            if (SpinCount <= 0) {
                throw new ArgumentOutOfRangeException(nameof(SpinCount), "Spin count must be greater than zero.");
            }

            if (KeyBits != 128 && KeyBits != 192 && KeyBits != 256) {
                throw new ArgumentOutOfRangeException(nameof(KeyBits), "KeyBits must be 128, 192, or 256.");
            }
        }
    }

    /// <summary>
    /// Encrypts and decrypts Office Open XML packages using Microsoft Office Agile encryption.
    /// </summary>
    internal static class OfficeEncryption {
        private const int SaltSize = 16;
        private const int BlockSize = 16;
        private const int SegmentSize = 4096;
        private const string EncryptionNamespace = "http://schemas.microsoft.com/office/2006/encryption";
        private const string PasswordNamespace = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

        private static readonly byte[] VerifierHashInputBlockKey = { 0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79 };
        private static readonly byte[] VerifierHashValueBlockKey = { 0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e };
        private static readonly byte[] EncryptedKeyValueBlockKey = { 0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6 };
        private static readonly byte[] HmacKeyBlockKey = { 0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6 };
        private static readonly byte[] HmacValueBlockKey = { 0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33 };

        /// <summary>
        /// Returns true when the byte array appears to be an encrypted Office package.
        /// </summary>
        public static bool IsEncryptedPackage(byte[] bytes) {
            if (!CompoundFile.TryRead(bytes, out var streams)) {
                return false;
            }

            return streams.ContainsKey("EncryptionInfo") && streams.ContainsKey("EncryptedPackage");
        }

        /// <summary>
        /// Encrypts a plain Open XML package into the encrypted Office container format.
        /// </summary>
        public static byte[] EncryptPackage(byte[] packageBytes, string password, OfficeEncryptionOptions? options = null) {
            if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
            if (password == null) throw new ArgumentNullException(nameof(password));

            options ??= OfficeEncryptionOptions.Default;
            options.Validate();

            string hashName = GetHashName(options.HashAlgorithm);
            int hashSize = GetHashSize(options.HashAlgorithm);

            byte[]? keyDataSalt = null;
            byte[]? passwordSalt = null;
            byte[]? verifier = null;
            byte[]? verifierHash = null;
            byte[]? secretKey = null;

            try {
                keyDataSalt = RandomBytes(SaltSize);
                passwordSalt = RandomBytes(SaltSize);
                verifier = RandomBytes(SaltSize);
                secretKey = RandomBytes(options.KeyBits / 8);

                byte[] encryptedVerifierHashInput = EncryptWithPasswordDerivedKey(verifier, password, passwordSalt, options.SpinCount, hashName, options.KeyBits, VerifierHashInputBlockKey);
                verifierHash = Hash(verifier, hashName);
                byte[] encryptedVerifierHashValue = EncryptWithPasswordDerivedKey(verifierHash, password, passwordSalt, options.SpinCount, hashName, options.KeyBits, VerifierHashValueBlockKey);
                byte[] encryptedKeyValue = EncryptWithPasswordDerivedKey(secretKey, password, passwordSalt, options.SpinCount, hashName, options.KeyBits, EncryptedKeyValueBlockKey);

                byte[] encryptedPackage = EncryptPackagePayload(packageBytes, secretKey, keyDataSalt, hashName);
                byte[] encryptedHmacKey;
                byte[] encryptedHmacValue;
                GenerateIntegrityParameters(encryptedPackage, secretKey, keyDataSalt, hashName, out encryptedHmacKey, out encryptedHmacValue);

                byte[] encryptionInfo = BuildEncryptionInfo(
                    options,
                    hashName,
                    hashSize,
                    keyDataSalt,
                    passwordSalt,
                    encryptedVerifierHashInput,
                    encryptedVerifierHashValue,
                    encryptedKeyValue,
                    encryptedHmacKey,
                    encryptedHmacValue);

                return CompoundFile.Write(new Dictionary<string, byte[]>(StringComparer.Ordinal) {
                    ["EncryptionInfo"] = encryptionInfo,
                    ["EncryptedPackage"] = encryptedPackage
                });
            } finally {
                Clear(keyDataSalt, passwordSalt, verifier, verifierHash, secretKey);
            }
        }

        /// <summary>
        /// Decrypts an encrypted Office package and returns the inner Open XML package bytes.
        /// </summary>
        public static byte[] DecryptPackage(byte[] encryptedPackageBytes, string password) {
            if (encryptedPackageBytes == null) throw new ArgumentNullException(nameof(encryptedPackageBytes));
            if (password == null) throw new ArgumentNullException(nameof(password));

            if (!CompoundFile.TryRead(encryptedPackageBytes, out var streams) ||
                !streams.TryGetValue("EncryptionInfo", out var encryptionInfoBytes) ||
                !streams.TryGetValue("EncryptedPackage", out var encryptedPackage)) {
                throw new InvalidDataException("The document is not an encrypted Office package.");
            }

            var descriptor = AgileDescriptor.Parse(encryptionInfoBytes);
            byte[]? secretKey = null;

            try {
                secretKey = DecryptSecretKey(descriptor, password);
                VerifyIntegrity(descriptor, encryptedPackage, secretKey);
                return DecryptPackagePayload(encryptedPackage, secretKey, descriptor.KeyDataSaltValue, descriptor.KeyDataHashAlgorithm);
            } finally {
                Clear(secretKey);
            }
        }

        /// <summary>
        /// Encrypts package bytes and writes the encrypted Office container to a stream.
        /// </summary>
        public static void EncryptPackageToStream(byte[] packageBytes, string password, Stream destination, OfficeEncryptionOptions? options = null) {
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));

            var encryptedBytes = EncryptPackage(packageBytes, password, options);
            PrepareDestination(destination);
            destination.Write(encryptedBytes, 0, encryptedBytes.Length);
            try { destination.Flush(); } catch (NotSupportedException) { }
            if (destination.CanSeek) {
                destination.Seek(0, SeekOrigin.Begin);
            }
        }

        private static byte[] DecryptSecretKey(AgileDescriptor descriptor, string password) {
            byte[]? verifierKey = null;
            byte[]? verifierIv = null;
            byte[]? paddedVerifier = null;
            byte[]? verifier = null;
            byte[]? verifierHashKey = null;
            byte[]? verifierHashIv = null;
            byte[]? decryptedVerifierHash = null;
            byte[]? expectedHash = null;
            byte[]? keyKey = null;
            byte[]? keyIv = null;
            byte[]? decryptedKey = null;

            try {
                verifierKey = DeriveKey(
                    password,
                    descriptor.PasswordSaltValue,
                    descriptor.SpinCount,
                    descriptor.PasswordHashAlgorithm,
                    descriptor.PasswordKeyBits,
                    VerifierHashInputBlockKey);

                verifierIv = GenerateIv(descriptor.PasswordSaltValue, null, descriptor.PasswordHashAlgorithm);
                paddedVerifier = DecryptAes(descriptor.EncryptedVerifierHashInput, verifierKey, verifierIv);
                verifier = TrimTrailingPadding(paddedVerifier);

                verifierHashKey = DeriveKey(
                    password,
                    descriptor.PasswordSaltValue,
                    descriptor.SpinCount,
                    descriptor.PasswordHashAlgorithm,
                    descriptor.PasswordKeyBits,
                    VerifierHashValueBlockKey);

                verifierHashIv = GenerateIv(descriptor.PasswordSaltValue, null, descriptor.PasswordHashAlgorithm);
                decryptedVerifierHash = DecryptAes(descriptor.EncryptedVerifierHashValue, verifierHashKey, verifierHashIv);
                expectedHash = Hash(verifier, descriptor.PasswordHashAlgorithm);

                if (!ConstantTimeEquals(expectedHash, 0, decryptedVerifierHash, 0, descriptor.PasswordHashSize)) {
                    throw new CryptographicException("The password is incorrect.");
                }

                keyKey = DeriveKey(
                    password,
                    descriptor.PasswordSaltValue,
                    descriptor.SpinCount,
                    descriptor.PasswordHashAlgorithm,
                    descriptor.PasswordKeyBits,
                    EncryptedKeyValueBlockKey);

                keyIv = GenerateIv(descriptor.PasswordSaltValue, null, descriptor.PasswordHashAlgorithm);
                decryptedKey = DecryptAes(descriptor.EncryptedKeyValue, keyKey, keyIv);
                byte[] secretKey = new byte[descriptor.KeyDataKeyBits / 8];
                Buffer.BlockCopy(decryptedKey, 0, secretKey, 0, secretKey.Length);
                return secretKey;
            } finally {
                Clear(verifierKey, verifierIv, paddedVerifier, verifier, verifierHashKey, verifierHashIv, decryptedVerifierHash, expectedHash, keyKey, keyIv, decryptedKey);
            }
        }

        private static void VerifyIntegrity(AgileDescriptor descriptor, byte[] encryptedPackage, byte[] secretKey) {
            byte[]? hmacKeyIv = null;
            byte[]? hmacKey = null;
            byte[]? hmacKeyTrimmed = null;
            byte[]? actualHmac = null;
            byte[]? expectedHmacIv = null;
            byte[]? expectedHmac = null;

            try {
                hmacKeyIv = GenerateIv(descriptor.KeyDataSaltValue, HmacKeyBlockKey, descriptor.KeyDataHashAlgorithm);
                hmacKey = DecryptAes(
                    descriptor.EncryptedHmacKey,
                    secretKey,
                    hmacKeyIv);

                hmacKeyTrimmed = hmacKey.Take(descriptor.KeyDataSaltSize).ToArray();
                actualHmac = ComputeHmac(encryptedPackage, hmacKeyTrimmed, descriptor.KeyDataHashAlgorithm);
                expectedHmacIv = GenerateIv(descriptor.KeyDataSaltValue, HmacValueBlockKey, descriptor.KeyDataHashAlgorithm);
                expectedHmac = DecryptAes(
                    descriptor.EncryptedHmacValue,
                    secretKey,
                    expectedHmacIv);

                if (!ConstantTimeEquals(actualHmac, 0, expectedHmac, 0, actualHmac.Length)) {
                    throw new CryptographicException("The encrypted package integrity check failed.");
                }
            } finally {
                Clear(hmacKeyIv, hmacKey, hmacKeyTrimmed, actualHmac, expectedHmacIv, expectedHmac);
            }
        }

        private static byte[] EncryptWithPasswordDerivedKey(byte[] data, string password, byte[] salt, int spinCount, string hashName, int keyBits, byte[] blockKey) {
            byte[]? key = null;
            byte[]? iv = null;
            byte[]? padded = null;

            try {
                key = DeriveKey(password, salt, spinCount, hashName, keyBits, blockKey);
                iv = GenerateIv(salt, null, hashName);
                padded = PadToBlock(data);
                return EncryptAes(padded, key, iv);
            } finally {
                Clear(key, iv, padded);
            }
        }

        private static byte[] EncryptPackagePayload(byte[] packageBytes, byte[] secretKey, byte[] keyDataSalt, string hashName) {
            using var output = new MemoryStream(packageBytes.Length + BlockSize + 8);
            WriteUInt64(output, (ulong)packageBytes.Length);

            int segment = 0;
            for (int offset = 0; offset < packageBytes.Length; offset += SegmentSize, segment++) {
                int count = Math.Min(SegmentSize, packageBytes.Length - offset);
                byte[] plain = new byte[count];
                Buffer.BlockCopy(packageBytes, offset, plain, 0, count);
                byte[]? block = null;
                byte[]? iv = null;
                byte[]? encrypted = null;

                try {
                    block = count % BlockSize == 0 ? plain : PadToBlock(plain);
                    iv = GenerateIv(keyDataSalt, UInt32Bytes((uint)segment), hashName);
                    encrypted = EncryptAes(block, secretKey, iv);
                    output.Write(encrypted, 0, encrypted.Length);
                } finally {
                    Clear(plain, block, iv, encrypted);
                }
            }

            return output.ToArray();
        }

        private static byte[] DecryptPackagePayload(byte[] encryptedPackage, byte[] secretKey, byte[] keyDataSalt, string hashName) {
            if (encryptedPackage.Length < 8) {
                throw new InvalidDataException("EncryptedPackage stream is too small.");
            }

            ulong packageSize = ReadUInt64(encryptedPackage, 0);
            if (packageSize > int.MaxValue) {
                throw new NotSupportedException("Encrypted packages larger than 2 GB are not supported by this API.");
            }

            var packageBytes = new byte[(int)packageSize];
            int encryptedOffset = 8;
            int plainOffset = 0;
            int segment = 0;

            while (plainOffset < packageBytes.Length) {
                int remainingPlain = packageBytes.Length - plainOffset;
                int plainCount = Math.Min(SegmentSize, remainingPlain);
                int encryptedCount = RoundUp(plainCount, BlockSize);

                if (encryptedOffset + encryptedCount > encryptedPackage.Length) {
                    throw new InvalidDataException("EncryptedPackage stream ended before all package data was read.");
                }

                byte[]? encryptedSegment = null;
                byte[]? iv = null;
                byte[]? decrypted = null;

                try {
                    encryptedSegment = new byte[encryptedCount];
                    Buffer.BlockCopy(encryptedPackage, encryptedOffset, encryptedSegment, 0, encryptedCount);
                    iv = GenerateIv(keyDataSalt, UInt32Bytes((uint)segment), hashName);
                    decrypted = DecryptAes(encryptedSegment, secretKey, iv);
                    Buffer.BlockCopy(decrypted, 0, packageBytes, plainOffset, plainCount);
                } finally {
                    Clear(encryptedSegment, iv, decrypted);
                }

                encryptedOffset += encryptedCount;
                plainOffset += plainCount;
                segment++;
            }

            return packageBytes;
        }

        private static void GenerateIntegrityParameters(byte[] encryptedPackage, byte[] secretKey, byte[] keyDataSalt, string hashName, out byte[] encryptedHmacKey, out byte[] encryptedHmacValue) {
            byte[]? hmacKey = null;
            byte[]? hmacValue = null;
            byte[]? paddedHmacKey = null;
            byte[]? paddedHmacValue = null;
            byte[]? hmacKeyIv = null;
            byte[]? hmacValueIv = null;

            try {
                hmacKey = RandomBytes(SaltSize);
                hmacValue = ComputeHmac(encryptedPackage, hmacKey, hashName);

                paddedHmacKey = PadToBlock(hmacKey);
                hmacKeyIv = GenerateIv(keyDataSalt, HmacKeyBlockKey, hashName);
                encryptedHmacKey = EncryptAes(
                    paddedHmacKey,
                    secretKey,
                    hmacKeyIv);

                paddedHmacValue = PadToBlock(hmacValue);
                hmacValueIv = GenerateIv(keyDataSalt, HmacValueBlockKey, hashName);
                encryptedHmacValue = EncryptAes(
                    paddedHmacValue,
                    secretKey,
                    hmacValueIv);
            } finally {
                Clear(hmacKey, hmacValue, paddedHmacKey, paddedHmacValue, hmacKeyIv, hmacValueIv);
            }
        }

        private static byte[] BuildEncryptionInfo(
            OfficeEncryptionOptions options,
            string hashName,
            int hashSize,
            byte[] keyDataSalt,
            byte[] passwordSalt,
            byte[] encryptedVerifierHashInput,
            byte[] encryptedVerifierHashValue,
            byte[] encryptedKeyValue,
            byte[] encryptedHmacKey,
            byte[] encryptedHmacValue) {
            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<encryption xmlns=\"" + EncryptionNamespace + "\" xmlns:p=\"" + PasswordNamespace + "\">" +
                "<keyData saltSize=\"16\" blockSize=\"16\" keyBits=\"" + options.KeyBits.ToString(CultureInfo.InvariantCulture) + "\" hashSize=\"" + hashSize.ToString(CultureInfo.InvariantCulture) + "\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"" + hashName + "\" saltValue=\"" + Convert.ToBase64String(keyDataSalt) + "\"/>" +
                "<dataIntegrity encryptedHmacKey=\"" + Convert.ToBase64String(encryptedHmacKey) + "\" encryptedHmacValue=\"" + Convert.ToBase64String(encryptedHmacValue) + "\"/>" +
                "<keyEncryptors><keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">" +
                "<p:encryptedKey spinCount=\"" + options.SpinCount.ToString(CultureInfo.InvariantCulture) + "\" saltSize=\"16\" blockSize=\"16\" keyBits=\"" + options.KeyBits.ToString(CultureInfo.InvariantCulture) + "\" hashSize=\"" + hashSize.ToString(CultureInfo.InvariantCulture) + "\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"" + hashName + "\" saltValue=\"" + Convert.ToBase64String(passwordSalt) + "\" encryptedVerifierHashInput=\"" + Convert.ToBase64String(encryptedVerifierHashInput) + "\" encryptedVerifierHashValue=\"" + Convert.ToBase64String(encryptedVerifierHashValue) + "\" encryptedKeyValue=\"" + Convert.ToBase64String(encryptedKeyValue) + "\"/>" +
                "</keyEncryptor></keyEncryptors></encryption>";

            byte[] xmlBytes = Encoding.UTF8.GetBytes(xml);
            using var output = new MemoryStream(xmlBytes.Length + 8);
            WriteUInt16(output, 4);
            WriteUInt16(output, 4);
            WriteUInt32(output, 0x00000040);
            output.Write(xmlBytes, 0, xmlBytes.Length);
            return output.ToArray();
        }

        private static byte[] DeriveKey(string password, byte[] salt, int spinCount, string hashName, int keyBits, byte[] blockKey) {
            byte[]? passwordBytes = null;
            byte[]? initialInput = null;
            byte[]? hash = null;
            byte[]? iterationInput = null;
            byte[]? finalInput = null;
            byte[]? finalHash = null;

            try {
                passwordBytes = Encoding.Unicode.GetBytes(password);
                initialInput = Concat(salt, passwordBytes);
                hash = Hash(initialInput, hashName);

                for (uint i = 0; i < spinCount; i++) {
                    iterationInput = Concat(UInt32Bytes(i), hash);
                    byte[] nextHash = Hash(iterationInput, hashName);
                    Clear(iterationInput, hash);
                    iterationInput = null;
                    hash = nextHash;
                }

                finalInput = Concat(hash, blockKey);
                finalHash = Hash(finalInput, hashName);
                int keyBytes = keyBits / 8;
                byte[] result = new byte[keyBytes];
                int copy = Math.Min(finalHash.Length, result.Length);
                Buffer.BlockCopy(finalHash, 0, result, 0, copy);
                for (int i = copy; i < result.Length; i++) {
                    result[i] = 0x36;
                }
                return result;
            } finally {
                Clear(passwordBytes, initialInput, hash, iterationInput, finalInput, finalHash);
            }
        }

        private static byte[] GenerateIv(byte[] salt, byte[]? blockKey, string hashName) {
            byte[]? ivInput = null;
            byte[]? iv = null;

            try {
                if (blockKey == null) {
                    iv = salt;
                } else {
                    ivInput = Concat(salt, blockKey);
                    iv = Hash(ivInput, hashName);
                }

                byte[] result = new byte[BlockSize];
                int copy = Math.Min(iv.Length, result.Length);
                Buffer.BlockCopy(iv, 0, result, 0, copy);
                for (int i = copy; i < result.Length; i++) {
                    result[i] = 0x36;
                }
                return result;
            } finally {
                if (blockKey != null) {
                    Clear(ivInput, iv);
                }
            }
        }

        private static byte[] EncryptAes(byte[] data, byte[] key, byte[] iv) {
            using var aes = Aes.Create();
            aes.KeySize = key.Length * 8;
            aes.BlockSize = BlockSize * 8;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            aes.Key = key;
            aes.IV = iv;
            using var encryptor = aes.CreateEncryptor();
            return encryptor.TransformFinalBlock(data, 0, data.Length);
        }

        private static byte[] DecryptAes(byte[] data, byte[] key, byte[] iv) {
            using var aes = Aes.Create();
            aes.KeySize = key.Length * 8;
            aes.BlockSize = BlockSize * 8;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            aes.Key = key;
            aes.IV = iv;
            using var decryptor = aes.CreateDecryptor();
            return decryptor.TransformFinalBlock(data, 0, data.Length);
        }

        private static byte[] Hash(byte[] data, string hashName) {
            using HashAlgorithm algorithm = CreateHash(hashName);
            return algorithm.ComputeHash(data);
        }

        private static HashAlgorithm CreateHash(string hashName) {
            return NormalizeHashName(hashName) switch {
                "SHA1" => SHA1.Create(),
                "SHA256" => SHA256.Create(),
                "SHA384" => SHA384.Create(),
                "SHA512" => SHA512.Create(),
                _ => throw new NotSupportedException($"Hash algorithm '{hashName}' is not supported.")
            };
        }

        private static byte[] ComputeHmac(byte[] data, byte[] key, string hashName) {
            using HMAC hmac = NormalizeHashName(hashName) switch {
                "SHA1" => new HMACSHA1(key),
                "SHA256" => new HMACSHA256(key),
                "SHA384" => new HMACSHA384(key),
                "SHA512" => new HMACSHA512(key),
                _ => throw new NotSupportedException($"Hash algorithm '{hashName}' is not supported.")
            };
            return hmac.ComputeHash(data);
        }

        private static string GetHashName(OfficeEncryptionHashAlgorithm algorithm) {
            return algorithm switch {
                OfficeEncryptionHashAlgorithm.Sha1 => "SHA1",
                OfficeEncryptionHashAlgorithm.Sha256 => "SHA256",
                OfficeEncryptionHashAlgorithm.Sha384 => "SHA384",
                OfficeEncryptionHashAlgorithm.Sha512 => "SHA512",
                _ => throw new NotSupportedException($"Hash algorithm '{algorithm}' is not supported.")
            };
        }

        private static int GetHashSize(OfficeEncryptionHashAlgorithm algorithm) {
            return algorithm switch {
                OfficeEncryptionHashAlgorithm.Sha1 => 20,
                OfficeEncryptionHashAlgorithm.Sha256 => 32,
                OfficeEncryptionHashAlgorithm.Sha384 => 48,
                OfficeEncryptionHashAlgorithm.Sha512 => 64,
                _ => throw new NotSupportedException($"Hash algorithm '{algorithm}' is not supported.")
            };
        }

        private static string NormalizeHashName(string hashName) {
            return hashName.Replace("-", string.Empty).ToUpperInvariant();
        }

        private static byte[] RandomBytes(int length) {
            byte[] bytes = new byte[length];
            using var rng = RandomNumberGenerator.Create();
            rng.GetBytes(bytes);
            return bytes;
        }

        private static void Clear(params byte[]?[] buffers) {
            foreach (var buffer in buffers) {
                if (buffer != null) {
                    Array.Clear(buffer, 0, buffer.Length);
                }
            }
        }

        private static byte[] PadToBlock(byte[] data) {
            int paddedLength = RoundUp(data.Length, BlockSize);
            if (paddedLength == data.Length) {
                byte[] copy = new byte[data.Length];
                Buffer.BlockCopy(data, 0, copy, 0, data.Length);
                return copy;
            }

            byte[] padded = new byte[paddedLength];
            Buffer.BlockCopy(data, 0, padded, 0, data.Length);
            return padded;
        }

        private static byte[] TrimTrailingPadding(byte[] data) {
            int length = data.Length;
            while (length > 0 && data[length - 1] == 0) {
                length--;
            }

            byte[] result = new byte[length];
            Buffer.BlockCopy(data, 0, result, 0, length);
            return result;
        }

        private static bool ConstantTimeEquals(byte[] expected, int expectedOffset, byte[] actual, int actualOffset, int count) {
            if (expectedOffset < 0 || actualOffset < 0 || count < 0 ||
                expectedOffset + count > expected.Length ||
                actualOffset + count > actual.Length) {
                return false;
            }

            int diff = 0;
            for (int i = 0; i < count; i++) {
                diff |= expected[expectedOffset + i] ^ actual[actualOffset + i];
            }

            return diff == 0;
        }

        private static byte[] Concat(byte[] first, byte[] second) {
            byte[] result = new byte[first.Length + second.Length];
            Buffer.BlockCopy(first, 0, result, 0, first.Length);
            Buffer.BlockCopy(second, 0, result, first.Length, second.Length);
            return result;
        }

        private static int RoundUp(int value, int multiple) {
            if (value == 0) return 0;
            int remainder = value % multiple;
            return remainder == 0 ? value : value + multiple - remainder;
        }

        private static byte[] UInt32Bytes(uint value) {
            return new[] {
                (byte)(value & 0xff),
                (byte)((value >> 8) & 0xff),
                (byte)((value >> 16) & 0xff),
                (byte)((value >> 24) & 0xff)
            };
        }

        private static ushort ReadUInt16(byte[] data, int offset) {
            return (ushort)(data[offset] | (data[offset + 1] << 8));
        }

        private static uint ReadUInt32(byte[] data, int offset) {
            return (uint)(data[offset] |
                (data[offset + 1] << 8) |
                (data[offset + 2] << 16) |
                (data[offset + 3] << 24));
        }

        private static ulong ReadUInt64(byte[] data, int offset) {
            uint low = ReadUInt32(data, offset);
            uint high = ReadUInt32(data, offset + 4);
            return low | ((ulong)high << 32);
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
            stream.WriteByte((byte)((value >> 16) & 0xff));
            stream.WriteByte((byte)((value >> 24) & 0xff));
        }

        private static void WriteUInt64(Stream stream, ulong value) {
            WriteUInt32(stream, (uint)(value & 0xffffffff));
            WriteUInt32(stream, (uint)(value >> 32));
        }

        private static void PrepareDestination(Stream destination) {
            if (!destination.CanSeek) return;
            destination.Seek(0, SeekOrigin.Begin);
            destination.SetLength(0);
        }

        private sealed class AgileDescriptor {
            public byte[] KeyDataSaltValue = Array.Empty<byte>();
            public int KeyDataSaltSize;
            public int KeyDataKeyBits;
            public string KeyDataHashAlgorithm = "SHA512";
            public byte[] PasswordSaltValue = Array.Empty<byte>();
            public int PasswordKeyBits;
            public int PasswordHashSize;
            public string PasswordHashAlgorithm = "SHA512";
            public int SpinCount;
            public byte[] EncryptedVerifierHashInput = Array.Empty<byte>();
            public byte[] EncryptedVerifierHashValue = Array.Empty<byte>();
            public byte[] EncryptedKeyValue = Array.Empty<byte>();
            public byte[] EncryptedHmacKey = Array.Empty<byte>();
            public byte[] EncryptedHmacValue = Array.Empty<byte>();

            public static AgileDescriptor Parse(byte[] encryptionInfoBytes) {
                if (encryptionInfoBytes.Length < 8) {
                    throw new InvalidDataException("EncryptionInfo stream is too small.");
                }

                ushort major = ReadUInt16(encryptionInfoBytes, 0);
                ushort minor = ReadUInt16(encryptionInfoBytes, 2);
                if (major != 4 || minor != 4) {
                    throw new NotSupportedException("Only Office Agile encryption is supported.");
                }

                string xml = Encoding.UTF8.GetString(encryptionInfoBytes, 8, encryptionInfoBytes.Length - 8).TrimEnd('\0', ' ', '\r', '\n', '\t');
                var document = XDocument.Parse(xml);
                XNamespace ns = EncryptionNamespace;
                XNamespace p = PasswordNamespace;

                var keyData = document.Root?.Element(ns + "keyData") ?? throw new InvalidDataException("EncryptionInfo is missing keyData.");
                var dataIntegrity = document.Root.Element(ns + "dataIntegrity") ?? throw new InvalidDataException("EncryptionInfo is missing dataIntegrity.");
                var encryptedKey = document.Root
                    .Element(ns + "keyEncryptors")?
                    .Elements(ns + "keyEncryptor")
                    .Select(e => e.Element(p + "encryptedKey"))
                    .FirstOrDefault(e => e != null) ?? throw new InvalidDataException("EncryptionInfo is missing password encryptedKey.");

                return new AgileDescriptor {
                    KeyDataSaltSize = ReadRequiredInt(keyData, "saltSize"),
                    KeyDataKeyBits = ReadRequiredInt(keyData, "keyBits"),
                    KeyDataHashAlgorithm = ReadRequiredString(keyData, "hashAlgorithm"),
                    KeyDataSaltValue = Convert.FromBase64String(ReadRequiredString(keyData, "saltValue")),
                    EncryptedHmacKey = Convert.FromBase64String(ReadRequiredString(dataIntegrity, "encryptedHmacKey")),
                    EncryptedHmacValue = Convert.FromBase64String(ReadRequiredString(dataIntegrity, "encryptedHmacValue")),
                    SpinCount = ReadRequiredInt(encryptedKey, "spinCount"),
                    PasswordKeyBits = ReadRequiredInt(encryptedKey, "keyBits"),
                    PasswordHashSize = ReadRequiredInt(encryptedKey, "hashSize"),
                    PasswordHashAlgorithm = ReadRequiredString(encryptedKey, "hashAlgorithm"),
                    PasswordSaltValue = Convert.FromBase64String(ReadRequiredString(encryptedKey, "saltValue")),
                    EncryptedVerifierHashInput = Convert.FromBase64String(ReadRequiredString(encryptedKey, "encryptedVerifierHashInput")),
                    EncryptedVerifierHashValue = Convert.FromBase64String(ReadRequiredString(encryptedKey, "encryptedVerifierHashValue")),
                    EncryptedKeyValue = Convert.FromBase64String(ReadRequiredString(encryptedKey, "encryptedKeyValue"))
                };
            }

            private static string ReadRequiredString(XElement element, string name) {
                return element.Attribute(name)?.Value ?? throw new InvalidDataException($"EncryptionInfo is missing '{name}'.");
            }

            private static int ReadRequiredInt(XElement element, string name) {
                return int.Parse(ReadRequiredString(element, name), CultureInfo.InvariantCulture);
            }
        }

        private sealed class CompoundFile {
            private const int SectorSize = 512;
            private const int MiniSectorSize = 64;
            private const int DirectoryEntrySize = 128;
            private const int MiniStreamCutoff = 4096;
            private const uint FreeSect = 0xffffffff;
            private const uint EndOfChain = 0xfffffffe;
            private const uint FatSect = 0xfffffffd;
            private const uint DifSect = 0xfffffffc;

            public static byte[] Write(Dictionary<string, byte[]> streams) {
                var streamInfos = streams
                    .OrderBy(kvp => kvp.Key, StringComparer.Ordinal)
                    .Select(kvp => new StreamInfo(kvp.Key, kvp.Value))
                    .ToList();

                var miniStreams = streamInfos.Where(s => s.Data.Length < MiniStreamCutoff).ToList();
                var regularStreams = streamInfos.Where(s => s.Data.Length >= MiniStreamCutoff).ToList();

                byte[] miniStreamBytes = BuildMiniStream(miniStreams);
                byte[] miniFatBytes = BuildMiniFat(miniStreams);

                var regularChains = new List<RegularChain>();
                foreach (var info in regularStreams) {
                    regularChains.Add(new RegularChain(info, SplitIntoSectors(info.Data)));
                }

                var miniStreamSectors = SplitIntoSectors(miniStreamBytes);
                var miniFatSectors = SplitIntoSectors(miniFatBytes);
                int directorySectorCount = SplitIntoSectors(BuildDirectory(streamInfos, miniStreamBytes.Length, 0)).Count;

                int dataSectorCount = regularChains.Sum(c => c.Sectors.Count) + miniStreamSectors.Count + miniFatSectors.Count + directorySectorCount;
                int fatSectorCount = 0;
                int difatSectorCount = 0;
                while (true) {
                    int nextFatCount = CeilingDiv(dataSectorCount + fatSectorCount + difatSectorCount, SectorSize / 4);
                    int nextDifatCount = nextFatCount <= 109 ? 0 : CeilingDiv(nextFatCount - 109, 127);
                    if (nextFatCount == fatSectorCount && nextDifatCount == difatSectorCount) {
                        break;
                    }

                    fatSectorCount = nextFatCount;
                    difatSectorCount = nextDifatCount;
                }

                var sectors = new List<byte[]>();
                foreach (var chain in regularChains) {
                    chain.StartSector = sectors.Count;
                    sectors.AddRange(chain.Sectors);
                    chain.Info.StartSector = (uint)chain.StartSector;
                }

                int miniStreamStart = sectors.Count;
                sectors.AddRange(miniStreamSectors);

                int miniFatStart = sectors.Count;
                sectors.AddRange(miniFatSectors);

                int directoryStart = sectors.Count;
                for (int i = 0; i < directorySectorCount; i++) {
                    sectors.Add(new byte[SectorSize]);
                }

                var directorySectors = SplitIntoSectors(BuildDirectory(streamInfos, miniStreamBytes.Length, miniStreamStart));
                for (int i = 0; i < directorySectors.Count; i++) {
                    sectors[directoryStart + i] = directorySectors[i];
                }

                int fatStart = sectors.Count;
                for (int i = 0; i < fatSectorCount; i++) {
                    sectors.Add(new byte[SectorSize]);
                }

                int difatStart = sectors.Count;
                for (int i = 0; i < difatSectorCount; i++) {
                    sectors.Add(new byte[SectorSize]);
                }

                uint[] fat = Enumerable.Repeat(FreeSect, sectors.Count).ToArray();
                foreach (var chain in regularChains) {
                    MarkChain(fat, chain.StartSector, chain.Sectors.Count);
                }
                MarkChain(fat, miniStreamStart, miniStreamSectors.Count);
                MarkChain(fat, miniFatStart, miniFatSectors.Count);
                MarkChain(fat, directoryStart, directorySectors.Count);
                for (int i = 0; i < fatSectorCount; i++) fat[fatStart + i] = FatSect;
                for (int i = 0; i < difatSectorCount; i++) fat[difatStart + i] = DifSect;

                WriteFatSectors(sectors, fat, fatStart, fatSectorCount);
                WriteDifatSectors(sectors, fatStart, fatSectorCount, difatStart, difatSectorCount);

                using var output = new MemoryStream(512 + sectors.Count * SectorSize);
                WriteHeader(output, fatStart, fatSectorCount, directoryStart, miniFatStart, miniFatSectors.Count, difatStart, difatSectorCount);
                foreach (var sector in sectors) {
                    output.Write(sector, 0, sector.Length);
                }

                return output.ToArray();
            }

            public static bool TryRead(byte[] bytes, out Dictionary<string, byte[]> streams) {
                streams = new Dictionary<string, byte[]>(StringComparer.Ordinal);
                try {
                    if (bytes.Length < SectorSize) return false;
                    byte[] signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
                    for (int i = 0; i < signature.Length; i++) {
                        if (bytes[i] != signature[i]) return false;
                    }

                    int sectorShift = ReadUInt16(bytes, 30);
                    int sectorSize = 1 << sectorShift;
                    if (sectorSize != SectorSize) return false;

                    int fatSectorCount = (int)ReadUInt32(bytes, 44);
                    uint directoryStart = ReadUInt32(bytes, 48);
                    uint miniCutoff = ReadUInt32(bytes, 56);
                    uint miniFatStart = ReadUInt32(bytes, 60);
                    int miniFatSectorCount = (int)ReadUInt32(bytes, 64);
                    uint firstDifat = ReadUInt32(bytes, 68);
                    int difatSectorCount = (int)ReadUInt32(bytes, 72);

                    List<uint> fatSectorIds = ReadDifat(bytes, firstDifat, difatSectorCount, fatSectorCount);
                    uint[] fat = ReadFat(bytes, fatSectorIds);
                    byte[] directoryBytes = ReadRegularStream(bytes, fat, directoryStart, long.MaxValue);
                    var entries = ReadDirectoryEntries(directoryBytes);
                    var root = entries.FirstOrDefault(e => e.ObjectType == 5);
                    byte[] miniStream = root == null || root.StartSector == EndOfChain
                        ? Array.Empty<byte>()
                        : ReadRegularStream(bytes, fat, root.StartSector, root.Size);
                    uint[] miniFat = miniFatStart == EndOfChain || miniFatSectorCount == 0
                        ? Array.Empty<uint>()
                        : BytesToUInt32Array(ReadRegularStream(bytes, fat, miniFatStart, miniFatSectorCount * SectorSize));

                    foreach (var entry in entries.Where(e => e.ObjectType == 2)) {
                        byte[] data = entry.Size < miniCutoff
                            ? ReadMiniStream(miniStream, miniFat, entry.StartSector, entry.Size)
                            : ReadRegularStream(bytes, fat, entry.StartSector, entry.Size);
                        streams[entry.Name] = data;
                    }

                    return true;
                } catch {
                    streams = new Dictionary<string, byte[]>(StringComparer.Ordinal);
                    return false;
                }
            }

            private static byte[] BuildMiniStream(List<StreamInfo> miniStreams) {
                using var output = new MemoryStream();
                uint nextMiniSector = 0;
                foreach (var stream in miniStreams) {
                    stream.StartSector = nextMiniSector;
                    int sectors = CeilingDiv(stream.Data.Length, MiniSectorSize);
                    byte[] padded = new byte[sectors * MiniSectorSize];
                    Buffer.BlockCopy(stream.Data, 0, padded, 0, stream.Data.Length);
                    output.Write(padded, 0, padded.Length);
                    nextMiniSector += (uint)sectors;
                }

                return output.ToArray();
            }

            private static byte[] BuildMiniFat(List<StreamInfo> miniStreams) {
                var entries = new List<uint>();
                uint current = 0;
                foreach (var stream in miniStreams) {
                    int sectors = CeilingDiv(stream.Data.Length, MiniSectorSize);
                    for (int i = 0; i < sectors; i++) {
                        entries.Add(i == sectors - 1 ? EndOfChain : current + 1);
                        current++;
                    }
                }

                using var output = new MemoryStream();
                foreach (uint entry in entries) {
                    OfficeEncryption.WriteUInt32(output, entry);
                }

                return output.ToArray();
            }

            private static byte[] BuildDirectory(List<StreamInfo> streams, int miniStreamSize, int miniStreamStart) {
                using var output = new MemoryStream();
                WriteDirectoryEntry(output, "Root Entry", 5, EndOfChain, EndOfChain, streams.Count > 0 ? 1u : EndOfChain, miniStreamSize > 0 ? (uint)miniStreamStart : EndOfChain, (ulong)miniStreamSize);
                for (int i = 0; i < streams.Count; i++) {
                    uint left = EndOfChain;
                    uint right = i + 1 < streams.Count ? (uint)(i + 2) : EndOfChain;
                    WriteDirectoryEntry(output, streams[i].Name, 2, left, right, EndOfChain, streams[i].StartSector, (ulong)streams[i].Data.Length);
                }

                int remainder = (int)(output.Length % SectorSize);
                if (remainder != 0) {
                    output.Write(new byte[SectorSize - remainder], 0, SectorSize - remainder);
                }

                return output.ToArray();
            }

            private static List<byte[]> SplitIntoSectors(byte[] data) {
                if (data.Length == 0) return new List<byte[]>();
                int count = CeilingDiv(data.Length, SectorSize);
                var sectors = new List<byte[]>(count);
                for (int i = 0; i < count; i++) {
                    byte[] sector = new byte[SectorSize];
                    int offset = i * SectorSize;
                    int length = Math.Min(SectorSize, data.Length - offset);
                    Buffer.BlockCopy(data, offset, sector, 0, length);
                    sectors.Add(sector);
                }

                return sectors;
            }

            private static void MarkChain(uint[] fat, int start, int count) {
                if (count <= 0 || start < 0) return;
                for (int i = 0; i < count; i++) {
                    fat[start + i] = i == count - 1 ? EndOfChain : (uint)(start + i + 1);
                }
            }

            private static void WriteFatSectors(List<byte[]> sectors, uint[] fat, int fatStart, int fatSectorCount) {
                int index = 0;
                for (int i = 0; i < fatSectorCount; i++) {
                    using var stream = new MemoryStream(sectors[fatStart + i], writable: true);
                    for (int j = 0; j < SectorSize / 4; j++) {
                        OfficeEncryption.WriteUInt32(stream, index < fat.Length ? fat[index++] : FreeSect);
                    }
                }
            }

            private static void WriteDifatSectors(List<byte[]> sectors, int fatStart, int fatSectorCount, int difatStart, int difatSectorCount) {
                int fatIndex = 109;
                for (int i = 0; i < difatSectorCount; i++) {
                    using var stream = new MemoryStream(sectors[difatStart + i], writable: true);
                    for (int j = 0; j < 127; j++) {
                        if (fatIndex < fatSectorCount) {
                            OfficeEncryption.WriteUInt32(stream, (uint)(fatStart + fatIndex));
                            fatIndex++;
                        } else {
                            OfficeEncryption.WriteUInt32(stream, FreeSect);
                        }
                    }
                    OfficeEncryption.WriteUInt32(stream, i + 1 < difatSectorCount ? (uint)(difatStart + i + 1) : EndOfChain);
                }
            }

            private static void WriteHeader(Stream output, int fatStart, int fatSectorCount, int directoryStart, int miniFatStart, int miniFatSectorCount, int difatStart, int difatSectorCount) {
                byte[] header = new byte[SectorSize];
                using var stream = new MemoryStream(header, writable: true);
                stream.Write(new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }, 0, 8);
                stream.Position = 24;
                OfficeEncryption.WriteUInt16(stream, 0x003e);
                OfficeEncryption.WriteUInt16(stream, 0x0003);
                OfficeEncryption.WriteUInt16(stream, 0xfffe);
                OfficeEncryption.WriteUInt16(stream, 0x0009);
                OfficeEncryption.WriteUInt16(stream, 0x0006);
                stream.Position = 44;
                OfficeEncryption.WriteUInt32(stream, (uint)fatSectorCount);
                OfficeEncryption.WriteUInt32(stream, (uint)directoryStart);
                OfficeEncryption.WriteUInt32(stream, 0);
                OfficeEncryption.WriteUInt32(stream, MiniStreamCutoff);
                OfficeEncryption.WriteUInt32(stream, miniFatSectorCount > 0 ? (uint)miniFatStart : EndOfChain);
                OfficeEncryption.WriteUInt32(stream, (uint)miniFatSectorCount);
                OfficeEncryption.WriteUInt32(stream, difatSectorCount > 0 ? (uint)difatStart : EndOfChain);
                OfficeEncryption.WriteUInt32(stream, (uint)difatSectorCount);
                for (int i = 0; i < 109; i++) {
                    OfficeEncryption.WriteUInt32(stream, i < fatSectorCount && i < 109 ? (uint)(fatStart + i) : FreeSect);
                }
                output.Write(header, 0, header.Length);
            }

            private static void WriteDirectoryEntry(Stream output, string name, byte objectType, uint leftSibling, uint rightSibling, uint childId, uint startSector, ulong size) {
                byte[] entry = new byte[DirectoryEntrySize];
                byte[] nameBytes = Encoding.Unicode.GetBytes(name + '\0');
                if (nameBytes.Length > 64) {
                    throw new InvalidOperationException("Compound file directory entry name is too long.");
                }
                Buffer.BlockCopy(nameBytes, 0, entry, 0, nameBytes.Length);
                WriteUInt16(entry, 64, (ushort)nameBytes.Length);
                entry[66] = objectType;
                entry[67] = 1;
                WriteUInt32(entry, 68, leftSibling);
                WriteUInt32(entry, 72, rightSibling);
                WriteUInt32(entry, 76, childId);
                WriteUInt32(entry, 116, startSector);
                WriteUInt64(entry, 120, size);
                output.Write(entry, 0, entry.Length);
            }

            private static List<uint> ReadDifat(byte[] bytes, uint firstDifat, int difatSectorCount, int fatSectorCount) {
                var result = new List<uint>(fatSectorCount);
                for (int i = 0; i < 109 && result.Count < fatSectorCount; i++) {
                    uint sector = ReadUInt32(bytes, 76 + i * 4);
                    if (sector != FreeSect) result.Add(sector);
                }

                uint next = firstDifat;
                for (int d = 0; d < difatSectorCount && next != EndOfChain && result.Count < fatSectorCount; d++) {
                    int offset = SectorOffset(next);
                    for (int i = 0; i < 127 && result.Count < fatSectorCount; i++) {
                        uint sector = ReadUInt32(bytes, offset + i * 4);
                        if (sector != FreeSect) result.Add(sector);
                    }
                    next = ReadUInt32(bytes, offset + 127 * 4);
                }

                return result;
            }

            private static uint[] ReadFat(byte[] bytes, List<uint> fatSectorIds) {
                var entries = new List<uint>(fatSectorIds.Count * (SectorSize / 4));
                foreach (uint sector in fatSectorIds) {
                    int offset = SectorOffset(sector);
                    for (int i = 0; i < SectorSize / 4; i++) {
                        entries.Add(ReadUInt32(bytes, offset + i * 4));
                    }
                }
                return entries.ToArray();
            }

            private static byte[] ReadRegularStream(byte[] bytes, uint[] fat, uint startSector, long size) {
                if (startSector == EndOfChain) return Array.Empty<byte>();

                using var output = new MemoryStream();
                uint sector = startSector;
                var visited = new HashSet<uint>();
                while (sector != EndOfChain && sector != FreeSect) {
                    if (sector >= fat.Length || !visited.Add(sector)) {
                        throw new InvalidDataException("Invalid compound file sector chain.");
                    }

                    int offset = SectorOffset(sector);
                    output.Write(bytes, offset, SectorSize);
                    sector = fat[sector];
                    if (output.Length >= size && size != long.MaxValue) {
                        break;
                    }
                }

                byte[] data = output.ToArray();
                if (size != long.MaxValue && data.LongLength > size) {
                    Array.Resize(ref data, (int)size);
                }
                return data;
            }

            private static byte[] ReadMiniStream(byte[] miniStream, uint[] miniFat, uint startSector, long size) {
                if (startSector == EndOfChain) return Array.Empty<byte>();

                using var output = new MemoryStream();
                uint sector = startSector;
                var visited = new HashSet<uint>();
                while (sector != EndOfChain && sector != FreeSect) {
                    if (sector >= miniFat.Length || !visited.Add(sector)) {
                        throw new InvalidDataException("Invalid compound file mini sector chain.");
                    }

                    int offset = checked((int)sector * MiniSectorSize);
                    output.Write(miniStream, offset, Math.Min(MiniSectorSize, miniStream.Length - offset));
                    sector = miniFat[sector];
                    if (output.Length >= size) {
                        break;
                    }
                }

                byte[] data = output.ToArray();
                if (data.LongLength > size) {
                    Array.Resize(ref data, (int)size);
                }
                return data;
            }

            private static List<DirectoryEntry> ReadDirectoryEntries(byte[] directoryBytes) {
                var result = new List<DirectoryEntry>();
                for (int offset = 0; offset + DirectoryEntrySize <= directoryBytes.Length; offset += DirectoryEntrySize) {
                    ushort nameLength = ReadUInt16(directoryBytes, offset + 64);
                    byte objectType = directoryBytes[offset + 66];
                    if (objectType == 0 || nameLength < 2 || nameLength > 64) continue;
                    string name = Encoding.Unicode.GetString(directoryBytes, offset, nameLength - 2);
                    result.Add(new DirectoryEntry(
                        name,
                        objectType,
                        ReadUInt32(directoryBytes, offset + 116),
                        (long)ReadUInt64(directoryBytes, offset + 120)));
                }
                return result;
            }

            private static uint[] BytesToUInt32Array(byte[] bytes) {
                uint[] result = new uint[bytes.Length / 4];
                for (int i = 0; i < result.Length; i++) {
                    result[i] = ReadUInt32(bytes, i * 4);
                }
                return result;
            }

            private static int SectorOffset(uint sector) {
                return checked(SectorSize + (int)sector * SectorSize);
            }

            private static int CeilingDiv(int value, int divisor) {
                return value == 0 ? 0 : ((value - 1) / divisor) + 1;
            }

            private static void WriteUInt16(byte[] buffer, int offset, ushort value) {
                buffer[offset] = (byte)(value & 0xff);
                buffer[offset + 1] = (byte)((value >> 8) & 0xff);
            }

            private static void WriteUInt32(byte[] buffer, int offset, uint value) {
                buffer[offset] = (byte)(value & 0xff);
                buffer[offset + 1] = (byte)((value >> 8) & 0xff);
                buffer[offset + 2] = (byte)((value >> 16) & 0xff);
                buffer[offset + 3] = (byte)((value >> 24) & 0xff);
            }

            private static void WriteUInt64(byte[] buffer, int offset, ulong value) {
                WriteUInt32(buffer, offset, (uint)(value & 0xffffffff));
                WriteUInt32(buffer, offset + 4, (uint)(value >> 32));
            }

            private sealed class StreamInfo {
                public StreamInfo(string name, byte[] data) {
                    Name = name;
                    Data = data;
                }

                public string Name { get; }
                public byte[] Data { get; }
                public uint StartSector { get; set; } = EndOfChain;
            }

            private sealed class RegularChain {
                public RegularChain(StreamInfo info, List<byte[]> sectors) {
                    Info = info;
                    Sectors = sectors;
                }

                public StreamInfo Info { get; }
                public List<byte[]> Sectors { get; }
                public int StartSector { get; set; } = -1;
            }

            private sealed class DirectoryEntry {
                public DirectoryEntry(string name, byte objectType, uint startSector, long size) {
                    Name = name;
                    ObjectType = objectType;
                    StartSector = startSector;
                    Size = size;
                }

                public string Name { get; }
                public byte ObjectType { get; }
                public uint StartSector { get; }
                public long Size { get; }
            }
        }
    }
}
