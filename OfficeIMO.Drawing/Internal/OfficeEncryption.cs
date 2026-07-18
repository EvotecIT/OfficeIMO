#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Xml.Linq;

namespace OfficeIMO.Drawing.Internal {
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
        public static OfficeEncryptionOptions Default => new OfficeEncryptionOptions();

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
    internal static partial class OfficeEncryption {
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
        public static byte[] DecryptPackage(byte[] encryptedPackageBytes,
            string password) => DecryptPackage(encryptedPackageBytes,
            password, CancellationToken.None);

        /// <summary>
        /// Decrypts an encrypted Office package and observes cancellation while deriving keys and decrypting payload segments.
        /// </summary>
        public static byte[] DecryptPackage(byte[] encryptedPackageBytes,
            string password, CancellationToken cancellationToken) {
            return DecryptPackage(encryptedPackageBytes, password,
                cancellationToken, maximumDecryptedPackageBytes: null);
        }

        internal static byte[] DecryptPackage(byte[] encryptedPackageBytes,
            string password, CancellationToken cancellationToken,
            long? maximumDecryptedPackageBytes) {
            if (encryptedPackageBytes == null) throw new ArgumentNullException(nameof(encryptedPackageBytes));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (maximumDecryptedPackageBytes.HasValue
                && maximumDecryptedPackageBytes.Value < 1) {
                throw new ArgumentOutOfRangeException(
                    nameof(maximumDecryptedPackageBytes));
            }
            cancellationToken.ThrowIfCancellationRequested();

            long compoundByteLimit = Math.Max(1L,
                encryptedPackageBytes.LongLength);
            int directoryEntryLimit = Math.Max(4, Math.Min(65536,
                checked((int)Math.Min(int.MaxValue,
                    compoundByteLimit / 128L + 1L))));
            var compoundOptions = new OfficeCompoundReadOptions(
                maxDirectoryEntries: directoryEntryLimit,
                maxStreamCount: Math.Max(2, Math.Min(32768,
                    directoryEntryLimit)),
                maxStreamBytes: compoundByteLimit,
                maxTotalStreamBytes: compoundByteLimit);
            if (!OfficeCompoundFileReader.TryRead(encryptedPackageBytes,
                    compoundOptions, out OfficeCompoundFile? compound,
                    out string? compoundError)
                || compound == null
                || !compound.Streams.TryGetValue("EncryptionInfo",
                    out byte[]? encryptionInfoBytes)
                || !compound.Streams.TryGetValue("EncryptedPackage",
                    out byte[]? encryptedPackage)) {
                throw new InvalidDataException(compoundError
                    ?? "The document is not an encrypted Office package.");
            }

            cancellationToken.ThrowIfCancellationRequested();
            var descriptor = AgileDescriptor.Parse(encryptionInfoBytes);
            byte[]? secretKey = null;

            try {
                secretKey = DecryptSecretKey(descriptor, password,
                    cancellationToken);
                VerifyIntegrity(descriptor, encryptedPackage, secretKey,
                    cancellationToken);
                return DecryptPackagePayload(encryptedPackage, secretKey,
                    descriptor.KeyDataSaltValue,
                    descriptor.KeyDataHashAlgorithm, cancellationToken,
                    maximumDecryptedPackageBytes);
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

        private static byte[] DecryptSecretKey(AgileDescriptor descriptor,
            string password, CancellationToken cancellationToken) {
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
                    VerifierHashInputBlockKey, cancellationToken);

                verifierIv = GenerateIv(descriptor.PasswordSaltValue, null, descriptor.PasswordHashAlgorithm);
                verifier = DecryptAes(descriptor.EncryptedVerifierHashInput, verifierKey, verifierIv);

                verifierHashKey = DeriveKey(
                    password,
                    descriptor.PasswordSaltValue,
                    descriptor.SpinCount,
                    descriptor.PasswordHashAlgorithm,
                    descriptor.PasswordKeyBits,
                    VerifierHashValueBlockKey, cancellationToken);

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
                    EncryptedKeyValueBlockKey, cancellationToken);

                keyIv = GenerateIv(descriptor.PasswordSaltValue, null, descriptor.PasswordHashAlgorithm);
                decryptedKey = DecryptAes(descriptor.EncryptedKeyValue, keyKey, keyIv);
                byte[] secretKey = new byte[descriptor.KeyDataKeyBits / 8];
                Buffer.BlockCopy(decryptedKey, 0, secretKey, 0, secretKey.Length);
                return secretKey;
            } finally {
                Clear(verifierKey, verifierIv, paddedVerifier, verifier, verifierHashKey, verifierHashIv, decryptedVerifierHash, expectedHash, keyKey, keyIv, decryptedKey);
            }
        }

        private static void VerifyIntegrity(AgileDescriptor descriptor,
            byte[] encryptedPackage, byte[] secretKey,
            CancellationToken cancellationToken) {
            byte[]? hmacKeyIv = null;
            byte[]? hmacKey = null;
            byte[]? hmacKeyTrimmed = null;
            byte[]? actualHmac = null;
            byte[]? expectedHmacIv = null;
            byte[]? expectedHmac = null;

            try {
                cancellationToken.ThrowIfCancellationRequested();
                hmacKeyIv = GenerateIv(descriptor.KeyDataSaltValue, HmacKeyBlockKey, descriptor.KeyDataHashAlgorithm);
                hmacKey = DecryptAes(
                    descriptor.EncryptedHmacKey,
                    secretKey,
                    hmacKeyIv);

                hmacKeyTrimmed = hmacKey.Take(descriptor.KeyDataSaltSize).ToArray();
                actualHmac = ComputeHmac(encryptedPackage, hmacKeyTrimmed, descriptor.KeyDataHashAlgorithm);
                cancellationToken.ThrowIfCancellationRequested();
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

        private static byte[] DecryptPackagePayload(byte[] encryptedPackage,
            byte[] secretKey, byte[] keyDataSalt, string hashName,
            CancellationToken cancellationToken,
            long? maximumDecryptedPackageBytes) {
            cancellationToken.ThrowIfCancellationRequested();
            if (encryptedPackage.Length < 8) {
                throw new InvalidDataException("EncryptedPackage stream is too small.");
            }

            ulong packageSize = ReadUInt64(encryptedPackage, 0);
            if (maximumDecryptedPackageBytes.HasValue
                && packageSize > unchecked((ulong)
                    maximumDecryptedPackageBytes.Value)) {
                throw new InvalidDataException(
                    $"The decrypted Office package declares {packageSize} bytes, exceeding the configured maximum of {maximumDecryptedPackageBytes.Value} bytes.");
            }
            if (packageSize > int.MaxValue) {
                throw new NotSupportedException("Encrypted packages larger than 2 GB are not supported by this API.");
            }

            var packageBytes = new byte[(int)packageSize];
            int encryptedOffset = 8;
            int plainOffset = 0;
            int segment = 0;

            while (plainOffset < packageBytes.Length) {
                cancellationToken.ThrowIfCancellationRequested();
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

        private static byte[] DeriveKey(string password, byte[] salt,
            int spinCount, string hashName, int keyBits, byte[] blockKey,
            CancellationToken cancellationToken = default) {
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
                    if ((i & 1023U) == 0U) {
                        cancellationToken.ThrowIfCancellationRequested();
                    }
                    iterationInput = Concat(UInt32Bytes(i), hash);
                    byte[] nextHash = Hash(iterationInput, hashName);
                    Clear(iterationInput, hash);
                    iterationInput = null;
                    hash = nextHash;
                }

                cancellationToken.ThrowIfCancellationRequested();
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

    }
}
