using System.Security.Cryptography;

namespace OfficeIMO.OpenDocument;

internal static class OdfPasswordEncryption {
    internal const string Aes256Cbc = "http://www.w3.org/2001/04/xmlenc#aes256-cbc";
    internal const string Sha256 = "http://www.w3.org/2000/09/xmldsig#sha256";
    internal const string Sha256XmlEncryptionAlias = "http://www.w3.org/2001/04/xmlenc#sha256";
    internal const string Sha256OneKilobyte = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#sha256-1k";
    internal const string Pbkdf2 = "PBKDF2";
    private const int KeySize = 32;
    private const int SaltSize = 16;
    private const int IvSize = 16;
    private const int MaxPasswordBytes = 1024;

    internal static byte[] CreateStartKey(string password) {
        if (password == null) throw new ArgumentNullException(nameof(password));
        byte[] passwordBytes = Encoding.UTF8.GetBytes(password);
        if (passwordBytes.Length == 0) throw new ArgumentException("An ODF encryption password cannot be empty.", nameof(password));
        if (passwordBytes.Length > MaxPasswordBytes) {
            throw new ArgumentException($"An ODF encryption password cannot exceed {MaxPasswordBytes} UTF-8 bytes.", nameof(password));
        }
        try {
            using var sha256 = SHA256.Create();
            return sha256.ComputeHash(passwordBytes);
        } finally {
            Array.Clear(passwordBytes, 0, passwordBytes.Length);
        }
    }

    internal static void ValidateIterationCount(int iterationCount) {
        if (iterationCount < 10000 || iterationCount > 10000000) {
            throw new ArgumentOutOfRangeException(nameof(iterationCount), iterationCount,
                "ODF PBKDF2 iteration count must be between 10,000 and 10,000,000.");
        }
    }

    internal static OdfEncryptedEntry Encrypt(byte[] plaintext, byte[] startKey, int iterationCount) {
        byte[] compressed = OdfZipWriter.Deflate(plaintext);
        byte[] salt = RandomBytes(SaltSize);
        byte[] iv = RandomBytes(IvSize);
        byte[] key = DeriveKey(startKey, salt, iterationCount, KeySize);
        byte[] padded = ApplyW3cPadding(compressed, IvSize);
        byte[] ciphertext;
        try {
            using (Aes aes = Aes.Create()) {
                aes.KeySize = 256;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                aes.Key = key;
                aes.IV = iv;
                using ICryptoTransform transform = aes.CreateEncryptor();
                ciphertext = transform.TransformFinalBlock(padded, 0, padded.Length);
            }
        } finally {
            Array.Clear(key, 0, key.Length);
        }
        using var sha256 = SHA256.Create();
        int checksumLength = Math.Min(1024, compressed.Length);
        byte[] checksum = sha256.ComputeHash(compressed, 0, checksumLength);
        return new OdfEncryptedEntry(ciphertext, salt, iv, checksum, plaintext.LongLength, iterationCount);
    }

    internal static byte[] Decrypt(byte[] ciphertext, byte[] startKey, byte[] salt, byte[] iv,
        int iterationCount, byte[] expectedChecksum, long expectedSize, long maxBytes, string entryPath) {
        try {
            ValidateIterationCount(iterationCount);
        } catch (ArgumentOutOfRangeException ex) {
            throw Failure("Encrypted ODF entry uses a PBKDF2 iteration count outside the supported security policy.",
                OdfEncryptionFailureReason.UnsupportedProfile, entryPath, ex);
        }
        if (expectedSize > maxBytes) {
            throw Failure($"Decrypted ODF entry exceeds its configured uncompressed read budget ({maxBytes}).",
                OdfEncryptionFailureReason.ResourceLimitExceeded, entryPath);
        }
        if (salt.Length == 0 || salt.Length > 1024 || iv.Length != IvSize || expectedChecksum.Length != 32 ||
            ciphertext.Length == 0 || ciphertext.Length % IvSize != 0) {
            throw Failure("Encrypted ODF entry metadata is invalid.", OdfEncryptionFailureReason.InvalidEncryptedPackage, entryPath);
        }

        byte[] key = DeriveKey(startKey, salt, iterationCount, KeySize);
        byte[] padded;
        try {
            try {
                using Aes aes = Aes.Create();
                aes.KeySize = 256;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                aes.Key = key;
                aes.IV = iv;
                using ICryptoTransform transform = aes.CreateDecryptor();
                padded = transform.TransformFinalBlock(ciphertext, 0, ciphertext.Length);
            } catch (CryptographicException ex) {
                throw Failure("Encrypted ODF entry could not be decrypted.",
                    OdfEncryptionFailureReason.InvalidEncryptedPackage, entryPath, ex);
            }
        } finally {
            Array.Clear(key, 0, key.Length);
        }

        byte[] compressed = RemoveW3cPadding(padded, entryPath);
        using (var sha256 = SHA256.Create()) {
            int checksumLength = Math.Min(1024, compressed.Length);
            byte[] actualChecksum = sha256.ComputeHash(compressed, 0, checksumLength);
            if (!FixedTimeEquals(actualChecksum, expectedChecksum)) {
                throw Failure("The supplied ODF password is incorrect.", OdfEncryptionFailureReason.IncorrectPassword, entryPath);
            }
        }

        try {
            byte[] plaintext = OdfZipWriter.Inflate(compressed, maxBytes, entryPath);
            if (expectedSize >= 0 && plaintext.LongLength != expectedSize) {
                throw Failure("Decrypted ODF entry size does not match its manifest metadata.",
                    OdfEncryptionFailureReason.InvalidEncryptedPackage, entryPath);
            }
            return plaintext;
        } catch (OdfEncryptedPackageException) {
            throw;
        } catch (Exception ex) when (ex is InvalidDataException || ex is IOException) {
            throw Failure("Decrypted ODF entry contains invalid compressed data.",
                OdfEncryptionFailureReason.InvalidEncryptedPackage, entryPath, ex);
        }
    }

    private static byte[] DeriveKey(byte[] startKey, byte[] salt, int iterationCount, int keySize) {
        ValidateIterationCount(iterationCount);
#if NET8_0_OR_GREATER
        return Rfc2898DeriveBytes.Pbkdf2(startKey, salt, iterationCount, HashAlgorithmName.SHA1, keySize);
#else
        using var derivation = new Rfc2898DeriveBytes(startKey, salt, iterationCount);
        return derivation.GetBytes(keySize);
#endif
    }

    private static byte[] ApplyW3cPadding(byte[] value, int blockSize) {
        int padding = blockSize - value.Length % blockSize;
        byte[] output = new byte[value.Length + padding];
        Buffer.BlockCopy(value, 0, output, 0, value.Length);
        if (padding > 1) {
            byte[] random = RandomBytes(padding - 1);
            Buffer.BlockCopy(random, 0, output, value.Length, random.Length);
        }
        output[output.Length - 1] = (byte)padding;
        return output;
    }

    private static byte[] RemoveW3cPadding(byte[] value, string entryPath) {
        if (value.Length == 0) {
            throw Failure("Encrypted ODF entry has no padded payload.", OdfEncryptionFailureReason.InvalidEncryptedPackage, entryPath);
        }
        int padding = value[value.Length - 1];
        if (padding < 1 || padding > IvSize || padding > value.Length) {
            throw Failure("Encrypted ODF entry has invalid W3C block padding.",
                OdfEncryptionFailureReason.IncorrectPassword, entryPath);
        }
        byte[] output = new byte[value.Length - padding];
        Buffer.BlockCopy(value, 0, output, 0, output.Length);
        return output;
    }

    private static byte[] RandomBytes(int length) {
        var bytes = new byte[length];
        using RandomNumberGenerator random = RandomNumberGenerator.Create();
        random.GetBytes(bytes);
        return bytes;
    }

    private static bool FixedTimeEquals(byte[] left, byte[] right) {
        if (left.Length != right.Length) return false;
        int difference = 0;
        for (int i = 0; i < left.Length; i++) difference |= left[i] ^ right[i];
        return difference == 0;
    }

    private static OdfEncryptedPackageException Failure(string message, OdfEncryptionFailureReason reason,
        string entryPath, Exception? innerException = null) =>
        new OdfEncryptedPackageException(message, reason, entryPath, innerException);
}

internal sealed class OdfEncryptedEntry {
    internal OdfEncryptedEntry(byte[] ciphertext, byte[] salt, byte[] initializationVector, byte[] checksum,
        long originalSize, int iterationCount) {
        Ciphertext = ciphertext;
        Salt = salt;
        InitializationVector = initializationVector;
        Checksum = checksum;
        OriginalSize = originalSize;
        IterationCount = iterationCount;
    }

    internal byte[] Ciphertext { get; }
    internal byte[] Salt { get; }
    internal byte[] InitializationVector { get; }
    internal byte[] Checksum { get; }
    internal long OriginalSize { get; }
    internal int IterationCount { get; }
}
