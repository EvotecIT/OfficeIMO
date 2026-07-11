using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

#pragma warning disable CA1850 // Static HashData is unavailable on netstandard2.0 and net472.
internal sealed partial class PdfStandardSecurityHandler {
    private static PdfStandardSecurityHandler CreateModern(
        PdfDictionary encryptionDictionary,
        string? password,
        bool passwordWasSupplied,
        int revision) {
        byte[] ownerEntry = GetRequiredBytes(encryptionDictionary, "O");
        byte[] userEntry = GetRequiredBytes(encryptionDictionary, "U");
        byte[] ownerEncryptedFileKey = GetRequiredBytes(encryptionDictionary, "OE");
        byte[] userEncryptedFileKey = GetRequiredBytes(encryptionDictionary, "UE");
        byte[] encryptedPermissions = GetRequiredBytes(encryptionDictionary, "Perms");
        int permissions = GetRequiredPermissions(encryptionDictionary);
        bool encryptMetadata = encryptionDictionary.Get<PdfBoolean>("EncryptMetadata")?.Value ?? true;
        ValidateModernEntries(ownerEntry, userEntry, ownerEncryptedFileKey, userEncryptedFileKey, encryptedPermissions);
        byte[] passwordBytes = NormalizeModernPassword(passwordWasSupplied ? password ?? string.Empty : string.Empty);

        byte[]? fileKey = TryAuthenticateModernUser(passwordBytes, userEntry, userEncryptedFileKey, revision);
        if (fileKey is null) {
            fileKey = TryAuthenticateModernOwner(passwordBytes, ownerEntry, ownerEncryptedFileKey, userEntry, revision);
        }

        if (fileKey is null) {
            if (!passwordWasSupplied) {
                throw new PdfPasswordRequiredException("Encrypted PDF requires a password.");
            }

            throw new PdfInvalidPasswordException("The supplied PDF password is invalid.");
        }

        ValidateModernPermissions(fileKey, encryptedPermissions, permissions, encryptMetadata);
        PdfCryptMethod streamMethod = ResolveCryptMethod(encryptionDictionary, "StmF", version: 5);
        PdfCryptMethod stringMethod = ResolveCryptMethod(encryptionDictionary, "StrF", version: 5);
        return new PdfStandardSecurityHandler(fileKey, revision, 32, streamMethod, stringMethod, encryptMetadata);
    }

    private static byte[]? TryAuthenticateModernUser(byte[] password, byte[] userEntry, byte[] encryptedFileKey, int revision) {
        byte[] validationHash = ComputeModernHash(password, SliceModern(userEntry, 32, 40), Array.Empty<byte>(), revision);
        if (!StartsWith(userEntry, validationHash, 32)) {
            return null;
        }

        byte[] key = ComputeModernHash(password, SliceModern(userEntry, 40, 48), Array.Empty<byte>(), revision);
        return DecryptAes256NoPadding(key, encryptedFileKey);
    }

    private static byte[]? TryAuthenticateModernOwner(
        byte[] password,
        byte[] ownerEntry,
        byte[] encryptedFileKey,
        byte[] userEntry,
        int revision) {
        byte[] validationHash = ComputeModernHash(password, SliceModern(ownerEntry, 32, 40), userEntry, revision);
        if (!StartsWith(ownerEntry, validationHash, 32)) {
            return null;
        }

        byte[] key = ComputeModernHash(password, SliceModern(ownerEntry, 40, 48), userEntry, revision);
        return DecryptAes256NoPadding(key, encryptedFileKey);
    }

    private static byte[] ComputeModernHash(byte[] password, byte[] salt, byte[] userEntry, int revision) =>
        revision == 5
            ? Sha256Modern(PdfObjectBytes.Concat(password, salt, userEntry))
            : ComputeRevision6HashModern(password, salt, userEntry);

    private static byte[] ComputeRevision6HashModern(byte[] password, byte[] salt, byte[] userEntry) {
        byte[] key = Sha256Modern(PdfObjectBytes.Concat(password, salt, userEntry));
        int round = 0;
        byte lastByte;
        do {
            byte[] block = PdfObjectBytes.Concat(password, key, userEntry);
            var repeated = new byte[block.Length * 64];
            for (int i = 0; i < 64; i++) {
                Buffer.BlockCopy(block, 0, repeated, i * block.Length, block.Length);
            }

            byte[] aesKey = SliceModern(key, 0, 16);
            byte[] iv = SliceModern(key, 16, 32);
            byte[] encrypted = TransformAesCbcNoPadding(aesKey, iv, repeated, encrypt: true);
            int selector = 0;
            for (int i = 0; i < 16; i++) {
                selector = ((selector << 8) + encrypted[i]) % 3;
            }

            key = selector == 0 ? Sha256Modern(encrypted) : selector == 1 ? Sha384Modern(encrypted) : Sha512Modern(encrypted);
            lastByte = encrypted[encrypted.Length - 1];
            round++;
        } while (round < 64 || lastByte > round - 32);

        return SliceModern(key, 0, 32);
    }

    private static byte[] NormalizeModernPassword(string password) {
        byte[] bytes = Encoding.UTF8.GetBytes((password ?? string.Empty).Normalize(NormalizationForm.FormKC));
        return bytes.Length <= 127 ? bytes : SliceModern(bytes, 0, 127);
    }

    private static byte[] DecryptAes256NoPadding(byte[] key, byte[] data) =>
        TransformAesCbcNoPadding(key, new byte[16], data, encrypt: false);

    private static byte[] TransformAesCbcNoPadding(byte[] key, byte[] iv, byte[] data, bool encrypt) {
        using Aes aes = Aes.Create();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;
        aes.Key = key;
        aes.IV = iv;
        using ICryptoTransform transform = encrypt ? aes.CreateEncryptor() : aes.CreateDecryptor();
        return transform.TransformFinalBlock(data, 0, data.Length);
    }

    private static void ValidateModernPermissions(byte[] fileKey, byte[] encrypted, int permissions, bool encryptMetadata) {
        using Aes aes = Aes.Create();
        aes.Mode = CipherMode.ECB;
        aes.Padding = PaddingMode.None;
        aes.Key = fileKey;
        using ICryptoTransform decryptor = aes.CreateDecryptor();
        byte[] plain = decryptor.TransformFinalBlock(encrypted, 0, encrypted.Length);
        int storedPermissions = plain[0] | (plain[1] << 8) | (plain[2] << 16) | (plain[3] << 24);
        bool markerValid = plain[9] == (byte)'a' && plain[10] == (byte)'d' && plain[11] == (byte)'b';
        bool metadataValid = plain[8] == (encryptMetadata ? (byte)'T' : (byte)'F');
        if (storedPermissions != permissions || !markerValid || !metadataValid) {
            throw new PdfUnsupportedEncryptionException("AES-256 Standard security permissions validation failed.");
        }
    }

    private static void ValidateModernEntries(byte[] owner, byte[] user, byte[] ownerKey, byte[] userKey, byte[] permissions) {
        if (owner.Length != 48 || user.Length != 48 || ownerKey.Length != 32 || userKey.Length != 32 || permissions.Length != 16) {
            throw new PdfUnsupportedEncryptionException("AES-256 Standard security dictionary entries have invalid lengths.");
        }
    }

    private static byte[] SliceModern(byte[] value, int start, int end) {
        var result = new byte[end - start];
        Buffer.BlockCopy(value, start, result, 0, result.Length);
        return result;
    }

    private static byte[] Sha256Modern(byte[] data) {
        using SHA256 sha = SHA256.Create();
        return sha.ComputeHash(data);
    }

    private static byte[] Sha384Modern(byte[] data) {
        using SHA384 sha = SHA384.Create();
        return sha.ComputeHash(data);
    }

    private static byte[] Sha512Modern(byte[] data) {
        using SHA512 sha = SHA512.Create();
        return sha.ComputeHash(data);
    }
}
#pragma warning restore CA1850
