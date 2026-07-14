using System.Globalization;
using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

#pragma warning disable CA1850 // Static HashData is unavailable on netstandard2.0 and net472.
internal static partial class PdfStandardSecurityWriter {
    private static PdfEncryptionAssembly EncryptAes256(IReadOnlyList<byte[]> sourceObjects, PdfStandardEncryptionOptions options, long objectMemoryLimitBytes) {
        byte[] fileId = CreateFileId();
        byte[] fileKey = RandomBytes(32);
        byte[] userPassword = NormalizeModernPassword(options.UserPassword);
        byte[] ownerPassword = NormalizeModernPassword(options.OwnerPassword ?? options.UserPassword);
        byte[] userValidationSalt = RandomBytes(8);
        byte[] userKeySalt = RandomBytes(8);
        byte[] userHash = ComputeRevision6Hash(userPassword, userValidationSalt, Array.Empty<byte>());
        byte[] userEntry = PdfObjectBytes.Concat(userHash, userValidationSalt, userKeySalt);
        byte[] userEncryptionKey = ComputeRevision6Hash(userPassword, userKeySalt, Array.Empty<byte>());
        byte[] userEncryptedFileKey = EncryptAes256NoPadding(userEncryptionKey, fileKey);

        byte[] ownerValidationSalt = RandomBytes(8);
        byte[] ownerKeySalt = RandomBytes(8);
        byte[] ownerHash = ComputeRevision6Hash(ownerPassword, ownerValidationSalt, userEntry);
        byte[] ownerEntry = PdfObjectBytes.Concat(ownerHash, ownerValidationSalt, ownerKeySalt);
        byte[] ownerEncryptionKey = ComputeRevision6Hash(ownerPassword, ownerKeySalt, userEntry);
        byte[] ownerEncryptedFileKey = EncryptAes256NoPadding(ownerEncryptionKey, fileKey);
        byte[] encryptedPermissions = EncryptPermissions(fileKey, options.Permissions, options.EncryptMetadata);

        int encryptionObjectNumber = sourceObjects.Count + 1;
        var objects = new PdfObjectStore(objectMemoryLimitBytes);
        try {
            for (int i = 0; i < sourceObjects.Count; i++) {
                objects.Add(EncryptAesIndirectObject(sourceObjects[i], i + 1, fileKey, options.EncryptMetadata, deriveObjectKey: false));
            }

            objects.Add(PdfObjectBytes.WrapIndirectObject(
                encryptionObjectNumber,
                BuildAes256EncryptionDictionary(
                    ownerEntry,
                    userEntry,
                    ownerEncryptedFileKey,
                    userEncryptedFileKey,
                    encryptedPermissions,
                    options.Permissions,
                    options.EncryptMetadata)));
            return new PdfEncryptionAssembly(objects, encryptionObjectNumber, fileId, objects);
        } catch {
            objects.Dispose();
            throw;
        }
    }

    private static byte[] ComputeRevision6Hash(byte[] password, byte[] salt, byte[] userEntry) {
        byte[] key = Sha256(PdfObjectBytes.Concat(password, salt, userEntry));
        int round = 0;
        byte lastByte;
        do {
            byte[] block = PdfObjectBytes.Concat(password, key, userEntry);
            var repeated = new byte[block.Length * 64];
            for (int i = 0; i < 64; i++) {
                Buffer.BlockCopy(block, 0, repeated, i * block.Length, block.Length);
            }

            byte[] aesKey = Take(key, 16);
            var iv = new byte[16];
            Buffer.BlockCopy(key, 16, iv, 0, 16);
            byte[] encrypted = EncryptAesNoPadding(aesKey, iv, repeated);
            int selector = 0;
            for (int i = 0; i < 16; i++) {
                selector = ((selector << 8) + encrypted[i]) % 3;
            }

            key = selector == 0 ? Sha256(encrypted) : selector == 1 ? Sha384(encrypted) : Sha512(encrypted);
            lastByte = encrypted[encrypted.Length - 1];
            round++;
        } while (round < 64 || lastByte > round - 32);

        return Take(key, 32);
    }

    private static byte[] NormalizeModernPassword(string password) {
        string normalized = (password ?? string.Empty).Normalize(NormalizationForm.FormKC);
        byte[] bytes = Encoding.UTF8.GetBytes(normalized);
        return bytes.Length <= 127 ? bytes : Take(bytes, 127);
    }

    private static byte[] EncryptAes256NoPadding(byte[] key, byte[] data) =>
        EncryptAesNoPadding(key, new byte[16], data);

    private static byte[] EncryptAesNoPadding(byte[] key, byte[] iv, byte[] data) {
        using Aes aes = Aes.Create();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;
        aes.Key = key;
        aes.IV = iv;
        using ICryptoTransform encryptor = aes.CreateEncryptor();
        return encryptor.TransformFinalBlock(data, 0, data.Length);
    }

    private static byte[] EncryptPermissions(byte[] fileKey, int permissions, bool encryptMetadata) {
        byte[] plain = RandomBytes(16);
        unchecked {
            plain[0] = (byte)(permissions & 0xFF);
            plain[1] = (byte)((permissions >> 8) & 0xFF);
            plain[2] = (byte)((permissions >> 16) & 0xFF);
            plain[3] = (byte)((permissions >> 24) & 0xFF);
        }

        plain[4] = 0xFF;
        plain[5] = 0xFF;
        plain[6] = 0xFF;
        plain[7] = 0xFF;
        plain[8] = encryptMetadata ? (byte)'T' : (byte)'F';
        plain[9] = (byte)'a';
        plain[10] = (byte)'d';
        plain[11] = (byte)'b';

        // The PDF permissions entry is exactly one AES block, so CBC with a zero IV is equivalent to ECB here.
        return EncryptAesNoPadding(fileKey, new byte[16], plain);
    }

    private static byte[] RandomBytes(int length) {
        var bytes = new byte[length];
        using RandomNumberGenerator rng = RandomNumberGenerator.Create();
        rng.GetBytes(bytes);
        return bytes;
    }

    private static byte[] Sha256(byte[] data) {
        using SHA256 sha = SHA256.Create();
        return sha.ComputeHash(data);
    }

    private static byte[] Sha384(byte[] data) {
        using SHA384 sha = SHA384.Create();
        return sha.ComputeHash(data);
    }

    private static byte[] Sha512(byte[] data) {
        using SHA512 sha = SHA512.Create();
        return sha.ComputeHash(data);
    }

    private static string BuildAes256EncryptionDictionary(
        byte[] ownerEntry,
        byte[] userEntry,
        byte[] ownerEncryptedFileKey,
        byte[] userEncryptedFileKey,
        byte[] encryptedPermissions,
        int permissions,
        bool encryptMetadata) =>
        "<< /Filter /Standard /V 5 /R 6 /Length 256" +
        " /O " + PdfSyntaxEscaper.HexString(ownerEntry) +
        " /U " + PdfSyntaxEscaper.HexString(userEntry) +
        " /OE " + PdfSyntaxEscaper.HexString(ownerEncryptedFileKey) +
        " /UE " + PdfSyntaxEscaper.HexString(userEncryptedFileKey) +
        " /Perms " + PdfSyntaxEscaper.HexString(encryptedPermissions) +
        " /P " + permissions.ToString(CultureInfo.InvariantCulture) +
        " /EncryptMetadata " + (encryptMetadata ? "true" : "false") +
        " /CF << /StdCF << /AuthEvent /DocOpen /CFM /AESV3 /Length 32 >> >>" +
        " /StmF /StdCF /StrF /StdCF >>";
}
#pragma warning restore CA1850
