using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfStandardSecurityHandler {
    private static readonly byte[] PasswordPadding = new byte[] {
        0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41,
        0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08,
        0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80,
        0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A
    };

    private readonly byte[] _fileKey;
    private readonly int _revision;
    private readonly int _keyLengthBytes;
    private readonly PdfCryptMethod _streamMethod;
    private readonly PdfCryptMethod _stringMethod;
    private readonly bool _encryptMetadata;
    private readonly PdfPasswordAuthenticationRole _authenticationRole;

    private PdfStandardSecurityHandler(
        byte[] fileKey,
        int revision,
        int keyLengthBytes,
        PdfCryptMethod streamMethod,
        PdfCryptMethod stringMethod,
        bool encryptMetadata,
        PdfPasswordAuthenticationRole authenticationRole) {
        _fileKey = fileKey;
        _revision = revision;
        _keyLengthBytes = keyLengthBytes;
        _streamMethod = streamMethod;
        _stringMethod = stringMethod;
        _encryptMetadata = encryptMetadata;
        _authenticationRole = authenticationRole;
    }

    internal PdfPasswordAuthenticationRole AuthenticationRole => _authenticationRole;

    public static PdfStandardSecurityHandler Create(PdfDictionary encryptionDictionary, byte[] fileId, string? password, bool passwordWasSupplied) {
        string filter = encryptionDictionary.Get<PdfName>("Filter")?.Name ?? string.Empty;
        if (!string.Equals(filter, "Standard", StringComparison.Ordinal)) {
            throw new PdfUnsupportedEncryptionException("Only PDF Standard password encryption is supported.");
        }

        int version = GetRequiredInt(encryptionDictionary, "V");
        int revision = GetRequiredInt(encryptionDictionary, "R");
        if ((revision == 5 || revision == 6) && version == 5) {
            return CreateModern(encryptionDictionary, password, passwordWasSupplied, revision);
        }

        if (revision < 2 || revision > 4 || version < 1 || version > 4) {
            throw new PdfUnsupportedEncryptionException("Only PDF Standard security handler revisions 2 through 6 are supported.");
        }

        byte[] ownerEntry = GetRequiredBytes(encryptionDictionary, "O");
        byte[] userEntry = GetRequiredBytes(encryptionDictionary, "U");
        int permissions = GetRequiredPermissions(encryptionDictionary);
        int keyLengthBits = revision == 2 ? 40 : (GetOptionalInt(encryptionDictionary, "Length") ?? 128);
        int keyLengthBytes = Math.Max(5, Math.Min(16, keyLengthBits / 8));
        bool encryptMetadata = encryptionDictionary.Get<PdfBoolean>("EncryptMetadata")?.Value ?? true;
        PdfCryptMethod streamMethod = ResolveCryptMethod(encryptionDictionary, "StmF", version);
        PdfCryptMethod stringMethod = ResolveCryptMethod(encryptionDictionary, "StrF", version);

        string passwordCandidate = passwordWasSupplied ? password ?? string.Empty : string.Empty;
        if (TryAuthenticateOwnerPassword(passwordCandidate, revision, keyLengthBytes, ownerEntry, userEntry, permissions, fileId, encryptMetadata, out byte[] fileKey)) {
            return new PdfStandardSecurityHandler(fileKey, revision, keyLengthBytes, streamMethod, stringMethod, encryptMetadata, PdfPasswordAuthenticationRole.Owner);
        }

        if (TryAuthenticateUserPassword(passwordCandidate, revision, keyLengthBytes, ownerEntry, userEntry, permissions, fileId, encryptMetadata, out fileKey)) {
            return new PdfStandardSecurityHandler(fileKey, revision, keyLengthBytes, streamMethod, stringMethod, encryptMetadata, PdfPasswordAuthenticationRole.User);
        }

        if (!passwordWasSupplied) {
            throw new PdfPasswordRequiredException("Encrypted PDF requires a password.");
        }

        throw new PdfInvalidPasswordException("The supplied PDF password is invalid.");
    }

    public PdfObject DecryptObject(int objectNumber, int generation, PdfObject value) {
        if (value is PdfStringObj text) {
            return DecryptString(objectNumber, generation, text);
        }

        if (value is PdfArray array) {
            var decrypted = new PdfArray();
            for (int i = 0; i < array.Items.Count; i++) {
                decrypted.Items.Add(DecryptObject(objectNumber, generation, array.Items[i]));
            }

            return decrypted;
        }

        if (value is PdfDictionary dictionary) {
            return DecryptDictionary(objectNumber, generation, dictionary);
        }

        if (value is PdfStream stream) {
            PdfDictionary streamDictionary = (PdfDictionary)DecryptDictionary(objectNumber, generation, stream.Dictionary);
            byte[] data = ShouldSkipStreamData(streamDictionary)
                ? stream.Data
                : DecryptData(objectNumber, generation, stream.Data, _streamMethod);
            return new PdfStream(streamDictionary, data, stream.DecodingFailed, stream.DecodingError);
        }

        return value;
    }

    private PdfStringObj DecryptString(int objectNumber, int generation, PdfStringObj text) {
        byte[] decrypted = DecryptData(objectNumber, generation, text.RawBytes, _stringMethod);
        return new PdfStringObj(
            decrypted,
            text.UseTextStringEncoding,
            text.EncodedTokenLength);
    }

    private PdfDictionary DecryptDictionary(int objectNumber, int generation, PdfDictionary dictionary) {
        var decrypted = new PdfDictionary();
        foreach (var item in dictionary.Items) {
            decrypted.Items[item.Key] = DecryptObject(objectNumber, generation, item.Value);
        }

        return decrypted;
    }

    private bool ShouldSkipStreamData(PdfDictionary dictionary) {
        if (dictionary.Get<PdfName>("Type")?.Name == "XRef") {
            return true;
        }

        if (!_encryptMetadata && dictionary.Get<PdfName>("Type")?.Name == "Metadata") {
            return true;
        }

        return _streamMethod == PdfCryptMethod.Identity;
    }

    private byte[] DecryptData(int objectNumber, int generation, byte[] data, PdfCryptMethod method) {
        if (method == PdfCryptMethod.Identity || data.Length == 0) {
            return data;
        }

        if (method == PdfCryptMethod.AesV3) {
            return DecryptAesV2(_fileKey, data);
        }

        byte[] objectKey = ComputeObjectKey(objectNumber, generation, method == PdfCryptMethod.AesV2);
        if (method == PdfCryptMethod.Rc4) {
            return Rc4.Transform(objectKey, data);
        }

        if (method == PdfCryptMethod.AesV2) {
            return DecryptAesV2(objectKey, data);
        }

        throw new PdfUnsupportedEncryptionException("Unsupported PDF crypt filter method.");
    }

    private byte[] ComputeObjectKey(int objectNumber, int generation, bool aes) {
        byte[] buffer = new byte[_fileKey.Length + 5 + (aes ? 4 : 0)];
        Buffer.BlockCopy(_fileKey, 0, buffer, 0, _fileKey.Length);
        int offset = _fileKey.Length;
        buffer[offset++] = (byte)(objectNumber & 0xFF);
        buffer[offset++] = (byte)((objectNumber >> 8) & 0xFF);
        buffer[offset++] = (byte)((objectNumber >> 16) & 0xFF);
        buffer[offset++] = (byte)(generation & 0xFF);
        buffer[offset++] = (byte)((generation >> 8) & 0xFF);
        if (aes) {
            buffer[offset++] = 0x73;
            buffer[offset++] = 0x41;
            buffer[offset++] = 0x6C;
            buffer[offset] = 0x54;
        }

        byte[] digest = Md5(buffer);
        int length = Math.Min(_keyLengthBytes + 5, 16);
        var key = new byte[length];
        Buffer.BlockCopy(digest, 0, key, 0, length);
        return key;
    }

    private static bool TryAuthenticateUserPassword(
        string password,
        int revision,
        int keyLengthBytes,
        byte[] ownerEntry,
        byte[] userEntry,
        int permissions,
        byte[] fileId,
        bool encryptMetadata,
        out byte[] fileKey) {
        fileKey = ComputeFileKey(password, revision, keyLengthBytes, ownerEntry, permissions, fileId, encryptMetadata);
        byte[] expected = ComputeUserEntry(revision, fileKey, fileId);
        return revision == 2
            ? StartsWith(userEntry, expected, 32)
            : StartsWith(userEntry, expected, 16);
    }

    private static bool TryAuthenticateUserPasswordBytes(
        byte[] passwordBytes,
        int revision,
        int keyLengthBytes,
        byte[] ownerEntry,
        byte[] userEntry,
        int permissions,
        byte[] fileId,
        bool encryptMetadata,
        out byte[] fileKey) {
        fileKey = ComputeFileKeyFromPasswordBytes(passwordBytes, revision, keyLengthBytes, ownerEntry, permissions, fileId, encryptMetadata);
        byte[] expected = ComputeUserEntry(revision, fileKey, fileId);
        return revision == 2
            ? StartsWith(userEntry, expected, 32)
            : StartsWith(userEntry, expected, 16);
    }

    private static bool TryAuthenticateOwnerPassword(
        string password,
        int revision,
        int keyLengthBytes,
        byte[] ownerEntry,
        byte[] userEntry,
        int permissions,
        byte[] fileId,
        bool encryptMetadata,
        out byte[] fileKey) {
        fileKey = Array.Empty<byte>();
        byte[] ownerKey = ComputeOwnerPasswordKey(password, revision, keyLengthBytes);
        byte[] userPasswordBytes = revision == 2
            ? Rc4.Transform(ownerKey, ownerEntry)
            : DecryptOwnerEntryRevision3Or4(ownerKey, ownerEntry);
        return TryAuthenticateUserPasswordBytes(TrimPadding(userPasswordBytes), revision, keyLengthBytes, ownerEntry, userEntry, permissions, fileId, encryptMetadata, out fileKey);
    }

    private static byte[] ComputeFileKey(string password, int revision, int keyLengthBytes, byte[] ownerEntry, int permissions, byte[] fileId, bool encryptMetadata) {
        return ComputeFileKeyFromPasswordBytes(EncodePassword(password), revision, keyLengthBytes, ownerEntry, permissions, fileId, encryptMetadata);
    }

    private static byte[] ComputeFileKeyFromPasswordBytes(byte[] passwordBytes, int revision, int keyLengthBytes, byte[] ownerEntry, int permissions, byte[] fileId, bool encryptMetadata) {
        byte[] padded = PadPasswordBytes(passwordBytes);
        var buffer = new List<byte>(padded.Length + ownerEntry.Length + 16 + fileId.Length + 4);
        buffer.AddRange(padded);
        buffer.AddRange(ownerEntry);
        AppendInt32LittleEndian(buffer, permissions);
        buffer.AddRange(fileId);
        if (revision >= 4 && !encryptMetadata) {
            buffer.AddRange(new byte[] { 0xFF, 0xFF, 0xFF, 0xFF });
        }

        byte[] digest = Md5(buffer.ToArray());
        if (revision >= 3) {
            byte[] current = Take(digest, keyLengthBytes);
            for (int i = 0; i < 50; i++) {
                current = Md5(Take(current, keyLengthBytes));
            }

            digest = current;
        }

        return Take(digest, keyLengthBytes);
    }

    private static byte[] ComputeUserEntry(int revision, byte[] fileKey, byte[] fileId) {
        if (revision == 2) {
            return Rc4.Transform(fileKey, PasswordPadding);
        }

        var buffer = new List<byte>(PasswordPadding.Length + fileId.Length);
        buffer.AddRange(PasswordPadding);
        buffer.AddRange(fileId);
        byte[] value = Take(Md5(buffer.ToArray()), 16);
        value = Rc4.Transform(fileKey, value);
        for (int i = 1; i <= 19; i++) {
            value = Rc4.Transform(XorKey(fileKey, i), value);
        }

        var result = new byte[32];
        Buffer.BlockCopy(value, 0, result, 0, Math.Min(16, value.Length));
        return result;
    }

    private static byte[] ComputeOwnerPasswordKey(string password, int revision, int keyLengthBytes) {
        byte[] digest = Md5(PadPassword(password));
        if (revision >= 3) {
            for (int i = 0; i < 50; i++) {
                digest = Md5(Take(digest, keyLengthBytes));
            }
        }

        return Take(digest, keyLengthBytes);
    }

    private static byte[] DecryptOwnerEntryRevision3Or4(byte[] ownerKey, byte[] ownerEntry) {
        byte[] current = Take(ownerEntry, ownerEntry.Length);
        for (int i = 19; i >= 0; i--) {
            current = Rc4.Transform(XorKey(ownerKey, i), current);
        }

        return current;
    }

    private static PdfCryptMethod ResolveCryptMethod(PdfDictionary encryptionDictionary, string key, int version) {
        if (version < 4) {
            return PdfCryptMethod.Rc4;
        }

        string filterName = encryptionDictionary.Get<PdfName>(key)?.Name ?? "Identity";
        if (string.Equals(filterName, "Identity", StringComparison.Ordinal)) {
            return PdfCryptMethod.Identity;
        }

        PdfDictionary? cryptFilters = encryptionDictionary.Get<PdfDictionary>("CF");
        PdfDictionary? filter = cryptFilters?.Get<PdfDictionary>(filterName);
        string cfm = filter?.Get<PdfName>("CFM")?.Name ?? "V2";
        switch (cfm) {
            case "None":
            case "Identity":
                return PdfCryptMethod.Identity;
            case "V2":
                return PdfCryptMethod.Rc4;
            case "AESV2":
                return PdfCryptMethod.AesV2;
            case "AESV3":
                return PdfCryptMethod.AesV3;
            default:
                throw new PdfUnsupportedEncryptionException("Unsupported PDF crypt filter method /" + cfm + ".");
        }
    }

    private static byte[] DecryptAesV2(byte[] key, byte[] data) {
        if (data.Length < 16 || (data.Length % 16) != 0) {
            throw new PdfUnsupportedEncryptionException("Invalid AESV2 encrypted stream length.");
        }

        byte[] iv = new byte[16];
        Buffer.BlockCopy(data, 0, iv, 0, iv.Length);
        using Aes aes = Aes.Create();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;
        aes.Key = key;
        aes.IV = iv;
        using ICryptoTransform decryptor = aes.CreateDecryptor();
        byte[] decrypted = decryptor.TransformFinalBlock(data, 16, data.Length - 16);
        return RemovePkcs7Padding(decrypted);
    }

    private static byte[] RemovePkcs7Padding(byte[] data) {
        if (data.Length == 0) {
            return data;
        }

        int count = data[data.Length - 1];
        if (count <= 0 || count > 16 || count > data.Length) {
            return data;
        }

        for (int i = data.Length - count; i < data.Length; i++) {
            if (data[i] != count) {
                return data;
            }
        }

        return Take(data, data.Length - count);
    }

    private static byte[] PadPassword(string password) {
        return PadPasswordBytes(EncodePassword(password));
    }

    private static byte[] PadPasswordBytes(byte[] passwordBytes) {
        var padded = new byte[32];
        int copy = Math.Min(passwordBytes.Length, 32);
        Buffer.BlockCopy(passwordBytes, 0, padded, 0, copy);
        if (copy < 32) {
            Buffer.BlockCopy(PasswordPadding, 0, padded, copy, 32 - copy);
        }

        return padded;
    }

    private static byte[] EncodePassword(string password) {
        if (string.IsNullOrEmpty(password)) {
            return Array.Empty<byte>();
        }

        return PdfWinAnsiEncoding.CanEncode(password, out _)
            ? PdfWinAnsiEncoding.Encode(password)
            : Encoding.UTF8.GetBytes(password);
    }

    private static byte[] TrimPadding(byte[] value) {
        int length = value.Length;
        for (int i = 0; i <= value.Length - PasswordPadding.Length; i++) {
            bool match = true;
            for (int j = 0; j < PasswordPadding.Length; j++) {
                if (value[i + j] != PasswordPadding[j]) {
                    match = false;
                    break;
                }
            }

            if (match) {
                length = i;
                break;
            }
        }

        return Take(value, length);
    }

    private static int GetRequiredInt(PdfDictionary dictionary, string key) {
        PdfNumber? number = dictionary.Get<PdfNumber>(key);
        if (number is null) {
            throw new PdfUnsupportedEncryptionException("PDF encryption dictionary is missing /" + key + ".");
        }

        return (int)number.Value;
    }

    private static int GetRequiredPermissions(PdfDictionary dictionary) {
        PdfNumber? number = dictionary.Get<PdfNumber>("P");
        if (number is null) {
            throw new PdfUnsupportedEncryptionException("PDF encryption dictionary is missing /P.");
        }

        double value = number.Value;
        if (value >= 0D && value <= uint.MaxValue && value > int.MaxValue) {
            return unchecked((int)(uint)value);
        }

        return (int)value;
    }

    private static int? GetOptionalInt(PdfDictionary dictionary, string key) {
        return dictionary.Get<PdfNumber>(key) is PdfNumber number ? (int)number.Value : null;
    }

    private static byte[] GetRequiredBytes(PdfDictionary dictionary, string key) {
        PdfStringObj? value = dictionary.Get<PdfStringObj>(key);
        if (value is null) {
            throw new PdfUnsupportedEncryptionException("PDF encryption dictionary is missing /" + key + ".");
        }

        return value.RawBytes;
    }

    private static bool StartsWith(byte[] actual, byte[] expected, int count) {
        if (actual.Length < count || expected.Length < count) {
            return false;
        }

        for (int i = 0; i < count; i++) {
            if (actual[i] != expected[i]) {
                return false;
            }
        }

        return true;
    }

    private static byte[] XorKey(byte[] key, int value) {
        var result = new byte[key.Length];
        for (int i = 0; i < key.Length; i++) {
            result[i] = (byte)(key[i] ^ value);
        }

        return result;
    }

    private static byte[] Take(byte[] value, int count) {
        var result = new byte[count];
        Buffer.BlockCopy(value, 0, result, 0, Math.Min(value.Length, count));
        return result;
    }

    private static void AppendInt32LittleEndian(List<byte> buffer, int value) {
        unchecked {
            buffer.Add((byte)(value & 0xFF));
            buffer.Add((byte)((value >> 8) & 0xFF));
            buffer.Add((byte)((value >> 16) & 0xFF));
            buffer.Add((byte)((value >> 24) & 0xFF));
        }
    }

    private static byte[] Md5(byte[] data) {
#pragma warning disable CA5351, CA1850
        using MD5 md5 = MD5.Create();
        return md5.ComputeHash(data);
#pragma warning restore CA5351, CA1850
    }

    private enum PdfCryptMethod {
        Identity,
        Rc4,
        AesV2,
        AesV3
    }

    private static class Rc4 {
        public static byte[] Transform(byte[] key, byte[] data) {
            var state = new byte[256];
            for (int i = 0; i < state.Length; i++) {
                state[i] = (byte)i;
            }

            int j = 0;
            for (int i = 0; i < 256; i++) {
                j = (j + state[i] + key[i % key.Length]) & 0xFF;
                Swap(state, i, j);
            }

            var result = new byte[data.Length];
            int x = 0;
            int y = 0;
            for (int i = 0; i < data.Length; i++) {
                x = (x + 1) & 0xFF;
                y = (y + state[x]) & 0xFF;
                Swap(state, x, y);
                byte k = state[(state[x] + state[y]) & 0xFF];
                result[i] = (byte)(data[i] ^ k);
            }

            return result;
        }

        private static void Swap(byte[] state, int left, int right) {
            byte temp = state[left];
            state[left] = state[right];
            state[right] = temp;
        }
    }
}
