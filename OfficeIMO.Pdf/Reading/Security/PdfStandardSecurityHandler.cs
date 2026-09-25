using System.Threading;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Security;

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
    private readonly IOfficeAesCryptographyProvider? _aesCryptographyProvider;

    private PdfStandardSecurityHandler(
        byte[] fileKey,
        int revision,
        int keyLengthBytes,
        PdfCryptMethod streamMethod,
        PdfCryptMethod stringMethod,
        bool encryptMetadata,
        PdfPasswordAuthenticationRole authenticationRole,
        IOfficeAesCryptographyProvider? aesCryptographyProvider) {
        _fileKey = fileKey;
        _revision = revision;
        _keyLengthBytes = keyLengthBytes;
        _streamMethod = streamMethod;
        _stringMethod = stringMethod;
        _encryptMetadata = encryptMetadata;
        _authenticationRole = authenticationRole;
        _aesCryptographyProvider = aesCryptographyProvider;
    }

    internal PdfPasswordAuthenticationRole AuthenticationRole => _authenticationRole;

    public static PdfStandardSecurityHandler Create(
        PdfDictionary encryptionDictionary,
        byte[] fileId,
        string? password,
        bool passwordWasSupplied,
        IOfficeAesCryptographyProvider? aesCryptographyProvider,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        string filter = encryptionDictionary.Get<PdfName>("Filter")?.Name ?? string.Empty;
        if (!string.Equals(filter, "Standard", StringComparison.Ordinal)) {
            throw new PdfUnsupportedEncryptionException("Only PDF Standard password encryption is supported.");
        }

        int version = GetRequiredInt(encryptionDictionary, "V");
        int revision = GetRequiredInt(encryptionDictionary, "R");
        if ((revision == 5 || revision == 6) && version == 5) {
            return CreateModern(encryptionDictionary, password, passwordWasSupplied, revision, aesCryptographyProvider, cancellationToken);
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
        if (TryAuthenticateOwnerPassword(passwordCandidate, revision, keyLengthBytes, ownerEntry, userEntry, permissions, fileId, encryptMetadata, out byte[] fileKey, cancellationToken)) {
            return new PdfStandardSecurityHandler(
                fileKey,
                revision,
                keyLengthBytes,
                streamMethod,
                stringMethod,
                encryptMetadata,
                PdfPasswordAuthenticationRole.Owner,
                aesCryptographyProvider);
        }

        if (TryAuthenticateUserPassword(passwordCandidate, revision, keyLengthBytes, ownerEntry, userEntry, permissions, fileId, encryptMetadata, out fileKey, cancellationToken)) {
            return new PdfStandardSecurityHandler(
                fileKey,
                revision,
                keyLengthBytes,
                streamMethod,
                stringMethod,
                encryptMetadata,
                PdfPasswordAuthenticationRole.User,
                aesCryptographyProvider);
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (!passwordWasSupplied) {
            throw new PdfPasswordRequiredException("Encrypted PDF requires a password.");
        }

        throw new PdfInvalidPasswordException("The supplied PDF password is invalid.");
    }

    public PdfObject DecryptObject(int objectNumber, int generation, PdfObject value, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (value is PdfStringObj text) {
            return DecryptString(objectNumber, generation, text, cancellationToken);
        }

        if (value is PdfArray array) {
            var decrypted = new PdfArray { HasIncompleteSyntax = array.HasIncompleteSyntax };
            for (int i = 0; i < array.Items.Count; i++) {
                decrypted.Items.Add(DecryptObject(objectNumber, generation, array.Items[i], cancellationToken));
            }

            return decrypted;
        }

        if (value is PdfDictionary dictionary) {
            return DecryptDictionary(objectNumber, generation, dictionary, cancellationToken);
        }

        if (value is PdfStream stream) {
            PdfDictionary streamDictionary = (PdfDictionary)DecryptDictionary(objectNumber, generation, stream.Dictionary, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] sourceData = stream.GetData(cancellationToken);
            byte[] data = ShouldSkipStreamData(streamDictionary)
                ? sourceData
                : DecryptData(objectNumber, generation, sourceData, _streamMethod, cancellationToken);
            return new PdfStream(streamDictionary, data, stream.DecodingFailed, stream.DecodingError) {
                HasIncompleteSyntax = stream.HasIncompleteSyntax
            };
        }

        return value;
    }

    private PdfStringObj DecryptString(int objectNumber, int generation, PdfStringObj text, CancellationToken cancellationToken) {
        if (_stringMethod == PdfCryptMethod.Identity) return text;
        byte[] decrypted = DecryptData(objectNumber, generation, text.RawBytes, _stringMethod, cancellationToken);
        PdfStringObj result = PdfStringObj.FromParsedBytes(
            decrypted,
            PdfTextString.Decode(decrypted, cancellationToken),
            text.UseTextStringEncoding,
            text.EncodedTokenLength);
        result.HasIncompleteSyntax = text.HasIncompleteSyntax;
        return result;
    }

    private PdfDictionary DecryptDictionary(int objectNumber, int generation, PdfDictionary dictionary, CancellationToken cancellationToken) {
        var decrypted = new PdfDictionary { HasIncompleteSyntax = dictionary.HasIncompleteSyntax };
        foreach (var item in dictionary.Items) {
            decrypted.Items[item.Key] = DecryptObject(objectNumber, generation, item.Value, cancellationToken);
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

    private byte[] DecryptData(int objectNumber, int generation, byte[] data, PdfCryptMethod method, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (method == PdfCryptMethod.Identity || data.Length == 0) {
            return data;
        }

        if (method == PdfCryptMethod.AesV3) {
            return DecryptAesV2(_fileKey, data, cancellationToken);
        }

        byte[] objectKey = ComputeObjectKey(objectNumber, generation, method == PdfCryptMethod.AesV2);
        if (method == PdfCryptMethod.Rc4) {
            return Rc4.Transform(objectKey, data, cancellationToken);
        }

        if (method == PdfCryptMethod.AesV2) {
            return DecryptAesV2(objectKey, data, cancellationToken);
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
        out byte[] fileKey,
        CancellationToken cancellationToken) {
        fileKey = ComputeFileKey(password, revision, keyLengthBytes, ownerEntry, permissions, fileId, encryptMetadata, cancellationToken);
        byte[] expected = ComputeUserEntry(revision, fileKey, fileId, cancellationToken);
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
        out byte[] fileKey,
        CancellationToken cancellationToken) {
        fileKey = ComputeFileKeyFromPasswordBytes(passwordBytes, revision, keyLengthBytes, ownerEntry, permissions, fileId, encryptMetadata, cancellationToken);
        byte[] expected = ComputeUserEntry(revision, fileKey, fileId, cancellationToken);
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
        out byte[] fileKey,
        CancellationToken cancellationToken) {
        fileKey = Array.Empty<byte>();
        byte[] ownerKey = ComputeOwnerPasswordKey(password, revision, keyLengthBytes, cancellationToken);
        byte[] userPasswordBytes = revision == 2
            ? Rc4.Transform(ownerKey, ownerEntry, cancellationToken)
            : DecryptOwnerEntryRevision3Or4(ownerKey, ownerEntry, cancellationToken);
        return TryAuthenticateUserPasswordBytes(TrimPadding(userPasswordBytes, cancellationToken), revision, keyLengthBytes, ownerEntry, userEntry, permissions, fileId, encryptMetadata, out fileKey, cancellationToken);
    }

    private static byte[] ComputeFileKey(string password, int revision, int keyLengthBytes, byte[] ownerEntry, int permissions, byte[] fileId, bool encryptMetadata, CancellationToken cancellationToken) {
        return ComputeFileKeyFromPasswordBytes(EncodePassword(password, cancellationToken), revision, keyLengthBytes, ownerEntry, permissions, fileId, encryptMetadata, cancellationToken);
    }

    private static byte[] ComputeFileKeyFromPasswordBytes(byte[] passwordBytes, int revision, int keyLengthBytes, byte[] ownerEntry, int permissions, byte[] fileId, bool encryptMetadata, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] padded = PadPasswordBytes(passwordBytes);
        byte[] permissionBytes = {
            (byte)permissions, (byte)(permissions >> 8), (byte)(permissions >> 16), (byte)(permissions >> 24)
        };
        byte[] metadataBytes = revision >= 4 && !encryptMetadata
            ? new byte[] { 0xFF, 0xFF, 0xFF, 0xFF }
            : Array.Empty<byte>();
        byte[] digest = Md5Parts(cancellationToken, padded, ownerEntry, permissionBytes, fileId, metadataBytes);
        if (revision >= 3) {
            byte[] current = Take(digest, keyLengthBytes, cancellationToken);
            for (int i = 0; i < 50; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                current = Md5(Take(current, keyLengthBytes, cancellationToken));
            }

            digest = current;
        }

        return Take(digest, keyLengthBytes, cancellationToken);
    }

    private static byte[] ComputeUserEntry(int revision, byte[] fileKey, byte[] fileId, CancellationToken cancellationToken) {
        if (revision == 2) {
            return Rc4.Transform(fileKey, PasswordPadding, cancellationToken);
        }

        byte[] value = Take(Md5Parts(cancellationToken, PasswordPadding, fileId), 16, cancellationToken);
        value = Rc4.Transform(fileKey, value, cancellationToken);
        for (int i = 1; i <= 19; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            value = Rc4.Transform(XorKey(fileKey, i), value, cancellationToken);
        }

        var result = new byte[32];
        Buffer.BlockCopy(value, 0, result, 0, Math.Min(16, value.Length));
        return result;
    }

    private static byte[] ComputeOwnerPasswordKey(string password, int revision, int keyLengthBytes, CancellationToken cancellationToken) {
        byte[] digest = Md5(PadPassword(password, cancellationToken));
        if (revision >= 3) {
            for (int i = 0; i < 50; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                digest = Md5(Take(digest, keyLengthBytes, cancellationToken));
            }
        }

        return Take(digest, keyLengthBytes, cancellationToken);
    }

    private static byte[] DecryptOwnerEntryRevision3Or4(byte[] ownerKey, byte[] ownerEntry, CancellationToken cancellationToken) {
        byte[] current = Take(ownerEntry, ownerEntry.Length, cancellationToken);
        for (int i = 19; i >= 0; i--) {
            cancellationToken.ThrowIfCancellationRequested();
            current = Rc4.Transform(XorKey(ownerKey, i), current, cancellationToken);
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

    private byte[] DecryptAesV2(byte[] key, byte[] data, CancellationToken cancellationToken) {
        if (data.Length < 16 || (data.Length % 16) != 0) {
            throw new PdfUnsupportedEncryptionException("Invalid AESV2 encrypted stream length.");
        }

        byte[] iv = new byte[16];
        Buffer.BlockCopy(data, 0, iv, 0, iv.Length);
        var ciphertext = new byte[data.Length - 16];
        for (int offset = 0; offset < ciphertext.Length; offset += 65536) {
            cancellationToken.ThrowIfCancellationRequested();
            Buffer.BlockCopy(data, 16 + offset, ciphertext, offset, Math.Min(65536, ciphertext.Length - offset));
        }
        byte[] decrypted = PdfAesCryptography.DecryptNoPadding(key, iv, ciphertext, _aesCryptographyProvider, cancellationToken);
        return RemovePkcs7Padding(decrypted, cancellationToken);
    }

    private static byte[] RemovePkcs7Padding(byte[] data, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
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

        return Take(data, data.Length - count, cancellationToken);
    }

    private static byte[] PadPassword(string password, CancellationToken cancellationToken) {
        return PadPasswordBytes(EncodePassword(password, cancellationToken));
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

    private static byte[] EncodePassword(string password, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrEmpty(password)) {
            return Array.Empty<byte>();
        }

        return PdfLegacyPasswordEncoding.EncodePrefix(password, cancellationToken);
    }

    private static byte[] TrimPadding(byte[] value, CancellationToken cancellationToken) {
        int length = value.Length;
        for (int i = 0; i <= value.Length - PasswordPadding.Length; i++) {
            if ((i & 65535) == 0) cancellationToken.ThrowIfCancellationRequested();
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

        return Take(value, length, cancellationToken);
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

    private static byte[] Take(byte[] value, int count, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var result = new byte[count];
        int length = Math.Min(value.Length, count);
        for (int offset = 0; offset < length; offset += 65536) {
            cancellationToken.ThrowIfCancellationRequested();
            Buffer.BlockCopy(value, offset, result, offset, Math.Min(65536, length - offset));
        }
        cancellationToken.ThrowIfCancellationRequested();
        return result;
    }

    private static byte[] Md5(byte[] data) {
#pragma warning disable CA5351, CA1850 // Reading PDF standard-security revisions 2-4 requires their MD5 key-derivation algorithm; use Create for legacy targets.
        using MD5 md5 = MD5.Create();
        return md5.ComputeHash(data);
#pragma warning restore CA5351, CA1850
    }

    private static byte[] Md5Parts(CancellationToken cancellationToken, params byte[][] parts) {
#pragma warning disable CA5351, CA1850 // PDF Standard revisions 2-4 require MD5; stream large dictionary values to observe cancellation.
        using MD5 md5 = MD5.Create();
        foreach (byte[] part in parts) {
            for (int offset = 0; offset < part.Length;) {
                cancellationToken.ThrowIfCancellationRequested();
                int count = Math.Min(65536, part.Length - offset);
                md5.TransformBlock(part, offset, count, part, offset);
                offset += count;
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        md5.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        cancellationToken.ThrowIfCancellationRequested();
        return md5.Hash!;
#pragma warning restore CA5351, CA1850
    }

    private enum PdfCryptMethod {
        Identity,
        Rc4,
        AesV2,
        AesV3
    }

    private static class Rc4 {
        public static byte[] Transform(byte[] key, byte[] data, CancellationToken cancellationToken = default) {
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
                if ((i & 0xFFFF) == 0) cancellationToken.ThrowIfCancellationRequested();
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
