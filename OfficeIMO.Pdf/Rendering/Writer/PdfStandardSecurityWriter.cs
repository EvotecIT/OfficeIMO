using System.Globalization;
using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

internal sealed class PdfEncryptionAssembly {
    public PdfEncryptionAssembly(IReadOnlyList<byte[]> objects, int encryptionObjectNumber, byte[] fileId) {
        Objects = objects;
        EncryptionObjectNumber = encryptionObjectNumber;
        FileId = fileId;
    }

    public IReadOnlyList<byte[]> Objects { get; }
    public int EncryptionObjectNumber { get; }
    public byte[] FileId { get; }
}

internal static class PdfStandardSecurityWriter {
    private static readonly byte[] PasswordPadding = new byte[] {
        0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41,
        0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08,
        0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80,
        0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A
    };

    private const int Revision = 3;
    private const int KeyLengthBytes = 16;

    internal static PdfEncryptionAssembly Encrypt(IReadOnlyList<byte[]> sourceObjects, PdfStandardEncryptionOptions options) {
        Guard.NotNull(sourceObjects, nameof(sourceObjects));
        Guard.NotNull(options, nameof(options));

        byte[] fileId = CreateFileId();
        string ownerPassword = options.OwnerPassword ?? options.UserPassword;
        byte[] ownerEntry = ComputeOwnerEntry(ownerPassword, options.UserPassword);
        byte[] fileKey = ComputeFileKey(options.UserPassword, ownerEntry, options.Permissions, fileId);
        byte[] userEntry = ComputeUserEntry(fileKey, fileId);

        int encryptionObjectNumber = sourceObjects.Count + 1;
        var objects = new List<byte[]>(sourceObjects.Count + 1);
        for (int i = 0; i < sourceObjects.Count; i++) {
            objects.Add(EncryptIndirectObject(sourceObjects[i], i + 1, fileKey));
        }

        objects.Add(PdfObjectBytes.WrapIndirectObject(encryptionObjectNumber, BuildEncryptionDictionary(ownerEntry, userEntry, options.Permissions)));
        return new PdfEncryptionAssembly(objects, encryptionObjectNumber, fileId);
    }

    private static byte[] EncryptIndirectObject(byte[] objectBytes, int objectNumber, byte[] fileKey) {
        int bodyStart = IndexOf(objectBytes, PdfEncoding.Latin1GetBytes("obj\n"), 0);
        int bodyEnd = LastIndexOf(objectBytes, PdfEncoding.Latin1GetBytes("endobj\n"));
        if (bodyStart < 0 || bodyEnd <= bodyStart) {
            return objectBytes;
        }

        bodyStart += 4;
        var body = new byte[bodyEnd - bodyStart];
        Buffer.BlockCopy(objectBytes, bodyStart, body, 0, body.Length);
        return PdfObjectBytes.WrapIndirectObject(objectNumber, EncryptObjectBody(body, objectNumber, fileKey));
    }

    private static byte[] EncryptObjectBody(byte[] body, int objectNumber, byte[] fileKey) {
        byte[] streamMarker = PdfEncoding.Latin1GetBytes("\nstream\n");
        byte[] endStreamMarker = PdfEncoding.Latin1GetBytes("\nendstream");
        int streamMarkerIndex = IndexOf(body, streamMarker, 0);
        if (streamMarkerIndex < 0) {
            return EncryptPdfStrings(body, objectNumber, fileKey);
        }

        int streamDataStart = streamMarkerIndex + streamMarker.Length;
        int streamDataEnd = LastIndexOf(body, endStreamMarker);
        if (streamDataEnd < streamDataStart) {
            return EncryptPdfStrings(body, objectNumber, fileKey);
        }

        byte[] prefix = Slice(body, 0, streamDataStart);
        byte[] streamData = Slice(body, streamDataStart, streamDataEnd);
        byte[] suffix = Slice(body, streamDataEnd, body.Length);
        byte[] encryptedPrefix = EncryptPdfStrings(prefix, objectNumber, fileKey);
        byte[] encryptedStream = Rc4.Transform(ComputeObjectKey(fileKey, objectNumber, 0), streamData);
        return PdfObjectBytes.Concat(encryptedPrefix, encryptedStream, suffix);
    }

    private static byte[] EncryptPdfStrings(byte[] input, int objectNumber, byte[] fileKey) {
        using var output = new MemoryStream(input.Length);
        int index = 0;
        while (index < input.Length) {
            byte current = input[index];
            if (current == (byte)'(' && TryReadLiteralString(input, index, out int literalEnd, out byte[] literalBytes)) {
                WriteEncryptedHexString(output, literalBytes, objectNumber, fileKey);
                index = literalEnd + 1;
                continue;
            }

            if (current == (byte)'<' && index + 1 < input.Length && input[index + 1] != (byte)'<' &&
                TryReadHexString(input, index, out int hexEnd, out byte[] hexBytes)) {
                WriteEncryptedHexString(output, hexBytes, objectNumber, fileKey);
                index = hexEnd + 1;
                continue;
            }

            output.WriteByte(current);
            index++;
        }

        return output.ToArray();
    }

    private static bool TryReadLiteralString(byte[] input, int start, out int end, out byte[] value) {
        int depth = 0;
        bool escaped = false;
        for (int i = start; i < input.Length; i++) {
            byte current = input[i];
            if (escaped) {
                escaped = false;
                continue;
            }

            if (current == (byte)'\\') {
                escaped = true;
                continue;
            }

            if (current == (byte)'(') {
                depth++;
                continue;
            }

            if (current == (byte)')') {
                depth--;
                if (depth == 0) {
                    string inner = PdfEncoding.Latin1GetString(Slice(input, start + 1, i));
                    value = PdfStringParser.ParseLiteralToBytes(inner);
                    end = i;
                    return true;
                }
            }
        }

        value = Array.Empty<byte>();
        end = -1;
        return false;
    }

    private static bool TryReadHexString(byte[] input, int start, out int end, out byte[] value) {
        for (int i = start + 1; i < input.Length; i++) {
            if (input[i] == (byte)'>') {
                if (TryParseHexBytes(input, start + 1, i, out value)) {
                    end = i;
                    return true;
                }

                break;
            }
        }

        value = Array.Empty<byte>();
        end = -1;
        return false;
    }

    private static bool TryParseHexBytes(byte[] input, int start, int end, out byte[] bytes) {
        var nibbles = new List<int>();
        for (int i = start; i < end; i++) {
            byte current = input[i];
            if (IsWhiteSpace(current)) {
                continue;
            }

            int value = HexValue(current);
            if (value < 0) {
                bytes = Array.Empty<byte>();
                return false;
            }

            nibbles.Add(value);
        }

        if ((nibbles.Count % 2) != 0) {
            nibbles.Add(0);
        }

        bytes = new byte[nibbles.Count / 2];
        for (int i = 0; i < bytes.Length; i++) {
            bytes[i] = (byte)((nibbles[i * 2] << 4) | nibbles[i * 2 + 1]);
        }

        return true;
    }

    private static void WriteEncryptedHexString(Stream output, byte[] value, int objectNumber, byte[] fileKey) {
        byte[] encrypted = value.Length == 0
            ? value
            : Rc4.Transform(ComputeObjectKey(fileKey, objectNumber, 0), value);
        byte[] hex = PdfEncoding.Latin1GetBytes(PdfSyntaxEscaper.HexString(encrypted));
        output.Write(hex, 0, hex.Length);
    }

    private static string BuildEncryptionDictionary(byte[] ownerEntry, byte[] userEntry, int permissions) {
        return "<< /Filter /Standard /V 2 /R " + Revision.ToString(CultureInfo.InvariantCulture) +
            " /Length 128 /O " + PdfSyntaxEscaper.HexString(ownerEntry) +
            " /U " + PdfSyntaxEscaper.HexString(userEntry) +
            " /P " + permissions.ToString(CultureInfo.InvariantCulture) +
            " >>";
    }

    private static byte[] ComputeOwnerEntry(string ownerPassword, string userPassword) {
        byte[] ownerKey = ComputeOwnerPasswordKey(ownerPassword);
        byte[] value = Rc4.Transform(ownerKey, PadPassword(userPassword));
        for (int i = 1; i <= 19; i++) {
            value = Rc4.Transform(XorKey(ownerKey, i), value);
        }

        return value;
    }

    private static byte[] ComputeFileKey(string userPassword, byte[] ownerEntry, int permissions, byte[] fileId) {
        byte[] padded = PadPassword(userPassword);
        var buffer = new List<byte>(padded.Length + ownerEntry.Length + 4 + fileId.Length);
        buffer.AddRange(padded);
        buffer.AddRange(ownerEntry);
        AppendInt32LittleEndian(buffer, permissions);
        buffer.AddRange(fileId);

        byte[] digest = Md5(buffer.ToArray());
        byte[] current = Take(digest, KeyLengthBytes);
        for (int i = 0; i < 50; i++) {
            current = Md5(current);
        }

        return Take(current, KeyLengthBytes);
    }

    private static byte[] ComputeUserEntry(byte[] fileKey, byte[] fileId) {
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

    private static byte[] ComputeOwnerPasswordKey(string password) {
        byte[] digest = Md5(PadPassword(password));
        for (int i = 0; i < 50; i++) {
            digest = Md5(Take(digest, KeyLengthBytes));
        }

        return Take(digest, KeyLengthBytes);
    }

    private static byte[] ComputeObjectKey(byte[] fileKey, int objectNumber, int generation) {
        var buffer = new byte[fileKey.Length + 5];
        Buffer.BlockCopy(fileKey, 0, buffer, 0, fileKey.Length);
        int offset = fileKey.Length;
        buffer[offset++] = (byte)(objectNumber & 0xFF);
        buffer[offset++] = (byte)((objectNumber >> 8) & 0xFF);
        buffer[offset++] = (byte)((objectNumber >> 16) & 0xFF);
        buffer[offset++] = (byte)(generation & 0xFF);
        buffer[offset] = (byte)((generation >> 8) & 0xFF);

        byte[] digest = Md5(buffer);
        return Take(digest, Math.Min(fileKey.Length + 5, 16));
    }

    private static byte[] PadPassword(string password) {
        byte[] passwordBytes = EncodePassword(password);
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

    private static byte[] CreateFileId() {
        var fileId = new byte[16];
        using RandomNumberGenerator rng = RandomNumberGenerator.Create();
        rng.GetBytes(fileId);
        return fileId;
    }

    private static byte[] XorKey(byte[] key, int value) {
        var result = new byte[key.Length];
        for (int i = 0; i < key.Length; i++) {
            result[i] = (byte)(key[i] ^ value);
        }

        return result;
    }

    private static byte[] Md5(byte[] data) {
#pragma warning disable CA5351, CA1850
        using MD5 md5 = MD5.Create();
        return md5.ComputeHash(data);
#pragma warning restore CA5351, CA1850
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

    private static byte[] Slice(byte[] source, int start, int end) {
        int length = end - start;
        var result = new byte[length];
        Buffer.BlockCopy(source, start, result, 0, length);
        return result;
    }

    private static int IndexOf(byte[] source, byte[] pattern, int start) {
        if (pattern.Length == 0) {
            return start;
        }

        for (int i = start; i <= source.Length - pattern.Length; i++) {
            bool match = true;
            for (int j = 0; j < pattern.Length; j++) {
                if (source[i + j] != pattern[j]) {
                    match = false;
                    break;
                }
            }

            if (match) {
                return i;
            }
        }

        return -1;
    }

    private static int LastIndexOf(byte[] source, byte[] pattern) {
        if (pattern.Length == 0) {
            return source.Length;
        }

        for (int i = source.Length - pattern.Length; i >= 0; i--) {
            bool match = true;
            for (int j = 0; j < pattern.Length; j++) {
                if (source[i + j] != pattern[j]) {
                    match = false;
                    break;
                }
            }

            if (match) {
                return i;
            }
        }

        return -1;
    }

    private static bool IsWhiteSpace(byte value) =>
        value == 0x00 ||
        value == 0x09 ||
        value == 0x0A ||
        value == 0x0C ||
        value == 0x0D ||
        value == 0x20;

    private static int HexValue(byte value) {
        if (value >= (byte)'0' && value <= (byte)'9') {
            return value - (byte)'0';
        }

        if (value >= (byte)'A' && value <= (byte)'F') {
            return value - (byte)'A' + 10;
        }

        if (value >= (byte)'a' && value <= (byte)'f') {
            return value - (byte)'a' + 10;
        }

        return -1;
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
                result[i] = (byte)(data[i] ^ state[(state[x] + state[y]) & 0xFF]);
            }

            return result;
        }

        private static void Swap(byte[] state, int left, int right) {
            byte value = state[left];
            state[left] = state[right];
            state[right] = value;
        }
    }
}
