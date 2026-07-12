using System.Globalization;
using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

internal static partial class PdfStandardSecurityWriter {
    private static PdfEncryptionAssembly EncryptAes128(IReadOnlyList<byte[]> sourceObjects, PdfStandardEncryptionOptions options) {
        byte[] fileId = CreateFileId();
        string ownerPassword = options.OwnerPassword ?? options.UserPassword;
        byte[] ownerEntry = ComputeOwnerEntry(ownerPassword, options.UserPassword);
        byte[] fileKey = ComputeAes128FileKey(options.UserPassword, ownerEntry, options.Permissions, fileId, options.EncryptMetadata);
        byte[] userEntry = ComputeUserEntry(fileKey, fileId);
        int encryptionObjectNumber = sourceObjects.Count + 1;
        var objects = new List<byte[]>(sourceObjects.Count + 1);
        for (int i = 0; i < sourceObjects.Count; i++) {
            objects.Add(EncryptAesIndirectObject(sourceObjects[i], i + 1, fileKey, options.EncryptMetadata, deriveObjectKey: true));
        }

        objects.Add(PdfObjectBytes.WrapIndirectObject(
            encryptionObjectNumber,
            BuildAes128EncryptionDictionary(ownerEntry, userEntry, options.Permissions, options.EncryptMetadata)));
        return new PdfEncryptionAssembly(objects, encryptionObjectNumber, fileId);
    }

    private static byte[] ComputeAes128FileKey(
        string userPassword,
        byte[] ownerEntry,
        int permissions,
        byte[] fileId,
        bool encryptMetadata) {
        byte[] padded = PadPassword(userPassword);
        var buffer = new List<byte>(padded.Length + ownerEntry.Length + fileId.Length + 8);
        buffer.AddRange(padded);
        buffer.AddRange(ownerEntry);
        AppendInt32LittleEndian(buffer, permissions);
        buffer.AddRange(fileId);
        if (!encryptMetadata) {
            buffer.AddRange(new byte[] { 0xFF, 0xFF, 0xFF, 0xFF });
        }

        byte[] current = Take(Md5(buffer.ToArray()), KeyLengthBytes);
        for (int i = 0; i < 50; i++) {
            current = Md5(current);
        }

        return Take(current, KeyLengthBytes);
    }

    private static byte[] EncryptAesIndirectObject(byte[] objectBytes, int objectNumber, byte[] fileKey, bool encryptMetadata, bool deriveObjectKey) {
        int bodyStart = IndexOf(objectBytes, PdfEncoding.Latin1GetBytes("obj\n"), 0);
        int bodyEnd = LastIndexOf(objectBytes, PdfEncoding.Latin1GetBytes("endobj\n"));
        if (bodyStart < 0 || bodyEnd <= bodyStart) {
            return objectBytes;
        }

        bodyStart += 4;
        byte[] body = Slice(objectBytes, bodyStart, bodyEnd);
        return PdfObjectBytes.WrapIndirectObject(objectNumber, EncryptAesObjectBody(body, objectNumber, fileKey, encryptMetadata, deriveObjectKey));
    }

    private static byte[] EncryptAesObjectBody(byte[] body, int objectNumber, byte[] fileKey, bool encryptMetadata, bool deriveObjectKey) {
        byte[] streamMarker = PdfEncoding.Latin1GetBytes("\nstream\n");
        byte[] endStreamMarker = PdfEncoding.Latin1GetBytes("\nendstream");
        int streamMarkerIndex = IndexOf(body, streamMarker, 0);
        byte[] objectKey = deriveObjectKey ? ComputeAes128ObjectKey(fileKey, objectNumber, generation: 0) : fileKey;
        if (streamMarkerIndex < 0) {
            return EncryptAesPdfStrings(body, objectKey);
        }

        int streamDataStart = streamMarkerIndex + streamMarker.Length;
        int streamDataEnd = LastIndexOf(body, endStreamMarker);
        if (streamDataEnd < streamDataStart) {
            return EncryptAesPdfStrings(body, objectKey);
        }

        byte[] prefix = Slice(body, 0, streamDataStart);
        byte[] streamData = Slice(body, streamDataStart, streamDataEnd);
        byte[] suffix = Slice(body, streamDataEnd, body.Length);
        if (!encryptMetadata && IndexOf(prefix, PdfEncoding.Latin1GetBytes("/Type /Metadata"), 0) >= 0) {
            return PdfObjectBytes.Concat(EncryptAesPdfStrings(prefix, objectKey), streamData, suffix);
        }

        byte[] encryptedStream = EncryptAesCbc(objectKey, streamData);
        prefix = ReplaceDirectStreamLength(prefix, encryptedStream.Length);
        byte[] encryptedPrefix = EncryptAesPdfStrings(prefix, objectKey);
        return PdfObjectBytes.Concat(encryptedPrefix, encryptedStream, suffix);
    }

    private static byte[] EncryptAesPdfStrings(byte[] input, byte[] objectKey) {
        using var output = new MemoryStream(input.Length + 32);
        int index = 0;
        while (index < input.Length) {
            byte current = input[index];
            if (current == (byte)'(' && TryReadLiteralString(input, index, out int literalEnd, out byte[] literalBytes)) {
                WriteAesEncryptedHexString(output, literalBytes, objectKey);
                index = literalEnd + 1;
                continue;
            }

            if (current == (byte)'<' && index + 1 < input.Length && input[index + 1] != (byte)'<' &&
                TryReadHexString(input, index, out int hexEnd, out byte[] hexBytes)) {
                WriteAesEncryptedHexString(output, hexBytes, objectKey);
                index = hexEnd + 1;
                continue;
            }

            output.WriteByte(current);
            index++;
        }

        return output.ToArray();
    }

    private static void WriteAesEncryptedHexString(Stream output, byte[] value, byte[] objectKey) {
        byte[] encrypted = EncryptAesCbc(objectKey, value);
        byte[] hex = PdfEncoding.Latin1GetBytes(PdfSyntaxEscaper.HexString(encrypted));
        output.Write(hex, 0, hex.Length);
    }

    private static byte[] EncryptAesCbc(byte[] key, byte[] data) {
        var iv = new byte[16];
        using (RandomNumberGenerator rng = RandomNumberGenerator.Create()) {
            rng.GetBytes(iv);
        }

        using Aes aes = Aes.Create();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.PKCS7;
        aes.Key = key;
        aes.IV = iv;
        using ICryptoTransform encryptor = aes.CreateEncryptor();
        byte[] ciphertext = encryptor.TransformFinalBlock(data, 0, data.Length);
        return PdfObjectBytes.Concat(iv, ciphertext);
    }

    private static byte[] ComputeAes128ObjectKey(byte[] fileKey, int objectNumber, int generation) {
        var buffer = new byte[fileKey.Length + 9];
        Buffer.BlockCopy(fileKey, 0, buffer, 0, fileKey.Length);
        int offset = fileKey.Length;
        buffer[offset++] = (byte)(objectNumber & 0xFF);
        buffer[offset++] = (byte)((objectNumber >> 8) & 0xFF);
        buffer[offset++] = (byte)((objectNumber >> 16) & 0xFF);
        buffer[offset++] = (byte)(generation & 0xFF);
        buffer[offset++] = (byte)((generation >> 8) & 0xFF);
        buffer[offset++] = 0x73;
        buffer[offset++] = 0x41;
        buffer[offset++] = 0x6C;
        buffer[offset] = 0x54;
        byte[] digest = Md5(buffer);
        return Take(digest, Math.Min(fileKey.Length + 5, 16));
    }

    private static byte[] ReplaceDirectStreamLength(byte[] prefix, int encryptedLength) {
        byte[] marker = PdfEncoding.Latin1GetBytes("/Length");
        int markerIndex = IndexOf(prefix, marker, 0);
        if (markerIndex < 0) {
            throw new InvalidOperationException("AES-encrypted PDF streams require a direct /Length entry.");
        }

        int numberStart = markerIndex + marker.Length;
        while (numberStart < prefix.Length && IsWhiteSpace(prefix[numberStart])) {
            numberStart++;
        }

        int numberEnd = numberStart;
        while (numberEnd < prefix.Length && prefix[numberEnd] >= (byte)'0' && prefix[numberEnd] <= (byte)'9') {
            numberEnd++;
        }

        if (numberEnd == numberStart) {
            throw new InvalidOperationException("AES-encrypted PDF streams require a numeric direct /Length entry.");
        }

        byte[] replacement = PdfEncoding.Latin1GetBytes(encryptedLength.ToString(CultureInfo.InvariantCulture));
        using var output = new MemoryStream(prefix.Length - (numberEnd - numberStart) + replacement.Length);
        output.Write(prefix, 0, numberStart);
        output.Write(replacement, 0, replacement.Length);
        output.Write(prefix, numberEnd, prefix.Length - numberEnd);
        return output.ToArray();
    }

    private static string BuildAes128EncryptionDictionary(
        byte[] ownerEntry,
        byte[] userEntry,
        int permissions,
        bool encryptMetadata) =>
        "<< /Filter /Standard /V 4 /R 4 /Length 128" +
        " /O " + PdfSyntaxEscaper.HexString(ownerEntry) +
        " /U " + PdfSyntaxEscaper.HexString(userEntry) +
        " /P " + permissions.ToString(CultureInfo.InvariantCulture) +
        " /EncryptMetadata " + (encryptMetadata ? "true" : "false") +
        " /CF << /StdCF << /AuthEvent /DocOpen /CFM /AESV2 /Length 16 >> >>" +
        " /StmF /StdCF /StrF /StdCF >>";
}
