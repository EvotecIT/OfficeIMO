namespace OfficeIMO.Email.Store;

internal static class PstDeflate {
    internal static byte[] Decode(byte[] compressed, int expectedLength) {
        if (expectedLength <= 0) throw new InvalidDataException("A compressed OST block has no decoded length.");
        if (compressed == null) throw new ArgumentNullException(nameof(compressed));
        bool zlibWrapped = HasZlibWrapper(compressed);
        int payloadOffset = zlibWrapped ? 2 : 0;
        int payloadLength = zlibWrapped ? compressed.Length - 6 : compressed.Length;
        if (payloadLength < 0) throw new InvalidDataException("A compressed OST block is truncated.");
        var result = new byte[expectedLength];
        using (var source = new MemoryStream(
            compressed, payloadOffset, payloadLength, writable: false, publiclyVisible: true))
        using (var deflate = new System.IO.Compression.DeflateStream(
            source, System.IO.Compression.CompressionMode.Decompress, leaveOpen: false)) {
            int total = 0;
            while (total < result.Length) {
                int read = deflate.Read(result, total, result.Length - total);
                if (read == 0) break;
                total += read;
            }
            if (total != result.Length || deflate.ReadByte() != -1) {
                throw new InvalidDataException("A compressed OST block did not produce its declared decoded length.");
            }
        }
        if (zlibWrapped) ValidateAdler32(compressed, result);
        return result;
    }

    private static bool HasZlibWrapper(byte[] bytes) {
        if (bytes.Length < 6) return false;
        int header = bytes[0] << 8 | bytes[1];
        bool deflate = (bytes[0] & 0x0F) == 8;
        bool validWindow = (bytes[0] >> 4) <= 7;
        bool validCheck = header % 31 == 0;
        if (!deflate || !validWindow || !validCheck) return false;
        if ((bytes[1] & 0x20) != 0) {
            throw new NotSupportedException("Preset dictionaries are not supported for compressed OST blocks.");
        }
        return true;
    }

    private static void ValidateAdler32(byte[] compressed, byte[] decoded) {
        int offset = compressed.Length - 4;
        uint expected = (uint)(compressed[offset] << 24 | compressed[offset + 1] << 16 |
            compressed[offset + 2] << 8 | compressed[offset + 3]);
        const uint modulus = 65521;
        uint a = 1;
        uint b = 0;
        for (int index = 0; index < decoded.Length; index++) {
            a = (a + decoded[index]) % modulus;
            b = (b + a) % modulus;
        }
        uint actual = b << 16 | a;
        if (actual != expected) {
            throw new InvalidDataException("A compressed OST block has an invalid Adler-32 checksum.");
        }
    }
}
