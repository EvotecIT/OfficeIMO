namespace OfficeIMO.Email.Store;

internal static class PstDeflate {
    internal static byte[] Decode(byte[] compressed, int expectedLength) {
        if (expectedLength <= 0) throw new InvalidDataException("A compressed OST block has no decoded length.");
        var result = new byte[expectedLength];
        using (var source = new MemoryStream(compressed, writable: false))
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
        return result;
    }
}
