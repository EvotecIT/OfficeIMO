namespace OfficeIMO.OneNote;

// The CAB adapter keeps OneNote diagnostics while the compression codec is shared.
internal static class OneNoteLzxDecoder {
    internal static byte[] Decompress(IReadOnlyList<byte[]> chunks, IReadOnlyList<int> sizes,
        int windowBits, long maxOutputBytes) {
        try { return OfficeLzxDecoder.Decompress(chunks, sizes, windowBits, maxOutputBytes); }
        catch (OfficeLzxException exception) {
            throw new OneNoteFormatException("ONENOTE_CAB_" + exception.Code, exception.Message);
        }
    }
}
