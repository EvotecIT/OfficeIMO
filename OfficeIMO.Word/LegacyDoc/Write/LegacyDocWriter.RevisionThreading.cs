namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static byte[] CreateRevisionThreading(int authorCount) {
            // MS-DOC requires all six RmdThreading STTBs for the Word 2000 format.
            // Ordinary document authors have empty message IDs and personal styles;
            // both arrays remain parallel to SttbfRMark, including its Unknown author.
            using var stream = new MemoryStream();
            WriteEmptyRevisionThreadingStrings(stream, authorCount, 8);
            WriteEmptyRevisionThreadingStrings(stream, authorCount, 0);
            WriteEmptyRevisionThreadingStrings(stream, 0, 2);
            WriteEmptyRevisionThreadingStrings(stream, 0, 0);
            WriteEmptyRevisionThreadingStrings(stream, 0, 2);
            WriteEmptyRevisionThreadingStrings(stream, 0, 0);
            return stream.ToArray();
        }

        private static void WriteEmptyRevisionThreadingStrings(MemoryStream stream, int count, ushort extraLength) {
            WriteUInt16(stream, 0xFFFF);
            WriteUInt16(stream, checked((ushort)count));
            WriteUInt16(stream, extraLength);
            for (int index = 0; index < count; index++) {
                WriteUInt16(stream, 0);
                for (int extra = 0; extra < extraLength; extra++) stream.WriteByte(0);
            }
        }
    }
}
