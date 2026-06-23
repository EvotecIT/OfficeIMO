namespace OfficeIMO.Excel.LegacyXls.Compound {
    internal sealed class LegacyCompoundFile {
        internal LegacyCompoundFile(
            IReadOnlyDictionary<string, byte[]> streams,
            IReadOnlyList<LegacyCompoundFileEntry> entries) {
            Streams = streams;
            Entries = entries;
        }

        internal IReadOnlyDictionary<string, byte[]> Streams { get; }

        internal IReadOnlyList<LegacyCompoundFileEntry> Entries { get; }
    }
}
