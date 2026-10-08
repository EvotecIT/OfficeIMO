namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static IReadOnlyDictionary<string, ushort> CreateNoteParagraphStyleIndexes(
            IReadOnlyDictionary<string, ushort> styleIndexes, string noteStyleId) {
            var noteStyleIndexes = new Dictionary<string, ushort>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, ushort> styleIndex in styleIndexes) {
                noteStyleIndexes.Add(styleIndex.Key, styleIndex.Value);
            }

            // Keep the existing DOC note default while resolving authored styles
            // through the same stylesheet as the body and other stories.
            noteStyleIndexes[noteStyleId] = NoteTextParagraphStyleIndex;
            return noteStyleIndexes;
        }
    }
}
