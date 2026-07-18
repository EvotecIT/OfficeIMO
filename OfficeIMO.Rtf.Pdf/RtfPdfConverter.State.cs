namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private sealed class PdfRenderState {
        private readonly RtfDocument _document;
        private readonly IReadOnlyDictionary<int, OfficeIMO.Pdf.PdfStandardFont> _fontSlots;
        private readonly Dictionary<string, int> _listCounters = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly List<PdfNoteReference> _noteReferences = new List<PdfNoteReference>();

        public PdfRenderState(RtfDocument document, IReadOnlyDictionary<int, OfficeIMO.Pdf.PdfStandardFont> fontSlots) {
            _document = document;
            _fontSlots = fontSlots;
        }

        public IReadOnlyList<PdfNoteReference> NoteReferences => _noteReferences.AsReadOnly();

        public OfficeIMO.Pdf.PdfStandardFont? ResolveFont(int? fontId, bool bold, bool italic) {
            if (!fontId.HasValue || !_fontSlots.TryGetValue(fontId.Value, out OfficeIMO.Pdf.PdfStandardFont font)) {
                return null;
            }

            return OfficeIMO.Pdf.PdfStandardFontMapper.GetStyledFont(font, bold, italic);
        }

        public void AddNote(RtfNote note, string marker) {
            _noteReferences.Add(new PdfNoteReference(note, marker, _noteReferences.Count + 1));
        }

        public int NextDecimalMarker(RtfParagraph paragraph) {
            string key = GetListCounterKey(paragraph);
            if (!_listCounters.TryGetValue(key, out int value)) {
                value = GetListStart(paragraph);
            }

            _listCounters[key] = value + 1;
            return value;
        }

        public void AdvanceDecimalList(RtfParagraph paragraph, string markerText) {
            if (TryReadLeadingIntegerMarker(markerText, out int value)) {
                _listCounters[GetListCounterKey(paragraph)] = value + 1;
                return;
            }

            NextDecimalMarker(paragraph);
        }

        private int GetListStart(RtfParagraph paragraph) {
            int levelIndex = paragraph.ListLevel ?? 0;
            RtfListOverride? listOverride = paragraph.ListId.HasValue
                ? _document.ListOverrides.FirstOrDefault(item => item.Id == paragraph.ListId.Value)
                : null;
            RtfListLevelOverride? levelOverride = listOverride?.LevelOverrides.ElementAtOrDefault(levelIndex);
            if (levelOverride?.StartAt.HasValue == true) {
                return levelOverride.StartAt.Value;
            }

            int? definitionId = paragraph.ListDefinitionId ?? listOverride?.ListId;
            RtfListDefinition? definition = definitionId.HasValue
                ? _document.ListDefinitions.FirstOrDefault(item => item.Id == definitionId.Value)
                : null;
            RtfListLevel? level = definition?.Levels.FirstOrDefault(item => item.LevelIndex == levelIndex)
                ?? definition?.Levels.ElementAtOrDefault(levelIndex);
            return level?.StartAt ?? 1;
        }

        private static string GetListCounterKey(RtfParagraph paragraph) {
            int listId = paragraph.ListId ?? paragraph.ListDefinitionId ?? 0;
            int level = paragraph.ListLevel ?? 0;
            return listId.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" +
                   level.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    private sealed class PdfNoteReference {
        public PdfNoteReference(RtfNote note, string marker, int ordinal) {
            Note = note;
            Marker = marker;
            Ordinal = ordinal;
        }

        public RtfNote Note { get; }

        public string Marker { get; }

        public int Ordinal { get; }
    }
}
