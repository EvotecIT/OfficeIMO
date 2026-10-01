namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private sealed class PdfRenderState {
        private readonly IReadOnlyDictionary<int, OfficeIMO.Pdf.PdfStandardFont> _fontSlots;
        private readonly RtfListNumbering _numbering;
        private readonly List<PdfNoteReference> _noteReferences = new List<PdfNoteReference>();

        public PdfRenderState(RtfDocument document, IReadOnlyDictionary<int, OfficeIMO.Pdf.PdfStandardFont> fontSlots) {
            _numbering = new RtfListNumbering(document);
            _fontSlots = fontSlots;
        }

        public IReadOnlyList<PdfNoteReference> NoteReferences => _noteReferences.AsReadOnly();

        public bool InColumns { get; set; }

        public OfficeIMO.Pdf.PdfStandardFont? ResolveFont(int? fontId, bool bold, bool italic) {
            if (!fontId.HasValue || !_fontSlots.TryGetValue(fontId.Value, out OfficeIMO.Pdf.PdfStandardFont font)) {
                return null;
            }

            return OfficeIMO.Pdf.PdfStandardFontMapper.GetStyledFont(font, bold, italic);
        }

        public void AddNote(RtfNote note, string marker) {
            _noteReferences.Add(new PdfNoteReference(note, marker, _noteReferences.Count + 1));
        }

        public RtfListMarker? NextListMarker(RtfParagraph paragraph) => _numbering.Next(paragraph);
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
