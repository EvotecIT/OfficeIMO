namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private readonly struct PdfFontBinding {
        public PdfFontBinding(OfficeIMO.Pdf.PdfStandardFont font, string? namedFamily = null) {
            Font = font;
            NamedFamily = namedFamily;
        }
        public OfficeIMO.Pdf.PdfStandardFont Font { get; }
        public string? NamedFamily { get; }
    }

    private sealed class PdfRenderState {
        private readonly IReadOnlyDictionary<int, PdfFontBinding> _fontSlots;
        private readonly int? _defaultFontId;
        private readonly RtfListNumbering _numbering;
        private readonly List<PdfNoteReference> _noteReferences = new List<PdfNoteReference>();

        public PdfRenderState(RtfDocument document, IReadOnlyDictionary<int, PdfFontBinding> fontSlots,
            OfficeIMO.Pdf.PdfOptions fontOptions) {
            _numbering = new RtfListNumbering(document);
            _fontSlots = fontSlots;
            _defaultFontId = document.Settings.DefaultFontId;
            FontOptions = fontOptions;
        }

        public OfficeIMO.Pdf.PdfOptions FontOptions { get; set; }

        public IReadOnlyList<PdfNoteReference> NoteReferences => _noteReferences.AsReadOnly();

        public bool InColumns { get; set; }

        // Body section controls remain active until a later body paragraph resets them.
        public RtfParagraph? PreviousBodyParagraph { get; set; }

        public RtfParagraph CreateImplicitBodyParagraph(RtfDocument document) {
            RtfParagraph paragraph = PreviousBodyParagraph == null
                ? new RtfParagraph()
                : document.GetParagraphFormatting(PreviousBodyParagraph).CopyEmptyFormattingView();
            paragraph.UseParsedParagraphDefaults = true;
            return paragraph;
        }

        /// <summary>Shares operation font slots while isolating running-story numbering and note references from body flow.</summary>
        public PdfRenderState CreateRunningStoryState(RtfDocument document, OfficeIMO.Pdf.PdfOptions fontOptions) =>
            new PdfRenderState(document, _fontSlots, fontOptions);

        public OfficeIMO.Pdf.PdfStandardFont? ResolveFont(int? fontId, bool bold, bool italic) {
            if (!TryResolveFontBinding(fontId, out PdfFontBinding binding)) {
                return null;
            }

            return OfficeIMO.Pdf.PdfStandardFontMapper.GetStyledFont(binding.Font, bold, italic);
        }

        public string? ResolveFontFamily(int? fontId) {
            return TryResolveFontBinding(fontId, out PdfFontBinding binding) && binding.NamedFamily != null
                ? OfficeIMO.Drawing.OfficeFontFamilyParser.Quote(binding.NamedFamily) : null;
        }

        private bool TryResolveFontBinding(int? fontId, out PdfFontBinding binding) {
            if (fontId.HasValue && _fontSlots.TryGetValue(fontId.Value, out binding)) return true;
            if (_defaultFontId.HasValue && _fontSlots.TryGetValue(_defaultFontId.Value, out binding)) return true;
            binding = default;
            return false;
        }

        public OfficeIMO.Pdf.PdfTextRun CreateDefaultRun(string text, bool bold = false, double? fontSize = null) =>
            new OfficeIMO.Pdf.PdfTextRun(text, bold: bold, fontSize: fontSize,
                font: ResolveFont(null, bold, italic: false), fontFamily: ResolveFontFamily(null));

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
