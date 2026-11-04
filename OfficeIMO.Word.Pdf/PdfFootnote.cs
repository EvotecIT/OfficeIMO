namespace OfficeIMO.Word.Pdf {
    internal sealed class PdfFootnote {
        public string Label { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
        public bool IsEndnote { get; set; }
        public long ReferenceId { get; set; }
        public DocumentFormat.OpenXml.Wordprocessing.FootnoteEndnoteType? Source { get; set; }
    }
}

