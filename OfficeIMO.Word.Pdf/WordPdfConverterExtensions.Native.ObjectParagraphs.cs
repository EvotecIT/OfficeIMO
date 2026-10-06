namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Applies paragraph spacing once around the objects that actually enter flow.</summary>
    private sealed class NativeObjectParagraphSpacing {
        private readonly INativePdfFlow _flow;
        private readonly double _before;
        private readonly double _after;
        private bool _started;

        internal NativeObjectParagraphSpacing(INativePdfFlow flow, double before, double after) {
            _flow = flow;
            _before = before;
            _after = after;
        }

        internal void BeforeFlowObject() {
            if (_started) return;
            _flow.ParagraphSpacingBefore(_before);
            _started = true;
        }

        internal void Complete() {
            if (_started) _flow.ParagraphSpacingAfter(_after);
        }
    }
}
