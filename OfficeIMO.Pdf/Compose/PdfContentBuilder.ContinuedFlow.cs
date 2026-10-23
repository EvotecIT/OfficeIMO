namespace OfficeIMO.Pdf;

// Kept with the existing receiver: captured callbacks share its document and font resources.
public sealed partial class PdfContentBuilder {
    internal void ContinuedFlow(Action<PdfContentBuilder> build, Action<PdfContentBuilder>? separator,
        Action<PdfContentBuilder>? continuationSeparator, Action<PdfContentBuilder>? continuationNotice,
        bool reserveFinalNotice, bool allowContinuationSeparatorInBottomMargin = false) =>
        _doc.ContinuedFlow(build, separator, continuationSeparator, continuationNotice, reserveFinalNotice,
            allowContinuationSeparatorInBottomMargin);
}
