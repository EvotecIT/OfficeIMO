namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Describes content that could not be faithfully mapped during PDF export.
    /// </summary>
    public sealed class PdfExportWarning {
        /// <summary>
        /// Stable warning code suitable for assertions and wrapper routing.
        /// </summary>
        public string Code { get; }

        /// <summary>
        /// Document location or feature area where the warning was produced.
        /// </summary>
        public string Source { get; }

        /// <summary>
        /// Human-readable warning message.
        /// </summary>
        public string Message { get; }

        /// <summary>
        /// Creates a new export warning.
        /// </summary>
        /// <param name="code">Stable warning code.</param>
        /// <param name="source">Document location or feature area.</param>
        /// <param name="message">Human-readable warning message.</param>
        public PdfExportWarning(string code, string source, string message) {
            Code = code ?? throw new System.ArgumentNullException(nameof(code));
            Source = source ?? throw new System.ArgumentNullException(nameof(source));
            Message = message ?? throw new System.ArgumentNullException(nameof(message));
        }

        /// <summary>
        /// Returns a readable representation of the warning.
        /// </summary>
        public override string ToString() => Code + ": " + Message;

        /// <summary>
        /// Converts this Word-specific warning to the shared PDF conversion warning contract.
        /// </summary>
        public OfficeIMO.Pdf.PdfConversionWarning ToConversionWarning() =>
            new OfficeIMO.Pdf.PdfConversionWarning("OfficeIMO.Word.Pdf", Code, Source, Message,
                OfficeIMO.Pdf.PdfConversionWarningSeverity.Warning, ResolveLossKind());

        // These paths skip content, whereas wrapping, font substitution and chart
        // appearance fallbacks retain content with reduced fidelity. Both PDF and
        // image acceptance policies must see the same distinction.
        private OfficeIMO.OfficeConversionLossKind ResolveLossKind() => Code switch {
            "NativeBodyChartUnsupported" or
            "NativeBodyContentControlUnsupported" or
            "NativeBodyElementUnsupported" or
            "NativeBodyEmbeddedDocumentUnsupported" or
            "NativeBodyEquationUnsupported" or
            "NativeBodyImageUnavailable" or
            "NativeBodyImageUnsupported" or
            "NativeBodySmartArtUnsupported" or
            "NativeHeaderFooterContentControlUnsupported" or
            "NativeHeaderFooterEmbeddedDocumentUnsupported" or
            "NativeHeaderFooterEquationUnsupported" or
            "NativeHeaderFooterImageUnsupported" or
            "NativeHeaderFooterShapeUnsupported" or
            "NativeHeaderFooterSmartArtUnsupported" or
            "NativeHeaderFooterTextBoxUnsupported" or
            "NativeFootnoteContentOmitted" or
            "NativeFootnoteSeparatorMarkOmitted" or
            "NativeFootnoteSeparatorDecorationOmitted" or
            "NativeFieldResultUnavailable" or
            "NativeMixedDrawableUnsupported" or
            "NativeShapeGroupUnsupported" or
            "NativeWatermarkImageUnsupported" or
            "NativeAnchoredImageClipped" => OfficeIMO.OfficeConversionLossKind.Omission,
            _ => OfficeIMO.OfficeConversionLossKind.Approximation
        };
    }
}
