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
            "native-body-chart-unsupported" or
            "native-body-content-control-unsupported" or
            "native-body-element-unsupported" or
            "native-body-embedded-document-unsupported" or
            "native-body-equation-unsupported" or
            "native-body-image-unavailable" or
            "NativeBodyImageUnsupported" or
            "native-body-smart-art-unsupported" or
            "native-header-footer-content-control-unsupported" or
            "native-header-footer-embedded-document-unsupported" or
            "native-header-footer-equation-unsupported" or
            "NativeHeaderFooterImageUnsupported" or
            "native-header-footer-shape-unsupported" or
            "native-header-footer-smart-art-unsupported" or
            "native-header-footer-text-box-unsupported" or
            "native-footnote-content-omitted" or
            "native-footnote-separator-mark-omitted" or
            "native-footnote-separator-decoration-omitted" or
            "native-field-result-unavailable" or
            "native-mixed-drawable-unsupported" or
            "native-shape-group-unsupported" or
            "native-watermark-image-unsupported" or
            "native-anchored-image-clipped" or
            "native-paragraph-bar-tab-omitted" or
            "native-bookmark-target-unprojected" or
            "native-bookmark-target-missing" => OfficeIMO.OfficeConversionLossKind.Omission,
            _ => OfficeIMO.OfficeConversionLossKind.Approximation
        };
    }
}
