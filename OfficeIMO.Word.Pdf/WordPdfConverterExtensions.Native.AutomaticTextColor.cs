namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // The paragraph whose content is currently being built inside a painted shading panel, and that fill.
        // Set by the renderer that paints the panel, so automatic text follows the route actually taken.
        [ThreadStatic] private static OfficeXmlNode _nativePaintedPanelParagraph;
        [ThreadStatic] private static global::OfficeIMO.OfficeColor? _nativePaintedPanelFill;

        /// <summary>Marks the paragraph fill painted behind content built until the scope is disposed.</summary>
        private static NativePaintedPanelScope EnterNativePaintedPanel(WordParagraph paragraph, global::OfficeIMO.OfficeColor? fill) {
            var scope = new NativePaintedPanelScope(_nativePaintedPanelParagraph, _nativePaintedPanelFill);
            _nativePaintedPanelParagraph = paragraph.Node();
            _nativePaintedPanelFill = fill;
            return scope;
        }

        private readonly struct NativePaintedPanelScope : IDisposable {
            private readonly OfficeXmlNode _paragraph;
            private readonly global::OfficeIMO.OfficeColor? _fill;

            internal NativePaintedPanelScope(OfficeXmlNode paragraph, global::OfficeIMO.OfficeColor? fill) {
                _paragraph = paragraph;
                _fill = fill;
            }

            public void Dispose() {
                _nativePaintedPanelParagraph = _paragraph;
                _nativePaintedPanelFill = _fill;
            }
        }

        /// <summary>
        /// Resolves Word's automatic text color against the fill painted behind the paragraph: its shading
        /// panel when the renderer paints one, otherwise the containing table cell. Returns white on a dark
        /// fill and <see langword="null"/> (the ordinary black default) otherwise.
        /// </summary>
        private static global::OfficeIMO.OfficeColor? ResolveNativeAutomaticTextColor(WordParagraph paragraph) {
            OfficeXmlNode node = paragraph.Node();
            global::OfficeIMO.OfficeColor? background = !node.IsNone && node.Equals(_nativePaintedPanelParagraph)
                    ? _nativePaintedPanelFill
                    : null;
            if (!background.HasValue && paragraph._document != null) {
                OfficeXmlNode cellShading = paragraph.ContainingTableCellShadingNode();
                background = cellShading.IsNone ? null : ParseNativeColor(WordConverterNodes.ResolveShadingFillColorHex(cellShading, paragraph._document));
            }

            return background is { } fill ? ResolveNativeAutomaticTextColor(fill) : null;
        }

        /// <summary>Returns white for automatic text on a dark backdrop and <see langword="null"/> (black) otherwise.</summary>
        private static global::OfficeIMO.OfficeColor? ResolveNativeAutomaticTextColor(global::OfficeIMO.OfficeColor fill) {
            // Word switches automatic text on encoded RGB brightness, not WCAG contrast; Word-produced
            // DOC and DOCX references cover both sides of this threshold.
            var rgb = fill.ToNormalizedRgb();
            double brightness = 0.299D * rgb.R + 0.587D * rgb.G + 0.114D * rgb.B;
            return brightness < 60D / 256D ? global::OfficeIMO.OfficeColor.White : null;
        }
    }
}
