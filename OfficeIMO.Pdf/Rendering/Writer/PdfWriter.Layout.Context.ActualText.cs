using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderLogicalText(string actualText, double anchorX, double anchorY, Action drawPaint) {
            // One replacement owns both the paint and its invisible anchor. Artifact marking
            // alone does not stop independent readers from extracting the painted glyphs again.
            int? markedContentId = RegisterTextStructureElement("Span", _canvasStructureParentElement);
            sb.Append("/Span << /ActualText ").Append(PdfSyntaxEscaper.TextString(actualText));
            if (markedContentId.HasValue) {
                sb.Append(" /MCID ").Append(markedContentId.Value.ToString(CultureInfo.InvariantCulture));
            }
            sb.Append(" >> BDC\n");
            PdfStandardFont font = ChooseNormal(currentOpts.DefaultFont);
            string fontResource = GetFontResourceName(font, null, font);
            // Readers can derive replacement geometry from the first and last glyph.
            // Matching anchors keep arbitrary painted baselines out of that geometry.
            WriteLogicalTextAnchor(font, fontResource, anchorX, anchorY).RestoreState();
            bool previousAccessibility = _suppressCanvasAccessibilityWrappers;
            bool previousStructure = _suppressCanvasStructureRegistration;
            bool previousActualTextChildren = _suppressCanvasActualTextChildren;
            _suppressCanvasAccessibilityWrappers = true;
            _suppressCanvasStructureRegistration = true;
            _suppressCanvasActualTextChildren = true;
            try {
                drawPaint();
            } finally {
                _suppressCanvasAccessibilityWrappers = previousAccessibility;
                _suppressCanvasStructureRegistration = previousStructure;
                _suppressCanvasActualTextChildren = previousActualTextChildren;
            }
            var content = WriteLogicalTextAnchor(font, fontResource, anchorX, anchorY);
            // Close the replacement while its anchor font and text state are active.
            // Restoring the painted text state first changes reader spacing heuristics.
            sb.Append("EMC\n");
            content.RestoreState();
            MarkSimpleFont(font);
            pageDirty = true;
        }

        private ContentStreamBuilder WriteLogicalTextAnchor(PdfStandardFont font, string fontResource, double anchorX, double anchorY) {
            var content = new ContentStreamBuilder(sb)
                .SaveState()
                .BeginText()
                .Font(fontResource, 1D)
                .TextRenderingMode(3)
                .TextMatrix(anchorX, anchorY);
            content.ShowText(EncodeActualTextAnchor(font, currentOpts), 1D);
            return content.EndText();
        }

    }
}
