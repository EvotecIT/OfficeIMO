using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderLogicalText(string actualText, double anchorX, double anchorY, Action drawPaint,
            double width = 0D, double height = 0D, long? logicalOrder = null) {
            // One replacement owns both the paint and its invisible anchor. Artifact marking
            // alone does not stop independent readers from extracting the painted glyphs again.
            bool paintOnly = actualText.Length == 0;
            bool hasBounds = width > 0D && height > 0D;
            int? markedContentId = paintOnly
                ? null : RegisterTextStructureElement("Span", _canvasStructureParentElement, logicalOrder: logicalOrder);
            // Secondary paint has no second structure-tree owner. Classify it as
            // an artifact while retaining its empty replacement and interactions.
            if (paintOnly) sb.Append("/Artifact << /Type /Layout >> BDC\n");
            sb.Append("/Span << /ActualText ").Append(PdfSyntaxEscaper.TextString(actualText));
            // Other readers use the standard invisible glyph geometry. This optional
            // owner hint lets our reader distinguish it from a legacy point carrier
            // whose geometry should be promoted to the first visible paint run.
            if (hasBounds) sb.Append(" /OfficeIMOLogicalBounds true");
            if (markedContentId.HasValue) {
                sb.Append(" /MCID ").Append(markedContentId.Value.ToString(CultureInfo.InvariantCulture));
            }
            sb.Append(" >> BDC\n");
            PdfStandardFont font = ChooseNormal(currentOpts.DefaultFont);
            string fontResource = GetFontResourceName(font, null, font);
            // Readers can derive replacement geometry from the first and last glyph.
            // Matching anchors keep arbitrary painted baselines out of that geometry.
            if (actualText.Length > 0)
                WriteLogicalTextAnchor(sb, currentOpts, font, fontResource, anchorX, anchorY, actualText, width, height).RestoreState();
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
            var content = actualText.Length > 0
                ? WriteLogicalTextAnchor(sb, currentOpts, font, fontResource, anchorX, anchorY, actualText, width, height)
                : null;
            // Close the replacement while its anchor font and text state are active.
            // Restoring the painted text state first changes reader spacing heuristics.
            sb.Append("EMC\n");
            if (paintOnly) sb.Append("EMC\n");
            content?.RestoreState();
            if (actualText.Length > 0) MarkSimpleFont(font);
            pageDirty = true;
        }


    }
}
