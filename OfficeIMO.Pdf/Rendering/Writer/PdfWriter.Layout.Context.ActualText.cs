using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderLogicalText(string actualText, double anchorX, double anchorY, Action drawPaint,
            double width = 0D, double height = 0D) {
            // One replacement owns both the paint and its invisible anchor. Artifact marking
            // alone does not stop independent readers from extracting the painted glyphs again.
            bool paintOnly = actualText.Length == 0;
            bool hasBounds = width > 0D && height > 0D;
            int? markedContentId = paintOnly
                ? null : RegisterTextStructureElement("Span", _canvasStructureParentElement);
            // Secondary paint has no second structure-tree owner. Classify it as
            // an artifact while retaining its empty replacement and interactions.
            if (paintOnly) sb.Append("/Artifact BMC\n");
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
                WriteLogicalTextAnchor(font, fontResource, anchorX, anchorY, actualText, width, height).RestoreState();
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
                ? WriteLogicalTextAnchor(font, fontResource, anchorX, anchorY, actualText, width, height)
                : null;
            // Close the replacement while its anchor font and text state are active.
            // Restoring the painted text state first changes reader spacing heuristics.
            sb.Append("EMC\n");
            if (paintOnly) sb.Append("EMC\n");
            content?.RestoreState();
            if (actualText.Length > 0) MarkSimpleFont(font);
            pageDirty = true;
        }

        private ContentStreamBuilder WriteLogicalTextAnchor(PdfStandardFont font, string fontResource, double anchorX, double anchorY,
            string actualText, double width, double height) {
            bool hasBounds = width > 0D && height > 0D;
            double fontSize = hasBounds ? height : 1D;
            int anchorCount = hasBounds ? CountLogicalAnchorScalars(actualText) : 1;
            PdfTextShowCommand anchor = hasBounds
                ? EncodeBoundedLogicalTextAnchor(font, currentOpts, anchorCount)
                : EncodeActualTextAnchor(font, currentOpts, anchorCount);
            var content = new ContentStreamBuilder(sb)
                .SaveState()
                .BeginText()
                .Font(fontResource, fontSize)
                .TextRenderingMode(3)
                .TextMatrix(anchorX, anchorY);
            if (hasBounds) {
                // The configured embedded font can have different space metrics
                // from its standard-font fallback. Use the encoded run's advance.
                content.WordSpacing(0D).TextRise(0D).HorizontalTextScaling(ResolveLogicalAnchorScaling(anchor, width, fontSize));
            }
            content.ShowText(anchor, fontSize);
            return content.EndText();
        }

        private static PdfTextShowCommand EncodeBoundedLogicalTextAnchor(PdfStandardFont font, PdfOptions options, int count) {
            PdfTextShowCommand anchor = EncodeActualTextAnchor(font, options, count);
            // Font positioning styles painted glyphs, not the caller's explicit
            // semantic rectangle. Emit one nominal run so its first text-show
            // operation already carries the complete replacement geometry.
            double? nominalAdvance = anchor.PositionedGlyphs == null
                ? anchor.AdvanceWidth1000
                : anchor.PositionedGlyphs.Sum(glyph => (double)glyph.NominalWidth1000);
            return new PdfTextShowCommand(anchor.GlyphHex,
                advanceWidth1000: nominalAdvance, wordSpaceCount: anchor.WordSpaceCount,
                unitsPerEm: anchor.UnitsPerEm, fontMetricScale: anchor.FontMetricScale);
        }

        private static double ResolveLogicalAnchorScaling(PdfTextShowCommand anchor, double width, double height) {
            double advance = anchor.AdvanceWidth1000.GetValueOrDefault() * height / 1000D;
            if (advance <= 0D) throw new InvalidOperationException("The logical text anchor font requires a positive space advance.");
            return width / advance * 100D;
        }

        private static int CountLogicalAnchorScalars(string text) {
            int count = 0;
            for (int index = 0; index < text.Length; index++, count++) {
                if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length &&
                    char.IsLowSurrogate(text[index + 1])) index++;
            }
            return count;
        }

    }
}
