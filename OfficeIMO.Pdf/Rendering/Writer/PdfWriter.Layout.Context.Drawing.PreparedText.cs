using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private OfficeCffOperationBudget? preparedDrawingCffBudget;

        // A retained run belongs to its exact program, before profile aliases or fresh shaping can replace it.
        private void DrawPreparedDrawingTextAt(OfficeDrawingText text, OfficePreparedTextShaping prepared,
            double originX, double originTopY) {
            cancellationToken.ThrowIfCancellationRequested();
            if (text.WrapText || text.ShrinkToFit || text.StackedText || text.HasPadding || text.HasParagraphIndent ||
                text.VerticalAlignment != OfficeTextVerticalAlignment.Top || text.TextDirection == OfficeTextDirection.TopToBottom ||
                text.Font.IsUnderline || text.Font.IsStrikethrough || text.UnderlineStyle != OfficeTextDecorationStyle.None ||
                text.StrikethroughStyle != OfficeTextDecorationStyle.None || text.FontMetricScale != 1D ||
                text.OutlineStrokeColor.HasValue || text.FillSuppressed ||
                !string.Equals(text.RasterText, prepared.Text, StringComparison.Ordinal) ||
                !string.Equals(text.Text, prepared.Text, StringComparison.Ordinal) ||
                !string.IsNullOrEmpty(text.FontPalette) && text.FontPalette != "normal")
                throw new NotSupportedException(
                    "Prepared PDF glyphs require one unwrapped positioned horizontal monochrome slice without decorations or specialized text geometry.");

            double size = text.Font.Size * text.BaselineScale;
            double measuredAdvance = prepared.Measure(size);
            double advance = text.TextAdvanceWidth ?? measuredAdvance;
            if (double.IsNaN(advance) || double.IsInfinity(advance) || advance < 0D)
                throw new NotSupportedException("Prepared PDF glyphs require a finite nonnegative positioned advance.");
            double scaleX = measuredAdvance > 0D ? advance / measuredAdvance : 1D;
            double left = OfficeTextPlacement.ResolveLineLeft(0D, text.Width, advance, text.Alignment);
            double baselineOffset = prepared.Program is IOfficeFontBaselineMetrics metrics
                ? metrics.BaselineOffset(size) : prepared.Program.LineHeight(size) * .8D;
            const int MaximumOutlinePoints = 1_000_000;
            List<List<OfficePoint>> contours = prepared.GetContours(0D, -baselineOffset, size,
                MaximumOutlinePoints, cancellationToken, preparedDrawingCffBudget ??= new OfficeCffOperationBudget());
            var commands = new List<OfficePathCommand>();
            foreach (List<OfficePoint> contour in contours) {
                cancellationToken.ThrowIfCancellationRequested();
                if (contour.Count < 3) continue;
                for (int index = 0; index < contour.Count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var point = new OfficePoint(left + contour[index].X * scaleX,
                        text.Font.Size + text.BaselineOffset + contour[index].Y);
                    commands.Add(index == 0 ? OfficePathCommand.MoveTo(point) : OfficePathCommand.LineTo(point));
                }
                commands.Add(OfficePathCommand.Close());
            }
            OfficeShape? shape = commands.Count == 0 ? null : OfficeShape.Path(text.Width, text.Height, commands);
            if (shape != null) {
                shape.FillColor = text.Color ?? OfficeColor.Black;
                shape.FillRule = OfficeFillRule.NonZero;
                shape.StrokeColor = null;
            }
            double frameX = originX + text.X;
            double frameTopY = originTopY - text.Y;
            void Paint() {
                var content = new ContentStreamBuilder(sb).SaveState();
                if (text.ClipToFrame)
                    content.Rectangle(frameX, frameTopY - text.Height, text.Width, text.Height).ClipPath().EndPath();
                if (shape != null) DrawShapeGeometryAt(shape, frameX, frameTopY - text.Height);
                content.RestoreState();
            }
            void PaintLogical() {
                if (_suppressCanvasActualTextChildren) { Paint(); return; }
                bool previousStructure = _suppressCanvasStructureRegistration;
                _suppressCanvasStructureRegistration |= _suppressCanvasAccessibilityWrappers;
                try {
                    RenderLogicalText(prepared.Text, frameX + left,
                        frameTopY - text.Font.Size - text.BaselineOffset, Paint, advance, size);
                } finally {
                    _suppressCanvasStructureRegistration = previousStructure;
                }
            }
            if (text.HasFrameTransform) {
                OfficeTransform transform = ToTopLeftPageTransform(
                    text.CreateFrameTransform().CreateDestinationTransform(), originX, originTopY);
                RenderEffectGroup(transform, 1D, PaintLogical);
            } else PaintLogical();
            pageDirty = true;
        }
    }
}
