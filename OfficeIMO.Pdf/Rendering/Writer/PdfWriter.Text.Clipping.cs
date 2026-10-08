using System.Text;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static void WriteClippedRichParagraph(StringBuilder sb, RichParagraphBlock block, System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, System.Collections.Generic.List<double> lineHeights, PdfOptions opts, double startY, double fontSize, double defaultLeading, System.Collections.Generic.List<LinkAnnotation> annots, double clipX, double clipY, double clipWidth, double clipHeight, double? xOverride = null, double? widthOverride = null, double? firstLineXOverride = null, double? firstLineWidthOverride = null, string? structureType = null, int? markedContentId = null, LayoutResult.Page? structurePage = null, System.Collections.Generic.IReadOnlyList<PdfAlign?>? lineAlignments = null, System.Collections.Generic.IReadOnlyList<double>? lineXOffsets = null, System.Collections.Generic.IReadOnlyList<double>? lineWidths = null, bool suppressActualText = false, System.Collections.Generic.IReadOnlyList<double>? lineTopGaps = null, PdfStandardFont? baselineFont = null, bool clipToFrame = true) {
        // Prevent link coalescing from extending a newly clipped annotation into a prior text frame.
        annots.Add(new LinkAnnotation());
        int annotationStart = annots.Count;
        if (clipToFrame) {
            new ContentStreamBuilder(sb)
                .SaveState()
                .Rectangle(clipX, clipY, clipWidth, clipHeight)
                .ClipPath()
                .EndPath();
        }

        WriteRichParagraph(sb, block, lines, lineHeights, opts, startY, fontSize, defaultLeading, annots, xOverride, widthOverride, firstLineXOverride, firstLineWidthOverride, structureType, markedContentId, structurePage, lineAlignments, lineXOffsets, lineWidths, suppressActualText, lineTopGaps, baselineFont);
        if (clipToFrame) ClipLinkAnnotations(annots, annotationStart, clipX, clipY, clipWidth, clipHeight);
        annots.RemoveAt(annotationStart - 1);

        if (clipToFrame) new ContentStreamBuilder(sb).RestoreState();
    }

    private static void ClipLinkAnnotations(System.Collections.Generic.List<LinkAnnotation> annots, int annotationStart, double clipX, double clipY, double clipWidth, double clipHeight) {
        double clipRight = clipX + Math.Max(0D, clipWidth);
        double clipTop = clipY + Math.Max(0D, clipHeight);
        for (int index = annots.Count - 1; index >= annotationStart; index--) {
            LinkAnnotation annotation = annots[index];
            double x1 = Math.Max(annotation.X1, clipX);
            double y1 = Math.Max(annotation.Y1, clipY);
            double x2 = Math.Min(annotation.X2, clipRight);
            double y2 = Math.Min(annotation.Y2, clipTop);
            if (x2 <= x1 || y2 <= y1) {
                annots.RemoveAt(index);
                continue;
            }

            annotation.X1 = x1;
            annotation.Y1 = y1;
            annotation.X2 = x2;
            annotation.Y2 = y2;
        }
    }
}
