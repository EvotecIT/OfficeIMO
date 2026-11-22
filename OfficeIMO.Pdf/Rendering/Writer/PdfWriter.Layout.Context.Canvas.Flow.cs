using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>
        /// Lays out a positioned frame's blocks with the ordinary flow engine on a virtual page as wide
        /// as the real one and tall enough for the content, then moves that paint into the frame
        /// rectangle. The surrounding flow cursor does not move.
        /// </summary>
        private void RenderCanvasFlow(PdfCanvasFlowItem item) {
            EnsurePage();
            double pageWidth = currentOpts.PageWidth.ToPoints();
            double pageHeight = currentOpts.PageHeight.ToPoints();
            PdfCreateOptions frame = currentOpts.CreateRunningContentFrame(0D, 0D);
            frame.MarginLeft = OfficeLength.Points(item.X);
            frame.MarginRight = OfficeLength.Points(Math.Max(0D, pageWidth - item.X - item.Width));
            // The frame never paints running content, so the page header and footer offsets do not apply to it.
            frame.HeaderOffsetY = OfficeLength.Zero;
            frame.FooterOffsetY = OfficeLength.Zero;
            double measured = MeasureBlockSequence(item.Blocks, item.X, item.Width, frame.DefaultFontSize.ToPoints()) ?? pageHeight;
            double virtualHeight = Math.Ceiling(Math.Max(item.Height, measured)) + 1D;
            frame.PageHeight = OfficeLength.Points(virtualHeight);

            using var story = new LayoutContext(frame, cancellationToken: cancellationToken,
                sharedPageContents: pageContents, isRunningContent: isRunningContent,
                drawingPatternAnnotations: drawingPatternAnnotations, pictureAnchors: pictureAnchors, svgForms: svgForms);
            story.currentPage = currentPage;
            story.width = item.Width;
            story.yStart = virtualHeight;
            story.y = virtualHeight;
            story.currentVisiblePageNumber = currentVisiblePageNumber;
            story.currentPageGroupId = currentPageGroupId;
            LayoutResult.Page page = currentPage!;
            int firstImage = page.Images.Count;
            int firstAnnotation = page.Annotations.Count, firstTextAnnotation = page.TextAnnotations.Count;
            int firstFreeTextAnnotation = page.FreeTextAnnotations.Count, firstHighlightAnnotation = page.HighlightAnnotations.Count;
            int firstFormField = page.FormFields.Count, firstBookmark = page.Bookmarks.Count, firstDestination = page.NamedDestinations.Count;
            story.ProcessBlocks(item.Blocks.ToList());
            if (!ReferenceEquals(story.currentPage, currentPage) || story.pages.Count > 0)
                throw new InvalidOperationException("Positioned frame content left its frame page.");
            if (isRunningContent)
                for (int index = firstImage; index < page.Images.Count; index++)
                    runningAssets.Share(page.Images[index], cancellationToken);
            story.ApplyDeferredTextLayers();
            usedBold |= story.usedBold;
            usedItalic |= story.usedItalic;
            usedBoldItalic |= story.usedBoldItalic;
            double clipBottom = pageHeight - item.ClipY - item.ClipHeight;
            double shift = pageHeight - item.Y - virtualHeight;
            // Interactive regions and destinations were placed on the virtual page; move them into the frame.
            // Enclosing canvas effects transform rectangles afterwards, but destinations only when created.
            OfficeTransform toFrame = OfficeTransform.Translate(0D, shift);
            TransformCanvasRectangles(page.Annotations, firstAnnotation, toFrame);
            TransformCanvasRectangles(page.TextAnnotations, firstTextAnnotation, toFrame);
            TransformCanvasRectangles(page.FreeTextAnnotations, firstFreeTextAnnotation, toFrame);
            TransformCanvasRectangles(page.HighlightAnnotations, firstHighlightAnnotation, toFrame);
            TransformCanvasRectangles(page.FormFields, firstFormField, toFrame);
            // Content outside the frame is not painted, so its interactive regions must not remain either.
            ClipFrameRectangles(page.Annotations, firstAnnotation, item.ClipX, clipBottom, item.ClipX + item.ClipWidth, clipBottom + item.ClipHeight,
                link => (link.X1, link.Y1, link.X2, link.Y2), (link, box) => { link.X1 = box.X1; link.Y1 = box.Y1; link.X2 = box.X2; link.Y2 = box.Y2; });
            ClipFrameRectangles(page.TextAnnotations, firstTextAnnotation, item.ClipX, clipBottom, item.ClipX + item.ClipWidth, clipBottom + item.ClipHeight,
                note => (note.X1, note.Y1, note.X2, note.Y2), null);
            ClipFrameRectangles(page.FreeTextAnnotations, firstFreeTextAnnotation, item.ClipX, clipBottom, item.ClipX + item.ClipWidth, clipBottom + item.ClipHeight,
                note => (note.X1, note.Y1, note.X2, note.Y2), null);
            ClipFrameRectangles(page.HighlightAnnotations, firstHighlightAnnotation, item.ClipX, clipBottom, item.ClipX + item.ClipWidth, clipBottom + item.ClipHeight,
                note => (note.X1, note.Y1, note.X2, note.Y2), null);
            ClipFrameRectangles(page.FormFields, firstFormField, item.ClipX, clipBottom, item.ClipX + item.ClipWidth, clipBottom + item.ClipHeight,
                field => (field.X1, field.Y1, field.X2, field.Y2), null);
            TransformCanvasPageImageBounds(page.Images, firstImage, toFrame);
            OfficeTransform toPage = toFrame.Then(_canvasEffectToPage);
            for (int index = firstBookmark; index < page.Bookmarks.Count; index++) {
                OfficePoint point = toPage.TransformPoint(new OfficePoint(page.Bookmarks[index].X, page.Bookmarks[index].Y));
                page.Bookmarks[index].X = point.X;
                page.Bookmarks[index].Y = point.Y;
            }
            for (int index = firstDestination; index < page.NamedDestinations.Count; index++) {
                OfficePoint point = toPage.TransformPoint(new OfficePoint(page.NamedDestinations[index].X, page.NamedDestinations[index].Y));
                page.NamedDestinations[index].X = point.X;
                page.NamedDestinations[index].Y = point.Y;
            }
            sb.Append("q\n")
                .Append(Format(item.ClipX)).Append(' ').Append(Format(clipBottom)).Append(' ')
                .Append(Format(item.ClipWidth)).Append(' ').Append(Format(item.ClipHeight)).Append(" re W n\n")
                .Append("1 0 0 1 0 ").Append(Format(shift)).Append(" cm\n")
                .Append(story.sb).Append("Q\n");
            pageDirty = true;

            static string Format(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Drops regions wholly outside the frame. A link that is partly inside is trimmed to the visible
        /// part; other annotations and fields have a fixed appearance, so they stay only when wholly inside.
        /// </summary>
        private static void ClipFrameRectangles<T>(List<T> items, int start, double left, double bottom, double right, double top,
            Func<T, (double X1, double Y1, double X2, double Y2)> read, Action<T, (double X1, double Y1, double X2, double Y2)>? trim) {
            for (int index = items.Count - 1; index >= start; index--) {
                var (x1, y1, x2, y2) = read(items[index]);
                double minX = Math.Min(x1, x2), maxX = Math.Max(x1, x2), minY = Math.Min(y1, y2), maxY = Math.Max(y1, y2);
                double clippedLeft = Math.Max(minX, left), clippedRight = Math.Min(maxX, right);
                double clippedBottom = Math.Max(minY, bottom), clippedTop = Math.Min(maxY, top);
                bool inside = minX >= left - 0.001D && maxX <= right + 0.001D && minY >= bottom - 0.001D && maxY <= top + 0.001D;
                if (inside) continue;
                if (trim == null || clippedRight - clippedLeft <= 0.001D || clippedTop - clippedBottom <= 0.001D) {
                    items.RemoveAt(index);
                    continue;
                }
                trim(items[index], (clippedLeft, clippedBottom, clippedRight, clippedTop));
            }
        }
    }
}
