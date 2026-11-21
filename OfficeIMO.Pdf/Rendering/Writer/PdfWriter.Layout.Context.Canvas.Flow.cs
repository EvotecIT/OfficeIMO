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
            if (story.behindTextCanvases.Count > 0) {
                var behind = new StringBuilder();
                foreach (var canvas in story.behindTextCanvases.OrderBy(entry => entry.ZOrder)) behind.Append(canvas.Content);
                story.sb.Insert(0, behind.ToString());
            }
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
    }
}
