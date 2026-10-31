using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private readonly record struct PictureAnchorKey(object Source, string Story, int StoryPage, int Occurrence);
    private sealed record PictureAnchorPlacement(PictureAnchorKey Key, int Page, double StoryOrigin,
        bool ParagraphY, double X, double Top, double Width, double Height, PdfImageTextExclusion? Exclusion);
    private sealed class PictureAnchorState {
        internal readonly Dictionary<PictureAnchorKey, PictureAnchorPlacement> Working;
        internal readonly Dictionary<PictureAnchorKey, PictureAnchorPlacement> Observed = new();
        internal readonly Dictionary<PictureParagraphPlanKey, PictureParagraphPlan> Paragraphs = new();
        internal PictureAnchorState(IReadOnlyDictionary<PictureAnchorKey, PictureAnchorPlacement>? previous = null) =>
            Working = previous == null ? new() : previous.ToDictionary(pair => pair.Key, pair => pair.Value);
    }
    private readonly record struct PictureParagraphPlanKey(IPdfBlock Block, string Story, int Page,
        double X, double Width, double FirstLineX, double FirstLineWidth, double Top, double Size);
    private sealed record PictureParagraphPlan(List<List<RichSeg>> Lines, List<double> Heights,
        TextExclusionBounds[] Exclusions, PdfOptions Fonts);

    private static bool PictureAnchorsEqual(IReadOnlyDictionary<PictureAnchorKey, PictureAnchorPlacement>? previous,
        IReadOnlyDictionary<PictureAnchorKey, PictureAnchorPlacement> actual) {
        if (previous == null) return actual.Count == 0;
        if (previous.Count != actual.Count) return false;
        foreach (var entry in actual) {
            if (!previous.TryGetValue(entry.Key, out PictureAnchorPlacement? before)) return false;
            PictureAnchorPlacement after = entry.Value;
            if (before.Page != after.Page || before.ParagraphY != after.ParagraphY ||
                Math.Abs(before.X - after.X) > .001D || Math.Abs(before.Top - after.Top) > .001D ||
                before.Width != after.Width || before.Height != after.Height) return false;
        }
        return true;
    }

    private sealed partial class LayoutContext {
        private readonly PictureAnchorState pictureAnchors;
        private readonly Dictionary<PdfCanvasAnchoredImageItem, PictureAnchorKey> pictureAnchorKeys = new();
        private readonly Dictionary<object, int> pictureSourceOccurrences = new();
        private string pictureStory = "Body";
        private int? picturePhysicalPage;
        private double? pictureStoryOrigin;
        private double? paragraphImageAnchorTop;
        private int PicturePage => picturePhysicalPage ?? pages.Count + 1;
        private int PictureStoryPage => isRunningContent ? PicturePage : 0;

        private void SetPictureStory(string story, int physicalPage, double origin) {
            pictureStory = story;
            picturePhysicalPage = physicalPage;
            pictureStoryOrigin = origin;
        }

        private PictureAnchorKey GetPictureAnchorKey(PdfCanvasAnchoredImageItem item) {
            if (pictureAnchorKeys.TryGetValue(item, out PictureAnchorKey key)) return key;
            pictureSourceOccurrences.TryGetValue(item.SourceIdentity, out int occurrence);
            pictureSourceOccurrences[item.SourceIdentity] = occurrence + 1;
            key = new PictureAnchorKey(item.SourceIdentity, pictureStory, PictureStoryPage, occurrence);
            pictureAnchorKeys.Add(item, key);
            return key;
        }

        private PictureAnchorPlacement ResolvePictureAnchor(PdfCanvasAnchoredImageItem item, double paragraphTop) {
            PdfAnchoredImagePosition position = item.Position;
            double x = position.X + (position.MarginX ? PageAnchorOptions.MarginLeft - position.AuthoredLeftMargin : 0D);
            double top = position.ParagraphY ? paragraphTop - position.Y : currentOpts.PageHeight - position.Y;
            return new PictureAnchorPlacement(GetPictureAnchorKey(item), PicturePage,
                pictureStoryOrigin ?? yStart, position.ParagraphY, x, top, item.Width, item.Height, item.Exclusion);
        }

        private static bool HasParagraphPictures(PdfParagraphStyle? style) =>
            HasAnchoredPictures(style?.AnchoredCanvas);

        private static bool HasAnchoredPictures(PdfCanvasBlock? canvas) =>
            canvas?.Items.Any(item => item is PdfCanvasAnchoredImageItem) == true;

        private void RegisterParagraphPictures(PdfCanvasBlock? canvas, double paragraphTop) {
            if (canvas == null) return;
            foreach (PdfCanvasAnchoredImageItem item in canvas.Items.OfType<PdfCanvasAnchoredImageItem>()) {
                PictureAnchorPlacement placement = ResolvePictureAnchor(item, paragraphTop);
                if (placement.Exclusion != null) pictureAnchors.Working[placement.Key] = placement;
            }
        }

        private IEnumerable<PictureAnchorPlacement> CurrentPictureExclusions() {
            foreach (PictureAnchorPlacement placement in pictureAnchors.Working.Values) {
                if (placement.Page != PicturePage || placement.Key.Story != pictureStory ||
                    placement.Key.StoryPage != PictureStoryPage || placement.Exclusion == null) continue;
                double translation = isRunningContent && placement.ParagraphY
                    ? (pictureStoryOrigin ?? yStart) - placement.StoryOrigin : 0D;
                yield return translation == 0D ? placement : placement with { Top = placement.Top + translation };
            }
        }

        private bool HasPictureExclusions => CurrentPictureExclusions().Any();

        /// <summary>Uses the same resolved picture rectangle for registration and final paint.</summary>
        private void RenderAnchoredPicture(PdfCanvasAnchoredImageItem item) {
            PictureAnchorPlacement placement = ResolvePictureAnchor(item, paragraphImageAnchorTop ?? paragraphCanvasTop ?? y);
            if (placement.Exclusion != null) {
                pictureAnchors.Working[placement.Key] = placement;
                pictureAnchors.Observed[placement.Key] = placement;
            }
            OfficeTransform translation = OfficeTransform.Translate(placement.X, currentOpts.PageHeight - placement.Top);
            if (item.Content.Items.Count == 1 && item.Content.Items[0] is PdfCanvasBehindTextItem behind) {
                // Deferred paint must capture the translation together with the image token.
                // Capturing only the inner image would leave its source coordinates at (0,0).
                var translated = new PdfCanvasEffectItem(behind.Items, translation, 1D, OfficeBlendMode.Normal);
                RenderBehindTextCanvas(new PdfCanvasBehindTextItem(new PdfCanvasItem[] { translated }, behind.ZOrder));
            } else {
                RenderCanvasEffect(new PdfCanvasEffectItem(item.Content.Items, translation, 1D, OfficeBlendMode.Normal));
            }
        }

        /// <summary>Separates the pre-clearance picture origin from legacy first-line canvas placement.</summary>
        private void RenderPictureParagraphCanvas(PdfCanvasBlock canvas, double paragraphTop, double textTop) {
            double? saved = paragraphImageAnchorTop;
            paragraphImageAnchorTop = paragraphTop;
            try { RenderParagraphCanvas(canvas, textTop); }
            finally { paragraphImageAnchorTop = saved; }
        }

        private (List<List<RichSeg>> Lines, List<double> LineHeights) WrapPictureParagraph(
            RichParagraphBlock paragraph, double frameX, double frameWidth, double fontSize, double paragraphTop) {
            PdfParagraphStyle? style = EffectiveParagraphStyle(paragraph);
            double leading = GetParagraphLeading(style, fontSize);
            var textFrame = GetParagraphTextFrame(style, frameX, frameWidth);
            return WrapPictureText(paragraph, style?.AnchoredCanvas, textFrame, fontSize, paragraphTop,
                rowLayout => WrapRichRunsCoreWithFirstLineOrigin(paragraph.Runs, textFrame.Width, fontSize,
                    ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth,
                    textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(style), currentOpts,
                    style?.TabStops.ToArray(), lineSpacing: style?.LineSpacing,
                    textWhitespaceMode: style?.TextWhitespaceMode,
                    wordJustificationMode: style?.UseWordJustificationSpacing == true && paragraph.Align == PdfAlign.Justify ? style.JustificationMode : null,
                    rowLayout: rowLayout, cancellationToken: cancellationToken));
        }

        /// <summary>Retains heading policy while composing the same picture-aware rows used by paragraph flow.</summary>
        private (List<List<RichSeg>> Lines, List<double> LineHeights) WrapHeadingForFrame(
            HeadingBlock heading, IReadOnlyList<PdfTextRun> runs, PdfHeadingStyle? style,
            double frameX, double frameWidth, double fontSize, double leading, double top) {
            if (!HasAnchoredPictures(style?.AnchoredCanvas) && !HasPictureExclusions)
                return WrapRichHeadingRuns(runs, frameWidth, fontSize, ChooseNormal(currentOpts.DefaultFont),
                    leading, null, DefaultParagraphTabStopWidth, currentOpts, style, heading.Align,
                    cancellationToken: cancellationToken);
            return WrapPictureText(heading, style?.AnchoredCanvas,
                (frameX, frameWidth, frameX, frameWidth), fontSize, top,
                rowLayout => WrapRichHeadingRuns(runs, frameWidth, fontSize, ChooseNormal(currentOpts.DefaultFont),
                    leading, null, DefaultParagraphTabStopWidth, currentOpts, style, heading.Align, rowLayout,
                    cancellationToken));
        }

        /// <summary>Owns source exclusions and retained rows for both paragraph and heading composition.</summary>
        private (List<List<RichSeg>> Lines, List<double> LineHeights) WrapPictureText(
            IPdfBlock source, PdfCanvasBlock? canvas,
            (double X, double Width, double FirstLineX, double FirstLineWidth) textFrame,
            double fontSize, double top,
            Func<Func<int, double, double, double, RichTextRowFrame>,
                (List<List<RichSeg>> Lines, List<double> LineHeights)> compose) {
            RegisterParagraphPictures(canvas, top);
            TextExclusionBounds[] exclusions = GetCurrentTextExclusions().ToArray();
            double origin = isRunningContent ? pictureStoryOrigin ?? yStart : 0D;
            var relative = exclusions.Select(region => region with { Top = region.Top - origin, Bottom = region.Bottom - origin }).ToArray();
            var key = new PictureParagraphPlanKey(source, pictureStory, PicturePage,
                textFrame.X, textFrame.Width, textFrame.FirstLineX, textFrame.FirstLineWidth, top - origin, fontSize);
            if (pictureAnchors.Paragraphs.TryGetValue(key, out PictureParagraphPlan? plan) &&
                relative.SequenceEqual(plan.Exclusions) && currentOpts.HasSameMeasurementFonts(plan.Fonts))
                return (plan.Lines, plan.Heights);
            var wrapped = compose((row, completed, required, minimumWidth) => {
                double left = row == 0 ? textFrame.FirstLineX : textFrame.X;
                double available = row == 0 ? textFrame.FirstLineWidth : textFrame.Width;
                RichTextRowFrame frame = ResolveTextRowFrame(left, available, top - completed,
                    required, minimumWidth, exclusions);
                return new RichTextRowFrame(frame.Intervals.Select(slot =>
                    new RichTextInterval(slot.X - textFrame.X, slot.Width)).ToArray(), frame.Gap);
            });
            pictureAnchors.Paragraphs[key] = new PictureParagraphPlan(wrapped.Lines, wrapped.LineHeights, relative,
                currentOpts.CreateTextMeasurementSnapshot());
            return wrapped;
        }

        private double MeasurePictureParagraph(RichParagraphBlock paragraph, double frameX, double frameWidth,
            double fontSize, bool firstVisualOnly = false) {
            PdfParagraphStyle? style = EffectiveParagraphStyle(paragraph);
            fontSize = style?.FontSize ?? currentOpts.DefaultFontSize;
            double leading = GetParagraphLeading(style, fontSize);
            double before = ResolveTopLevelSpacingBefore(GetParagraphSpacingBefore(style));
            var wrapped = WrapPictureParagraph(paragraph, frameX, frameWidth, fontSize, y - before);
            return before + (firstVisualOnly
                ? wrapped.Lines.Count == 0 ? 0D : GetRichLineFrameFitHeight(wrapped.Lines[0],
                    wrapped.LineHeights[0], ((RichLine)wrapped.Lines[0]).FragmentTopGap)
                : wrapped.LineHeights.Sum() + GetParagraphSpacingAfter(style, leading));
        }
    }
}
