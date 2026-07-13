using System.Diagnostics;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Markup.PowerPoint;

internal sealed partial class OfficeMarkupPowerPointExporter {
    private static IEnumerable<OfficeMarkupSlideBlock> GetSlides(OfficeMarkupDocument document) {
        var pendingBlocks = new List<OfficeMarkupBlock>();
        foreach (var block in document.Blocks) {
            if (block is OfficeMarkupSlideBlock slide) {
                if (pendingBlocks.Count > 0) {
                    yield return CreateImplicitSlide(pendingBlocks);
                    pendingBlocks.Clear();
                }

                yield return slide;
            } else {
                pendingBlocks.Add(block);
            }
        }

        if (pendingBlocks.Count > 0) {
            yield return CreateImplicitSlide(pendingBlocks);
        }
    }

    private static OfficeMarkupSlideBlock CreateImplicitSlide(IReadOnlyList<OfficeMarkupBlock> blocks) {
        var slide = new OfficeMarkupSlideBlock();
        var startIndex = 0;
        if (blocks.Count > 0 && blocks[0] is OfficeMarkupHeadingBlock heading && heading.Level == 1) {
            slide.Title = heading.Text;
            startIndex = 1;
        }

        for (int index = startIndex; index < blocks.Count; index++) {
            slide.Blocks.Add(blocks[index]);
        }

        return slide;
    }

    private static void ExportSlide(
        PowerPointPresentation presentation,
        PowerPointDeckComposer deck,
        OfficeMarkupSlideBlock slideBlock,
        MarkupToPowerPointOptions options,
        SlideCanvasMetrics metrics,
        OfficeMarkupStyleResolver styleResolver) {
        if (TryExportDesignerSlide(presentation, deck, slideBlock, styleResolver, metrics, out var designedSlide)) {
            ApplyTransition(designedSlide!, slideBlock.Transition);
            ApplyBackground(designedSlide!, slideBlock.Background, deck.Design.Theme, options, metrics);
            if (!string.IsNullOrWhiteSpace(slideBlock.Notes)) {
                designedSlide!.Notes.Text = slideBlock.Notes!;
            }

            return;
        }

        var useDesignerCanvas = ShouldUseDesignerCanvas(slideBlock);
        var slide = useDesignerCanvas
            ? deck.ComposeSlide(composer => {
                if (ShouldRenderSlideTitle(slideBlock)) {
                    composer.AddTitle(slideBlock.Title!);
                }
            }, seed: slideBlock.Title ?? slideBlock.Layout ?? "slide", dark: IsDarkBackground(slideBlock.Background, deck.Design.Theme))
            : presentation.AddSlide();

        ApplyTransition(slide, slideBlock.Transition);
        ApplyBackground(slide, slideBlock.Background, deck.Design.Theme, options, metrics);

        if (!useDesignerCanvas && !HasExplicitBackground(slideBlock)) {
            AddMarkupCanvas(slide, deck.Design.Theme, slideBlock, metrics);
        }

        if (!useDesignerCanvas && ShouldRenderSlideTitle(slideBlock)) {
            var title = slide.AddTitleInches(
                slideBlock.Title!,
                metrics.Horizontal(TitleLeft),
                metrics.Vertical(TitleTop),
                metrics.Horizontal(TitleWidth),
                metrics.Vertical(TitleHeight));
            ApplyTextStyle(title, styleResolver.Resolve("slide-title"));
        }

        var cursor = useDesignerCanvas
            ? new LayoutCursor(
                metrics.Horizontal(DesignerBodyLeft),
                metrics.Vertical(DesignerBodyTop),
                metrics.Horizontal(DesignerBodyWidth),
                metrics.Vertical(DesignerBodyHeight))
            : new LayoutCursor(
                metrics.Horizontal(BodyLeft),
                metrics.Vertical(BodyTop),
                metrics.Horizontal(BodyWidth),
                metrics.Vertical(BodyHeight));
        for (int index = 0; index < slideBlock.Blocks.Count; index++) {
            var block = slideBlock.Blocks[index];
            if (IsColumns(block)) {
                index = ExportColumns(slide, slideBlock.Blocks, index, options, metrics, cursor, styleResolver);
                if (!HasExplicitPlacement(block)) {
                    cursor.MoveToBottom();
                }
                continue;
            }

            ExportBlock(slide, block, options, metrics, cursor, styleResolver);
        }

        if (!string.IsNullOrWhiteSpace(slideBlock.Notes)) {
            slide.Notes.Text = slideBlock.Notes!;
        }
    }

    private static bool TryExportDesignerSlide(
        PowerPointPresentation presentation,
        PowerPointDeckComposer deck,
        OfficeMarkupSlideBlock slideBlock,
        OfficeMarkupStyleResolver styleResolver,
        SlideCanvasMetrics metrics,
        out PowerPointSlide? slide) {
        if (TryExportMarkupSummarySlide(presentation, deck.Design.Theme, slideBlock, styleResolver, metrics, out slide)
            || TryExportDesignerSectionSlide(deck, slideBlock, out slide)
            || TryExportDesignerProcessSlide(deck, slideBlock, out slide)
            || TryExportDesignerTwoColumnSlide(deck, slideBlock, styleResolver, out slide)
            || TryExportDesignerCardGridSlide(deck, slideBlock, out slide)) {
            return true;
        }

        return false;
    }

    private static bool TryExportMarkupSummarySlide(
        PowerPointPresentation presentation,
        PowerPointDesignTheme theme,
        OfficeMarkupSlideBlock slideBlock,
        OfficeMarkupStyleResolver styleResolver,
        SlideCanvasMetrics metrics,
        out PowerPointSlide? slide) {
        slide = null;
        var layout = Normalize(slideBlock.Layout ?? string.Empty);
        if (layout != "titleandcontent" && layout != "content") {
            return false;
        }

        if (HasExplicitBackground(slideBlock) || HasAnyExplicitPlacement(slideBlock.Blocks)) {
            return false;
        }

        var list = slideBlock.Blocks.OfType<OfficeMarkupListBlock>().FirstOrDefault();
        if (list == null || list.Items.Count < 2 || list.Items.Count > 4) {
            return false;
        }

        var supported = slideBlock.Blocks.All(block => block is OfficeMarkupListBlock || block is OfficeMarkupParagraphBlock);
        if (!supported) {
            return false;
        }

        var items = list.Items.Select(item => item.Text.Trim()).Where(item => item.Length > 0).ToList();
        if (items.Count < 2) {
            return false;
        }

        slide = presentation.AddSlide();
        AddMarkupCanvas(slide, theme, slideBlock, metrics, strong: true);

        var eyebrow = slide.AddTextBoxInches("OfficeIMO Markup", metrics.Horizontal(0.72), metrics.Vertical(0.42), metrics.Horizontal(2.2), metrics.Vertical(0.28));
        ApplyTextStyle(eyebrow, new OfficeMarkupResolvedStyle {
            FontName = theme.BodyFontName,
            FontSize = 9,
            Bold = true,
            TextColor = theme.AccentColor
        });

        var title = slide.AddTitleInches(slideBlock.Title ?? "Summary", metrics.Horizontal(0.72), metrics.Vertical(0.72), metrics.Horizontal(5.7), metrics.Vertical(0.78));
        ApplyTextStyle(title, new OfficeMarkupResolvedStyle {
            FontName = theme.HeadingFontName,
            FontSize = 30,
            Bold = true,
            TextColor = theme.PrimaryTextColor
        });

        var paragraph = slideBlock.Blocks.OfType<OfficeMarkupParagraphBlock>().FirstOrDefault();
        if (paragraph != null && !string.IsNullOrWhiteSpace(paragraph.Text)) {
            var subtitle = slide.AddTextBoxInches(paragraph.Text.Trim(), metrics.Horizontal(0.76), metrics.Vertical(1.48), metrics.Horizontal(5.4), metrics.Vertical(0.52));
            ApplyTextStyle(subtitle, styleResolver.Resolve("lead"));
        }

        AddSummaryCards(slide, theme, items, metrics);
        return true;
    }

    private static bool TryExportDesignerSectionSlide(
        PowerPointDeckComposer deck,
        OfficeMarkupSlideBlock slideBlock,
        out PowerPointSlide? slide) {
        slide = null;
        var layout = Normalize(slideBlock.Layout ?? string.Empty);
        if (layout != "section" && layout != "title") {
            return false;
        }

        if (HasExplicitBackground(slideBlock) || HasAnyExplicitPlacement(slideBlock.Blocks)) {
            return false;
        }

        var subtitle = ExtractSingleSubtitle(slideBlock.Blocks);
        if (slideBlock.Blocks.Count > 0 && subtitle == null) {
            return false;
        }

        slide = deck.AddSectionSlide(slideBlock.Title ?? "OfficeIMO Markup", subtitle, seed: slideBlock.Title);
        return true;
    }

    private static bool TryExportDesignerProcessSlide(
        PowerPointDeckComposer deck,
        OfficeMarkupSlideBlock slideBlock,
        out PowerPointSlide? slide) {
        slide = null;
        var layout = Normalize(slideBlock.Layout ?? string.Empty);
        if (layout != "process" && layout != "timeline") {
            return false;
        }

        if (HasExplicitBackground(slideBlock) || HasAnyExplicitPlacement(slideBlock.Blocks)) {
            return false;
        }

        var list = slideBlock.Blocks.OfType<OfficeMarkupListBlock>().FirstOrDefault();
        if (list == null || list.Items.Count < 2 || list.Items.Count > 8) {
            return false;
        }

        var subtitle = slideBlock.Blocks.OfType<OfficeMarkupParagraphBlock>().FirstOrDefault()?.Text;
        var steps = list.Items.Select((item, index) => CreateProcessStep(item.Text, index + 1)).ToList();
        slide = deck.AddProcessSlide(slideBlock.Title ?? "Process", subtitle, steps, seed: slideBlock.Title);
        return true;
    }

    private static bool TryExportDesignerCardGridSlide(
        PowerPointDeckComposer deck,
        OfficeMarkupSlideBlock slideBlock,
        out PowerPointSlide? slide) {
        slide = null;
        var layout = Normalize(slideBlock.Layout ?? string.Empty);
        var cards = slideBlock.Blocks.OfType<OfficeMarkupCardBlock>().ToList();
        if (cards.Count == 0 || (layout != "cards" && layout != "cardgrid" && cards.Count < 2)) {
            return false;
        }

        if (HasExplicitBackground(slideBlock) || HasAnyExplicitPlacement(slideBlock.Blocks)) {
            return false;
        }

        var allowed = slideBlock.Blocks.All(block => block is OfficeMarkupCardBlock || block is OfficeMarkupParagraphBlock);
        if (!allowed) {
            return false;
        }

        var subtitle = slideBlock.Blocks.OfType<OfficeMarkupParagraphBlock>().FirstOrDefault()?.Text;
        var designerCards = cards.Select(CreateCardContent).ToList();
        slide = deck.AddCardGridSlide(slideBlock.Title ?? "Highlights", subtitle, designerCards, seed: slideBlock.Title);
        return true;
    }

    private static bool TryExportDesignerTwoColumnSlide(
        PowerPointDeckComposer deck,
        OfficeMarkupSlideBlock slideBlock,
        OfficeMarkupStyleResolver styleResolver,
        out PowerPointSlide? slide) {
        slide = null;
        var layout = Normalize(slideBlock.Layout ?? string.Empty);
        if (layout != "twocolumns" && layout != "twocolumn" && layout != "twocolumnlayout" && layout != "comparison") {
            return false;
        }

        if (HasExplicitBackground(slideBlock) || HasSemanticDesignerPlacementOverrides(slideBlock.Blocks)) {
            return false;
        }

        if (!TryCollectSemanticColumns(slideBlock.Blocks, out var subtitle, out _, out var columns)) {
            return false;
        }

        if (columns.Count != 2 || columns.Any(column => column.Count == 0 || !SupportsDesignerColumnBlocks(column))) {
            return false;
        }

        slide = deck.ComposeSlide(composer => {
            composer.AddTitle(slideBlock.Title ?? "Overview", subtitle);
            var regions = composer.ContentColumns(2, gutterCm: 0.78, topCm: 4.0, bottomMarginCm: 1.05, horizontalMarginCm: 1.55);
            for (var index = 0; index < regions.Length && index < columns.Count; index++) {
                AddDesignerColumnRegion(composer.Slide, deck.Design.Theme, regions[index], columns[index], styleResolver, index);
            }
        }, seed: slideBlock.Title ?? slideBlock.Layout ?? "two-columns", dark: IsDarkBackground(slideBlock.Background, deck.Design.Theme));

        return true;
    }

    private static int ExportColumns(
        PowerPointSlide slide,
        IList<OfficeMarkupBlock> blocks,
        int startIndex,
        MarkupToPowerPointOptions options,
        SlideCanvasMetrics metrics,
        LayoutCursor cursor,
        OfficeMarkupStyleResolver styleResolver) {
        var columnsBlock = blocks[startIndex];
        var columns = new List<List<OfficeMarkupBlock>>();
        var index = startIndex + 1;
        while (index < blocks.Count) {
            var current = blocks[index];
            if (!IsColumn(current)) {
                break;
            }

            var columnBlocks = new List<OfficeMarkupBlock>();
            var body = GetColumnBody(current);
            if (!string.IsNullOrWhiteSpace(body)) {
                columnBlocks.AddRange(ParseLightweightMarkdown(body));
            }

            index++;
            while (index < blocks.Count && !IsColumn(blocks[index])) {
                columnBlocks.Add(blocks[index]);
                index++;
            }

            columns.Add(columnBlocks);
        }

        if (columns.Count == 0) {
            return startIndex;
        }

        var region = ResolveBox(GetPlacement(columnsBlock), columnsBlock.Attributes, cursor, cursor.Height, metrics);
        var gap = ResolveGap(columnsBlock, metrics);
        var columnWidth = (region.Width - (gap * (columns.Count - 1))) / columns.Count;
        for (int columnIndex = 0; columnIndex < columns.Count; columnIndex++) {
            var columnCursor = new LayoutCursor(
                region.Left + (columnIndex * (columnWidth + gap)),
                region.Top,
                columnWidth,
                region.Height);
            foreach (var block in columns[columnIndex]) {
                ExportBlock(slide, block, options, metrics, columnCursor, styleResolver);
            }
        }

        return index - 1;
    }
}
