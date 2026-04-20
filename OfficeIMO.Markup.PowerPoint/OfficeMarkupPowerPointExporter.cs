using System.Diagnostics;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Markup.PowerPoint;

public sealed class OfficeMarkupPowerPointExporter {
    private const double SlideWidth = 10.0;
    private const double SlideHeight = 5.625;

    private const double TitleLeft = 0.55;
    private const double TitleTop = 0.32;
    private const double TitleWidth = 8.9;
    private const double TitleHeight = 0.65;

    private const double BodyLeft = 0.75;
    private const double BodyTop = 1.15;
    private const double BodyWidth = 8.5;
    private const double BodyHeight = 3.85;

    private const double DesignerBodyLeft = 0.72;
    private const double DesignerBodyTop = 1.72;
    private const double DesignerBodyWidth = 8.55;
    private const double DesignerBodyHeight = 2.95;

    private readonly struct SlideCanvasMetrics {
        public SlideCanvasMetrics(double width, double height) {
            Width = width > 0 ? width : SlideWidth;
            Height = height > 0 ? height : SlideHeight;
        }

        public double Width { get; }
        public double Height { get; }
        private double ScaleX => Width / SlideWidth;
        private double ScaleY => Height / SlideHeight;

        public double Horizontal(double value) => value * ScaleX;
        public double Vertical(double value) => value * ScaleY;
    }

    public void Export(OfficeMarkupDocument document, OfficeMarkupPowerPointExportOptions options) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (document.Profile != OfficeMarkupProfile.Presentation) {
            throw new InvalidOperationException("PowerPoint export requires the Presentation OfficeIMO markup profile.");
        }

        if (string.IsNullOrWhiteSpace(options.OutputPath)) {
            throw new InvalidOperationException("PowerPoint export requires an output path.");
        }

        var directory = Path.GetDirectoryName(Path.GetFullPath(options.OutputPath));
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        var styleResolver = OfficeMarkupStyleResolver.Create(document);
        var metrics = new SlideCanvasMetrics(options.SlideWidthInches, options.SlideHeightInches);
        using PowerPointPresentation presentation = PowerPointPresentation.Create(options.OutputPath);
        presentation.SlideSize.SetSizeInches(metrics.Width, metrics.Height);
        var deck = presentation.UseDesigner(CreateDeckDesign(document), applyTheme: true);
        string? activeSection = null;
        foreach (var slideBlock in GetSlides(document)) {
            ExportSlide(presentation, deck, slideBlock, options, metrics, styleResolver);
            var section = NormalizeSectionName(slideBlock.Section);
            if (section != null && !string.Equals(activeSection, section, StringComparison.Ordinal)) {
                presentation.AddSection(section, presentation.Slides.Count - 1);
                activeSection = section;
            }
        }

        presentation.Save();
    }

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
        OfficeMarkupPowerPointExportOptions options,
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
        OfficeMarkupPowerPointExportOptions options,
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

    private static void ExportBlock(
        PowerPointSlide slide,
        OfficeMarkupBlock block,
        OfficeMarkupPowerPointExportOptions options,
        SlideCanvasMetrics metrics,
        LayoutCursor cursor,
        OfficeMarkupStyleResolver styleResolver) {
        switch (block) {
            case OfficeMarkupHeadingBlock heading:
                AddText(slide, heading.Text, cursor, height: heading.Level <= 2 ? 0.45 : 0.34, styleResolver.Resolve(heading));
                break;
            case OfficeMarkupParagraphBlock paragraph:
                AddText(slide, paragraph.Text, cursor, height: EstimateTextHeight(paragraph.Text), styleResolver.Resolve(paragraph));
                break;
            case OfficeMarkupListBlock list:
                AddList(slide, list, cursor, styleResolver.Resolve("body"));
                break;
            case OfficeMarkupImageBlock image:
                AddImage(slide, image, cursor, options, metrics);
                break;
            case OfficeMarkupTableBlock table:
                AddTable(slide, table, cursor);
                break;
            case OfficeMarkupDiagramBlock diagram:
                AddDiagram(slide, diagram, cursor, options, metrics, styleResolver);
                break;
            case OfficeMarkupChartBlock chart:
                AddChart(slide, chart, cursor, options, metrics);
                break;
            case OfficeMarkupTextBoxBlock textBox:
                AddTextBox(slide, textBox, cursor, metrics, styleResolver);
                break;
            case OfficeMarkupCardBlock card:
                AddCard(slide, card, cursor, metrics, styleResolver);
                break;
            case OfficeMarkupColumnsBlock:
            case OfficeMarkupColumnBlock:
                break;
            case OfficeMarkupExtensionBlock extension:
                ExportExtension(slide, extension, options, metrics, cursor, styleResolver);
                break;
            default:
                if (options.IncludeUnsupportedBlocksAsText) {
                    AddText(slide, block.Kind.ToString(), cursor, height: 0.4, styleResolver.Resolve("caption"));
                }

                break;
        }
    }

    private static void ExportExtension(
        PowerPointSlide slide,
        OfficeMarkupExtensionBlock extension,
        OfficeMarkupPowerPointExportOptions options,
        SlideCanvasMetrics metrics,
        LayoutCursor cursor,
        OfficeMarkupStyleResolver styleResolver) {
        switch (Normalize(extension.Command)) {
            case "textbox":
                AddTextBox(slide, extension.Body, null, extension.Attributes, cursor, metrics, styleResolver.Resolve(extension));
                break;
            case "card":
                AddCard(slide, GetAttribute(extension.Attributes, "title"), extension.Body, null, extension.Attributes, cursor, metrics, styleResolver.Resolve(extension));
                break;
            case "column":
            case "left":
            case "right":
            case "columns":
                break;
            default:
                if (options.IncludeUnsupportedBlocksAsText && !string.IsNullOrWhiteSpace(extension.Body)) {
                    AddText(slide, extension.Body.Trim(), cursor, height: EstimateTextHeight(extension.Body), styleResolver.Resolve(extension));
                }

                break;
        }
    }

    private static void AddText(PowerPointSlide slide, string text, LayoutCursor cursor, double height, OfficeMarkupResolvedStyle? style = null) {
        if (string.IsNullOrWhiteSpace(text)) {
            return;
        }

        var actualHeight = Math.Max(0.28, Math.Min(height, cursor.RemainingHeight));
        var textBox = slide.AddTextBoxInches(text.Trim(), cursor.Left, cursor.Top, cursor.Width, actualHeight);
        ApplyTextStyle(textBox, style);
        cursor.Advance(actualHeight);
    }

    private static void AddTextBox(PowerPointSlide slide, OfficeMarkupTextBoxBlock textBox, LayoutCursor cursor, SlideCanvasMetrics metrics, OfficeMarkupStyleResolver styleResolver) =>
        AddTextBox(slide, textBox.Text, textBox.Placement, textBox.Attributes, cursor, metrics, styleResolver.Resolve(textBox));

    private static void AddTextBox(
        PowerPointSlide slide,
        string text,
        OfficeMarkupPlacement? placement,
        IDictionary<string, string> attributes,
        LayoutCursor cursor,
        SlideCanvasMetrics metrics,
        OfficeMarkupResolvedStyle? style) {
        var box = ResolveBox(placement, attributes, cursor, 0.62, metrics);
        var textBox = slide.AddTextBoxInches(text.Trim(), box.Left, box.Top, box.Width, box.Height);
        ApplyTextStyle(textBox, style);
        if (!HasExplicitPlacement(placement, attributes)) {
            cursor.Advance(box.Height);
        }
    }

    private static void AddCard(PowerPointSlide slide, OfficeMarkupCardBlock card, LayoutCursor cursor, SlideCanvasMetrics metrics, OfficeMarkupStyleResolver styleResolver) =>
        AddCard(slide, card.Title, card.Body, card.Placement, card.Attributes, cursor, metrics, styleResolver.Resolve(card));

    private static void AddCard(
        PowerPointSlide slide,
        string? title,
        string body,
        OfficeMarkupPlacement? placement,
        IDictionary<string, string> attributes,
        LayoutCursor cursor,
        SlideCanvasMetrics metrics,
        OfficeMarkupResolvedStyle? style) {
        var text = string.IsNullOrWhiteSpace(title)
            ? body.Trim()
            : title!.Trim() + Environment.NewLine + body.Trim();
        if (string.IsNullOrWhiteSpace(text)) {
            return;
        }

        var box = ResolveBox(placement, attributes, cursor, Math.Min(1.25, cursor.RemainingHeight), metrics);
        AddPanel(slide, box, style, "OfficeIMO Markup Card Panel");
        var textBox = slide.AddTextBoxInches(text, box.Left, box.Top, box.Width, box.Height);
        ApplyTextStyle(textBox, style);
        if (!HasExplicitPlacement(placement, attributes)) {
            cursor.Advance(box.Height);
        }
    }

    private static void AddList(PowerPointSlide slide, OfficeMarkupListBlock list, LayoutCursor cursor, OfficeMarkupResolvedStyle? style) {
        var items = list.Items.Select(item => item.Text).Where(text => !string.IsNullOrWhiteSpace(text)).ToList();
        if (items.Count == 0) {
            return;
        }

        var height = Math.Max(0.45, Math.Min(cursor.RemainingHeight, 0.28 * items.Count + 0.25));
        var box = slide.AddTextBoxInches(string.Empty, cursor.Left, cursor.Top, cursor.Width, height);
        box.Clear();
        if (list.Ordered) {
            box.SetNumberedList(items, list.Start);
        } else {
            box.SetBullets(items);
        }

        ApplyTextStyle(box, style);
        cursor.Advance(height);
    }

    private static void AddDiagram(
        PowerPointSlide slide,
        OfficeMarkupDiagramBlock diagram,
        LayoutCursor cursor,
        OfficeMarkupPowerPointExportOptions options,
        SlideCanvasMetrics metrics,
        OfficeMarkupStyleResolver styleResolver) {
        var box = ResolveBox(diagram.Placement, diagram.Attributes, cursor, Math.Min(2.4, cursor.RemainingHeight), metrics);
        if (ShouldAddVisualPanel(diagram.Attributes, defaultValue: true)) {
            AddVisualPanel(slide, box, metrics, "OfficeIMO Markup Diagram Panel");
        }

        if (OfficeMarkupMermaidRenderer.TryRenderPng(diagram, options, out var imagePath)) {
            try {
                AddPicture(slide, imagePath, box, GetAttribute(diagram.Attributes, "fit"));
                if (!HasExplicitPlacement(diagram.Placement, diagram.Attributes)) {
                    cursor.Advance(box.Height);
                }

                return;
            } finally {
                TryDelete(imagePath);
            }
        }

        if (options.IncludeUnsupportedBlocksAsText) {
            var text = IsMermaid(diagram.Language)
                ? "Mermaid diagram\nInstall or configure the Mermaid renderer to export this block as an image."
                : $"{diagram.Language} diagram";
            var textBox = slide.AddTextBoxInches(text.Trim(), box.Left, box.Top, box.Width, box.Height);
            ApplyTextStyle(textBox, styleResolver.Resolve("caption"));
            if (!HasExplicitPlacement(diagram.Placement, diagram.Attributes)) {
                cursor.Advance(box.Height);
            }
        }
    }

    private static bool ShouldRenderSlideTitle(OfficeMarkupSlideBlock slideBlock) {
        if (slideBlock == null || string.IsNullOrWhiteSpace(slideBlock.Title)) {
            return false;
        }

        // Blank slides are explicit-placement canvases; keep the semantic title in the AST
        // without drawing an extra title that collides with authored textboxes/charts.
        return !string.Equals(Normalize(slideBlock.Layout ?? string.Empty), "blank", StringComparison.Ordinal);
    }

    private static bool ShouldUseDesignerCanvas(OfficeMarkupSlideBlock slideBlock) {
        if (slideBlock == null || HasExplicitBackground(slideBlock)) {
            return false;
        }

        var layout = Normalize(slideBlock.Layout ?? string.Empty);
        return layout.Length == 0
            || layout == "titleandcontent"
            || layout == "content"
            || layout == "twocolumns";
    }

    private static PowerPointDeckDesign CreateDeckDesign(OfficeMarkupDocument document) {
        var accent = ToPowerPointColor(GetMetadata(document, "accent")
            ?? GetMetadata(document, "accent-color")
            ?? GetMetadata(document, "brand-color")
            ?? DefaultAccentForTheme(GetMetadata(document, "theme")))
            ?? "2563EB";
        var seed = GetMetadata(document, "title")
            ?? GetMetadata(document, "name")
            ?? "OfficeIMO Markup";
        var themeName = GetMetadata(document, "theme") ?? "OfficeIMO Markup";
        var footerLeft = GetMetadata(document, "footer-left") ?? "OfficeIMO Markup";
        var footerRight = GetMetadata(document, "footer-right") ?? GetMetadata(document, "author");
        var eyebrow = GetMetadata(document, "eyebrow");

        return PowerPointDeckDesign.FromBrand(
            accent,
            seed,
            ParseDesignMood(GetMetadata(document, "mood") ?? GetMetadata(document, "design-mood")),
            name: themeName,
            eyebrow: eyebrow,
            footerLeft: footerLeft,
            footerRight: footerRight);
    }

    private static PowerPointDesignMood ParseDesignMood(string? value) {
        switch (Normalize(value ?? string.Empty)) {
            case "editorial":
                return PowerPointDesignMood.Editorial;
            case "energetic":
            case "energy":
                return PowerPointDesignMood.Energetic;
            case "minimal":
            case "minimalist":
                return PowerPointDesignMood.Minimal;
            default:
                return PowerPointDesignMood.Corporate;
        }
    }

    private static string DefaultAccentForTheme(string? themeName) {
        switch (Normalize(themeName ?? string.Empty)) {
            case "evotecmodern":
                return "2563EB";
            case "modernblue":
                return "0098C8";
            default:
                return "2563EB";
        }
    }

    private static string? GetMetadata(OfficeMarkupDocument document, string name) =>
        document.Metadata.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value.Trim()
            : null;

    private static bool HasExplicitBackground(OfficeMarkupSlideBlock slideBlock) =>
        !string.IsNullOrWhiteSpace(slideBlock.Background);

    private static bool IsDarkBackground(string? background, PowerPointDesignTheme theme) {
        var color = ParseBackgroundColor(background, theme);
        if (string.IsNullOrWhiteSpace(color)) {
            return false;
        }

        var red = Convert.ToInt32(color!.Substring(0, 2), 16);
        var green = Convert.ToInt32(color.Substring(2, 2), 16);
        var blue = Convert.ToInt32(color.Substring(4, 2), 16);
        return ((red * 299) + (green * 587) + (blue * 114)) / 1000 < 128;
    }

    private static bool HasAnyExplicitPlacement(IEnumerable<OfficeMarkupBlock> blocks) =>
        blocks.Any(block => HasExplicitPlacement(block));

    private static bool TryCollectSemanticColumns(
        IList<OfficeMarkupBlock> blocks,
        out string? subtitle,
        out OfficeMarkupBlock? columnsBlock,
        out List<List<OfficeMarkupBlock>> columns) {
        subtitle = null;
        columnsBlock = null;
        columns = new List<List<OfficeMarkupBlock>>();

        if (blocks.Count == 0) {
            return false;
        }

        var index = 0;
        if (blocks[0] is OfficeMarkupParagraphBlock paragraph && !string.IsNullOrWhiteSpace(paragraph.Text)) {
            subtitle = paragraph.Text.Trim();
            index++;
        }

        if (index >= blocks.Count || !IsColumns(blocks[index])) {
            return false;
        }

        columnsBlock = blocks[index];
        index++;
        while (index < blocks.Count) {
            var current = blocks[index];
            if (!IsColumn(current)) {
                return false;
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

        return columns.Count > 0;
    }

    private static string? ExtractSingleSubtitle(IList<OfficeMarkupBlock> blocks) {
        if (blocks.Count == 0) {
            return null;
        }

        return blocks.Count == 1 && blocks[0] is OfficeMarkupParagraphBlock paragraph
            ? paragraph.Text
            : null;
    }

    private static PowerPointCardContent CreateCardContent(OfficeMarkupCardBlock card) {
        var title = string.IsNullOrWhiteSpace(card.Title) ? FirstNonEmptyLine(card.Body) ?? "Card" : card.Title!;
        var items = ExtractListLines(card.Body, title).ToList();
        return new PowerPointCardContent(title, items, ToPowerPointColor(GetAttribute(card.Attributes, "accent")));
    }

    private static PowerPointProcessStep CreateProcessStep(string text, int number) {
        var value = (text ?? string.Empty).Trim();
        var separator = value.IndexOf(':');
        if (separator > 0 && separator < value.Length - 1) {
            return new PowerPointProcessStep(
                value.Substring(0, separator).Trim(),
                value.Substring(separator + 1).Trim(),
                number.ToString(CultureInfo.InvariantCulture));
        }

        return new PowerPointProcessStep(value, " ", number.ToString(CultureInfo.InvariantCulture));
    }

    private static string? FirstNonEmptyLine(string value) =>
        value.Replace("\r\n", "\n").Replace('\r', '\n')
            .Split('\n')
            .Select(line => line.Trim().TrimStart('-', '*').Trim())
            .FirstOrDefault(line => !string.IsNullOrWhiteSpace(line));

    private static IEnumerable<string> ExtractListLines(string value, string title) {
        foreach (var rawLine in value.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')) {
            var line = rawLine.Trim();
            if (line.StartsWith("- ", StringComparison.Ordinal) || line.StartsWith("* ", StringComparison.Ordinal)) {
                yield return line.Substring(2).Trim();
            } else if (!string.IsNullOrWhiteSpace(line) && !string.Equals(line, title, StringComparison.OrdinalIgnoreCase)) {
                yield return line;
            }
        }
    }

    private static void AddMarkupCanvas(
        PowerPointSlide slide,
        PowerPointDesignTheme theme,
        OfficeMarkupSlideBlock slideBlock,
        SlideCanvasMetrics metrics,
        bool strong = false) {
        if (!HasExplicitBackground(slideBlock)) {
            slide.BackgroundColor = strong ? theme.SurfaceColor : theme.BackgroundColor;
        }

        var wash = slide.AddShapeInches(
            A.ShapeTypeValues.Parallelogram,
            metrics.Horizontal(strong ? 6.95 : 8.18),
            metrics.Vertical(-0.12),
            metrics.Horizontal(strong ? 2.95 : 1.96),
            metrics.Vertical(SlideHeight + 0.34),
            "OfficeIMO Markup Canvas Wash");
        wash.FillColor = theme.AccentLightColor;
        wash.FillTransparency = strong ? 8 : 24;
        wash.OutlineColor = theme.AccentLightColor;
        wash.OutlineWidthPoints = 0;

        var rail = slide.AddShapeInches(
            A.ShapeTypeValues.Rectangle,
            metrics.Horizontal(strong ? 9.72 : 9.82),
            0,
            metrics.Horizontal(strong ? 0.12 : 0.08),
            metrics.Height,
            "OfficeIMO Markup Canvas Rail");
        rail.FillColor = theme.AccentColor;
        rail.FillTransparency = strong ? 0 : 16;
        rail.OutlineColor = theme.AccentColor;
        rail.OutlineWidthPoints = 0;

        var rule = slide.AddShapeInches(
            A.ShapeTypeValues.Rectangle,
            metrics.Horizontal(0.72),
            metrics.Vertical(strong ? 1.78 : 0.98),
            metrics.Horizontal(strong ? 1.2 : 0.82),
            metrics.Vertical(0.035),
            "OfficeIMO Markup Canvas Rule");
        rule.FillColor = theme.AccentColor;
        rule.OutlineColor = theme.AccentColor;
        rule.OutlineWidthPoints = 0;
    }

    private static void AddSummaryCards(PowerPointSlide slide, PowerPointDesignTheme theme, IReadOnlyList<string> items, SlideCanvasMetrics metrics) {
        var count = Math.Min(4, items.Count);
        var gap = metrics.Horizontal(0.22);
        var left = metrics.Horizontal(0.72);
        var top = metrics.Vertical(2.28);
        var totalWidth = metrics.Horizontal(8.28);
        var cardWidth = (totalWidth - (gap * (count - 1))) / count;
        var cardHeight = metrics.Vertical(1.62);

        for (var index = 0; index < count; index++) {
            var x = left + index * (cardWidth + gap);
            var card = slide.AddShapeInches(
                A.ShapeTypeValues.Rectangle,
                x,
                top,
                cardWidth,
                cardHeight,
                "OfficeIMO Markup Summary Card");
            card.FillColor = theme.PanelColor;
            card.OutlineColor = theme.PanelBorderColor;
            card.OutlineWidthPoints = 0.75;

            var accent = slide.AddShapeInches(
                A.ShapeTypeValues.Rectangle,
                x,
                top,
                cardWidth,
                metrics.Vertical(0.08),
                "OfficeIMO Markup Summary Card Accent");
            accent.FillColor = index % 2 == 0 ? theme.AccentColor : theme.WarningColor;
            accent.OutlineColor = accent.FillColor;
            accent.OutlineWidthPoints = 0;

            var number = slide.AddTextBoxInches((index + 1).ToString("00", CultureInfo.InvariantCulture), x + metrics.Horizontal(0.18), top + metrics.Vertical(0.25), metrics.Horizontal(0.52), metrics.Vertical(0.3));
            ApplyTextStyle(number, new OfficeMarkupResolvedStyle {
                FontName = theme.HeadingFontName,
                FontSize = 10,
                Bold = true,
                TextColor = theme.AccentColor
            });

            var text = slide.AddTextBoxInches(items[index], x + metrics.Horizontal(0.18), top + metrics.Vertical(0.65), cardWidth - metrics.Horizontal(0.36), metrics.Vertical(0.68));
            ApplyTextStyle(text, new OfficeMarkupResolvedStyle {
                FontName = theme.BodyFontName,
                FontSize = 14,
                Bold = true,
                TextColor = theme.PrimaryTextColor
            });
        }
    }

    private static bool SupportsDesignerColumnBlocks(IEnumerable<OfficeMarkupBlock> blocks) =>
        blocks.All(block =>
            block is OfficeMarkupHeadingBlock
            || block is OfficeMarkupParagraphBlock
            || block is OfficeMarkupListBlock);

    private static bool HasSemanticDesignerPlacementOverrides(IEnumerable<OfficeMarkupBlock> blocks) =>
        blocks.Any(block => HasSemanticDesignerPlacementOverride(block));

    private static bool HasSemanticDesignerPlacementOverride(OfficeMarkupBlock block) {
        if (IsColumn(block)) {
            return HasSemanticDesignerColumnPlacementOverride(GetPlacement(block), block.Attributes);
        }

        return HasExplicitPlacement(block);
    }

    private static bool HasSemanticDesignerColumnPlacementOverride(OfficeMarkupPlacement? placement, IDictionary<string, string> attributes) =>
        placement?.HasValue == true
        || attributes.ContainsKey("x")
        || attributes.ContainsKey("y")
        || attributes.ContainsKey("h")
        || attributes.ContainsKey("height");

    private static void AddDesignerColumnRegion(
        PowerPointSlide slide,
        PowerPointDesignTheme theme,
        PowerPointLayoutBox bounds,
        IReadOnlyList<OfficeMarkupBlock> blocks,
        OfficeMarkupStyleResolver styleResolver,
        int columnIndex) {
        var textOnlyOptions = new OfficeMarkupPowerPointExportOptions {
            IncludeUnsupportedBlocksAsText = false,
            RenderMermaidDiagrams = false
        };

        var panel = slide.AddRectangleCm(bounds.LeftCm, bounds.TopCm, bounds.WidthCm, bounds.HeightCm,
            "OfficeIMO Markup Semantic Column Panel");
        panel.FillColor = theme.PanelColor;
        panel.OutlineColor = theme.PanelBorderColor;
        panel.OutlineWidthPoints = 0.85;

        var accent = slide.AddRectangleCm(bounds.LeftCm, bounds.TopCm, bounds.WidthCm, 0.14,
            "OfficeIMO Markup Semantic Column Accent");
        accent.FillColor = columnIndex % 2 == 0 ? theme.AccentColor : theme.Accent2Color;
        accent.OutlineColor = accent.FillColor;
        accent.OutlineWidthPoints = 0;

        var cursor = new LayoutCursor(
            bounds.LeftInches + 0.18,
            bounds.TopInches + 0.28,
            Math.Max(0.6, bounds.WidthInches - 0.36),
            Math.Max(0.8, bounds.HeightInches - 0.42));

        var startIndex = 0;
        if (blocks.Count > 0 && blocks[0] is OfficeMarkupHeadingBlock heading) {
            var headingBox = slide.AddTextBoxInches(heading.Text, cursor.Left, cursor.Top, cursor.Width, 0.42);
            ApplyTextStyle(headingBox, new OfficeMarkupResolvedStyle {
                FontName = theme.HeadingFontName,
                FontSize = 18,
                Bold = true,
                TextColor = theme.PrimaryTextColor
            });

            var rule = slide.AddShapeInches(
                A.ShapeTypeValues.Rectangle,
                cursor.Left,
                cursor.Top + 0.47,
                Math.Min(1.2, cursor.Width),
                0.03,
                "OfficeIMO Markup Semantic Column Rule");
            rule.FillColor = theme.AccentColor;
            rule.OutlineColor = theme.AccentColor;
            rule.OutlineWidthPoints = 0;

            cursor.Advance(0.58);
            startIndex = 1;
        }

        for (var index = startIndex; index < blocks.Count; index++) {
            ExportBlock(slide, blocks[index], textOnlyOptions, new SlideCanvasMetrics(SlideWidth, SlideHeight), cursor, styleResolver);
        }
    }

    private static void ApplyBackground(PowerPointSlide slide, string? background, PowerPointDesignTheme theme, OfficeMarkupPowerPointExportOptions options, SlideCanvasMetrics metrics) {
        var spec = ParseBackground(background, theme, options);
        if (!string.IsNullOrWhiteSpace(spec.GradientStartColor) && !string.IsNullOrWhiteSpace(spec.GradientEndColor)) {
            slide.SetBackgroundGradient(
                spec.GradientStartColor!,
                spec.GradientEndColor!,
                spec.GradientAngleDegrees ?? 135d);
        } else if (!string.IsNullOrWhiteSpace(spec.Color)) {
            slide.BackgroundColor = spec.Color;
        }

        if (!string.IsNullOrWhiteSpace(spec.ImagePath)) {
            slide.SetBackgroundImage(spec.ImagePath!);
        }

        if (!string.IsNullOrWhiteSpace(spec.OverlayColor)) {
            var overlay = slide.AddShapeInches(
                A.ShapeTypeValues.Rectangle,
                0,
                0,
                metrics.Width,
                metrics.Height,
                "OfficeIMO Markup Background Overlay");
            overlay.FillColor = spec.OverlayColor;
            overlay.FillTransparency = spec.OverlayTransparency;
            overlay.OutlineColor = spec.OverlayColor;
            overlay.OutlineWidthPoints = 0;
        }
    }

    private static string? ParseBackgroundColor(string? background, PowerPointDesignTheme? theme) {
        var spec = ParseBackground(background, theme, null);
        return spec.Color ?? spec.GradientStartColor;
    }

    private static OfficeMarkupBackgroundSpec ParseBackground(string? background, PowerPointDesignTheme? theme, OfficeMarkupPowerPointExportOptions? options) {
        if (string.IsNullOrWhiteSpace(background)) {
            return new OfficeMarkupBackgroundSpec();
        }

        var value = background!.Trim();
        var solid = TryExtractFunctionArgument(value, "solid");
        var gradient = TryExtractFunctionArgument(value, "gradient");
        var image = TryExtractFunctionArgument(value, "image");
        var angle = TryExtractNamedValue(value, "angle");
        var overlay = TryExtractNamedValue(value, "overlay");

        string? resolvedImage = null;
        if (!string.IsNullOrWhiteSpace(image)) {
            var candidate = image!.Trim().Trim('"', '\'');
            var resolved = ResolvePath(options, candidate);
            if (File.Exists(resolved)) {
                resolvedImage = resolved;
            }
        }

        TryParseGradient(gradient, theme, out var gradientStartColor, out var gradientEndColor);
        var spec = new OfficeMarkupBackgroundSpec {
            Color = !string.IsNullOrWhiteSpace(solid) ? ResolveThemeColor(solid, theme) : ResolveThemeColor(value, theme),
            GradientStartColor = gradientStartColor,
            GradientEndColor = gradientEndColor,
            GradientAngleDegrees = TryParseGradientAngle(angle, out var gradientAngleDegrees) ? gradientAngleDegrees : null,
            ImagePath = resolvedImage
        };

        if (TryParseOverlay(overlay, out var overlayColor, out var overlayTransparency)) {
            spec.OverlayColor = overlayColor;
            spec.OverlayTransparency = overlayTransparency;
        }

        return spec;
    }

    private static string ResolvePath(OfficeMarkupPowerPointExportOptions? options, string source) {
        if (Path.IsPathRooted(source) || options == null || string.IsNullOrWhiteSpace(options.BaseDirectory)) {
            return source;
        }

        return Path.Combine(options.BaseDirectory!, source);
    }

    private static string? TryExtractFunctionArgument(string value, string functionName) {
        var prefix = functionName + "(";
        var start = value.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
        if (start < 0) {
            return null;
        }

        start += prefix.Length;
        var end = value.IndexOf(')', start);
        if (end < 0) {
            return null;
        }

        return value.Substring(start, end - start).Trim();
    }

    private static string? TryExtractNamedValue(string value, string attributeName) {
        var start = value.IndexOf(attributeName + "=", StringComparison.OrdinalIgnoreCase);
        if (start < 0) {
            return null;
        }

        start += attributeName.Length + 1;
        if (start >= value.Length) {
            return null;
        }

        var remaining = value.Substring(start);
        if (remaining.StartsWith("rgba(", StringComparison.OrdinalIgnoreCase)) {
            var end = value.IndexOf(')', start);
            return end >= start ? value.Substring(start, end - start + 1).Trim() : null;
        }

        var nextSpace = value.IndexOf(' ', start);
        return (nextSpace >= 0 ? value.Substring(start, nextSpace - start) : value.Substring(start)).Trim();
    }

    private static bool TryParseOverlay(string? value, out string color, out int transparency) {
        color = string.Empty;
        transparency = 0;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        var normalized = value!.Trim();
        if (normalized.StartsWith("rgba(", StringComparison.OrdinalIgnoreCase) && normalized.EndsWith(")", StringComparison.Ordinal)) {
            var parts = normalized.Substring(5, normalized.Length - 6).Split(',');
            if (parts.Length == 4
                && int.TryParse(parts[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var red)
                && int.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var green)
                && int.TryParse(parts[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var blue)
                && double.TryParse(parts[3].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var alpha)) {
                red = Math.Max(0, Math.Min(255, red));
                green = Math.Max(0, Math.Min(255, green));
                blue = Math.Max(0, Math.Min(255, blue));
                alpha = Math.Max(0, Math.Min(1, alpha));
                color = $"{red:X2}{green:X2}{blue:X2}";
                transparency = (int)Math.Round((1 - alpha) * 100);
                return true;
            }
        }

        var hex = ToPowerPointColor(normalized);
        if (!string.IsNullOrWhiteSpace(hex)) {
            color = hex!;
            transparency = 0;
            return true;
        }

        return false;
    }

    private static bool TryParseGradient(
        string? value,
        PowerPointDesignTheme? theme,
        out string? startColor,
        out string? endColor) {
        startColor = null;
        endColor = null;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        var parts = value!.Split(',')
            .Select(part => ResolveThemeColor(part.Trim(), theme))
            .Where(color => !string.IsNullOrWhiteSpace(color))
            .Cast<string>()
            .ToList();
        if (parts.Count < 2) {
            return false;
        }

        startColor = parts[0];
        endColor = parts[1];
        return true;
    }

    private static bool TryParseGradientAngle(string? value, out double angleDegrees) {
        angleDegrees = 0;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        var normalized = value!.Trim();
        if (normalized.EndsWith("deg", StringComparison.OrdinalIgnoreCase)) {
            normalized = normalized.Substring(0, normalized.Length - 3).Trim();
        }

        return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out angleDegrees);
    }

    private static string? ResolveThemeColor(string? value, PowerPointDesignTheme? theme) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        var hex = ToPowerPointColor(value);
        if (!string.IsNullOrWhiteSpace(hex)) {
            return hex;
        }

        if (theme == null) {
            return null;
        }

        switch (Normalize(value!)) {
            case "primary":
                return theme.AccentDarkColor;
            case "accent":
            case "accent1":
            case "brand":
                return theme.AccentColor;
            case "accentdark":
                return theme.AccentDarkColor;
            case "accentlight":
                return theme.AccentLightColor;
            case "accent2":
            case "secondary":
                return theme.Accent2Color;
            case "accent3":
            case "tertiary":
                return theme.Accent3Color;
            case "warning":
                return theme.WarningColor;
            case "background":
            case "background1":
            case "bg1":
                return theme.BackgroundColor;
            case "surface":
            case "background2":
            case "bg2":
                return theme.SurfaceColor;
            case "panel":
                return theme.PanelColor;
            case "panelborder":
            case "border":
                return theme.PanelBorderColor;
            case "text":
            case "text1":
            case "foreground":
                return theme.PrimaryTextColor;
            case "text2":
            case "secondarytext":
                return theme.SecondaryTextColor;
            case "muted":
            case "mutedtext":
                return theme.MutedTextColor;
            case "white":
                return "FFFFFF";
            case "black":
                return "000000";
            default:
                return null;
        }
    }

    private sealed class OfficeMarkupBackgroundSpec {
        public string? Color { get; set; }
        public string? GradientStartColor { get; set; }
        public string? GradientEndColor { get; set; }
        public double? GradientAngleDegrees { get; set; }
        public string? ImagePath { get; set; }
        public string? OverlayColor { get; set; }
        public int? OverlayTransparency { get; set; }
    }

    private static bool IsMermaid(string language) =>
        string.Equals(language, "mermaid", StringComparison.OrdinalIgnoreCase);

    private static void ApplyTextStyle(PowerPointTextBox textBox, OfficeMarkupResolvedStyle? style) {
        textBox.SetTextMarginsInches(0.08, 0.04, 0.08, 0.04);

        if (style == null) {
            return;
        }

        if (!string.IsNullOrWhiteSpace(style.FontName)) {
            textBox.FontName = style.FontName;
        }

        if (style.FontSize != null) {
            textBox.FontSize = style.FontSize;
        }

        if (style.Bold != null) {
            textBox.Bold = style.Bold.Value;
        }

        if (style.Italic != null) {
            textBox.Italic = style.Italic.Value;
        }

        var textColor = ToPowerPointColor(style.TextColor);
        if (!string.IsNullOrWhiteSpace(textColor)) {
            textBox.Color = textColor;
        }

        var fillColor = ToPowerPointColor(style.FillColor);
        if (!string.IsNullOrWhiteSpace(fillColor)) {
            textBox.FillColor = fillColor;
        }

        var borderColor = ToPowerPointColor(style.BorderColor);
        if (!string.IsNullOrWhiteSpace(borderColor)) {
            textBox.OutlineColor = borderColor;
            textBox.OutlineWidthPoints = 0.75;
        }

        textBox.SetTextAutoFit(
            PowerPointTextAutoFit.Normal,
            new PowerPointTextAutoFitOptions(fontScalePercent: 82, lineSpaceReductionPercent: 18));
    }

    private static void AddPanel(PowerPointSlide slide, LayoutCursor box, OfficeMarkupResolvedStyle? style, string name) {
        var fillColor = ToPowerPointColor(style?.FillColor);
        var borderColor = ToPowerPointColor(style?.BorderColor);
        if (string.IsNullOrWhiteSpace(fillColor) && string.IsNullOrWhiteSpace(borderColor)) {
            return;
        }

        var panel = slide.AddShapeInches(A.ShapeTypeValues.Rectangle, box.Left, box.Top, box.Width, box.Height, name);
        if (!string.IsNullOrWhiteSpace(fillColor)) {
            panel.FillColor = fillColor;
        }

        if (!string.IsNullOrWhiteSpace(borderColor)) {
            panel.OutlineColor = borderColor;
            panel.OutlineWidthPoints = 0.75;
        }
    }

    private static bool ShouldAddChartPanel(OfficeMarkupChartBlock chart) =>
        !chart.Attributes.TryGetValue("panel", out var value) || !TryParseBool(value, out var parsed) || parsed;

    private static void AddChartPanel(PowerPointSlide slide, LayoutCursor box, SlideCanvasMetrics metrics) {
        const double padding = 0.12;
        var left = Math.Max(metrics.Horizontal(0.25), box.Left - metrics.Horizontal(padding));
        var top = Math.Max(metrics.Vertical(0.25), box.Top - metrics.Vertical(padding));
        var right = Math.Min(metrics.Width - metrics.Horizontal(0.25), box.Left + box.Width + metrics.Horizontal(padding));
        var bottom = Math.Min(metrics.Height - metrics.Vertical(0.25), box.Top + box.Height + metrics.Vertical(padding));

        var panel = slide.AddShapeInches(
            A.ShapeTypeValues.Rectangle,
            left,
            top,
            Math.Max(0.5, right - left),
            Math.Max(0.5, bottom - top),
            "OfficeIMO Markup Chart Panel");
        panel.FillColor = "F8FAFC";
        panel.OutlineColor = "D9E2EF";
        panel.OutlineWidthPoints = 0.75;
    }

    private static void ApplyChartStyle(PowerPointChart chart, OfficeMarkupChartBlock source, PowerPointChartData data) {
        var font = GetAttribute(source.Attributes, "font") ?? "Aptos";
        var textColor = ToPowerPointColor(GetAttribute(source.Attributes, "color")) ?? "172033";
        var gridColor = ToPowerPointColor(GetAttribute(source.Attributes, "grid-color")) ?? "E5E7EB";
        var borderColor = ToPowerPointColor(GetAttribute(source.Attributes, "border")) ?? "D9E2EF";
        var seriesColors = ResolveChartPalette(source);
        var normalizedType = Normalize(source.ChartType);

        chart.SetTitleTextStyle(fontSizePoints: 14, bold: true, color: textColor, fontName: font);
        chart.SetLegend(C.LegendPositionValues.Bottom, overlay: false);
        chart.SetLegendTextStyle(fontSizePoints: 9, color: "4B5563", fontName: font);
        chart.SetChartAreaStyle(fillColor: "FFFFFF", lineColor: borderColor, lineWidthPoints: 0.5);
        chart.SetPlotAreaStyle(fillColor: "FFFFFF", lineColor: "EEF2F7", lineWidthPoints: 0.5);

        for (var index = 0; index < data.Series.Count; index++) {
            var color = seriesColors[index % seriesColors.Count];
            if (normalizedType == "line") {
                chart.SetSeriesLineColor(index, color, widthPoints: 2.25);
            } else {
                chart.SetSeriesFillColor(index, color);
                chart.SetSeriesLineColor(index, color, widthPoints: 0.5);
            }
        }

        if (normalizedType == "pie" || normalizedType == "donut" || normalizedType == "doughnut") {
            chart.SetDataLabels(showValue: true, showCategoryName: false);
            chart.SetDataLabelTextStyle(fontSizePoints: 9, color: textColor, fontName: font);
            ApplyChartSemanticOptions(chart, source, normalizedType, font, textColor, gridColor);
            return;
        }

        chart.SetCategoryAxisLabelTextStyle(fontSizePoints: 9, color: "4B5563", fontName: font);
        chart.SetValueAxisLabelTextStyle(fontSizePoints: 9, color: "4B5563", fontName: font);
        chart.SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: gridColor, lineWidthPoints: 0.5);
        chart.ClearCategoryAxisGridlines();
        ApplyChartSemanticOptions(chart, source, normalizedType, font, textColor, gridColor);
    }

    private static void ApplyChartSemanticOptions(PowerPointChart chart, OfficeMarkupChartBlock source, string normalizedType, string font, string textColor, string gridColor) {
        if (GetAttribute(source.Attributes, "category-title", "categoryTitle", "x-title", "xTitle", "x-axis-title", "xAxisTitle") is { Length: > 0 } categoryTitle) {
            chart.SetCategoryAxisTitle(categoryTitle);
            chart.SetCategoryAxisTitleTextStyle(fontSizePoints: 10, bold: true, color: textColor, fontName: font);
        }

        if (GetAttribute(source.Attributes, "value-title", "valueTitle", "y-title", "yTitle", "y-axis-title", "yAxisTitle") is { Length: > 0 } valueTitle) {
            chart.SetValueAxisTitle(valueTitle);
            chart.SetValueAxisTitleTextStyle(fontSizePoints: 10, bold: true, color: textColor, fontName: font);
        }

        if (GetAttribute(source.Attributes, "category-format", "categoryFormat", "x-format", "xFormat", "category-number-format", "categoryNumberFormat") is { Length: > 0 } categoryFormat) {
            chart.SetCategoryAxisNumberFormat(categoryFormat);
        }

        if (GetAttribute(source.Attributes, "value-format", "valueFormat", "y-format", "yFormat", "value-number-format", "valueNumberFormat") is { Length: > 0 } valueFormat) {
            chart.SetValueAxisNumberFormat(valueFormat);
        }

        ApplyLegendOptions(chart, source);
        ApplyDataLabelOptions(chart, source, normalizedType, font, textColor);
        ApplyGridlineOptions(chart, source, normalizedType, gridColor);
    }

    private static void ApplyLegendOptions(PowerPointChart chart, OfficeMarkupChartBlock source) {
        var legend = GetAttribute(source.Attributes, "legend", "legend-position", "legendPosition");
        if (string.IsNullOrWhiteSpace(legend)) {
            return;
        }

        var normalized = Normalize(legend!);
        if (normalized is "false" or "none" or "hidden" or "off") {
            chart.HideLegend();
            return;
        }

        if (TryParseLegendPosition(legend!, out var position)) {
            chart.SetLegend(position);
        }
    }

    private static void ApplyDataLabelOptions(PowerPointChart chart, OfficeMarkupChartBlock source, string normalizedType, string font, string textColor) {
        var labels = GetAttribute(source.Attributes, "labels", "data-labels", "dataLabels");
        if (string.IsNullOrWhiteSpace(labels)) {
            return;
        }

        if (!IsTruthy(labels!)) {
            chart.ClearDataLabels();
            return;
        }

        var showPercent = normalizedType is "pie" or "donut" or "doughnut"
            && IsTruthy(GetAttribute(source.Attributes, "percent", "show-percent", "showPercent") ?? "false");
        chart.SetDataLabels(showValue: true, showCategoryName: false, showSeriesName: false, showLegendKey: false, showPercent: showPercent);

        var labelPosition = GetAttribute(source.Attributes, "label-position", "labelPosition", "data-label-position", "dataLabelPosition");
        if (TryParseDataLabelPosition(labelPosition, out var position)) {
            chart.SetDataLabelPosition(position);
        }

        var labelFormat = GetAttribute(source.Attributes, "label-format", "labelFormat", "data-label-format", "dataLabelFormat");
        if (!string.IsNullOrWhiteSpace(labelFormat)) {
            chart.SetDataLabelNumberFormat(labelFormat!);
        }

        chart.SetDataLabelTextStyle(fontSizePoints: 9, color: textColor, fontName: font);
    }

    private static void ApplyGridlineOptions(PowerPointChart chart, OfficeMarkupChartBlock source, string normalizedType, string gridColor) {
        if (normalizedType is "pie" or "donut" or "doughnut") {
            return;
        }

        var gridlines = GetAttribute(source.Attributes, "gridlines");
        var valueGridlines = GetAttribute(source.Attributes, "value-gridlines", "valueGridlines", "y-gridlines", "yGridlines") ?? gridlines;
        var categoryGridlines = GetAttribute(source.Attributes, "category-gridlines", "categoryGridlines", "x-gridlines", "xGridlines");

        if (!string.IsNullOrWhiteSpace(valueGridlines)) {
            if (IsTruthy(valueGridlines!)) {
                chart.SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: gridColor, lineWidthPoints: 0.5);
            } else {
                chart.ClearValueAxisGridlines();
            }
        }

        if (!string.IsNullOrWhiteSpace(categoryGridlines)) {
            if (IsTruthy(categoryGridlines!)) {
                chart.SetCategoryAxisGridlines(showMajor: true, showMinor: false, lineColor: gridColor, lineWidthPoints: 0.5);
            } else {
                chart.ClearCategoryAxisGridlines();
            }
        }
    }

    private static IReadOnlyList<string> ResolveChartPalette(OfficeMarkupChartBlock chart) {
        if (chart.Attributes.TryGetValue("palette", out var palette) && !string.IsNullOrWhiteSpace(palette)) {
            var colors = palette.Split(new[] { ',', ';', '|' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(ToPowerPointColor)
                .Where(color => !string.IsNullOrWhiteSpace(color))
                .Cast<string>()
                .ToList();
            if (colors.Count > 0) {
                return colors;
            }
        }

        return new[] { "2563EB", "F97316", "10B981", "A855F7", "EF4444", "14B8A6" };
    }

    private static bool TryParseBool(string value, out bool result) {
        if (bool.TryParse(value, out result)) {
            return true;
        }

        switch (Normalize(value)) {
            case "yes":
            case "y":
            case "1":
            case "on":
                result = true;
                return true;
            case "no":
            case "n":
            case "0":
            case "off":
                result = false;
                return true;
            default:
                result = false;
                return false;
        }
    }

    private static bool IsTruthy(string value) =>
        Normalize(value) is not ("false" or "no" or "off" or "none" or "hidden" or "0");

    private static bool TryParseLegendPosition(string value, out C.LegendPositionValues position) {
        switch (Normalize(value)) {
            case "left":
                position = C.LegendPositionValues.Left;
                return true;
            case "right":
                position = C.LegendPositionValues.Right;
                return true;
            case "top":
                position = C.LegendPositionValues.Top;
                return true;
            case "bottom":
                position = C.LegendPositionValues.Bottom;
                return true;
            case "corner":
            case "topright":
                position = C.LegendPositionValues.TopRight;
                return true;
            default:
                position = C.LegendPositionValues.Bottom;
                return false;
        }
    }

    private static bool TryParseDataLabelPosition(string? value, out C.DataLabelPositionValues position) {
        switch (Normalize(value ?? string.Empty)) {
            case "center":
                position = C.DataLabelPositionValues.Center;
                return true;
            case "insideend":
                position = C.DataLabelPositionValues.InsideEnd;
                return true;
            case "insidebase":
                position = C.DataLabelPositionValues.InsideBase;
                return true;
            case "outsideend":
                position = C.DataLabelPositionValues.OutsideEnd;
                return true;
            case "bestfit":
                position = C.DataLabelPositionValues.BestFit;
                return true;
            case "left":
                position = C.DataLabelPositionValues.Left;
                return true;
            case "right":
                position = C.DataLabelPositionValues.Right;
                return true;
            case "top":
                position = C.DataLabelPositionValues.Top;
                return true;
            case "bottom":
                position = C.DataLabelPositionValues.Bottom;
                return true;
            default:
                position = C.DataLabelPositionValues.OutsideEnd;
                return false;
        }
    }

    private static string? ToPowerPointColor(string? color) {
        if (string.IsNullOrWhiteSpace(color)) {
            return null;
        }

        color = color!.Trim();
        if (color.StartsWith("#", StringComparison.Ordinal)) {
            color = color.Substring(1);
        }

        return color.Length == 6 && color.All(IsHexDigit) ? color.ToUpperInvariant() : null;
    }

    private static bool IsHexDigit(char value) =>
        (value >= '0' && value <= '9')
        || (value >= 'a' && value <= 'f')
        || (value >= 'A' && value <= 'F');

    private static void AddChart(
        PowerPointSlide slide,
        OfficeMarkupChartBlock chart,
        LayoutCursor cursor,
        OfficeMarkupPowerPointExportOptions options,
        SlideCanvasMetrics metrics) {
        if (!TryCreateChartData(chart, out var data)) {
            if (options.IncludeUnsupportedBlocksAsText) {
                AddText(slide, $"Chart: {chart.Title ?? chart.ChartType}", cursor, height: 0.55);
            }

            return;
        }

        var box = ResolveBox(chart.Placement, chart.Attributes, cursor, Math.Min(2.4, cursor.RemainingHeight), metrics);
        if (ShouldAddChartPanel(chart)) {
            AddChartPanel(slide, box, metrics);
        }

        var nativeChart = AddNativeChart(slide, chart.ChartType, data, box);
        if (!string.IsNullOrWhiteSpace(chart.Title)) {
            nativeChart.SetTitle(chart.Title!);
        }

        ApplyChartStyle(nativeChart, chart, data);
        if (!HasExplicitPlacement(chart.Placement, chart.Attributes)) {
            cursor.Advance(box.Height);
        }
    }

    private static PowerPointChart AddNativeChart(
        PowerPointSlide slide,
        string chartType,
        PowerPointChartData data,
        LayoutCursor box) {
        switch (Normalize(chartType)) {
            case "line":
                return slide.AddLineChartInches(data, box.Left, box.Top, box.Width, box.Height);
            case "pie":
                return slide.AddPieChartInches(FirstSeriesOnly(data), box.Left, box.Top, box.Width, box.Height);
            case "donut":
            case "doughnut":
                return slide.AddDoughnutChartInches(FirstSeriesOnly(data), box.Left, box.Top, box.Width, box.Height);
            case "column":
            case "clusteredcolumn":
            case "bar":
            case "clusteredbar":
            default:
                return slide.AddChartInches(data, box.Left, box.Top, box.Width, box.Height);
        }
    }

    private static bool TryCreateChartData(OfficeMarkupChartBlock chart, out PowerPointChartData data) {
        data = PowerPointChartData.CreateDefault();
        if (chart.Data.Count < 2) {
            return false;
        }

        var headers = chart.Data[0].Select(cell => cell ?? string.Empty).ToList();
        if (headers.Count < 2) {
            return false;
        }

        var categories = new List<string>();
        var seriesValues = new List<List<double>>();
        for (int columnIndex = 1; columnIndex < headers.Count; columnIndex++) {
            seriesValues.Add(new List<double>());
        }

        for (int rowIndex = 1; rowIndex < chart.Data.Count; rowIndex++) {
            var row = chart.Data[rowIndex];
            if (row.Count == 0 || string.IsNullOrWhiteSpace(row[0])) {
                continue;
            }

            categories.Add(row[0]);
            for (int columnIndex = 1; columnIndex < headers.Count; columnIndex++) {
                var value = columnIndex < row.Count ? row[columnIndex] : string.Empty;
                seriesValues[columnIndex - 1].Add(ParseDouble(value));
            }
        }

        if (categories.Count == 0) {
            return false;
        }

        var series = new List<PowerPointChartSeries>();
        for (int index = 0; index < seriesValues.Count; index++) {
            var name = string.IsNullOrWhiteSpace(headers[index + 1]) ? $"Series {index + 1}" : headers[index + 1];
            series.Add(new PowerPointChartSeries(name, seriesValues[index]));
        }

        data = new PowerPointChartData(categories, series);
        return true;
    }

    private static PowerPointChartData FirstSeriesOnly(PowerPointChartData data) =>
        new PowerPointChartData(data.Categories, data.Series.Take(1));

    private static double ParseDouble(string value) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : 0d;

    private static void AddImage(
        PowerPointSlide slide,
        OfficeMarkupImageBlock image,
        LayoutCursor cursor,
        OfficeMarkupPowerPointExportOptions options,
        SlideCanvasMetrics metrics) {
        if (TryResolveFilePath(options, image.Source, out var path) && File.Exists(path)) {
            var box = ResolveBox(image.Placement, image.Attributes, cursor, Math.Min(2.2, cursor.RemainingHeight), metrics);
            if (ShouldAddVisualPanel(image.Attributes, defaultValue: false)) {
                AddVisualPanel(slide, box, metrics, "OfficeIMO Markup Image Panel");
            }

            AddPicture(slide, path, box, GetAttribute(image.Attributes, "fit"));
            if (!HasExplicitPlacement(image.Placement, image.Attributes)) {
                cursor.Advance(box.Height);
            }
        } else if (options.IncludeUnsupportedBlocksAsText) {
            AddText(slide, $"Image: {image.Source}", cursor, height: 0.4);
        }
    }

    private static bool TryResolveFilePath(
        OfficeMarkupPowerPointExportOptions? options,
        string source,
        out string path) {
        path = source;
        if (string.IsNullOrWhiteSpace(source)) {
            return false;
        }

        if (Uri.TryCreate(source, UriKind.Absolute, out var uri)) {
            if (!uri.IsFile) {
                return false;
            }

            path = uri.LocalPath;
        } else {
            path = ResolvePath(options, source);
        }

        try {
            path = Path.GetFullPath(path);
            return true;
        } catch (Exception) when (!Debugger.IsAttached) {
            return false;
        }
    }

    private static void AddTable(PowerPointSlide slide, OfficeMarkupTableBlock table, LayoutCursor cursor) {
        var rows = new List<IReadOnlyList<string>>();
        if (table.Headers.Count > 0) {
            rows.Add(table.Headers.ToList());
        }

        rows.AddRange(table.Rows);
        if (rows.Count == 0) {
            return;
        }

        var columns = rows.Max(row => row.Count);
        var height = Math.Min(cursor.RemainingHeight, Math.Max(0.8, 0.32 * rows.Count));
        var powerPointTable = slide.AddTableInches(rows.Count, Math.Max(1, columns), cursor.Left, cursor.Top, cursor.Width, height);
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            var row = rows[rowIndex];
            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                powerPointTable.GetCell(rowIndex, columnIndex).Text = row[columnIndex];
            }
        }

        cursor.Advance(height);
    }

    private static bool ShouldAddVisualPanel(IDictionary<string, string> attributes, bool defaultValue) {
        var value = GetAttribute(attributes, "panel", "frame", "background-panel");
        return string.IsNullOrWhiteSpace(value) ? defaultValue : IsTruthy(value!);
    }

    private static void AddVisualPanel(PowerPointSlide slide, LayoutCursor box, SlideCanvasMetrics metrics, string name) {
        const double padding = 0.08;
        var left = Math.Max(metrics.Horizontal(0.18), box.Left - metrics.Horizontal(padding));
        var top = Math.Max(metrics.Vertical(0.18), box.Top - metrics.Vertical(padding));
        var right = Math.Min(metrics.Width - metrics.Horizontal(0.18), box.Left + box.Width + metrics.Horizontal(padding));
        var bottom = Math.Min(metrics.Height - metrics.Vertical(0.18), box.Top + box.Height + metrics.Vertical(padding));
        var panel = slide.AddShapeInches(
            A.ShapeTypeValues.Rectangle,
            left,
            top,
            Math.Max(0.5, right - left),
            Math.Max(0.5, bottom - top),
            name);
        panel.FillColor = "FFFFFF";
        panel.FillTransparency = 4;
        panel.OutlineColor = "D9E2EF";
        panel.OutlineWidthPoints = 0.75;
    }

    private static void AddPicture(PowerPointSlide slide, string path, LayoutCursor box, string? fit) {
        switch (Normalize(fit ?? string.Empty)) {
            case "fill":
            case "stretch":
                slide.AddPictureInches(path, box.Left, box.Top, box.Width, box.Height);
                return;
            case "contain":
            default:
                AddPictureContained(slide, path, box);
                return;
        }
    }

    private static void AddPictureContained(PowerPointSlide slide, string path, LayoutCursor box) {
        var left = box.Left;
        var top = box.Top;
        var width = box.Width;
        var height = box.Height;

        if (TryReadImageSize(path, out var pixelWidth, out var pixelHeight) && pixelWidth > 0 && pixelHeight > 0) {
            var imageAspect = pixelWidth / (double)pixelHeight;
            var boxAspect = box.Width / box.Height;
            if (imageAspect > boxAspect) {
                height = box.Width / imageAspect;
                top = box.Top + ((box.Height - height) / 2.0);
            } else {
                width = box.Height * imageAspect;
                left = box.Left + ((box.Width - width) / 2.0);
            }
        }

        slide.AddPictureInches(path, left, top, width, height);
    }

    private static bool TryReadImageSize(string path, out int width, out int height) {
        width = 0;
        height = 0;

        try {
            using var stream = File.OpenRead(path);
            if (TryReadPngSize(stream, out width, out height)) {
                return true;
            }

            stream.Position = 0;
            return TryReadJpegSize(stream, out width, out height);
        } catch (IOException) {
            return false;
        } catch (UnauthorizedAccessException) {
            return false;
        }
    }

    private static bool TryReadPngSize(Stream stream, out int width, out int height) {
        width = 0;
        height = 0;

        var header = new byte[24];
        if (stream.Read(header, 0, header.Length) != header.Length || !IsPngHeader(header)) {
            return false;
        }

        width = ReadBigEndianInt32(header, 16);
        height = ReadBigEndianInt32(header, 20);
        return true;
    }

    private static bool TryReadJpegSize(Stream stream, out int width, out int height) {
        width = 0;
        height = 0;

        if (stream.ReadByte() != 0xFF || stream.ReadByte() != 0xD8) {
            return false;
        }

        while (stream.Position < stream.Length) {
            int prefix;
            do {
                prefix = stream.ReadByte();
            } while (prefix != -1 && prefix != 0xFF);

            if (prefix == -1) {
                return false;
            }

            int marker;
            do {
                marker = stream.ReadByte();
            } while (marker == 0xFF);

            if (marker == -1) {
                return false;
            }

            if (marker == 0xD8 || marker == 0xD9 || (marker >= 0xD0 && marker <= 0xD7) || marker == 0x01) {
                continue;
            }

            var segmentLength = ReadBigEndianUInt16(stream);
            if (segmentLength < 2) {
                return false;
            }

            if (IsJpegStartOfFrame(marker)) {
                if (segmentLength < 7) {
                    return false;
                }

                if (stream.ReadByte() == -1) {
                    return false;
                }

                height = ReadBigEndianUInt16(stream);
                width = ReadBigEndianUInt16(stream);
                return width > 0 && height > 0;
            }

            stream.Seek(segmentLength - 2, SeekOrigin.Current);
        }

        return false;
    }

    private static bool IsPngHeader(byte[] header) =>
        header.Length >= 24
        && header[0] == 0x89
        && header[1] == 0x50
        && header[2] == 0x4E
        && header[3] == 0x47
        && header[4] == 0x0D
        && header[5] == 0x0A
        && header[6] == 0x1A
        && header[7] == 0x0A;

    private static int ReadBigEndianInt32(byte[] value, int offset) =>
        (value[offset] << 24) | (value[offset + 1] << 16) | (value[offset + 2] << 8) | value[offset + 3];

    private static int ReadBigEndianUInt16(Stream stream) {
        var high = stream.ReadByte();
        var low = stream.ReadByte();
        return high < 0 || low < 0 ? -1 : (high << 8) | low;
    }

    private static bool IsJpegStartOfFrame(int marker) =>
        marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF;

    private static LayoutCursor ResolveBox(
        OfficeMarkupPlacement? placement,
        IDictionary<string, string> attributes,
        LayoutCursor fallback,
        double defaultHeight,
        SlideCanvasMetrics metrics) {
        if (!HasExplicitPlacement(placement, attributes)) {
            return new LayoutCursor(fallback.Left, fallback.Top, fallback.Width, Math.Min(metrics.Vertical(defaultHeight), fallback.RemainingHeight));
        }

        var left = ParsePercentOrInches(PlacementValue(placement, attributes, "x"), fallback.Left, metrics.Width);
        var top = ParsePercentOrInches(PlacementValue(placement, attributes, "y"), fallback.Top, metrics.Height);
        var width = ParsePercentOrInches(PlacementValue(placement, attributes, "w"), fallback.Width, metrics.Width);
        var height = ParsePercentOrInches(PlacementValue(placement, attributes, "h"), Math.Min(metrics.Vertical(defaultHeight), fallback.RemainingHeight), metrics.Height);
        return new LayoutCursor(left, top, width, height);
    }

    private static LayoutCursor ResolveBox(IDictionary<string, string> attributes, LayoutCursor fallback, double defaultHeight, SlideCanvasMetrics metrics) =>
        ResolveBox(null, attributes, fallback, defaultHeight, metrics);

    private static bool HasExplicitPlacement(OfficeMarkupBlock block) =>
        HasExplicitPlacement(GetPlacement(block), block.Attributes);

    private static bool HasExplicitPlacement(OfficeMarkupPlacement? placement, IDictionary<string, string> attributes) =>
        placement?.HasValue == true || HasExplicitPlacement(attributes);

    private static bool HasExplicitPlacement(IDictionary<string, string> attributes) =>
        attributes.ContainsKey("x")
        || attributes.ContainsKey("y")
        || attributes.ContainsKey("w")
        || attributes.ContainsKey("h")
        || attributes.ContainsKey("width")
        || attributes.ContainsKey("height");

    private static string? PlacementValue(OfficeMarkupPlacement? placement, IDictionary<string, string> attributes, string name) {
        if (placement != null) {
            switch (name) {
                case "x":
                    if (!string.IsNullOrWhiteSpace(placement.X)) {
                        return placement.X;
                    }
                    break;
                case "y":
                    if (!string.IsNullOrWhiteSpace(placement.Y)) {
                        return placement.Y;
                    }
                    break;
                case "w":
                    if (!string.IsNullOrWhiteSpace(placement.Width)) {
                        return placement.Width;
                    }
                    break;
                case "h":
                    if (!string.IsNullOrWhiteSpace(placement.Height)) {
                        return placement.Height;
                    }
                    break;
            }
        }

        if (attributes.TryGetValue(name, out var value)) {
            return value;
        }

        if (name == "w" && attributes.TryGetValue("width", out value)) {
            return value;
        }

        if (name == "h" && attributes.TryGetValue("height", out value)) {
            return value;
        }

        return null;
    }

    private static double ParsePercentOrInches(string? value, double fallback, double size) {
        if (string.IsNullOrWhiteSpace(value)) {
            return fallback;
        }

        value = value!.Trim();
        if (value.EndsWith("%", StringComparison.Ordinal)) {
            return double.TryParse(value.Substring(0, value.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out var percent)
                ? size * (percent / 100.0)
                : fallback;
        }

        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var inches) ? inches : fallback;
    }

    private static void ApplyTransition(PowerPointSlide slide, string? transition) {
        if (string.IsNullOrWhiteSpace(transition)) {
            return;
        }

        var resolvedTransition = OfficeMarkupTransitionResolver.Parse(transition);
        if (string.IsNullOrWhiteSpace(resolvedTransition.ResolvedIdentifier)) {
            return;
        }

        if (Enum.TryParse<SlideTransition>(resolvedTransition.ResolvedIdentifier, true, out var parsed)) {
            slide.Transition = parsed;
            ApplyTransitionAttributes(slide, resolvedTransition.Attributes);
        }
    }

    private static void ApplyTransitionAttributes(PowerPointSlide slide, IReadOnlyDictionary<string, string> attributes) {
        if (TryGetTransitionSpeed(attributes, out var speed)) {
            slide.TransitionSpeed = speed;
        }

        if (TryGetTransitionSeconds(attributes, out var durationSeconds, "duration", "dur")) {
            slide.TransitionDurationSeconds = durationSeconds;
        }

        if (TryGetTransitionBoolean(attributes, out var advanceOnClick, "advance-on-click", "advanceonclick", "advance-click", "onclick", "click")) {
            slide.TransitionAdvanceOnClick = advanceOnClick;
        }

        if (TryGetTransitionSeconds(attributes, out var advanceAfterSeconds, "advance-after", "advanceafter", "after", "delay")) {
            slide.TransitionAdvanceAfterSeconds = advanceAfterSeconds;
        }
    }

    private static bool TryGetTransitionSpeed(IReadOnlyDictionary<string, string> attributes, out SlideTransitionSpeed speed) {
        speed = default;
        var value = GetTransitionAttribute(attributes, "speed", "spd");
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        switch (NormalizeTransitionToken(value)) {
            case "slow":
                speed = SlideTransitionSpeed.Slow;
                return true;
            case "medium":
            case "med":
                speed = SlideTransitionSpeed.Medium;
                return true;
            case "fast":
                speed = SlideTransitionSpeed.Fast;
                return true;
            default:
                return false;
        }
    }

    private static bool TryGetTransitionSeconds(IReadOnlyDictionary<string, string> attributes, out double seconds, params string[] names) {
        seconds = default;
        var value = GetTransitionAttribute(attributes, names);
        return !string.IsNullOrWhiteSpace(value)
               && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out seconds);
    }

    private static bool TryGetTransitionBoolean(IReadOnlyDictionary<string, string> attributes, out bool enabled, params string[] names) {
        enabled = default;
        var value = GetTransitionAttribute(attributes, names);
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        switch (NormalizeTransitionToken(value)) {
            case "true":
            case "yes":
            case "on":
            case "1":
                enabled = true;
                return true;
            case "false":
            case "no":
            case "off":
            case "0":
                enabled = false;
                return true;
            default:
                return false;
        }
    }

    private static string? GetTransitionAttribute(IReadOnlyDictionary<string, string> attributes, params string[] names) {
        foreach (var name in names) {
            if (attributes.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)) {
                return value.Trim();
            }
        }

        return null;
    }

    private static string NormalizeTransitionToken(string? value) =>
        new string((value ?? string.Empty).Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());

    private static IReadOnlyList<OfficeMarkupBlock> ParseLightweightMarkdown(string body) {
        var blocks = new List<OfficeMarkupBlock>();
        foreach (var rawLine in body.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')) {
            var line = rawLine.Trim();
            if (line.Length == 0) {
                continue;
            }

            if (line.StartsWith("### ", StringComparison.Ordinal)) {
                blocks.Add(new OfficeMarkupHeadingBlock(3, line.Substring(4)));
            } else if (line.StartsWith("## ", StringComparison.Ordinal)) {
                blocks.Add(new OfficeMarkupHeadingBlock(2, line.Substring(3)));
            } else if (line.StartsWith("# ", StringComparison.Ordinal)) {
                blocks.Add(new OfficeMarkupHeadingBlock(1, line.Substring(2)));
            } else if (line.StartsWith("- ", StringComparison.Ordinal)) {
                var list = blocks.LastOrDefault() as OfficeMarkupListBlock;
                if (list == null) {
                    list = new OfficeMarkupListBlock(false);
                    blocks.Add(list);
                }

                list.Items.Add(new OfficeMarkupListItem(line.Substring(2)));
            } else {
                blocks.Add(new OfficeMarkupParagraphBlock(line));
            }
        }

        return blocks;
    }

    private static bool IsExtension(OfficeMarkupBlock block, string command) =>
        block is OfficeMarkupExtensionBlock extension
        && string.Equals(Normalize(extension.Command), Normalize(command), StringComparison.Ordinal);

    private static bool IsColumns(OfficeMarkupBlock block) =>
        block is OfficeMarkupColumnsBlock || IsExtension(block, "columns");

    private static bool IsColumn(OfficeMarkupBlock block) =>
        block is OfficeMarkupColumnBlock
        || IsExtension(block, "column")
        || IsExtension(block, "left")
        || IsExtension(block, "right");

    private static string GetColumnBody(OfficeMarkupBlock block) {
        if (block is OfficeMarkupColumnBlock column) {
            return column.Body;
        }

        return block is OfficeMarkupExtensionBlock extension ? extension.Body : string.Empty;
    }

    private static OfficeMarkupPlacement? GetPlacement(OfficeMarkupBlock block) {
        switch (block) {
            case OfficeMarkupImageBlock image:
                return image.Placement;
            case OfficeMarkupDiagramBlock diagram:
                return diagram.Placement;
            case OfficeMarkupChartBlock chart:
                return chart.Placement;
            case OfficeMarkupTextBoxBlock textBox:
                return textBox.Placement;
            case OfficeMarkupColumnsBlock columns:
                return columns.Placement;
            case OfficeMarkupCardBlock card:
                return card.Placement;
            default:
                return null;
        }
    }

    private static double ResolveGap(OfficeMarkupBlock block, SlideCanvasMetrics metrics) {
        string? value = null;
        if (block is OfficeMarkupColumnsBlock columns) {
            value = columns.Gap;
        }

        if (string.IsNullOrWhiteSpace(value)) {
            value = GetAttribute(block.Attributes, "gap");
        }

        return ParsePercentOrInches(value, metrics.Horizontal(0.28), metrics.Width);
    }

    private static string? GetAttribute(IDictionary<string, string> attributes, string name) =>
        attributes.TryGetValue(name, out var value) ? value : null;

    private static string? GetAttribute(IDictionary<string, string> attributes, params string[] names) {
        foreach (var name in names) {
            if (attributes.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)) {
                return value.Trim();
            }
        }

        return null;
    }

    private static string? NormalizeSectionName(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        return value!.Trim();
    }

    private static string Normalize(string value) => (value ?? string.Empty).Replace("-", string.Empty).ToLowerInvariant();

    private static double EstimateTextHeight(string text) {
        var lines = Math.Max(1, text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n').Length);
        return Math.Min(1.4, 0.3 * lines + 0.2);
    }

    private static void TryDelete(string path) {
        try {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        } catch (IOException ex) {
            Trace.TraceWarning($"OfficeIMO.Markup.PowerPoint could not delete temporary file '{path}': {ex.Message}");
        } catch (UnauthorizedAccessException ex) {
            Trace.TraceWarning($"OfficeIMO.Markup.PowerPoint could not delete temporary file '{path}': {ex.Message}");
        }
    }

    private sealed class LayoutCursor {
        public LayoutCursor(double left, double top, double width, double height) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
            InitialTop = top;
        }

        public double Left { get; }
        public double Top { get; private set; }
        public double Width { get; }
        public double Height { get; }
        public double Bottom => Top + Height;
        public double RemainingHeight => Math.Max(0.28, (InitialTop + Height) - Top);
        private double InitialTop { get; }

        public void Advance(double height) {
            Top += height + 0.12;
        }

        public void MoveToBottom() {
            Top = InitialTop + Height;
        }
    }
}
