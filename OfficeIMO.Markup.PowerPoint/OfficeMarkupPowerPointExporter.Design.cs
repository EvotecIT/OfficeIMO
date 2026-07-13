using System.Diagnostics;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Markup.PowerPoint;

internal sealed partial class OfficeMarkupPowerPointExporter {
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
        var textOnlyOptions = new MarkupToPowerPointOptions {
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
}
