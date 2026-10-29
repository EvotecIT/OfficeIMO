using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointLayoutEngine {
        private const double TileGap = 24D;

        // ------------------------------------------------------------------ Title, Section, Closing

        private static void RenderTitle(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            PowerPointLayoutGrid grid = context.Grid;
            PowerPointLayoutTheme theme = context.Theme;
            writer.Background(theme.Accent1);
            if (string.IsNullOrWhiteSpace(context.Deck.DeckTitle)) context.Deck.DeckTitle = slide.Title;
            string onAccentMuted = PowerPointLayoutTheme.Muted(theme.OnAccent, theme.Accent1, 0.72D);
            string? brand = context.Deck.FooterText;
            if (!string.IsNullOrWhiteSpace(brand)) {
                writer.Text("Brand", grid.Box(1, 8, 1, 1), context.Paragraph(brand!, PowerPointTextRole.Eyebrow, onAccentMuted));
            }

            PowerPointLayoutRect stack = grid.Box(1, 9, 3, 8);
            var items = new List<StackItem>();
            if (!string.IsNullOrWhiteSpace(slide.Eyebrow)) {
                items.Add(Stacked(context, "Eyebrow", slide.Eyebrow!, PowerPointTextRole.Eyebrow, stack.Width, 18D, onAccentMuted, balance: false));
            }

            // On 4:3 the display size steps down (x0.86) so the title stack keeps its line count.
            double? displaySize = grid.IsNarrow ? PowerPointLayoutTheme.RoundSize(context.Style(PowerPointTextRole.Display).Size * 0.86D) : null;
            items.Add(Stacked(context, "Title", slide.Title ?? string.Empty, PowerPointTextRole.Display, stack.Width, 22D, theme.OnAccent,
                displaySize, placeholder: "<p:ph type=\"ctrTitle\"/>", maxLines: 3));
            if (!string.IsNullOrWhiteSpace(slide.Subtitle)) {
                items.Add(Stacked(context, "Subtitle", slide.Subtitle!, PowerPointTextRole.Subtitle, grid.W(7), 0D, onAccentMuted,
                    placeholder: "<p:ph type=\"subTitle\" idx=\"1\"/>", maxLines: 2));
            }

            RenderStack(writer, stack, items);
            double ruleX = grid.X(10) - 12D;
            writer.Rect("Title Rule", new PowerPointLayoutRect(ruleX, grid.Y(3), 1D, grid.Y(11) - grid.Y(3) - 12D), theme.OnAccent, 0.28D);
            if (slide.Meta.Count > 0) {
                PowerPointTextStyleSpec caption = context.Style(PowerPointTextRole.Caption);
                List<PowerPointLayoutParagraph> lines = slide.Meta.Select(m => context.Paragraph(m, PowerPointTextRole.Caption, onAccentMuted, caption.Size + 1D)).ToList();
                writer.Text("Meta", grid.Box(10, 3, 9, 2), lines, anchor: "b");
            }
        }

        private static void RenderSection(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            PowerPointLayoutGrid grid = context.Grid;
            PowerPointLayoutTheme theme = context.Theme;
            writer.Background(theme.Ink);
            context.Deck.SectionCount++;
            context.Deck.CurrentSection = slide.Title;
            string number = context.Deck.SectionCount.ToString("00", CultureInfo.InvariantCulture);
            PowerPointTextStyleSpec numeral = context.Style(PowerPointTextRole.SectionNumeral);
            PowerPointLayoutRect box = grid.Box(8, 5, 1, 12);
            PowerPointLayoutParagraph numeralText = context.Paragraph(number, PowerPointTextRole.SectionNumeral, theme.OnInk, align: "r");
            numeralText.Runs[0].Alpha = 0.09D;
            writer.Text("Section Number", new PowerPointLayoutRect(box.X, box.Y, box.Width, grid.SlideHeight - box.Y), numeralText, anchor: "b", decorative: true);
            string onInkMuted = PowerPointLayoutTheme.Muted(theme.OnInk, theme.Ink, 0.64D);
            string eyebrowColor = PowerPointLayoutTheme.Contrast(theme.Accent1, theme.Ink) >= 3D ? theme.Accent1 : onInkMuted;
            if (PowerPointLayoutTheme.Contrast(eyebrowColor, theme.Ink) < 4.5D) eyebrowColor = onInkMuted;
            PowerPointLayoutRect stack = grid.Box(1, 8, 4, 7);
            var items = new List<StackItem> {
                Stacked(context, "Eyebrow", !string.IsNullOrWhiteSpace(slide.Eyebrow) ? slide.Eyebrow! : "Section " + number, PowerPointTextRole.Eyebrow,
                    stack.Width, 16D, eyebrowColor, balance: false),
                Stacked(context, "Title", slide.Title ?? string.Empty, PowerPointTextRole.SectionTitle, stack.Width, 18D, theme.OnInk,
                    grid.IsNarrow ? PowerPointLayoutTheme.RoundSize(context.Style(PowerPointTextRole.SectionTitle).Size * 0.9D) : null,
                    placeholder: "<p:ph type=\"title\"/>", maxLines: 3)
            };
            if (!string.IsNullOrWhiteSpace(slide.Subtitle)) {
                items.Add(Stacked(context, "Subtitle", slide.Subtitle!, PowerPointTextRole.Subtitle, stack.Width, 0D, onInkMuted,
                    placeholder: "<p:ph type=\"body\" idx=\"1\"/>", maxLines: 2));
            }

            RenderStack(writer, stack, items);
        }

        private static void RenderClosing(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            PowerPointLayoutGrid grid = context.Grid;
            PowerPointLayoutTheme theme = context.Theme;
            writer.Background(theme.Accent1);
            string onAccentMuted = PowerPointLayoutTheme.Muted(theme.OnAccent, theme.Accent1, 0.72D);
            PowerPointLayoutRect stack = grid.Box(1, 6, 3, 8);
            var items = new List<StackItem>();
            if (!string.IsNullOrWhiteSpace(slide.Eyebrow)) {
                items.Add(Stacked(context, "Eyebrow", slide.Eyebrow!, PowerPointTextRole.Eyebrow, stack.Width, 18D, onAccentMuted, balance: false));
            }

            double closingSize = PowerPointLayoutTheme.RoundSize(theme.Theme.TypeScale.Title.ToPoints() * 1.5D);
            items.Add(Stacked(context, "Title", slide.Title ?? string.Empty, PowerPointTextRole.Display, stack.Width, 20D, theme.OnAccent,
                closingSize, "<p:ph type=\"title\"/>", maxLines: 3, minimum: 32D));
            List<PowerPointLayoutBlock> blocks = slide.Blocks.Where(b => b.IsContentBlock).ToList();
            string? contact = blocks.OfType<PowerPointLayoutText>().FirstOrDefault()?.Paragraphs.FirstOrDefault();
            string? subtitle = slide.Subtitle ?? contact;
            if (!string.IsNullOrWhiteSpace(subtitle)) {
                // Each line (statement, contact) is its own frame so it wraps and measures independently.
                string[] parts = subtitle!.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < parts.Length; i++) {
                    items.Add(Stacked(context, i == 0 ? "Subtitle" : "Contact", parts[i].Trim(), PowerPointTextRole.Subtitle, stack.Width,
                        i == parts.Length - 1 ? 0D : 8D, onAccentMuted, maxLines: 3));
                }
            }

            RenderStack(writer, stack, items);
            PowerPointLayoutList? list = blocks.OfType<PowerPointLayoutList>().FirstOrDefault();
            if (list == null || list.Items.Count == 0) return;
            PowerPointLayoutRect box = grid.Box(8, 5, 3, 8);
            PowerPointTextStyleSpec body = context.Style(PowerPointTextRole.Body);
            PowerPointTextStyleSpec eyebrow = context.Style(PowerPointTextRole.Eyebrow);
            List<PowerPointLayoutListItem> steps = list.Items.Where(i => i.Level == 1).Take(4).ToList();
            var heights = steps.Select(s => 12D + eyebrow.Size * eyebrow.LineHeight + 6D + context.Height(s.Text, body, box.Width)).ToList();
            double y = box.Bottom - heights.Sum() - 18D * (steps.Count - 1);
            for (int i = 0; i < steps.Count; i++) {
                writer.Rect("Step Rule", new PowerPointLayoutRect(box.X, y, box.Width, 1D), theme.OnAccent, 0.28D);
                writer.Text("Step Number", new PowerPointLayoutRect(box.X, y + 12D, box.Width, eyebrow.Size * eyebrow.LineHeight + 2D),
                    context.Paragraph((i + 1).ToString("00", CultureInfo.InvariantCulture), PowerPointTextRole.Eyebrow, onAccentMuted));
                double textY = y + 12D + eyebrow.Size * eyebrow.LineHeight + 6D;
                writer.Text("Next Step", new PowerPointLayoutRect(box.X, textY, box.Width, context.Height(steps[i].Text, body, box.Width) + 4D),
                    context.Paragraph(steps[i].Text, PowerPointTextRole.Body, theme.OnAccent));
                y += heights[i] + 18D;
            }
        }

        // ------------------------------------------------------------------ tiles

        private static List<(double X, double Width)> Tiles(PowerPointLayoutGrid grid, int[] spans, bool fifths, int count) {
            var result = new List<(double, double)>();
            if (fifths) {
                double w = (grid.W(12) - 4D * grid.Gutter) / 5D;
                for (int i = 0; i < count; i++) result.Add((grid.X(1) + i * (w + grid.Gutter), w));
                return result;
            }

            int column = 1;
            foreach (int span in spans) {
                result.Add((grid.X(column), grid.W(span)));
                column += span;
            }

            return result;
        }

        private static string? ChangeColor(PowerPointLayoutContext context, PowerPointLayoutKpi kpi) {
            string? explicitColor = context.ResolveColor(kpi.Color);
            if (explicitColor != null) return explicitColor;
            PowerPointKpiTrend? trend = kpi.EffectiveTrend;
            if (!trend.HasValue || trend == PowerPointKpiTrend.Flat) return context.Theme.TextMuted;
            bool improved = kpi.GoodDirection == PowerPointKpiDirection.Up
                ? trend == PowerPointKpiTrend.Up : trend == PowerPointKpiTrend.Down;
            return improved ? context.Theme.Success : context.Theme.Danger;
        }

        /// <summary>Matches the change paragraph's enlarged caption font and its native line spacing.</summary>
        private static double MetricChangeHeight(PowerPointTextStyleSpec caption) =>
            (caption.Size + 1D) * caption.LineHeight + 2D;

        private static void RenderTakeaway(PowerPointLayoutContext context, PowerPointLayoutWriter writer, string? takeaway) {
            if (string.IsNullOrWhiteSpace(takeaway)) return;
            PowerPointLayoutGrid grid = context.Grid;
            PowerPointLayoutRect box = grid.Box(1, 8, 11, 2);
            writer.Rect("Takeaway Rule", new PowerPointLayoutRect(box.X, box.Y - 14D, grid.W(12), 1D), context.Theme.Hairline);
            writer.Text("Takeaway", box, context.Paragraph(takeaway!, PowerPointTextRole.Lede));
        }

        private static string? TakeawayOf(PowerPointLayoutSlide slide) =>
            slide.Takeaway ?? ContentBlocks(slide).OfType<PowerPointLayoutText>().FirstOrDefault()?.Paragraphs.FirstOrDefault();

        private static void RenderKpi(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            HeaderGeometry header = RenderHeader(context, writer, slide);
            DrawKpi(context, writer, slide, ContentZone(context, header), full: true);
        }

        internal static void DrawKpi(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide,
            PowerPointLayoutRect zone, bool full) {
            PowerPointLayoutGrid grid = PowerPointLayoutGrid.ForRegion(context.Grid, zone);
            List<PowerPointLayoutKpi> kpis = slide.Items<PowerPointLayoutKpi>().ToList();
            int n = kpis.Count;
            if (n == 0) return;
            bool twoRows = n == 6 || grid.IsNarrow && n > 3;
            List<List<PowerPointLayoutKpi>> rows = twoRows ? Balance(kpis, 2) : new List<List<PowerPointLayoutKpi>> { kpis };
            PowerPointTextStyleSpec value = context.Style(PowerPointTextRole.KpiValue);
            PowerPointTextStyleSpec label = context.Style(PowerPointTextRole.KpiLabel);
            PowerPointTextStyleSpec caption = context.Style(PowerPointTextRole.Caption);
            double rowHeight = twoRows ? 140D : grid.H(5);
            double top = twoRows ? zone.Y : grid.Y(zone.Height > 250D ? 2 : 1);
            bool first = true;
            for (int r = 0; r < rows.Count; r++) {
                List<PowerPointLayoutKpi> row = rows[r];
                int[] spans = row.Count switch { 1 => new[] { 6 }, 2 => new[] { 6, 6 }, 3 => new[] { 4, 4, 4 }, _ => new[] { 3, 3, 3, 3 } };
                List<(double X, double Width)> tiles = Tiles(grid, spans, row.Count == 5, row.Count);
                double y0 = top + r * (rowHeight + 26D);
                for (int i = 0; i < row.Count; i++) {
                    PowerPointLayoutKpi kpi = row[i];
                    (double x, double w) = tiles[i];
                    writer.Rect("Metric Rule", new PowerPointLayoutRect(x, y0, w, 2D), first ? context.Theme.Accent1 : context.Theme.Hairline);
                    first = false;
                    double size = value.Size * (n == 1 ? 1.5D : 1D);
                    // In a region smaller than the tile stack, the value gives way first (never below 20 pt).
                    double stack = 20D + 10D + label.Size * label.LineHeight + (string.IsNullOrWhiteSpace(kpi.Change) ? 0D : 12D + MetricChangeHeight(caption));
                    size = Math.Max(20D, Math.Min(size, Math.Floor((Math.Min(rowHeight, zone.Height) - stack) / value.LineHeight)));
                    while (size - 2D >= value.Minimum && context.Width(kpi.Value, value, size) > w - 4D) size -= 2D;
                    double y = y0 + 2D + 18D;
                    writer.Text("Metric Value", new PowerPointLayoutRect(x, y, w, size * 1.15D), context.Paragraph(kpi.Value, PowerPointTextRole.KpiValue, size: size));
                    y += size * value.LineHeight + 10D;
                    double labelHeight = Math.Min(2, context.Lines(kpi.Label, label, w - 8D)) * label.Size * label.LineHeight;
                    writer.Text("Metric Label", new PowerPointLayoutRect(x, y, w - 8D, labelHeight + 4D), context.Paragraph(kpi.Label, PowerPointTextRole.KpiLabel));
                    y += labelHeight + 12D;
                    if (!string.IsNullOrWhiteSpace(kpi.Change)) {
                        writer.Text("Metric Change", new PowerPointLayoutRect(x, y, w, MetricChangeHeight(caption)),
                            context.Paragraph(kpi.ChangeText, PowerPointTextRole.Caption, ChangeColor(context, kpi), caption.Size + 1D, bold: true, explicitColor: kpi.Color));
                    }
                }
            }

            if (full && !twoRows) RenderTakeaway(context, writer, TakeawayOf(slide));
        }

        private static void RenderCards(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            HeaderGeometry header = RenderHeader(context, writer, slide);
            DrawCards(context, writer, slide, ContentZone(context, header), full: true);
        }

        private static void RenderProcess(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            HeaderGeometry header = RenderHeader(context, writer, slide);
            DrawProcess(context, writer, slide, ContentZone(context, header), full: true);
        }

        internal static void DrawProcess(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide,
            PowerPointLayoutRect zone, bool full) {
            PowerPointLayoutGrid grid = PowerPointLayoutGrid.ForRegion(context.Grid, zone);
            PowerPointLayoutTheme theme = context.Theme;
            List<PowerPointLayoutStep> steps = slide.Items<PowerPointLayoutStep>().ToList();
            if (steps.Count == 0) return;
            int perRow = grid.IsNarrow ? 3 : 5;
            List<List<PowerPointLayoutStep>> rows = steps.Count <= perRow ? new List<List<PowerPointLayoutStep>> { steps } : Balance(steps, 2);
            const double d = 40D;
            PowerPointTextStyleSpec heading = context.Style(PowerPointTextRole.Heading);
            PowerPointTextStyleSpec small = context.Style(PowerPointTextRole.BodySmall);
            double rowHeight = rows.Count == 1 ? grid.H(6) : (zone.Height - TileGap) / 2D;
            double top = rows.Count == 1 ? grid.Y(zone.Height > 250D ? 2 : 1) : zone.Y;
            int index = 0;
            for (int r = 0; r < rows.Count; r++) {
                List<PowerPointLayoutStep> row = rows[r];
                int[] spans = row.Count switch { 1 => new[] { 6 }, 2 => new[] { 6, 6 }, 3 => new[] { 4, 4, 4 }, _ => new[] { 3, 3, 3, 3 } };
                List<(double X, double Width)> tiles = Tiles(grid, spans, row.Count == 5, row.Count);
                double y0 = top + r * (rowHeight + TileGap);
                for (int i = 0; i < row.Count; i++) {
                    PowerPointLayoutStep step = row[i];
                    (double x, double w) = tiles[i];
                    bool first = index == 0;
                    if (i < row.Count - 1) {
                        double next = tiles[i + 1].X;
                        writer.Rect("Step Connector", new PowerPointLayoutRect(x + d + 8D, y0 + d / 2D - 1D, next - x - d - 16D, 2D), theme.Hairline);
                    }

                    string number = step.Number ?? (slide.StepStart + index).ToString("00", CultureInfo.InvariantCulture);
                    double numberSize = 15D;
                    while (numberSize > 9D && PowerPointTextMeasure.Width(number, numberSize, theme.HeadingFont, true) > d - 8D) numberSize -= 1D;
                    PowerPointLayoutParagraph label = context.Paragraph(number, PowerPointTextRole.Heading, first ? theme.OnAccent : theme.Accent1, numberSize);
                    writer.Circle("Step Marker", new PowerPointLayoutRect(x, y0, d, d), first ? theme.Accent1 : theme.Background, theme.Accent1, 2D, label);
                    double textWidth = w - 8D;
                    double titleSize = context.WordFitSize(step.Title, heading, textWidth);
                    double titleHeight = context.Height(step.Title, heading, textWidth, titleSize);
                    var paragraphs = new List<PowerPointLayoutParagraph> { context.Paragraph(step.Title, PowerPointTextRole.Heading, size: titleSize) };
                    if (!string.IsNullOrWhiteSpace(step.Text)) {
                        PowerPointLayoutParagraph text = context.Paragraph(step.Text!, PowerPointTextRole.BodySmall);
                        text.SpaceBefore = 8D;
                        paragraphs.Add(text);
                    }

                    double height = Math.Max(titleHeight + 8D, rowHeight - d - 22D);
                    writer.Text("Step Text", new PowerPointLayoutRect(x, y0 + d + 22D, textWidth, height), paragraphs);
                    index++;
                }
            }

            if (full && rows.Count == 1) RenderTakeaway(context, writer, TakeawayOf(slide));
        }

        // ------------------------------------------------------------------ Quote, ImageFull

        private static void RenderQuote(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            PowerPointLayoutGrid grid = context.Grid;
            PowerPointLayoutTheme theme = context.Theme;
            writer.Background(theme.Surface);
            PowerPointLayoutQuote? quote = slide.Items<PowerPointLayoutQuote>().FirstOrDefault();
            string text = quote?.Text ?? ContentBlocks(slide).OfType<PowerPointLayoutText>().FirstOrDefault()?.Paragraphs.FirstOrDefault() ?? slide.Title ?? string.Empty;
            text = text.Trim().Trim('"', '“', '”');
            string? attribution = quote?.Attribution ?? slide.Subtitle;
            PowerPointLayoutRect markBox = grid.Box(1, 1, 2, 3);
            PowerPointLayoutParagraph mark = context.Paragraph("“", PowerPointTextRole.Display, theme.Accent1, 120D);
            mark.LineHeight = 1.0D;
            writer.Text("Quote Mark", new PowerPointLayoutRect(markBox.X - 4D, markBox.Y - 10D, markBox.Width + 40D, markBox.Height + 40D), mark, decorative: true);
            PowerPointLayoutRect box = grid.Box(2, 9, 4, 6);
            PowerPointTextStyleSpec style = context.Style(PowerPointTextRole.Quote);
            double size = style.Size;
            while (size - 2D >= style.Minimum && context.Lines(text, style, box.Width, size) > 5) size -= 2D;
            writer.Text("Quote", box, context.Paragraph(PowerPointTextMeasure.Balance(text, size, box.Width, theme.HeadingFont), PowerPointTextRole.Quote, size: size),
                placeholder: "<p:ph type=\"title\"/>");
            if (string.IsNullOrWhiteSpace(attribution)) return;
            PowerPointLayoutRect attributionBox = grid.Box(2, 7, 10, 2);
            writer.Rect("Attribution Rule", new PowerPointLayoutRect(attributionBox.X, attributionBox.Y - 14D, 48D, 2D), theme.Accent1);
            string[] parts = attribution!.Split(new[] { ", " }, 2, StringSplitOptions.None);
            PowerPointTextStyleSpec small = context.Style(PowerPointTextRole.BodySmall);
            var paragraph = new PowerPointLayoutParagraph(new PowerPointLayoutRun(parts[0], small, color: theme.Text, bold: true)) { LineHeight = small.LineHeight };
            if (parts.Length > 1) paragraph.Runs.Add(new PowerPointLayoutRun(", " + parts[1], small));
            writer.Text("Attribution", attributionBox, paragraph);
        }

        private static void RenderImageFull(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            PowerPointLayoutGrid grid = context.Grid;
            PowerPointLayoutTheme theme = context.Theme;
            writer.Background(theme.Ink);
            PowerPointLayoutImage? image = ContentBlocks(slide).OfType<PowerPointLayoutImage>().FirstOrDefault();
            var full = new PowerPointLayoutRect(0, 0, grid.SlideWidth, grid.SlideHeight);
            if (image != null) RenderImage(context, writer, image, full, cover: true);
            writer.Rect("Scrim", full, null, gradient: writer.Gradient(theme.Ink, 90, (0D, 0D), (0.35D, 0D), (1D, 0.78D)));
            writer.Rect("Scrim Left", full, null, gradient: writer.Gradient(theme.Ink, 0, (0D, 0.4D), (0.62D, 0D), (1D, 0D)));
            PowerPointLayoutRect stack = grid.Box(1, 8, 7, 5);
            var items = new List<StackItem>();
            if (!string.IsNullOrWhiteSpace(slide.Eyebrow)) {
                items.Add(Stacked(context, "Eyebrow", slide.Eyebrow!, PowerPointTextRole.Eyebrow, stack.Width, 12D,
                    PowerPointLayoutTheme.Blend("FFFFFF", theme.Ink, 0.75D), balance: false));
            }

            double titleSize = PowerPointLayoutTheme.RoundSize(theme.Theme.TypeScale.Heading1.ToPoints() * 1.5D);
            items.Add(Stacked(context, "Title", slide.Title ?? string.Empty, PowerPointTextRole.Title, stack.Width, 12D, "FFFFFF", titleSize,
                "<p:ph type=\"title\"/>", maxLines: 2, minimum: 28D));
            if (!string.IsNullOrWhiteSpace(slide.Subtitle)) {
                items.Add(Stacked(context, "Subtitle", slide.Subtitle!, PowerPointTextRole.Subtitle, stack.Width, 0D,
                    PowerPointLayoutTheme.Blend("FFFFFF", theme.Ink, 0.8D), maxLines: 2));
            }

            if (!string.IsNullOrWhiteSpace(slide.SourceNote)) {
                if (items.Count > 0) items[items.Count - 1] = WithGap(items[items.Count - 1], 12D);
                items.Add(Stacked(context, "Source", slide.SourceNote!, PowerPointTextRole.Caption, stack.Width, 0D,
                    PowerPointLayoutTheme.Blend("FFFFFF", theme.Ink, 0.8D), balance: false));
            }

            RenderStack(writer, stack, items);
        }

        private static StackItem WithGap(StackItem item, double gapAfter) =>
            new(item.Name, item.Paragraph, item.Height, gapAfter, item.Placeholder);

        // ------------------------------------------------------------------ CaseStudy, Logos, Coverage

        private static void RenderCaseStudy(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            HeaderGeometry header = RenderHeader(context, writer, slide);
            DrawCaseStudy(context, writer, slide, ContentZone(context, header), full: true);
        }

        internal static void DrawCaseStudy(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide,
            PowerPointLayoutRect zone, bool full) {
            PowerPointLayoutGrid grid = PowerPointLayoutGrid.ForRegion(context.Grid, zone);
            PowerPointLayoutTheme theme = context.Theme;
            PowerPointLayoutRect narrative = new PowerPointLayoutRect(grid.X(1), zone.Y, grid.W(7), zone.Height);
            PowerPointTextStyleSpec eyebrow = context.Style(PowerPointTextRole.Eyebrow);
            PowerPointTextStyleSpec body = context.Style(PowerPointTextRole.Body);
            List<PowerPointLayoutCard> cards = slide.Items<PowerPointLayoutCard>().Take(4).ToList();
            double bodySize = body.Size;
            double Total(double size) => cards.Sum(c => eyebrow.Size * eyebrow.LineHeight + 6D +
                                                       context.Height(CardText(c), body, narrative.Width, size)) + Math.Max(0, cards.Count - 1) * 33D;
            while (Total(bodySize) > narrative.Height && bodySize - 2D >= 14D) bodySize -= 2D;
            double y = narrative.Y;
            for (int i = 0; i < cards.Count; i++) {
                if (i > 0) {
                    writer.Rect("Narrative Rule", new PowerPointLayoutRect(narrative.X, y + 16D, narrative.Width, 1D), theme.Hairline);
                    y += 33D;
                }

                writer.Text("Narrative Label", new PowerPointLayoutRect(narrative.X, y, narrative.Width, eyebrow.Size * eyebrow.LineHeight + 2D),
                    context.Paragraph(cards[i].Title, PowerPointTextRole.Eyebrow));
                y += eyebrow.Size * eyebrow.LineHeight + 6D;
                string text = CardText(cards[i]);
                double h = context.Height(text, body, narrative.Width, bodySize);
                writer.Text("Narrative Text", new PowerPointLayoutRect(narrative.X, y, narrative.Width, h + 4D), context.Paragraph(text, PowerPointTextRole.Body, size: bodySize));
                y += h;
            }

            List<PowerPointLayoutKpi> kpis = slide.Items<PowerPointLayoutKpi>().Take(3).ToList();
            if (kpis.Count == 0) return;
            PowerPointLayoutRect panel = new PowerPointLayoutRect(grid.X(9), zone.Y, grid.W(4), zone.Height);
            writer.Rect("Results Panel", panel, theme.Surface);
            writer.Rect("Results Rule", new PowerPointLayoutRect(panel.X, panel.Y, panel.Width, 3D), theme.Accent1);
            PowerPointTextStyleSpec value = context.Style(PowerPointTextRole.KpiValue);
            PowerPointTextStyleSpec label = context.Style(PowerPointTextRole.KpiLabel);
            PowerPointTextStyleSpec caption = context.Style(PowerPointTextRole.Caption);
            PowerPointLayoutRect inner = panel.Inset(24D, 23D, 24D, 20D);
            double labelHeight = label.Size * label.LineHeight;
            double fit = (inner.Height - (kpis.Count - 1) * 26D) / kpis.Count - (4D + labelHeight + 6D + MetricChangeHeight(caption));
            double size = Math.Max(value.Minimum * 0.75D, Math.Min(PowerPointLayoutTheme.RoundSize(value.Size * 0.67D), Math.Floor(fit)));
            double itemHeight = size + 4D + labelHeight + 6D + MetricChangeHeight(caption);
            double gap = kpis.Count > 1 ? (inner.Height - kpis.Count * itemHeight) / (kpis.Count - 1) : 0D;
            double ky = inner.Y;
            for (int i = 0; i < kpis.Count; i++) {
                if (i > 0) writer.Rect("Results Divider", new PowerPointLayoutRect(inner.X, ky - Math.Min(12D, gap / 2D), inner.Width, 1D), theme.Hairline);
                PowerPointLayoutKpi kpi = kpis[i];
                double s = size;
                while (s - 2D >= 20D && context.Width(kpi.Value, value, s) > inner.Width) s -= 2D;
                writer.Text("Result Value", new PowerPointLayoutRect(inner.X, ky, inner.Width, s * 1.15D), context.Paragraph(kpi.Value, PowerPointTextRole.KpiValue, size: s));
                double ly = ky + s + 4D;
                writer.Text("Result Label", new PowerPointLayoutRect(inner.X, ly, inner.Width, labelHeight + 2D), context.Paragraph(kpi.Label, PowerPointTextRole.KpiLabel));
                if (!string.IsNullOrWhiteSpace(kpi.Change)) {
                    writer.Text("Result Change", new PowerPointLayoutRect(inner.X, ly + labelHeight + 6D, inner.Width, MetricChangeHeight(caption)),
                        context.Paragraph(kpi.ChangeText, PowerPointTextRole.Caption, ChangeColor(context, kpi), caption.Size + 1D, bold: true, explicitColor: kpi.Color));
                }

                ky += itemHeight + gap;
            }
        }

        private static string CardText(PowerPointLayoutCard card) =>
            !string.IsNullOrWhiteSpace(card.Text) ? card.Text! : string.Join("; ", card.Items);

        private static void RenderLogos(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            HeaderGeometry header = RenderHeader(context, writer, slide);
            DrawLogos(context, writer, slide, ContentZone(context, header), full: true);
        }

        internal static void DrawLogos(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide,
            PowerPointLayoutRect zone, bool full) {
            PowerPointLayoutGrid grid = PowerPointLayoutGrid.ForRegion(context.Grid, zone);
            PowerPointLayoutTheme theme = context.Theme;
            List<PowerPointLayoutCard> logos = slide.Items<PowerPointLayoutCard>().ToList();
            if (logos.Count == 0) return;
            int perRow = grid.IsNarrow ? 3 : logos.Count <= 4 ? logos.Count : logos.Count <= 8 ? 4 : 6;
            int span = 12 / Math.Max(1, perRow == 5 ? 4 : perRow);
            const double tileHeight = 96D;
            int rowCount = (int)Math.Ceiling(logos.Count / (double)perRow);
            double total = rowCount * tileHeight + (rowCount - 1) * TileGap;
            double y0 = zone.Y + Math.Max(0D, (zone.Height - total) / 2D - 10D);
            PowerPointTextStyleSpec heading = context.Style(PowerPointTextRole.Heading);
            for (int i = 0; i < logos.Count; i++) {
                int row = i / perRow, column = i % perRow;
                var tile = new PowerPointLayoutRect(grid.X(1 + column * span), y0 + row * (tileHeight + TileGap), grid.W(span), tileHeight);
                writer.Rect("Logo Tile", tile, theme.Background, outline: theme.Hairline, outlineWidth: 1D);
                PowerPointLayoutCard logo = logos[i];
                if (!string.IsNullOrWhiteSpace(logo.Image) && System.IO.File.Exists(logo.Image)) {
                    var max = new PowerPointLayoutRect(tile.X + tile.Width * 0.19D, tile.Y + tile.Height * 0.28D, tile.Width * 0.62D, tile.Height * 0.44D);
                    PowerPointPicture picture = RenderImage(context, writer, new PowerPointLayoutImage(logo.Image!, logo.Title), max, cover: false);
                    // Centre the contained logo in its tile.
                    double pw = picture.Width.ToPoints(), ph = picture.Height.ToPoints();
                    picture.Left = OfficeLength.Points(tile.X + (tile.Width - pw) / 2D);
                    picture.Top = OfficeLength.Points(tile.Y + (tile.Height - ph) / 2D);
                    if (!logo.Color.HasValue) picture.SetMonochrome(0.8D);
                } else {
                    // A name-only logo: keep every word whole by shrinking to the widest word.
                    double size = heading.Size;
                    double width = tile.Width - 24D;
                    while (size - 1D >= 11D && (PowerPointTextMeasure.LongestWord(logo.Title, size, theme.HeadingFont, true) > width ||
                                                context.Lines(logo.Title, heading, width, size) > 2)) size -= 1D;
                    writer.Text("Logo Name", tile.Inset(12D, 8D, 12D, 8D), context.Paragraph(logo.Title, PowerPointTextRole.Heading, theme.TextMuted, size, "ctr"),
                        anchor: "ctr");
                }
            }
        }

        private static void RenderCoverage(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            HeaderGeometry header = RenderHeader(context, writer, slide);
            DrawCoverage(context, writer, slide, ContentZone(context, header), full: true);
        }

        internal static void DrawCoverage(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide,
            PowerPointLayoutRect zone, bool full) {
            PowerPointLayoutGrid grid = PowerPointLayoutGrid.ForRegion(context.Grid, zone);
            PowerPointLayoutTheme theme = context.Theme;
            PowerPointLayoutRect map = new PowerPointLayoutRect(grid.X(1), zone.Y, grid.W(8), zone.Height);
            PowerPointLayoutImage? image = ContentBlocks(slide).OfType<PowerPointLayoutImage>().FirstOrDefault();
            if (image != null) {
                RenderImage(context, writer, image, map, cover: true);
            } else {
                writer.Rect("Map Surface", map, theme.Surface);
                for (double dy = map.Y + 12D; dy < map.Bottom - 6D; dy += 12D) {
                    writer.Line("Map Dots", map.X + 12D, dy, map.Right - 12D, dy, theme.Hairline, 1.4D, arrow: false, dotted: true, decorative: true);
                }
            }

            if (!string.IsNullOrWhiteSpace(slide.MapLabel)) {
                PowerPointTextStyleSpec caption = context.Style(PowerPointTextRole.Caption);
                writer.Text("Map Label", new PowerPointLayoutRect(map.X + 20D, map.Bottom - 16D - caption.Size * 1.4D, map.Width - 40D, caption.Size * 1.4D),
                    context.Paragraph(slide.MapLabel!, PowerPointTextRole.Caption));
            }

            List<PowerPointLayoutLocation> locations = slide.Items<PowerPointLayoutLocation>().ToList();
            int offset = slide.CardStart - 1;
            for (int i = 0; i < locations.Count; i++) {
                double px = map.X + locations[i].Left * map.Width, py = map.Y + locations[i].Top * map.Height;
                writer.Circle("Pin Halo", new PowerPointLayoutRect(px - 17D, py - 17D, 34D, 34D), theme.Accent1, null, 0D, null, 0.14D, decorative: true);
                PowerPointLayoutParagraph number = context.Paragraph((offset + i + 1).ToString(CultureInfo.InvariantCulture), PowerPointTextRole.Caption,
                    theme.OnAccent, 10D, bold: true);
                writer.Circle("Pin", new PowerPointLayoutRect(px - 11D, py - 11D, 22D, 22D), theme.Accent1, theme.Background, 2D, number);
            }

            PowerPointLayoutRect list = new PowerPointLayoutRect(grid.X(9), zone.Y, grid.W(4), zone.Height);
            bool truncated = locations.Count > 7;
            List<PowerPointLayoutLocation> shown = truncated ? locations.Take(6).ToList() : locations;
            if (truncated) {
                context.Deck.Report(PowerPointLayoutDiagnosticCodes.LocationsTruncated, OfficeDiagnosticSeverity.Information,
                    "The location list shows the first six of " + locations.Count + " locations; every pin stays numbered.", slide.Title);
            }

            PowerPointTextStyleSpec heading = context.Style(PowerPointTextRole.Heading);
            PowerPointTextStyleSpec small = context.Style(PowerPointTextRole.BodySmall);
            double y = list.Y;
            for (int i = 0; i < shown.Count; i++) {
                double textX = list.X + 28D, textW = list.Width - 28D;
                writer.Text("Location Number", new PowerPointLayoutRect(list.X, y + 3D, 24D, 14D),
                    context.Paragraph((offset + i + 1).ToString("00", CultureInfo.InvariantCulture), PowerPointTextRole.Caption, theme.Accent1, bold: true));
                var paragraphs = new List<PowerPointLayoutParagraph> { context.Paragraph(shown[i].Name, PowerPointTextRole.Heading, size: 16D) };
                double h = context.Height(shown[i].Name, heading, textW, 16D);
                if (!string.IsNullOrWhiteSpace(shown[i].Detail)) {
                    PowerPointLayoutParagraph detail = context.Paragraph(shown[i].Detail!, PowerPointTextRole.BodySmall);
                    detail.SpaceBefore = 2D;
                    paragraphs.Add(detail);
                    h += 2D + Math.Min(2, context.Lines(shown[i].Detail!, small, textW)) * small.Size * small.LineHeight;
                }

                writer.Text("Location", new PowerPointLayoutRect(textX, y, textW, h + 4D), paragraphs);
                y += h + 16D;
            }

            if (truncated) {
                writer.Text("More Locations", new PowerPointLayoutRect(list.X + 28D, y, list.Width - 28D, 20D),
                    context.Paragraph("and " + (locations.Count - 6).ToString(CultureInfo.CurrentCulture) + " more", PowerPointTextRole.BodySmall));
            }
        }
    }
}
