using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointLayoutEngine {
        internal static void DrawCards(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide,
            PowerPointLayoutRect zone, bool full) {
            PowerPointLayoutGrid grid = PowerPointLayoutGrid.ForRegion(context.Grid, zone);
            PowerPointLayoutTheme theme = context.Theme;
            List<PowerPointLayoutCard> cards = slide.Items<PowerPointLayoutCard>().ToList();
            if (cards.Count == 0) return;
            List<int[]> rows = CardRows(cards.Count, grid.IsNarrow);
            double rowHeight = rows.Count == 1 ? zone.Height : (zone.Height - TileGap) / 2D;
            PowerPointTextStyleSpec heading = context.Style(PowerPointTextRole.Heading);
            PowerPointTextStyleSpec eyebrow = context.Style(PowerPointTextRole.Eyebrow);
            int index = 0;
            for (int r = 0; r < rows.Count; r++) {
                List<(double X, double Width)> tiles = Tiles(grid, rows[r], false, rows[r].Length);
                double y = zone.Y + r * (rowHeight + TileGap);
                List<PowerPointLayoutCard> rowCards = cards.Skip(index).Take(tiles.Count).ToList();
                CardRowMetrics metrics = MeasureCardRow(context, rowCards, tiles, rowHeight, slide.CardBodySize);

                for (int i = 0; i < rowCards.Count; i++) {
                    PowerPointLayoutCard card = rowCards[i];
                    (double x, double w) = tiles[i];
                    string accent = context.ResolveColor(card.Color) ?? theme.Accent1;
                    writer.Rect("Card", new PowerPointLayoutRect(x, y, w, rowHeight), theme.Surface);
                    writer.Rect("Card Rule", new PowerPointLayoutRect(x, y, w, 3D), accent, explicitFill: card.Color);
                    int number = slide.CardStart + index + i;
                    writer.Text("Card Number", new PowerPointLayoutRect(x + 20D, y + 23D, w - 40D, eyebrow.Size * eyebrow.LineHeight + 2D),
                        context.Paragraph(number.ToString("00", CultureInfo.InvariantCulture), PowerPointTextRole.Eyebrow, theme.TextMuted));
                    double inner = w - 40D;
                    var paragraphs = new List<PowerPointLayoutParagraph>();
                    PowerPointLayoutParagraph title = context.Paragraph(card.Title, PowerPointTextRole.Heading, size: metrics.TitleSize);
                    title.SpaceAfter = Math.Max(0D, metrics.TitleHeight - CardTextHeight(context, card.Title, heading, inner, metrics.TitleSize));
                    paragraphs.Add(title);
                    if (!string.IsNullOrWhiteSpace(card.Text)) {
                        PowerPointLayoutParagraph text = context.Paragraph(card.Text!, PowerPointTextRole.BodySmall, size: metrics.BodySize);
                        text.SpaceBefore = 8D;
                        paragraphs.Add(text);
                    }

                    for (int k = 0; k < card.Items.Count; k++) {
                        PowerPointLayoutParagraph item = context.Paragraph(card.Items[k], PowerPointTextRole.BodySmall, theme.Text, metrics.BodySize);
                        item.SpaceBefore = k == 0 ? 10D : 3D;
                        item.Bullet = "–";
                        item.BulletColor = accent;
                        item.MarginLeft = 16D;
                        item.Indent = -16D;
                        paragraphs.Add(item);
                    }

                    if (!string.IsNullOrWhiteSpace(card.Image)) {
                        double imageTop = y + 23D + eyebrow.Size * eyebrow.LineHeight + 12D;
                        RenderImage(context, writer, new PowerPointLayoutImage(card.Image!, card.Title),
                            new PowerPointLayoutRect(x + 20D, imageTop, inner, metrics.ImageHeight), cover: false);
                    }
                    // Every heading starts on the same row baseline. A short body
                    // must not pull its heading down, nor a long one above the index.
                    writer.Text("Card Text", new PowerPointLayoutRect(x + 20D, y + metrics.TextOffset,
                        inner, metrics.AvailableHeight), paragraphs);

                }

                index += rowCards.Count;
            }
        }

        private static List<int[]> CardRows(int count, bool narrow) => count switch {
            1 => new List<int[]> { new[] { 6 } },
            2 => new List<int[]> { new[] { 6, 6 } },
            3 => new List<int[]> { new[] { 4, 4, 4 } },
            4 => narrow ? new List<int[]> { new[] { 6, 6 }, new[] { 6, 6 } } : new List<int[]> { new[] { 3, 3, 3, 3 } },
            5 => new List<int[]> { new[] { 4, 4, 4 }, new[] { 6, 6 } },
            _ => new List<int[]> { new[] { 4, 4, 4 }, new[] { 4, 4, 4 } }
        };

        private sealed class CardRowMetrics {
            internal double TitleSize;
            internal double TitleHeight;
            internal double BodySize;
            internal double ImageHeight;
            internal double TextOffset;
            internal double AvailableHeight;
            internal double Height;
            internal bool WidthFits;
            internal bool Fits => WidthFits && Height <= AvailableHeight + 0.01D;
        }

        // The saved rich-text renderer rounds each leading advance up to a
        // whole point. Budget every line at that advance, including one-line
        // bullets, so the rounding cannot accumulate outside the panel.
        private static double CardTextHeight(PowerPointLayoutContext context, string text, PowerPointTextStyleSpec style,
            double width, double size) => context.Lines(text, style, width / 1.06D, size) * Math.Ceiling(size * style.LineHeight);

        /// <summary>One measurement owns admission, shared heading geometry and the rendered card frames.</summary>
        private static CardRowMetrics MeasureCardRow(PowerPointLayoutContext context, IReadOnlyList<PowerPointLayoutCard> cards,
            IReadOnlyList<(double X, double Width)> tiles, double rowHeight, double? bodySize = null) {
            PowerPointTextStyleSpec heading = context.Style(PowerPointTextRole.Heading);
            PowerPointTextStyleSpec small = context.Style(PowerPointTextRole.BodySmall);
            PowerPointTextStyleSpec eyebrow = context.Style(PowerPointTextRole.Eyebrow);
            var metrics = new CardRowMetrics {
                TitleSize = cards.Select((card, index) => context.WordFitSize(card.Title, heading, tiles[index].Width - 40D, heading.Minimum)).Min(),
                BodySize = bodySize ?? small.Size,
                ImageHeight = cards.Any(card => !string.IsNullOrWhiteSpace(card.Image)) ? Math.Min(90D, rowHeight * 0.24D) : 0D
            };
            metrics.TitleHeight = cards.Select((card, index) => CardTextHeight(context, card.Title, heading, tiles[index].Width - 40D, metrics.TitleSize)).Max();
            metrics.TextOffset = 23D + eyebrow.Size * eyebrow.LineHeight + 16D + (metrics.ImageHeight > 0D ? metrics.ImageHeight + 12D : 0D);
            metrics.AvailableHeight = Math.Max(0D, rowHeight - 20D - metrics.TextOffset);
            double Height() => cards.Select((card, index) => {
                double width = tiles[index].Width - 40D;
                double height = metrics.TitleHeight;
                double lineHeight = Math.Ceiling(metrics.BodySize * small.LineHeight);
                bool wrapped = false;
                if (!string.IsNullOrWhiteSpace(card.Text)) {
                    double textHeight = CardTextHeight(context, card.Text!, small, width, metrics.BodySize);
                    height += 8D + textHeight;
                    wrapped |= textHeight > lineHeight + 0.01D;
                }
                if (card.Items.Count > 0) {
                    height += 10D + 3D * (card.Items.Count - 1);
                    foreach (string item in card.Items) {
                        double itemHeight = CardTextHeight(context, item, small, width - 16D, metrics.BodySize);
                        height += itemHeight;
                        wrapped |= itemHeight > lineHeight + 0.01D;
                    }
                }
                // Native font metrics can wrap one additional body line compared
                // with the deterministic layout measurer. Reserve that line before
                // accepting a row, rather than overflowing its saved text frame.
                if (wrapped) height += lineHeight;
                return height;
            }).Max();
            bool WidthFits() => cards.Select((card, index) => {
                double width = tiles[index].Width - 40D;
                return PowerPointTextMeasure.LongestWord(card.Title, metrics.TitleSize, context.FontOf(heading), heading.Bold) * 1.05D <= width &&
                    PowerPointTextMeasure.LongestWord(card.Text ?? string.Empty, metrics.BodySize, context.FontOf(small)) * 1.05D <= width &&
                    card.Items.All(item => PowerPointTextMeasure.LongestWord(item, metrics.BodySize, context.FontOf(small)) * 1.05D <= width - 16D);
            }).All(fits => fits);
            metrics.Height = Height();
            metrics.WidthFits = WidthFits();
            while ((!metrics.WidthFits || metrics.Height > metrics.AvailableHeight) && !bodySize.HasValue && metrics.BodySize - 1D >= small.Minimum) {
                metrics.BodySize -= 1D;
                metrics.Height = Height();
                metrics.WidthFits = WidthFits();
            }
            return metrics;
        }

        /// <summary>Reduces card density before rendering when minimum-size content needs a continuation.</summary>
        private static int CardsPerSlide(PowerPointLayoutContext context, PowerPointLayoutSlide slide) {
            List<PowerPointLayoutCard> cards = slide.Items<PowerPointLayoutCard>().ToList();
            int maximum = cards.Any(card => !string.IsNullOrWhiteSpace(card.Image)) ? 3 : 6;
            if (cards.Count == 0) return maximum;
            PowerPointLayoutRect zone = ContentZone(context, MeasureHeader(context, slide));
            PowerPointLayoutGrid grid = PowerPointLayoutGrid.ForRegion(context.Grid, zone);
            for (int perSlide = maximum; perSlide > 1; perSlide--) {
                int pages = (int)Math.Ceiling(cards.Count / (double)perSlide);
                bool fits = true;
                foreach (List<PowerPointLayoutCard> chunk in Balance(cards, pages)) {
                    List<int[]> rows = CardRows(chunk.Count, grid.IsNarrow);
                    double rowHeight = rows.Count == 1 ? zone.Height : (zone.Height - TileGap) / 2D;
                    int offset = 0;
                    foreach (int[] spans in rows) {
                        List<(double X, double Width)> tiles = Tiles(grid, spans, false, spans.Length);
                        List<PowerPointLayoutCard> row = chunk.Skip(offset).Take(tiles.Count).ToList();
                        CardRowMetrics metrics = MeasureCardRow(context, row, tiles, rowHeight);
                        if (!metrics.Fits) { fits = false; break; }
                        offset += row.Count;
                    }
                    if (!fits) break;
                }
                if (fits) return perSlide;
            }
            return 1;
        }
    }
}
