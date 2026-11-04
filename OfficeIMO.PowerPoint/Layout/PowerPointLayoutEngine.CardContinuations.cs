using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointLayoutEngine {
        /// <summary>Continues individual oversized cards after reducing the grid's density.</summary>
        private static IEnumerable<PowerPointLayoutSlide> SplitCards(PowerPointLayoutContext context, PowerPointLayoutSlide source) {
            var pages = new List<PowerPointLayoutSlide>();
            foreach (PowerPointLayoutSlide part in SplitTiles<PowerPointLayoutCard>(context, source, CardsPerSlide(context, source))) {
                List<PowerPointLayoutCard> cards = part.Items<PowerPointLayoutCard>().ToList();
                if (cards.Count != 1 || SingleCardFits(context, part, cards[0])) { pages.Add(part); continue; }
                foreach (PowerPointLayoutCard fragment in CardFragments(context, source, part, cards[0], out double bodySize)) {
                    PowerPointLayoutSlide page = part.CloneHeader();
                    page.Layout = PowerPointLayout.Cards;
                    page.CardStart = part.CardStart;
                    page.CardBodySize = bodySize;
                    page.Blocks.AddRange(part.Blocks.Where(block => block is not PowerPointLayoutCard));
                    page.Blocks.Add(fragment);
                    pages.Add(page);
                }
            }
            for (int i = 0; i < pages.Count; i++) {
                PowerPointLayoutSlide page = pages[i];
                page.Subtitle = i == 0 ? source.Subtitle : null;
                page.Notes = i == 0 ? source.Notes : null;
                page.Takeaway = i == 0 ? source.Takeaway : null;
                if (i == 0) page.ContinuedSuffix = source.ContinuedSuffix;
                else ContinueTitle(context, page, source.Title, i + 1, pages.Count);
                // Final continuation titles can change the live area's height. Reject
                // an impossible header instead of enlarging a text frame past its panel.
                var cards = page.Items<PowerPointLayoutCard>().ToList();
                if (cards.Count == 1 && !SingleCardFits(context, page, cards[0]))
                    throw new ArgumentException("The card heading or slide header leaves insufficient space for its content. Shorten the heading or header.");
                yield return page;
            }
        }

        private static bool SingleCardFits(PowerPointLayoutContext context, PowerPointLayoutSlide slide, PowerPointLayoutCard card) {
            PowerPointLayoutRect zone = ContentZone(context, MeasureHeader(context, slide));
            var tiles = Tiles(PowerPointLayoutGrid.ForRegion(context.Grid, zone), new[] { 6 }, false, 1);
            return MeasureCardRow(context, new[] { card }, tiles, zone.Height, slide.CardBodySize).Fits;
        }

        /// <summary>Preserves body and bullet order while splitting at measured line boundaries.</summary>
        private static List<PowerPointLayoutCard> CardFragments(PowerPointLayoutContext context, PowerPointLayoutSlide source,
            PowerPointLayoutSlide part, PowerPointLayoutCard card, out double bodySize) {
            PowerPointLayoutRect zone = ContentZone(context, MeasureHeader(context, part));
            PowerPointLayoutRect continued = ContentZone(context, MeasureHeader(context, Continuation(context, source, 2, 2)));
            double height = Math.Min(zone.Height, continued.Height);
            var tiles = Tiles(PowerPointLayoutGrid.ForRegion(context.Grid, zone), new[] { 6 }, false, 1);
            double width = tiles[0].Width - 40D;
            PowerPointTextStyleSpec small = context.Style(PowerPointTextRole.BodySmall);
            bodySize = small.Size;
            bool WordsFit(double size) => PowerPointTextMeasure.LongestWord(card.Text ?? string.Empty, size, context.FontOf(small)) * 1.05D <= width &&
                card.Items.All(item => PowerPointTextMeasure.LongestWord(item, size, context.FontOf(small)) * 1.05D <= width - 16D);
            while (!WordsFit(bodySize) && bodySize - 1D >= small.Minimum) bodySize -= 1D;
            if (!WordsFit(bodySize)) throw new ArgumentException("A card word is wider than its panel at the minimum text size. Shorten the word.");
            var lines = new List<(string Text, int Item)>();
            if (!string.IsNullOrWhiteSpace(card.Text)) lines.AddRange(PowerPointTextMeasure.Wrap(card.Text!, bodySize, width / 1.06D,
                context.FontOf(small)).Select(text => (text, -1)));
            for (int i = 0; i < card.Items.Count; i++) lines.AddRange(PowerPointTextMeasure.Wrap(card.Items[i], bodySize,
                (width - 16D) / 1.06D, context.FontOf(small)).Select(text => (text, i)));
            PowerPointLayoutCard Fragment(IEnumerable<(string Text, int Item)> content) {
                var units = content.ToList();
                string text = string.Join("\n", units.Where(line => line.Item < 0).Select(line => line.Text));
                var items = units.Where(line => line.Item >= 0).GroupBy(line => line.Item).Select(group => string.Join("\n", group.Select(line => line.Text)));
                return new PowerPointLayoutCard(card.Title, text, items, card.Image, card.Color);
            }
            var fragments = new List<PowerPointLayoutCard>();
            int start = 0;
            while (start < lines.Count) {
                int count = 0;
                while (start + count < lines.Count && MeasureCardRow(context, new[] { Fragment(lines.Skip(start).Take(count + 1)) },
                    tiles, height, bodySize).Fits) count++;
                if (count == 0) throw new ArgumentException("The card heading leaves insufficient space for a body line at the minimum text size. Shorten the heading.");
                fragments.Add(Fragment(lines.Skip(start).Take(count)));
                start += count;
            }
            if (fragments.Count == 0) throw new ArgumentException("The card heading does not fit at its minimum text size. Shorten the heading.");
            return fragments;
        }
    }
}
