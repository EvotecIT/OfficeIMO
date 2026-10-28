using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointLayoutEngine {
        private const double LedeAllowance = 76D;
        private const double TableCaptionAllowance = 22D;
        private const string NoStyleTableId = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}";

        // ------------------------------------------------------------------ blocks and lede

        private static bool IsVisualOrList(PowerPointLayoutBlock block) =>
            block is PowerPointLayoutList or PowerPointLayoutTable or PowerPointLayoutChart or PowerPointLayoutImage or PowerPointLayoutColumn or PowerPointLayoutDrawing;

        /// <summary>A paragraph immediately followed by a list, table, chart or image is that block's lede.</summary>
        internal static (PowerPointLayoutText? Lede, List<PowerPointLayoutBlock> Blocks) SplitLede(IEnumerable<PowerPointLayoutBlock> source) {
            List<PowerPointLayoutBlock> blocks = source.Where(b => b.IsContentBlock).ToList();
            if (blocks.Count >= 2 && blocks[0] is PowerPointLayoutText lede && IsVisualOrList(blocks[1]) && lede.Paragraphs.Count == 1) {
                return (lede, blocks.Skip(1).ToList());
            }

            return (null, blocks);
        }

        private static string Continued(PowerPointLayoutContext context, int part, int count) {
            if (!string.IsNullOrWhiteSpace(context.Deck.ContinuationFormat)) {
                return " " + string.Format(CultureInfo.CurrentCulture, context.Deck.ContinuationFormat!, string.Empty, part, count).Trim();
            }

            return count <= 2 ? " (continued)" : " (continued " + part + "/" + count + ")";
        }

        private static List<List<T>> Balance<T>(IReadOnlyList<T> items, int pages) {
            var result = new List<List<T>>();
            int baseSize = items.Count / pages, extra = items.Count % pages, index = 0;
            for (int page = 0; page < pages; page++) {
                int size = baseSize + (page < extra ? 1 : 0);
                result.Add(items.Skip(index).Take(size).ToList());
                index += size;
            }

            return result;
        }

        private static PowerPointLayoutSlide Continuation(PowerPointLayoutContext context, PowerPointLayoutSlide source, int part, int count) {
            PowerPointLayoutSlide copy = source.CloneHeader();
            copy.Subtitle = null;
            copy.Notes = null;
            copy.ContinuedSuffix = Continued(context, part, count);
            return copy;
        }

        // ------------------------------------------------------------------ overflow

        internal static IEnumerable<PowerPointLayoutSlide> Expand(PowerPointLayoutContext context, PowerPointLayoutSlide slide, PowerPointLayout layout) {
            slide.Layout = layout;
            (PowerPointLayoutText? lede, List<PowerPointLayoutBlock> blocks) = SplitLede(slide.Blocks);
            switch (layout) {
                case PowerPointLayout.TwoContent when blocks.Count > 2: {
                        context.Deck.Report(PowerPointLayoutDiagnosticCodes.SlideContentSplit, OfficeDiagnosticSeverity.Information,
                            "The slide has " + blocks.Count + " content blocks; the first two stay and the rest continue on the next slide.", slide.Title);
                        PowerPointLayoutSlide first = slide.CloneHeader();
                        first.Notes = slide.Notes;
                        first.Subtitle = slide.Subtitle;
                        if (lede != null) first.Blocks.Add(lede);
                        first.Blocks.AddRange(blocks.Take(2));
                        first.Layout = PowerPointLayout.TwoContent;
                        yield return first;
                        PowerPointLayoutSlide rest = Continuation(context, slide, 2, 2);
                        rest.Blocks.AddRange(blocks.Skip(2));
                        rest.Layout = PowerPointLayout.Auto;
                        PowerPointLayout next = Select(rest, false, false, context.Document.Slides.Count + 2);
                        foreach (PowerPointLayoutSlide s in Expand(context, rest, next)) yield return s;
                        yield break;
                    }
                case PowerPointLayout.TitleAndContent when blocks.Count == 1 && blocks[0] is PowerPointLayoutList list:
                    foreach (PowerPointLayoutSlide s in SplitList(context, slide, list, lede)) yield return s;
                    yield break;
                case PowerPointLayout.TitleAndContent when blocks.Count == 1 && blocks[0] is PowerPointLayoutTable table:
                    foreach (PowerPointLayoutSlide s in SplitTable(context, slide, table, lede, context.Grid.W(12))) yield return s;
                    yield break;
                case PowerPointLayout.TwoContent when blocks.Count == 2 && blocks.OfType<PowerPointLayoutTable>().Count() == 1: {
                        PowerPointLayoutTable twoTable = blocks.OfType<PowerPointLayoutTable>().Single();
                        PowerPointLayoutBlock other = blocks.Single(b => !ReferenceEquals(b, twoTable));
                        double width = context.Grid.W(IsTexty(other) ? 7 : 6);
                        foreach (PowerPointLayoutSlide s in SplitTable(context, slide, twoTable, lede, width)) yield return s;
                        yield break;
                    }
                case PowerPointLayout.TitleAndContent when blocks.Count == 1 && blocks[0] is PowerPointLayoutText text:
                    foreach (PowerPointLayoutSlide s in SplitText(context, slide, text)) yield return s;
                    yield break;
                case PowerPointLayout.Kpi:
                    foreach (PowerPointLayoutSlide s in SplitTiles<PowerPointLayoutKpi>(context, slide, 6)) yield return s;
                    yield break;
                case PowerPointLayout.Cards:
                    int cardLimit = CardsPerSlide(context, slide);
                    foreach (PowerPointLayoutSlide s in SplitTiles<PowerPointLayoutCard>(context, slide, cardLimit)) yield return s;
                    yield break;
                case PowerPointLayout.Logos:
                    foreach (PowerPointLayoutSlide s in SplitTiles<PowerPointLayoutCard>(context, slide, 12)) yield return s;
                    yield break;
                case PowerPointLayout.Process:
                    foreach (PowerPointLayoutSlide s in SplitTiles<PowerPointLayoutStep>(context, slide, 8)) yield return s;
                    yield break;
                case PowerPointLayout.Coverage:
                    foreach (PowerPointLayoutSlide s in SplitTiles<PowerPointLayoutLocation>(context, slide, 18)) yield return s;
                    yield break;
                default:
                    yield return slide;
                    yield break;
            }
        }

        private static IEnumerable<PowerPointLayoutSlide> SplitTiles<T>(PowerPointLayoutContext context, PowerPointLayoutSlide slide, int perSlide)
            where T : PowerPointLayoutBlock {
            List<T> items = slide.Items<T>().ToList();
            if (items.Count <= perSlide) {
                yield return slide;
                yield break;
            }

            int pages = (int)Math.Ceiling(items.Count / (double)perSlide);
            List<List<T>> chunks = Balance(items, pages);
            int start = 1;
            for (int page = 0; page < pages; page++) {
                PowerPointLayoutSlide part = page == 0 ? slide.CloneHeader() : Continuation(context, slide, page + 1, pages);
                if (page == 0) {
                    part.Subtitle = slide.Subtitle;
                    part.Notes = slide.Notes;
                    part.Takeaway = slide.Takeaway;
                }

                part.Layout = slide.Layout;
                part.Blocks.AddRange(slide.Blocks.Where(b => b is not T));
                part.Blocks.AddRange(chunks[page]);
                part.StepStart = part.CardStart = start;
                start += chunks[page].Count;
                yield return part;
            }
        }

        internal static double ListHeight(PowerPointLayoutContext context, IEnumerable<PowerPointLayoutListItem> items, double size, double width) {
            PowerPointTextStyleSpec body = context.Style(PowerPointTextRole.Body);
            double gap = context.Theme.ParagraphAfter * size / body.Size, height = 0D;
            int count = 0;
            foreach (PowerPointLayoutListItem item in items) {
                double itemSize = item.Level == 1 ? size : Math.Round(size * 0.86D);
                double indent = item.Level == 1 ? context.Theme.ListIndent : context.Theme.ListIndent + 22D;
                height += PowerPointTextMeasure.LineCount(item.Text, itemSize, width - indent, context.Theme.BodyFont) * itemSize * body.LineHeight + gap;
                count++;
            }

            return count == 0 ? 0D : height - gap;
        }

        private static PowerPointLayoutRect ContentZone(PowerPointLayoutContext context, HeaderGeometry header, int span = 12) =>
            context.Grid.Box(1, span, header.ContentRow, header.ContentRows);

        /// <summary>Writes the caption/source line at the bottom of the live area and returns the remaining zone.</summary>
        private static PowerPointLayoutRect RenderSourceNote(PowerPointLayoutContext context, PowerPointLayoutWriter writer,
            PowerPointLayoutSlide slide, PowerPointLayoutRect zone) {
            if (string.IsNullOrWhiteSpace(slide.SourceNote)) return zone;
            PowerPointTextStyleSpec caption = context.Style(PowerPointTextRole.Caption);
            double height = caption.Size * caption.LineHeight + 2D;
            writer.Text("Source", new PowerPointLayoutRect(zone.X, zone.Bottom - height, context.Grid.W(12), height),
                context.Paragraph(slide.SourceNote!, PowerPointTextRole.Caption));
            return zone.WithHeight(zone.Height - height - 12D);
        }

        private static IEnumerable<PowerPointLayoutSlide> SplitList(PowerPointLayoutContext context, PowerPointLayoutSlide slide,
            PowerPointLayoutList list, PowerPointLayoutText? lede) {
            PowerPointTextStyleSpec body = context.Style(PowerPointTextRole.Body);
            double width = context.Grid.W(8);
            double zone = ContentZone(context, MeasureHeader(context, slide)).Height - (lede != null ? LedeAllowance : 0D);
            for (double size = body.Size; size >= body.Minimum - 0.01D; size -= 2D) {
                if (ListHeight(context, list.Items, size, width) <= zone) {
                    slide.ListSize = size;
                    yield return slide;
                    yield break;
                }
            }

            // Split between level-1 items at the original size into the fewest balanced slides.
            var groups = new List<List<PowerPointLayoutListItem>>();
            foreach (PowerPointLayoutListItem item in list.Items) {
                if (item.Level > 1 && groups.Count > 0) groups[groups.Count - 1].Add(item);
                else groups.Add(new List<PowerPointLayoutListItem> { item });
            }

            List<List<List<PowerPointLayoutListItem>>> chunks = new();
            for (int pages = 2; pages <= Math.Max(2, groups.Count); pages++) {
                chunks = Balance(groups, pages);
                PowerPointLayoutSlide probe = Continuation(context, slide, 2, pages);
                double probeZone = ContentZone(context, MeasureHeader(context, probe)).Height;
                double firstZone = zone;
                bool fits = true;
                for (int i = 0; i < chunks.Count && fits; i++) {
                    fits = ListHeight(context, chunks[i].SelectMany(g => g), body.Size, width) <= (i == 0 ? firstZone : probeZone);
                }

                if (fits) break;
            }

            int start = 1;
            for (int page = 0; page < chunks.Count; page++) {
                List<PowerPointLayoutListItem> items = chunks[page].SelectMany(g => g).ToList();
                PowerPointLayoutSlide part = page == 0 ? slide.CloneHeader() : Continuation(context, slide, page + 1, chunks.Count);
                if (page == 0) {
                    part.Subtitle = slide.Subtitle;
                    part.Notes = slide.Notes;
                    if (lede != null) part.Blocks.Add(lede);
                }

                part.Layout = PowerPointLayout.TitleAndContent;
                part.ListSize = body.Size;
                part.ListStart = start;
                part.Blocks.Add(new PowerPointLayoutList(items, list.Numbered));
                start += chunks[page].Count;
                yield return part;
            }
        }

        private static IEnumerable<PowerPointLayoutSlide> SplitText(PowerPointLayoutContext context, PowerPointLayoutSlide slide, PowerPointLayoutText text) {
            PowerPointTextStyleSpec body = context.Style(PowerPointTextRole.Body);
            double width = context.Grid.W(8);
            double zone = ContentZone(context, MeasureHeader(context, slide)).Height;
            double Height(IEnumerable<string> paragraphs, double size) =>
                paragraphs.Sum(p => PowerPointTextMeasure.LineCount(p, size, width, context.Theme.BodyFont) * size * body.LineHeight + context.Theme.ParagraphAfter);
            for (double size = body.Size; size >= body.Minimum - 0.01D; size -= 2D) {
                if (Height(text.Paragraphs, size) <= zone + context.Theme.ParagraphAfter) {
                    slide.ListSize = size;
                    yield return slide;
                    yield break;
                }
            }

            // Split long paragraphs between sentences first, then balance paragraphs over slides.
            var units = new List<string>();
            foreach (string paragraph in text.Paragraphs) {
                if (Height(new[] { paragraph }, body.Size) <= zone) {
                    units.Add(paragraph);
                    continue;
                }

                context.Deck.Report(PowerPointLayoutDiagnosticCodes.TextSplitParagraph, OfficeDiagnosticSeverity.Information,
                    "A paragraph longer than the slide was split between sentences.", slide.Title);
                units.AddRange(Regex.Split(paragraph, @"(?<=[.!?])\s+").Where(s => s.Length > 0));
            }

            List<List<string>> chunks = new();
            for (int pages = 2; pages <= Math.Max(2, units.Count); pages++) {
                chunks = Balance(units, pages);
                if (chunks.All(c => Height(c, body.Size) <= zone + context.Theme.ParagraphAfter)) break;
            }

            for (int page = 0; page < chunks.Count; page++) {
                PowerPointLayoutSlide part = page == 0 ? slide.CloneHeader() : Continuation(context, slide, page + 1, chunks.Count);
                if (page == 0) {
                    part.Subtitle = slide.Subtitle;
                    part.Notes = slide.Notes;
                }

                part.Layout = PowerPointLayout.TitleAndContent;
                part.ListSize = body.Size;
                part.Blocks.Add(new PowerPointLayoutText(chunks[page]));
                yield return part;
            }
        }

        // ------------------------------------------------------------------ tables

        internal sealed class TableMetrics {
            internal double BodySize { get; set; }
            internal double HeaderSize { get; set; }
            internal double RowHeight { get; set; }
            internal double HeaderHeight { get; set; }
            internal double[] Widths { get; set; } = Array.Empty<double>();
            internal bool[] Numeric { get; set; } = Array.Empty<bool>();
        }

        private static readonly Regex NumberPattern = new(@"^\s*[\-−+]?[$€£]?\s?[\d.,]+\s?(%|[kKmM]|pt|h|ms)?\s*$", RegexOptions.CultureInvariant);

        internal static bool IsNumeric(string? value) => !string.IsNullOrWhiteSpace(value) && NumberPattern.IsMatch(value!);

        internal static TableMetrics MeasureTable(PowerPointLayoutContext context, PowerPointLayoutTable table, double width, double? bodySize = null) {
            PowerPointTextStyleSpec body = context.Style(PowerPointTextRole.TableBody);
            PowerPointTextStyleSpec header = context.Style(PowerPointTextRole.TableHeader);
            double pad = context.Theme.TableCellPadding;
            var metrics = new TableMetrics {
                BodySize = bodySize ?? body.Size,
                HeaderSize = Math.Min(header.Size, (bodySize ?? body.Size) - 1D)
            };
            metrics.RowHeight = metrics.BodySize * body.LineHeight + 2D * (pad + 2D);
            metrics.HeaderHeight = table.IncludeHeaders ? metrics.HeaderSize * header.LineHeight + 2D * (pad + 3D) : 0D;
            int columns = table.Headers.Count;
            var natural = new double[columns];
            metrics.Numeric = new bool[columns];
            for (int c = 0; c < columns; c++) {
                double w = table.IncludeHeaders ? PowerPointTextMeasure.Width(table.Headers[c], metrics.HeaderSize, context.Theme.BodyFont, true) : 0D;
                bool numeric = table.Rows.Count > 0;
                foreach (IReadOnlyList<string> row in table.Rows) {
                    string value = c < row.Count ? row[c] ?? string.Empty : string.Empty;
                    w = Math.Max(w, PowerPointTextMeasure.Width(value, metrics.BodySize, context.Theme.BodyFont));
                    if (!IsNumeric(value) && !string.IsNullOrWhiteSpace(value)) numeric = false;
                }

                metrics.Numeric[c] = numeric;
                natural[c] = w + 2D * (pad + 5D) + 4D;
            }

            double total = natural.Sum();
            if (total <= width) {
                // Spread the spare width over text columns so numbers stay compact.
                double spare = width - total;
                int text = metrics.Numeric.Count(n => !n);
                metrics.Widths = natural.Select((w, i) => w + (text == 0 ? spare / columns : metrics.Numeric[i] ? 0D : spare / text)).ToArray();
            } else {
                // Squeeze the widest columns first; no column narrower than its widest word.
                double minimum = width / columns * 0.5D;
                metrics.Widths = natural.Select(w => Math.Max(minimum, w * width / total)).ToArray();
                double scale = width / metrics.Widths.Sum();
                metrics.Widths = metrics.Widths.Select(w => w * scale).ToArray();
            }

            return metrics;
        }

        /// <summary>Rows a table row occupies toward the slide limit: one per line, wrapping at most two lines.</summary>
        private static int RowUnits(PowerPointLayoutContext context, TableMetrics metrics, IReadOnlyList<string> row) {
            double pad = context.Theme.TableCellPadding;
            int lines = 1;
            for (int c = 0; c < metrics.Widths.Length && c < row.Count; c++) {
                lines = Math.Max(lines, PowerPointTextMeasure.LineCount(row[c] ?? string.Empty, metrics.BodySize,
                    metrics.Widths[c] - 2D * (pad + 5D), context.Theme.BodyFont));
            }

            return Math.Min(2, lines);
        }

        private static int TableCapacity(PowerPointLayoutContext context, PowerPointLayoutSlide slide, TableMetrics metrics, bool lede) {
            double zone = ContentZone(context, MeasureHeader(context, slide)).Height - (lede ? LedeAllowance : 0D);
            return Math.Max(1, (int)Math.Floor((zone - metrics.HeaderHeight - TableCaptionAllowance) / metrics.RowHeight));
        }

        private static IEnumerable<PowerPointLayoutSlide> SplitTable(PowerPointLayoutContext context, PowerPointLayoutSlide slide,
            PowerPointLayoutTable table, PowerPointLayoutText? lede, double width) {
            TableMetrics metrics = MeasureTable(context, table, width);
            int maxColumns = context.Grid.IsNarrow ? 6 : 8;
            if (table.Headers.Count > maxColumns) {
                TableMetrics small = MeasureTable(context, table, width, context.Style(PowerPointTextRole.TableBody).Minimum);
                if (small.Widths.Sum() > width + 0.5D || table.Headers.Count > maxColumns + 4) {
                    context.Deck.Report(PowerPointLayoutDiagnosticCodes.TableTooWide, OfficeDiagnosticSeverity.Warning,
                        "The table has " + table.Headers.Count + " columns; choose fewer columns (-Property) to keep it readable.", slide.Title);
                }

                metrics = small;
            }

            List<int> units = table.Rows.Select(row => RowUnits(context, metrics, row)).ToList();
            int first = table.MaximumRows ?? TableCapacity(context, slide, metrics, lede != null);
            if (table.NoSplit || units.Sum() <= first && table.Rows.Count <= first) {
                if (table.NoSplit && units.Sum() > first) {
                    context.Deck.Report(PowerPointLayoutDiagnosticCodes.TableOverflow, OfficeDiagnosticSeverity.Warning,
                        "The table does not fit one slide; it was kept together because splitting was turned off.", slide.Title);
                }

                yield return slide;
                yield break;
            }

            PowerPointLayoutSlide probe = Continuation(context, slide, 2, 3);
            int limit = table.MaximumRows ?? Math.Min(first, TableCapacity(context, probe, metrics, false));
            int pages = (int)Math.Ceiling(table.Rows.Count / (double)limit);
            List<List<int>> chunks;
            while (true) {
                chunks = Balance(Enumerable.Range(0, table.Rows.Count).ToList(), pages);
                if (table.MaximumRows.HasValue || chunks.All(c => c.Sum(i => units[i]) <= limit)) break;
                pages++;
            }

            for (int page = 0; page < pages; page++) {
                List<int> rows = chunks[page];
                PowerPointLayoutSlide part = page == 0 ? slide.CloneHeader() : Continuation(context, slide, page + 1, pages);
                if (page == 0) {
                    part.Subtitle = slide.Subtitle;
                    part.Notes = slide.Notes;
                    if (lede != null) part.Blocks.Add(lede);
                }

                part.Layout = slide.Layout;
                part.TableRange = (rows[0] + 1, rows[rows.Count - 1] + 1, table.Rows.Count);
                PowerPointLayoutTable pageTable = table.WithRows(rows.Select(i => table.Rows[i]).ToList());
                foreach (PowerPointLayoutBlock block in slide.Blocks.Where(b => b.IsContentBlock && !ReferenceEquals(b, lede))) {
                    part.Blocks.Add(ReferenceEquals(block, table) ? pageTable : block);
                }

                part.SourceNote = slide.SourceNote;
                yield return part;
            }
        }

        private static string Emu(double points) => ((long)Math.Round(points * 12700D)).ToString(CultureInfo.InvariantCulture);

        /// <summary>Writes a native table styled from the theme table style; frame and grid always agree.</summary>
        internal static PowerPointTable RenderTable(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutTable table,
            PowerPointLayoutRect box, (int First, int Last, int Total)? range) {
            PowerPointLayoutTheme theme = context.Theme;
            TableMetrics metrics = MeasureTable(context, table, box.Width,
                table.Headers.Count > (context.Grid.IsNarrow ? 6 : 8) ? context.Style(PowerPointTextRole.TableBody).Minimum : null);
            double pad = theme.TableCellPadding;
            bool tintedHeader = PowerPointLayoutTheme.Contrast(theme.TableHeaderFill, theme.Background) < 1.6D;
            string hairline = Emu(Math.Max(0.25D, theme.TableBorderWidth));

            string Line(string tag, string? color, string width) => color == null
                ? "<a:" + tag + " w=\"0\"><a:noFill/></a:" + tag + ">"
                : "<a:" + tag + " w=\"" + width + "\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill>" + writer.ColorXml(color) + "</a:solidFill><a:prstDash val=\"solid\"/></a:" + tag + ">";

            string Cell(string text, bool header, int column, string? fill, double topPad, string bottomColor, string bottomWidth) {
                double size = header ? metrics.HeaderSize : metrics.BodySize;
                string color = header ? theme.TableHeaderText : theme.Text;
                if (!header && (text == "Outage" || text == "Failed" || text == "At risk")) color = theme.Danger;
                if (!header && (text == "Degraded" || text == "Warning")) color = theme.Warning;
                string align = metrics.Numeric[column] ? "r" : "l";
                var b = new StringBuilder("<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn=\"").Append(align).Append("\"><a:lnSpc><a:spcPct val=\"100000\"/></a:lnSpc><a:buNone/></a:pPr>");
                b.Append("<a:r><a:rPr lang=\"en-US\" sz=\"").Append(((int)Math.Round(size * 100D)).ToString(CultureInfo.InvariantCulture))
                    .Append("\" b=\"").Append(header ? '1' : '0').Append("\" dirty=\"0\"><a:solidFill>").Append(writer.ColorXml(color))
                    .Append("</a:solidFill><a:latin typeface=\"+mn-lt\"/></a:rPr><a:t>").Append(SecurityElement.Escape(text ?? string.Empty)).Append("</a:t></a:r></a:p></a:txBody>");
                b.Append("<a:tcPr marL=\"").Append(Emu(pad + 5D)).Append("\" marR=\"").Append(Emu(pad + 5D)).Append("\" marT=\"").Append(Emu(topPad))
                    .Append("\" marB=\"").Append(Emu(topPad)).Append("\" anchor=\"ctr\">");
                b.Append(Line("lnL", null, "0")).Append(Line("lnR", null, "0")).Append(Line("lnT", null, "0")).Append(Line("lnB", bottomColor, bottomWidth));
                b.Append(fill == null ? "<a:noFill/>" : "<a:solidFill>" + writer.ColorXml(fill) + "</a:solidFill>");
                b.Append("</a:tcPr></a:tc>");
                return b.ToString();
            }

            var heights = new List<double>();
            if (table.IncludeHeaders) heights.Add(metrics.HeaderHeight);
            var xml = new StringBuilder("<a:tbl xmlns:a=\"").Append(PowerPointLayoutMasterWriter.NsA).Append("\"><a:tblPr firstRow=\"").Append(table.IncludeHeaders ? '1' : '0').Append("\" bandRow=\"")
                .Append(theme.TableBanding ? '1' : '0').Append("\"><a:tableStyleId>").Append(NoStyleTableId).Append("</a:tableStyleId></a:tblPr><a:tblGrid>");
            foreach (double w in metrics.Widths) xml.Append("<a:gridCol w=\"").Append(Emu(w)).Append("\"/>");
            xml.Append("</a:tblGrid>");
            if (table.IncludeHeaders) {
                xml.Append("<a:tr h=\"").Append(Emu(metrics.HeaderHeight)).Append("\">");
                for (int c = 0; c < table.Headers.Count; c++) {
                    xml.Append(Cell(table.Headers[c], true, c, theme.TableHeaderFill, pad + 3D, tintedHeader ? theme.Accent1 : theme.TableHeaderFill,
                        tintedHeader ? Emu(1D) : "0"));
                }
                xml.Append("</a:tr>");
            }
            for (int r = 0; r < table.Rows.Count; r++) {
                IReadOnlyList<string> row = table.Rows[r];
                double height = metrics.RowHeight + (RowUnits(context, metrics, row) - 1) * metrics.BodySize * 1.2D;
                heights.Add(height);
                string? fill = theme.TableBanding && r % 2 == 1 ? theme.Surface : null;
                xml.Append("<a:tr h=\"").Append(Emu(height)).Append("\">");
                for (int c = 0; c < table.Headers.Count; c++) {
                    xml.Append(Cell(c < row.Count ? row[c] ?? string.Empty : string.Empty, false, c, fill, pad + 2D, theme.Hairline, hairline));
                }
                xml.Append("</a:tr>");
            }
            xml.Append("</a:tbl>");

            double width = metrics.Widths.Sum();
            double total = heights.Sum();
            PowerPointTable result = writer.Slide.AddTable(Math.Max(1, table.Rows.Count + (table.IncludeHeaders ? 1 : 0)), Math.Max(1, table.Headers.Count),
                PowerPointUnits.FromPoints(box.X), PowerPointUnits.FromPoints(box.Y), PowerPointUnits.FromPoints(width), PowerPointUnits.FromPoints(total));
            A.Table element = result.TableElement;
            A.GraphicData data = (A.GraphicData)element.Parent!;
            element.Remove();
            data.AppendChild(new A.Table(xml.ToString()));

            if (range.HasValue) {
                (int a, int b, int n) = range.Value;
                PowerPointLayoutParagraph caption = context.Paragraph(
                    "Rows " + a.ToString(CultureInfo.CurrentCulture) + "–" + b.ToString(CultureInfo.CurrentCulture) + " of " + n.ToString(CultureInfo.CurrentCulture),
                    PowerPointTextRole.Caption, align: "r");
                writer.Text("Table Range", new PowerPointLayoutRect(box.X, box.Y + total + 8D, width, 16D), caption);
            } else if (!string.IsNullOrWhiteSpace(table.Caption)) {
                writer.Text("Table Caption", new PowerPointLayoutRect(box.X, box.Y + total + 8D, width, 16D),
                    context.Paragraph(table.Caption!, PowerPointTextRole.Caption));
            }

            return result;
        }

        // ------------------------------------------------------------------ lists, text, charts, images

        internal static PowerPointTextBox RenderList(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutList list,
            PowerPointLayoutRect box, double? size = null, int start = 1, string? placeholder = null, string? color = null) {
            PowerPointTextStyleSpec body = context.Style(PowerPointTextRole.Body);
            PowerPointTextStyleSpec small = context.Style(PowerPointTextRole.BodySmall);
            double s = size ?? body.Size;
            double gap = context.Theme.ParagraphAfter * s / body.Size;
            var paragraphs = new List<PowerPointLayoutParagraph>();
            int number = start;
            foreach (PowerPointLayoutListItem item in list.Items) {
                bool first = item.Level == 1;
                double itemSize = first ? s : Math.Round(s * 0.86D);
                var run = new PowerPointLayoutRun(item.Text, first ? body : small, itemSize, first ? color ?? body.Color : small.Color);
                var paragraph = new PowerPointLayoutParagraph(run) {
                    Level = item.Level,
                    LineHeight = body.LineHeight,
                    SpaceAfter = gap,
                    MarginLeft = first ? context.Theme.ListIndent : context.Theme.ListIndent + 22D,
                    Indent = -(first ? context.Theme.ListIndent : 18D),
                    Bullet = first ? "▪" : "–",
                    BulletColor = first ? context.Theme.Accent1 : context.Theme.Hairline
                };
                if (first && list.Numbered) {
                    paragraph.NumberStart = number++;
                    paragraph.MarginLeft = 28D;
                    paragraph.Indent = -28D;
                }

                paragraphs.Add(paragraph);
            }

            return writer.Text("List", box, paragraphs, placeholder: placeholder);
        }

        internal static PowerPointTextBox RenderText(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutText text,
            PowerPointLayoutRect box, double? size = null, string? placeholder = null) {
            PowerPointTextStyleSpec body = context.Style(PowerPointTextRole.Body);
            double s = size ?? body.Size;
            List<PowerPointLayoutParagraph> paragraphs = text.Paragraphs.Select(p => {
                PowerPointLayoutParagraph paragraph = context.Paragraph(p, PowerPointTextRole.Body, size: s);
                paragraph.SpaceAfter = context.Theme.ParagraphAfter * s / body.Size;
                return paragraph;
            }).ToList();
            return writer.Text("Text", box, paragraphs, placeholder: placeholder);
        }

        internal static void RenderChart(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutChart chart, PowerPointLayoutRect box) {
            PowerPointChart result = writer.Slide.AddChart(chart.Kind, chart.Data, PowerPointUnits.FromPoints(box.X), PowerPointUnits.FromPoints(box.Y),
                PowerPointUnits.FromPoints(box.Width), PowerPointUnits.FromPoints(box.Height), new PowerPointChartAccessibilityOptions {
                    Name = "Chart",
                    AlternativeText = chart.AlternativeText ?? chart.Title ?? "Chart",
                    IncludeDataSummaryInAlternativeText = true
                });
            // The slide title carries the message; a chart keeps at most a short label-sized title.
            if (!string.IsNullOrWhiteSpace(chart.Title)) {
                result.SetTitle(chart.Title!);
                result.SetTitleTextStyle(fontSizePoints: context.Style(PowerPointTextRole.Caption).Size + 3D, bold: true, color: context.Theme.Text);
            } else {
                result.ClearTitle();
            }

            // The chart sits on the slide surface: no chart-area fill or frame.
            result.SetChartAreaStyle(noFill: true, noLine: true);
            if (chart.Kind is OfficeChartKind.Pie or OfficeChartKind.Doughnut) {
                ContrastSliceLabels(context, result, chart);
                return;
            }

            // Series follow the theme accents in order; lines are drawn at 2 pt.
            for (int i = 0; i < chart.Data.Series.Count && i < 6; i++) {
                OfficeChartSeries series = chart.Data.Series[i];
                // The shared writer owns explicit colours, fills and outline visibility.
                if (series.Color.HasValue || series.SeriesStyle != null) continue;
                OfficeChartKind kind = series.RenderKind ?? chart.Kind;
                bool line = kind is OfficeChartKind.Line or OfficeChartKind.LineStacked or OfficeChartKind.LineStacked100
                    or OfficeChartKind.Scatter or OfficeChartKind.Radar;
                string accent = context.Theme.Accent(i + 1);
                if (line) {
                    if (series.ConnectLine) result.SetSeriesLineColor(i, accent, series.StrokeWidth ?? 2D);
                }
                else result.SetSeriesFillColor(i, accent);
            }

            // One series needs no legend; several get theirs at the top, under the title.
            DocumentFormat.OpenXml.Drawing.Charts.Legend? legend = result.GetChart().GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Legend>();
            if (chart.Data.Series.Count <= 1) {
                legend?.Remove();
                result.Save();
            } else {
                result.SetLegend(OfficeChartLegendPosition.Top);
            }
        }

        /// <summary>
        ///     Pie and doughnut labels sit on their slice: each takes white or the theme text colour, whichever reads
        ///     on that slice's accent fill.
        /// </summary>
        private static void ContrastSliceLabels(PowerPointLayoutContext context, PowerPointChart result, PowerPointLayoutChart chart) {
            bool labelled = result.GetChart().Descendants<DocumentFormat.OpenXml.Drawing.Charts.DataLabels>().Any(labels =>
                labels.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ShowValue>()?.Val?.Value == true ||
                labels.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ShowPercent>()?.Val?.Value == true);
            if (!labelled || chart.Data.Series.Count == 0) return;
            int points = chart.Data.Categories.Count;
            for (int i = 0; i < points; i++) {
                string fill = context.Theme.Accent(i % 6 + 1);
                string color = PowerPointLayoutTheme.Contrast("FFFFFF", fill) >= 4.5D ? "FFFFFF" : context.Theme.Text;
                result.SetSeriesDataLabelTextStyleForPoint(0, i, color: color);
            }
        }

        /// <summary>Pictures keep their aspect ratio: contained in the box (top-left) or cover-cropped around the focal point.</summary>
        internal static PowerPointPicture RenderImage(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutImage image,
            PowerPointLayoutRect box, bool cover) {
            PowerPointPicture picture = writer.Slide.AddPicture(image.Path, PowerPointUnits.FromPoints(box.X), PowerPointUnits.FromPoints(box.Y),
                PowerPointUnits.FromPoints(box.Width), PowerPointUnits.FromPoints(box.Height));
            if (!PowerPointImageSize.TryRead(image.Path, out int pw, out int ph)) return picture;
            if (cover) {
                picture.Left = PowerPointUnits.FromPoints(box.X);
                picture.Top = PowerPointUnits.FromPoints(box.Y);
                picture.Width = PowerPointUnits.FromPoints(box.Width);
                picture.Height = PowerPointUnits.FromPoints(box.Height);
                picture.CropToFill(pw, ph, image.FocalX, image.FocalY);
            } else {
                double scale = Math.Min(box.Width / pw, box.Height / ph);
                picture.Left = PowerPointUnits.FromPoints(box.X);
                picture.Top = PowerPointUnits.FromPoints(box.Y);
                picture.Width = PowerPointUnits.FromPoints(pw * scale);
                picture.Height = PowerPointUnits.FromPoints(ph * scale);
            }

            if (!string.IsNullOrWhiteSpace(image.AlternativeText)) picture.Description = image.AlternativeText;
            RenderAnnotations(context, writer, picture, image);
            return picture;
        }

        /// <summary>Numbered markers placed on the visible (cropped) picture, with their labels beside them.</summary>
        private static void RenderAnnotations(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointPicture picture,
            PowerPointLayoutImage image) {
            if (image.Annotations.Count == 0) return;
            double left = PowerPointUnits.ToPoints(picture.Left), top = PowerPointUnits.ToPoints(picture.Top);
            double width = PowerPointUnits.ToPoints(picture.Width), height = PowerPointUnits.ToPoints(picture.Height);
            double visibleX = Math.Max(0.0001D, 1D - picture.CropLeftRatio - picture.CropRightRatio);
            double visibleY = Math.Max(0.0001D, 1D - picture.CropTopRatio - picture.CropBottomRatio);
            PowerPointTextStyleSpec caption = context.Style(PowerPointTextRole.Caption);
            for (int i = 0; i < image.Annotations.Count; i++) {
                (double ax, double ay, string label) = image.Annotations[i];
                double x = left + (ax - picture.CropLeftRatio) / visibleX * width;
                double y = top + (ay - picture.CropTopRatio) / visibleY * height;
                PowerPointLayoutParagraph number = context.Paragraph((i + 1).ToString(CultureInfo.InvariantCulture), PowerPointTextRole.Caption,
                    context.Theme.OnAccent, 10D, bold: true);
                writer.Circle("Screenshot Annotation " + (i + 1), new PowerPointLayoutRect(x - 11D, y - 11D, 22D, 22D), context.Theme.Accent1,
                    context.Theme.Background, 2D, number, shape: true);
                const double labelInset = 3D;
                var metrics = OfficeDrawingTextLayout.CreateMetrics(null);
                var layout = OfficeTextLayoutEngine.LayoutTextBlock(label, caption.Size, double.MaxValue, double.MaxValue,
                    OfficeDrawingTextLayout.ResolveLineHeightFactor(caption.Size * caption.LineHeight, caption.Size),
                    Math.Min(6D, caption.Size),
                    (text, size) => metrics.MeasureText(text, size, context.Theme.BodyFont, OfficeFontStyle.Regular), wrap: false);
                double labelWidth = Math.Ceiling(layout.Width) + 2D * labelInset + 10D;
                double labelHeight = Math.Max(18D, Math.Ceiling(Math.Max(layout.Height, OfficeDrawingTextLayout.PaintedHeight(layout))) + 2D * labelInset);
                writer.Text("Screenshot Annotation Label " + (i + 1), new PowerPointLayoutRect(x + 15D, y - labelHeight / 2D, labelWidth, labelHeight),
                    new[] { context.Paragraph(label, PowerPointTextRole.Caption, context.Theme.Text) }, anchor: "ctr", fill: context.Theme.Background, inset: labelInset);
            }
        }

    }
}
