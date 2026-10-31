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
            if ((layout == PowerPointLayout.TwoContent && blocks.Count is 1 or 2 && !blocks.Any(b => b is PowerPointLayoutTable)
                    || layout == PowerPointLayout.Comparison && blocks.Count == 2
                    || layout == PowerPointLayout.TitleAndContent && blocks.Count == 1 && blocks[0] is PowerPointLayoutColumn)
                && blocks.Any(HasNativeText)) {
                foreach (var part in SplitNativeColumns(context, slide, blocks, lede)) yield return part;
                yield break;
            }
            if (layout == PowerPointLayout.Closing || layout == PowerPointLayout.Quote) {
                foreach (var part in SplitNativeSpecial(context, slide)) yield return part;
                yield break;
            }
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
                    foreach (PowerPointLayoutSlide s in SplitCards(context, slide)) yield return s;
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
                height += item.NativeParagraph == null
                    ? PowerPointTextMeasure.LineCount(item.Text, itemSize, width - indent, context.Theme.BodyFont) * itemSize * body.LineHeight + gap
                    : PowerPointLayoutNativeTextMeasure.Height(context, item.NativeParagraph, itemSize, width, Math.Max(28D, indent), gap);
                count++;
            }

            return count == 0 ? 0D : height - (items.Last().NativeParagraph == null ? gap : 0D);
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
            if (list.NativeBody != null) {
                foreach (var part in SplitNativeContent(context, slide, list.NativeBody, list, lede)) yield return part;
                yield break;
            }
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
                part.Blocks.Add(new PowerPointLayoutList(items, list.Numbered) { NativeBody = list.NativeBody?.Slice(items.Select(i => i.NativeParagraph!).ToList()) });
                start += chunks[page].Count;
                yield return part;
            }
        }

        private static IEnumerable<PowerPointLayoutSlide> SplitText(PowerPointLayoutContext context, PowerPointLayoutSlide slide, PowerPointLayoutText text) {
            if (text.NativeBody != null) {
                foreach (var part in SplitNativeContent(context, slide, text.NativeBody)) yield return part;
                yield break;
            }
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
            var nativeUnits = new List<PowerPointParagraph>();
            int paragraphIndex = 0;
            foreach (string paragraph in text.Paragraphs) {
                var native = text.NativeBody?.Paragraphs[paragraphIndex++];
                if (Height(new[] { paragraph }, body.Size) <= zone) {
                    units.Add(paragraph);
                    if (native != null) nativeUnits.Add(native);
                    continue;
                }

                context.Deck.Report(PowerPointLayoutDiagnosticCodes.TextSplitParagraph, OfficeDiagnosticSeverity.Information,
                    "A paragraph longer than the slide was split between sentences.", slide.Title);
                var pieces = Regex.Split(paragraph, @"(?<=[.!?])\s+").Where(s => s.Length > 0).ToList();
                units.AddRange(pieces);
                if (native != null) nativeUnits.AddRange(PowerPointLayoutNativeText.SplitParagraph(native, pieces));
            }

            List<List<string>> chunks = new();
            for (int pages = 2; pages <= Math.Max(2, units.Count); pages++) {
                chunks = Balance(units, pages);
                if (chunks.All(c => Height(c, body.Size) <= zone + context.Theme.ParagraphAfter)) break;
            }

            int unitStart = 0;
            for (int page = 0; page < chunks.Count; page++) {
                PowerPointLayoutSlide part = page == 0 ? slide.CloneHeader() : Continuation(context, slide, page + 1, chunks.Count);
                if (page == 0) {
                    part.Subtitle = slide.Subtitle;
                    part.Notes = slide.Notes;
                }

                part.Layout = PowerPointLayout.TitleAndContent;
                part.ListSize = body.Size;
                part.Blocks.Add(new PowerPointLayoutText(chunks[page]) { NativeBody = text.NativeBody?.Slice(nativeUnits.Skip(unitStart).Take(chunks[page].Count).ToList()) });
                unitStart += chunks[page].Count;
                yield return part;
            }
        }

        // ------------------------------------------------------------------ tables

        internal sealed class TableMetrics {
            internal double BodySize { get; set; }
            internal double HeaderSize { get; set; }
            internal double RowHeight { get; set; }
            internal double HeaderHeight { get; set; }

            /// <summary>Top and bottom cell padding of body and header rows.</summary>
            internal double RowPad { get; set; }
            internal double HeaderPad { get; set; }
            internal double[] Widths { get; set; } = Array.Empty<double>();
            internal bool[] Numeric { get; set; } = Array.Empty<bool>();
        }

        private static readonly Regex NumberPattern = new(@"^\s*[\-−+]?[$€£]?\s?[\d.,]+\s?(%|[kKmM]|pt|h|ms)?\s*$", RegexOptions.CultureInvariant);

        internal static bool IsNumeric(string? value) => !string.IsNullOrWhiteSpace(value) && NumberPattern.IsMatch(value!);

        internal static TableMetrics MeasureTable(PowerPointLayoutContext context, PowerPointLayoutTable table, double width, double? bodySize = null) {
            PowerPointTextStyleSpec body = context.Style(PowerPointTextRole.TableBody);
            PowerPointTextStyleSpec header = context.Style(PowerPointTextRole.TableHeader);
            double pad = context.Theme.TableCellPadding;
            // Compact rows: two points smaller (not below the minimum) and a quarter of the vertical padding, the density
            // of a reference table, used when it keeps a long table on fewer slides.
            double size = bodySize ?? body.Size;
            if (table.Dense) size = Math.Min(size, Math.Max(body.Minimum, body.Size - 2D));
            var metrics = new TableMetrics {
                BodySize = size,
                HeaderSize = Math.Min(header.Size, size - 1D),
                RowPad = table.Dense ? pad * 0.25D + 1.5D : pad + 2D,
                HeaderPad = table.Dense ? pad * 0.5D + 1.5D : pad + 3D
            };
            metrics.RowHeight = metrics.BodySize * body.LineHeight + 2D * metrics.RowPad;
            metrics.HeaderHeight = table.IncludeHeaders ? metrics.HeaderSize * header.LineHeight + 2D * metrics.HeaderPad : 0D;
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
            PowerPointLayoutTable source = table;
            if (!table.MaximumRows.HasValue && !table.Dense) {
                // A long reference table reads better on fewer, fuller slides: switch to compact rows when that
                // saves a slide.
                PowerPointLayoutTable dense = table.WithRows(table.Rows);
                dense.Dense = true;
                TableMetrics compact = MeasureTable(context, dense, width, metrics.BodySize < context.Style(PowerPointTextRole.TableBody).Size ? metrics.BodySize : null);
                List<int> denseUnits = table.Rows.Select(row => RowUnits(context, compact, row)).ToList();
                int denseFirst = TableCapacity(context, slide, compact, lede != null);
                if (denseUnits.Sum() <= denseFirst && table.Rows.Count <= denseFirst) {
                    PowerPointLayoutSlide whole = slide.CloneHeader();
                    whole.Subtitle = slide.Subtitle;
                    whole.Notes = slide.Notes;
                    whole.Layout = slide.Layout;
                    whole.SourceNote = slide.SourceNote;
                    foreach (PowerPointLayoutBlock block in slide.Blocks) whole.Blocks.Add(ReferenceEquals(block, table) ? dense : block);
                    yield return whole;
                    yield break;
                }

                int denseLimit = Math.Min(denseFirst, TableCapacity(context, probe, compact, false));
                int densePages = (int)Math.Ceiling(table.Rows.Count / (double)denseLimit);
                if (densePages < pages) {
                    table = dense;
                    metrics = compact;
                    units = denseUnits;
                    limit = denseLimit;
                    pages = densePages;
                }
            }
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
                    part.Blocks.Add(ReferenceEquals(block, source) ? pageTable : page == 0 ? block : CopyRepeatedNativeContent(block));
                }

                part.SourceNote = slide.SourceNote;
                yield return part;
            }
        }

        /// <summary>Repeated table context is independent of the live nodes on the first physical slide.</summary>
        private static PowerPointLayoutBlock CopyRepeatedNativeContent(PowerPointLayoutBlock block) => block switch {
            PowerPointLayoutList list when list.NativeBody != null => new PowerPointLayoutList(Array.Empty<string>(), list.Numbered) {
                NativeBody = list.NativeBody.CopyForRepeat(), NativeStart = list.NativeStart
            },
            PowerPointLayoutText text when text.NativeBody != null => new PowerPointLayoutText(Array.Empty<string>()) {
                NativeBody = text.NativeBody.CopyForRepeat()
            },
            PowerPointLayoutChart chart when chart.NativeChart != null => CopyRepeatedNativeChart(chart),
            PowerPointLayoutImage image when image.NativePicture != null => CopyRepeatedNativeImage(image),
            PowerPointLayoutColumn column => new PowerPointLayoutColumn(column.Title, column.Blocks.Select(CopyRepeatedNativeContent).ToList()),
            _ => block
        };

        private static string Emu(double points) => ((long)Math.Round(points * 12700D)).ToString(CultureInfo.InvariantCulture);

        /// <summary>Writes a native table styled from the theme table style; frame and grid always agree.</summary>
        internal static PowerPointTable RenderTable(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutTable table,
            PowerPointLayoutRect box, (int First, int Last, int Total)? range) {
            PowerPointLayoutTheme theme = context.Theme;
            TableMetrics metrics = MeasureTable(context, table, box.Width,
                table.Headers.Count > (context.Grid.IsNarrow ? 6 : 8) ? context.Style(PowerPointTextRole.TableBody).Minimum : null);
            double rowPad = metrics.RowPad, headerPad = metrics.HeaderPad;
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
                    xml.Append(Cell(table.Headers[c], true, c, theme.TableHeaderFill, headerPad, tintedHeader ? theme.Accent1 : theme.TableHeaderFill,
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
                    xml.Append(Cell(c < row.Count ? row[c] ?? string.Empty : string.Empty, false, c, fill, rowPad, theme.Hairline, hairline));
                }
                xml.Append("</a:tr>");
            }
            xml.Append("</a:tbl>");

            double width = metrics.Widths.Sum();
            double total = heights.Sum();
            PowerPointTable result = writer.Slide.AddTable(Math.Max(1, table.Rows.Count + (table.IncludeHeaders ? 1 : 0)), Math.Max(1, table.Headers.Count),
                OfficeLength.Points(box.X), OfficeLength.Points(box.Y), OfficeLength.Points(width), OfficeLength.Points(total));
            result.Description = !string.IsNullOrWhiteSpace(table.Caption) ? table.Caption :
                "Table with " + table.Rows.Count.ToString(CultureInfo.InvariantCulture) + " data rows. Columns: " + string.Join(", ", table.Headers) + ".";
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

            return table.NativeBody?.Bind(writer.Slide, result) ?? result;
        }

        // ------------------------------------------------------------------ lists, text, charts, images

        internal static PowerPointTextBox RenderList(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutList list,
            PowerPointLayoutRect box, double? size = null, int start = 1, string? placeholder = null, string? color = null) {
            PowerPointTextStyleSpec body = context.Style(PowerPointTextRole.Body);
            PowerPointTextStyleSpec small = context.Style(PowerPointTextRole.BodySmall);
            double s = size ?? body.Size;
            double gap = context.Theme.ParagraphAfter * s / body.Size;
            var paragraphs = new List<PowerPointLayoutParagraph>();
            int number = list.NativeStart ?? start;
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

                paragraph.NativeParagraph = item.NativeParagraph;
                paragraph.NativeBody = list.NativeBody;
                paragraphs.Add(paragraph);
            }

            return writer.Text("List", box, paragraphs, placeholder: placeholder, nativeBody: list.NativeBody);
        }

        internal static PowerPointTextBox RenderText(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutText text,
            PowerPointLayoutRect box, double? size = null, string? placeholder = null) {
            PowerPointTextStyleSpec body = context.Style(PowerPointTextRole.Body);
            double s = size ?? body.Size;
            List<PowerPointLayoutParagraph> paragraphs = text.Paragraphs.Select((p, index) => {
                PowerPointLayoutParagraph paragraph = context.Paragraph(p, PowerPointTextRole.Body, size: s);
                paragraph.SpaceAfter = context.Theme.ParagraphAfter * s / body.Size;
                return PowerPointLayoutNativeText.Attach(paragraph, text.NativeBody, index);
            }).ToList();
            return writer.Text("Text", box, paragraphs, placeholder: placeholder, nativeBody: text.NativeBody);
        }

        internal static void RenderChart(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutChart chart, PowerPointLayoutRect box) {
            if (chart.NativeChart != null) {
                ArrangeNativeFigure(writer.Slide, chart.NativeChart, box);
                return;
            }
            // The chart label sits left-aligned above the plot as slide text (charts.title.align = left). A native
            // chart title is centred by every renderer, so the chart itself carries no title.
            if (!string.IsNullOrWhiteSpace(chart.Title)) {
                PowerPointTextStyleSpec caption = context.Style(PowerPointTextRole.Caption);
                double size = caption.Size + 3D;
                PowerPointLayoutParagraph label = context.Paragraph(chart.Title!, PowerPointTextRole.Caption, context.Theme.Text, size, bold: true);
                double height = Math.Ceiling(context.Height(chart.Title!, caption, box.Width, size)) + 2D;
                writer.Text("Chart Title", box.WithHeight(height), label);
                double offset = height + 8D;
                box = new PowerPointLayoutRect(box.X, box.Y + offset, box.Width, Math.Max(40D, box.Height - offset));
            }

            PowerPointChart result = writer.Slide.AddChart(chart.Kind, chart.Data, new PowerPointChartOptions {
                Left = OfficeLength.Points(box.X), Top = OfficeLength.Points(box.Y),
                Width = OfficeLength.Points(box.Width), Height = OfficeLength.Points(box.Height),
                Accessibility = new PowerPointChartAccessibilityOptions {
                    Name = "Chart",
                    AlternativeText = chart.AlternativeText ?? chart.Title ?? "Chart",
                    IncludeDataSummaryInAlternativeText = true
                }});
            result.ClearTitle();

            // The chart sits on the slide surface: no chart-area fill or frame.
            result.SetChartAreaStyle(noFill: true, noLine: true);
            if (chart.Kind is OfficeChartType.Pie or OfficeChartType.Doughnut) {
                // Readers get the share on each slice; the legend names the slices.
                result.SetDataLabels(showValue: false, showPercent: true);
                result.SetDataLabelNumberFormat("0%");
                result.SetDataLabelTextStyle(fontSize: OfficeLength.Points(SliceLabelSize), bold: true, color: SliceLabelColor(context, chart));
                result.SetLegend(OfficeChartLegendPosition.Bottom);
                result.SetLegendTextStyle(fontSize: OfficeLength.Points(ChartTextSize), color: context.Theme.TextMuted);
                return;
            }

            // Series follow the theme accents in order; lines are drawn at 2 pt.
            for (int i = 0; i < chart.Data.Series.Count && i < 6; i++) {
                OfficeSeries series = chart.Data.Series[i];
                // The shared writer owns explicit colours, fills and outline visibility.
                if (series.Color.HasValue || series.SeriesStyle != null) continue;
                OfficeChartType kind = series.RenderKind ?? chart.Kind;
                bool line = kind is OfficeChartType.Line or OfficeChartType.LineStacked or OfficeChartType.LineStacked100
                    or OfficeChartType.Scatter or OfficeChartType.Radar;
                string accent = context.Theme.Accent(i + 1);
                if (line) {
                    if (series.ConnectLine) result.SetSeriesLineColor(i, accent, OfficeLength.Points(series.StrokeWidth?.ToPoints() ?? 2D));
                }
                else result.SetSeriesFillColor(i, accent);
            }

            StyleCategoryChart(context, result, chart);

            // One series needs no legend; several get theirs at the top, under the title.
            DocumentFormat.OpenXml.Drawing.Charts.Legend? legend = result.GetChart().GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Legend>();
            if (chart.Data.Series.Count <= 1) {
                legend?.Remove();
                result.Save();
            } else {
                result.SetLegend(OfficeChartLegendPosition.Top);
                result.SetLegendTextStyle(fontSize: OfficeLength.Points(ChartTextSize), color: context.Theme.TextMuted);
            }
        }

        /// <summary>Axis and label text on layout charts (powerpoint-layouts.json, charts.axisText).</summary>
        private const double ChartTextSize = 11D;

        /// <summary>
        ///     The layout chart look for category charts: 11 pt muted axis text, hairline value gridlines on a rounded
        ///     scale of four or five steps, value labels when one series has few points, and straight lines with
        ///     open markers.
        /// </summary>
        private static void StyleCategoryChart(PowerPointLayoutContext context, PowerPointChart result, PowerPointLayoutChart chart) {
            if (chart.Kind is OfficeChartType.Scatter or OfficeChartType.Bubble or OfficeChartType.Radar) return;
            PowerPointLayoutTheme theme = context.Theme;
            result.SetCategoryAxisLabelTextStyle(fontSize: OfficeLength.Points(ChartTextSize), color: theme.TextMuted);
            result.SetValueAxisLabelTextStyle(fontSize: OfficeLength.Points(ChartTextSize), color: theme.TextMuted);
            result.SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: theme.Hairline, lineWidth: OfficeLength.Points(0.75D));
            UniformAxisText(result);

            bool secondary = chart.Data.Series.Any(series => series.AxisGroup == OfficeChartAxisGroup.Secondary);
            bool stacked100 = chart.Kind is OfficeChartType.ColumnStacked100 or OfficeChartType.BarStacked100 or
                OfficeChartType.LineStacked100 or OfficeChartType.AreaStacked100;
            bool single = chart.Data.Series.Count == 1 && chart.Data.Categories.Count <= 8 && !stacked100;
            if (!secondary && !stacked100 && NiceScale(chart, single) is (double minimum, double maximum, double step)) {
                result.SetValueAxisScale(minimum, maximum, step);
            }

            for (int i = 0; i < chart.Data.Series.Count && i < 6; i++) {
                OfficeSeries series = chart.Data.Series[i];
                OfficeChartType kind = series.RenderKind ?? chart.Kind;
                if (kind is OfficeChartType.Line or OfficeChartType.LineStacked or OfficeChartType.LineStacked100) {
                    if (series.Color == null && series.SeriesStyle == null) {
                        string accent = theme.Accent(i + 1);
                        if (series.ShowMarkers || single) result.SetSeriesMarker(i, OfficeChartMarkerShape.Circle, OfficeLength.Points(6), theme.Background, accent, OfficeLength.Points(1.5D));
                    }
                }
            }

            StraightLines(result);

            if (single) {
                OfficeChartType kind = chart.Data.Series[0].RenderKind ?? chart.Kind;
                bool line = kind is OfficeChartType.Line or OfficeChartType.LineStacked;
                bool area = kind is OfficeChartType.Area or OfficeChartType.AreaStacked;
                if (!area) {
                    result.SetDataLabels(showValue: true);
                    result.SetDataLabelPosition(line ? OfficeChartDataLabelPosition.Top : OfficeChartDataLabelPosition.OutsideEnd);
                    result.SetDataLabelTextStyle(fontSize: OfficeLength.Points(ChartTextSize), bold: true, color: theme.Text);
                }
            }
        }

        /// <summary>
        ///     Secondary axes (combo charts) take the same axis text as the primary ones, so the chart keeps one axis
        ///     text style and the shared chart renderer can draw it.
        /// </summary>
        private static void UniformAxisText(PowerPointChart result) {
            var axes = result.GetChart().PlotArea?.ChildElements
                .Where(axis => axis is DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis or DocumentFormat.OpenXml.Drawing.Charts.ValueAxis)
                .ToList();
            if (axes == null || axes.Count <= 2) return;
            DocumentFormat.OpenXml.OpenXmlElement? reference = axes[0].GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.TextProperties>();
            if (reference == null) return;
            foreach (DocumentFormat.OpenXml.OpenXmlElement axis in axes.Skip(1)) {
                DocumentFormat.OpenXml.Drawing.Charts.TextProperties? text = axis.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.TextProperties>();
                if (text == null) continue;
                text.InsertAfterSelf(reference.CloneNode(true));
                text.Remove();
            }
            result.Save();
        }

        /// <summary>Writes explicit unsmoothed lines, so no renderer curves a series between its points.</summary>
        private static void StraightLines(PowerPointChart result) {
            bool changed = false;
            foreach (var series in result.GetChart().Descendants<DocumentFormat.OpenXml.Drawing.Charts.LineChartSeries>().ToList()) {
                DocumentFormat.OpenXml.Drawing.Charts.Smooth? smooth = series.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Smooth>();
                if (smooth != null) {
                    if (smooth.Val?.Value == false) continue;
                    smooth.Val = false;
                } else {
                    smooth = new DocumentFormat.OpenXml.Drawing.Charts.Smooth { Val = false };
                    DocumentFormat.OpenXml.OpenXmlElement? extensions = series.ChildElements.FirstOrDefault(child => child.LocalName == "extLst");
                    if (extensions != null) series.InsertBefore(smooth, extensions);
                    else series.Append(smooth);
                }
                changed = true;
            }
            if (changed) result.Save();
        }

        /// <summary>
        ///     A value axis of four to six round steps. Bars and columns start at zero; a line whose values sit far
        ///     from zero gets a band around its data so the trend is readable. Headroom keeps value labels inside.
        /// </summary>
        internal static (double Minimum, double Maximum, double Step)? NiceScale(PowerPointLayoutChart chart, bool labelled) {
            bool stacked = chart.Kind is OfficeChartType.ColumnStacked or OfficeChartType.BarStacked or OfficeChartType.LineStacked or
                OfficeChartType.AreaStacked;
            var values = new List<double>();
            if (stacked) {
                for (int c = 0; c < chart.Data.Categories.Count; c++) {
                    double positive = 0D, negative = 0D;
                    foreach (OfficeSeries series in chart.Data.Series) {
                        if (c >= series.Values.Count || double.IsNaN(series.Values[c])) continue;
                        if (series.Values[c] >= 0D) positive += series.Values[c]; else negative += series.Values[c];
                    }
                    values.Add(positive);
                    values.Add(negative);
                }
            } else {
                values.AddRange(chart.Data.Series.SelectMany(series => series.Values).Where(value => !double.IsNaN(value) && !double.IsInfinity(value)));
            }

            if (values.Count == 0) return null;
            double min = values.Min(), max = values.Max();
            if (max <= min && max == 0D) return null;
            bool line = chart.Kind is OfficeChartType.Line or OfficeChartType.LineStacked;
            double headroom = labelled ? 1.12D : 1.04D;
            if (line && min > 0D && min >= max * 0.5D) {
                double span = Math.Max(max - min, max * 0.05D);
                double step = NiceStep(span * 1.6D / 4D);
                double low = Math.Max(0D, Math.Floor((min - span * 0.3D) / step) * step);
                double high = Math.Ceiling((max + span * (labelled ? 0.25D : 0.1D)) / step) * step;
                return (low, high, step);
            }

            double lowEnd = Math.Min(0D, min);
            double highEnd = Math.Max(0D, max);
            double unit = NiceStep((highEnd * headroom - lowEnd * headroom) / 4D);
            if (unit <= 0D) return null;
            double minimum = lowEnd < 0D ? Math.Floor(lowEnd * headroom / unit) * unit : 0D;
            double maximum = highEnd > 0D ? Math.Ceiling(highEnd * headroom / unit) * unit : 0D;
            if (maximum <= minimum) return null;
            return (minimum, maximum, unit);
        }

        private static double NiceStep(double raw) {
            if (raw <= 0D || double.IsNaN(raw) || double.IsInfinity(raw)) return 0D;
            double magnitude = Math.Pow(10D, Math.Floor(Math.Log10(raw)));
            double fraction = raw / magnitude;
            double nice = fraction <= 1D ? 1D : fraction <= 2D ? 2D : fraction <= 2.5D ? 2.5D : fraction <= 5D ? 5D : 10D;
            return nice * magnitude;
        }

        /// <summary>Slice labels are large text (14 pt bold), so 3:1 against every slice is readable.</summary>
        private const double SliceLabelSize = 14D;

        /// <summary>
        ///     One label colour for every slice (a per-slice colour is not portable across renderers): white or the
        ///     theme ink, whichever keeps the better worst-case contrast over the slice accents in use.
        /// </summary>
        private static string SliceLabelColor(PowerPointLayoutContext context, PowerPointLayoutChart chart) {
            int points = Math.Max(1, Math.Min(6, chart.Data.Categories.Count));
            string[] fills = Enumerable.Range(1, points).Select(context.Theme.Accent).ToArray();
            double white = fills.Min(fill => PowerPointLayoutTheme.Contrast("FFFFFF", fill));
            double ink = fills.Min(fill => PowerPointLayoutTheme.Contrast(context.Theme.Ink, fill));
            return white >= ink ? "FFFFFF" : context.Theme.Ink;
        }

        /// <summary>Pictures keep their aspect ratio: contained in the box (top-left) or cover-cropped around the focal point.</summary>
        internal static PowerPointPicture RenderImage(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutImage image,
            PowerPointLayoutRect box, bool cover) {
            if (image.NativePicture != null) return RenderNativeImage(context, writer, image, box, cover);
            PowerPointPicture picture = writer.Slide.AddImage(image.Path, OfficeLength.Points(box.X), OfficeLength.Points(box.Y),
                OfficeLength.Points(box.Width), OfficeLength.Points(box.Height));
            if (!PowerPointImageSize.TryRead(image.Path, out int pw, out int ph)) return picture;
            if (cover) {
                picture.Left = OfficeLength.Points(box.X);
                picture.Top = OfficeLength.Points(box.Y);
                picture.Width = OfficeLength.Points(box.Width);
                picture.Height = OfficeLength.Points(box.Height);
                picture.CropToFill(pw, ph, image.FocalX, image.FocalY);
            } else {
                double scale = Math.Min(box.Width / pw, box.Height / ph);
                picture.Left = OfficeLength.Points(box.X);
                picture.Top = OfficeLength.Points(box.Y);
                picture.Width = OfficeLength.Points(pw * scale);
                picture.Height = OfficeLength.Points(ph * scale);
            }

            if (!string.IsNullOrWhiteSpace(image.AlternativeText)) picture.Description = image.AlternativeText;
            RenderAnnotations(context, writer, picture, image);
            return picture;
        }

        /// <summary>Numbered markers placed on the visible (cropped) picture, with their labels beside them.</summary>
        private static void RenderAnnotations(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointPicture picture,
            PowerPointLayoutImage image) {
            if (image.Annotations.Count == 0) return;
            double left = picture.Left.ToPoints(), top = picture.Top.ToPoints();
            double width = picture.Width.ToPoints(), height = picture.Height.ToPoints();
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
