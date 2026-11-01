using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointLayoutEngine {
        /// <summary>Draws one content block in <paramref name="box" />; list and text keep their measure.</summary>
        internal static void RenderBlock(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutBlock block,
            PowerPointLayoutRect box, PowerPointLayoutSlide slide, string? placeholder = null, bool includeColumnTitle = true) {
            switch (block) {
                case PowerPointLayoutList list:
                    double listSize = slide.ListSize ?? FitListSize(context, list, box);
                    RenderList(context, writer, list, box, listSize, slide.ListStart, placeholder);
                    break;
                case PowerPointLayoutText text:
                    RenderText(context, writer, text, box, slide.ListSize ?? FitTextSize(context, text, box), placeholder);
                    break;
                case PowerPointLayoutTable table:
                    RenderTable(context, writer, table, box, slide.TableRange);
                    break;
                case PowerPointLayoutChart chart:
                    RenderChart(context, writer, chart, box);
                    break;
                case PowerPointLayoutImage image:
                    RenderImage(context, writer, image, box, cover: false);
                    break;
                case PowerPointLayoutColumn column:
                    (PowerPointLayoutText? lede, List<PowerPointLayoutBlock> inner) = SplitLede(column.Blocks);
                    PowerPointLayoutRect rest = box;
                    if (includeColumnTitle && !string.IsNullOrWhiteSpace(column.Title)) {
                        PowerPointTextStyleSpec heading = context.Style(PowerPointTextRole.Heading);
                        double titleHeight = context.Height(column.Title!, heading, rest.Width) + 4D;
                        writer.Text("Column Title", rest.WithHeight(titleHeight), context.Paragraph(column.Title!, PowerPointTextRole.Heading));
                        rest = rest.Inset(0, titleHeight + 12D, 0, 0);
                    }
                    if (lede != null) {
                        PowerPointTextStyleSpec ledeStyle = context.Style(PowerPointTextRole.BodySmall);
                        double h = NativeLedeHeight(context, lede, rest.Width, PowerPointTextRole.BodySmall,
                            context.Height(lede.Paragraphs[0], ledeStyle, rest.Width)) + 10D;
                        writer.Text("Column Lede", rest.WithHeight(h), PowerPointLayoutNativeText.Attach(context.Paragraph(lede.Paragraphs[0], PowerPointTextRole.BodySmall), lede.NativeBody));
                        rest = rest.Inset(0, h, 0, 0);
                    }

                    if (inner.Count > 0) RenderBlock(context, writer, inner[0], rest, slide, placeholder);
                    break;
                case PowerPointLayoutDrawing drawing:
                    drawing.Draw(writer.Slide, context, box);
                    break;
            }
        }

        /// <summary>Retains the designed allowance unless preserved native styling needs a taller lede frame.</summary>
        private static double NativeLedeHeight(PowerPointLayoutContext context, PowerPointLayoutText lede, double width,
            PowerPointTextRole role, double fallback) => lede.NativeBody?.Paragraphs.FirstOrDefault() is PowerPointParagraph paragraph
                ? Math.Max(fallback, PowerPointLayoutNativeTextMeasure.Height(context, paragraph, context.Style(role).Size, width, spaceAfter: 0D, role: role) + 4D)
                : fallback;

        private static double FitListSize(PowerPointLayoutContext context, PowerPointLayoutList list, PowerPointLayoutRect box) {
            PowerPointTextStyleSpec body = context.Style(PowerPointTextRole.Body);
            double size = body.Size;
            while (size - 2D >= 14D && ListHeight(context, list.Items, size, box.Width) > box.Height) size -= 2D;
            return size;
        }

        private static double FitTextSize(PowerPointLayoutContext context, PowerPointLayoutText text, PowerPointLayoutRect box) {
            PowerPointTextStyleSpec body = context.Style(PowerPointTextRole.Body);
            double size = body.Size;
            double Height(double s) => text.NativeBody == null
                ? text.Paragraphs.Sum(p => context.Height(p, body, box.Width, s) + context.Theme.ParagraphAfter)
                : text.NativeBody.Paragraphs.Sum(p => PowerPointLayoutNativeTextMeasure.Height(context, p, s, box.Width));
            while (size - 2D >= 14D && Height(size) > box.Height) size -= 2D;
            return size;
        }

        // ------------------------------------------------------------------ TitleAndContent, TwoContent, Comparison

        private static void RenderTitleAndContent(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            if (string.IsNullOrWhiteSpace(slide.Title) && slide.Blocks.Count > 0 && slide.Blocks.All(b => b is PowerPointLayoutDrawing)) {
                // A custom composition draws its own header through the composition context.
                foreach (PowerPointLayoutBlock drawing in slide.Blocks) RenderBlock(context, writer, drawing, context.Grid.Box(1, 12, 4, 9), slide);
                return;
            }

            HeaderGeometry header = RenderHeader(context, writer, slide);
            (PowerPointLayoutText? lede, List<PowerPointLayoutBlock> blocks) = SplitLede(slide.Blocks);
            PowerPointLayoutRect zone = RenderSourceNote(context, writer, slide, ContentZone(context, header));
            if (lede != null) {
                double height = NativeLedeHeight(context, lede, context.Grid.W(8), PowerPointTextRole.Lede, 64D);
                writer.Text("Lede", new PowerPointLayoutRect(zone.X, zone.Y, context.Grid.W(8), height), PowerPointLayoutNativeText.Attach(context.Paragraph(lede.Paragraphs[0], PowerPointTextRole.Lede), lede.NativeBody));
                zone = zone.Inset(0, height + 12D, 0, 0);
            }

            if (blocks.Count == 0) return;
            PowerPointLayoutBlock block = blocks[0];
            bool measure = block is PowerPointLayoutList or PowerPointLayoutText;
            if (measure) zone = zone.WithWidth(context.Grid.W(8));
            RenderBlock(context, writer, block, zone, slide,
                measure && writer.Slide.SlideLayoutName == "Title and Content" ? "<p:ph idx=\"1\"/>" : null);
        }

        private static bool IsTexty(PowerPointLayoutBlock block) {
            if (block is PowerPointLayoutColumn column) {
                List<PowerPointLayoutBlock> inner = SplitLede(column.Blocks).Blocks;
                return inner.Count == 0 || inner[0] is PowerPointLayoutList or PowerPointLayoutText;
            }

            return block is PowerPointLayoutList or PowerPointLayoutText;
        }

        /// <summary>Shared geometry for native continuation measurement and final column rendering.</summary>
        private static PowerPointLayoutRect[] TwoContentBoxes(PowerPointLayoutContext context, IReadOnlyList<PowerPointLayoutBlock> blocks,
            PowerPointLayoutRect zone) => TwoContentFrames(context, zone, IsTexty(blocks[0]), IsTexty(blocks[1]));

        private static PowerPointLayoutRect ComparisonBodyBox(PowerPointLayoutContext context, HeaderGeometry header, int index) =>
            context.Grid.Box(index == 0 ? 1 : 7, 6, header.ContentRow + 1, header.ContentRows - 1).Inset(0, 12, 0, 0);

        private static void RenderTwoContent(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            HeaderGeometry header = RenderHeader(context, writer, slide);
            (PowerPointLayoutText? lede, List<PowerPointLayoutBlock> blocks) = SplitLede(slide.Blocks);
            if (blocks.Count == 0) return;
            PowerPointLayoutRect zone = RenderSourceNote(context, writer, slide, ContentZone(context, header));
            if (lede != null) {
                double height = NativeLedeHeight(context, lede, zone.Width, PowerPointTextRole.Lede, 64D);
                writer.Text("Lede", zone.WithHeight(height), PowerPointLayoutNativeText.Attach(context.Paragraph(lede.Paragraphs[0], PowerPointTextRole.Lede), lede.NativeBody));
                zone = zone.Inset(0, height + 12D, 0, 0);
            }
            if (blocks.Count == 1) {
                RenderBlock(context, writer, blocks[0], zone, slide);
                return;
            }

            var boxes = TwoContentBoxes(context, blocks, zone);

            bool native = writer.Slide.SlideLayoutName == "Two Content";
            for (int i = 0; i < 2; i++) {
                PowerPointLayoutRect box = boxes[i];
                RenderBlock(context, writer, blocks[i], box, slide, native && IsTexty(blocks[i]) ? "<p:ph idx=\"" + (i + 1) + "\"/>" : null);
            }
        }

        private static void RenderComparison(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            HeaderGeometry header = RenderHeader(context, writer, slide);
            List<PowerPointLayoutColumn> columns = SplitLede(slide.Blocks).Blocks.OfType<PowerPointLayoutColumn>().Take(2).ToList();
            bool native = writer.Slide.SlideLayoutName == "Comparison";
            PowerPointTextStyleSpec heading = context.Style(PowerPointTextRole.Heading);
            for (int i = 0; i < columns.Count; i++) {
                int column = i == 0 ? 1 : 7;
                PowerPointLayoutRect head = context.Grid.Box(column, 6, header.ContentRow, 1);
                double size = heading.Size;
                while (size - 2D >= heading.Minimum && context.Width(columns[i].Title ?? string.Empty, heading, size) > head.Width) size -= 2D;
                writer.Text("Column Heading", new PowerPointLayoutRect(head.X, head.Y - 8D, head.Width, head.Height), context.Paragraph(columns[i].Title ?? string.Empty,
                    PowerPointTextRole.Heading, size: size), anchor: "b", placeholder: native ? "<p:ph type=\"body\" sz=\"quarter\" idx=\"" + (i == 0 ? 1 : 3) + "\"/>" : null);
                writer.Rect("Column Rule", new PowerPointLayoutRect(head.X, head.Bottom + 6D - 8D, head.Width, 3D), i == 0 ? context.Theme.Accent1 : context.Theme.Accent2);
                PowerPointLayoutRect body = ComparisonBodyBox(context, header, i);
                RenderBlock(context, writer, columns[i], body, slide, native ? "<p:ph sz=\"quarter\" idx=\"" + (i == 0 ? 2 : 4) + "\"/>" : null, includeColumnTitle: false);
            }
        }
    }
}
