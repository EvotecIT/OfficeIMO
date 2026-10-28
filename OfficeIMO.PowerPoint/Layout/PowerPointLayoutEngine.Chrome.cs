using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointLayoutEngine {
        private const double HeaderGap = 36D;
        private const double SubtitleGap = 10D;
        private static readonly string[] TitleSeparators = { ": ", " — ", " -- ", " - ", " – " };

        internal sealed class HeaderGeometry {
            internal double TitleY { get; set; }
            internal double TitleSize { get; set; }
            internal int TitleLines { get; set; }
            internal double SubtitleY { get; set; }
            internal int SubtitleLines { get; set; }
            internal double Bottom { get; set; }
            internal int ContentRow { get; set; }
            internal int ContentRows => PowerPointLayoutGrid.Rows + 1 - ContentRow;
            internal bool SuffixOnEyebrow { get; set; }
        }

        /// <summary>
        ///     Title overflow: fits in two lines, else splits at a colon or dash into title and subtitle, else
        ///     shrinks in 2 pt steps to the minimum and takes a third line. Never truncates.
        /// </summary>
        internal static void FitTitle(PowerPointLayoutContext context, PowerPointLayoutSlide slide) {
            if (slide.TitleSize.HasValue || string.IsNullOrWhiteSpace(slide.Title)) return;
            PowerPointTextStyleSpec style = context.Style(PowerPointTextRole.Title);
            double width = context.Grid.W(10);
            string title = slide.Title!;
            if (context.Lines(title, style, width) <= 2) {
                slide.TitleSize = style.Size;
                return;
            }

            if (string.IsNullOrWhiteSpace(slide.Subtitle)) {
                foreach (string separator in TitleSeparators) {
                    int at = title.IndexOf(separator, StringComparison.Ordinal);
                    if (at <= 0 || at + separator.Length >= title.Length) continue;
                    string head = title.Substring(0, at).Trim();
                    string tail = title.Substring(at + separator.Length).Trim();
                    slide.Title = head;
                    slide.Subtitle = char.ToUpper(tail[0], CultureInfo.CurrentCulture) + tail.Substring(1);
                    FitTitle(context, slide);
                    return;
                }
            }

            double size = style.Size;
            while (size - 2D >= style.Minimum && context.Lines(title, style, width, size) > 2) size -= 2D;
            if (context.Lines(title, style, width, size) > 2) {
                size = style.Minimum;
                bool tooLong = context.Lines(title, style, width, size) > 3;
                context.Deck.Report(tooLong ? PowerPointLayoutDiagnosticCodes.TitleTooLong : PowerPointLayoutDiagnosticCodes.TitleLong,
                    OfficeDiagnosticSeverity.Warning,
                    tooLong
                        ? "The slide title needs more than three lines at the minimum size; shorten it."
                        : "The slide title needs three lines at the minimum size; the content starts one row lower.",
                    title);
            }

            slide.TitleSize = size;
        }

        internal static HeaderGeometry MeasureHeader(PowerPointLayoutContext context, PowerPointLayoutSlide slide) {
            FitTitle(context, slide);
            PowerPointLayoutGrid grid = context.Grid;
            PowerPointTextStyleSpec title = context.Style(PowerPointTextRole.Title);
            PowerPointTextStyleSpec subtitle = context.Style(PowerPointTextRole.Subtitle);
            double size = slide.TitleSize ?? title.Size;
            string text = slide.Title ?? string.Empty;
            List<string> lines = PowerPointTextMeasure.Wrap(text, size, grid.W(10), context.Theme.HeadingFont, true, title.Tracking);
            var header = new HeaderGeometry { TitleY = grid.Y(2), TitleSize = size, TitleLines = string.IsNullOrWhiteSpace(text) ? 0 : lines.Count };
            if (!string.IsNullOrEmpty(slide.ContinuedSuffix) && lines.Count > 0) {
                double last = PowerPointTextMeasure.Width(lines[lines.Count - 1], size, context.Theme.HeadingFont, true, title.Tracking);
                double suffix = PowerPointTextMeasure.Width(slide.ContinuedSuffix!, subtitle.Size, context.Theme.BodyFont);
                header.SuffixOnEyebrow = (last + suffix) * 1.1D > grid.W(10);
            }

            double bottom = header.TitleY + header.TitleLines * size * title.LineHeight;
            if (!string.IsNullOrWhiteSpace(slide.Subtitle)) {
                header.SubtitleY = bottom + SubtitleGap;
                header.SubtitleLines = context.Lines(slide.Subtitle!, subtitle, grid.W(9));
                bottom = header.SubtitleY + header.SubtitleLines * subtitle.Size * subtitle.LineHeight;
            }

            header.Bottom = bottom;
            header.ContentRow = Math.Max(4, Math.Min(8, grid.FirstRowBelow(bottom, HeaderGap)));
            return header;
        }

        internal static string? Eyebrow(PowerPointLayoutContext context, PowerPointLayoutSlide slide, HeaderGeometry? header) {
            string? eyebrow = !string.IsNullOrWhiteSpace(slide.Eyebrow) ? slide.Eyebrow : context.Deck.CurrentSection;
            if (header?.SuffixOnEyebrow == true && !string.IsNullOrEmpty(slide.ContinuedSuffix)) {
                string suffix = slide.ContinuedSuffix!.Trim().Trim('(', ')');
                eyebrow = string.IsNullOrWhiteSpace(eyebrow) ? suffix : eyebrow + " · " + suffix;
            }

            return eyebrow;
        }

        /// <summary>The shared content-slide header: accent mark, eyebrow, title, subtitle.</summary>
        internal static HeaderGeometry RenderHeader(PowerPointLayoutContext context, PowerPointLayoutWriter writer,
            PowerPointLayoutSlide slide) {
            PowerPointLayoutGrid grid = context.Grid;
            if (!context.Document.OwnsLayoutTheme &&
                writer.Slide.GetLayoutPlaceholder(PowerPointPlaceholderType.Title)?.Bounds is PowerPointLayoutBox templateTitle) {
                return RenderTemplateHeader(context, writer, slide, templateTitle);
            }

            HeaderGeometry header = MeasureHeader(context, slide);
            writer.Rect("Accent Mark", new PowerPointLayoutRect(grid.X(1), grid.Y(1) - 12D, 28D, 3D), context.Theme.Accent1);
            string? eyebrow = Eyebrow(context, slide, header);
            if (!string.IsNullOrWhiteSpace(eyebrow)) {
                writer.Text("Eyebrow", grid.Box(1, 8, 1, 1), context.Paragraph(eyebrow!, PowerPointTextRole.Eyebrow));
            }

            PowerPointTextStyleSpec title = context.Style(PowerPointTextRole.Title);
            if (!string.IsNullOrWhiteSpace(slide.Title)) {
                string text = PowerPointTextMeasure.Balance(slide.Title!, header.TitleSize, grid.W(10), context.Theme.HeadingFont, true, title.Tracking);
                PowerPointLayoutParagraph paragraph = context.Paragraph(text, PowerPointTextRole.Title, size: header.TitleSize);
                if (!string.IsNullOrEmpty(slide.ContinuedSuffix) && !header.SuffixOnEyebrow) {
                    PowerPointTextStyleSpec subtitle = context.Style(PowerPointTextRole.Subtitle);
                    paragraph.Runs.Add(new PowerPointLayoutRun(slide.ContinuedSuffix!, subtitle, bold: false) { Tracking = 0D });
                }

                double height = header.TitleLines * header.TitleSize * title.LineHeight + 8D;
                writer.Text("Title", new PowerPointLayoutRect(grid.X(1), header.TitleY, grid.W(10), height), paragraph,
                    placeholder: "<p:ph type=\"title\"/>");
            }

            if (!string.IsNullOrWhiteSpace(slide.Subtitle)) {
                PowerPointTextStyleSpec subtitle = context.Style(PowerPointTextRole.Subtitle);
                string text = PowerPointTextMeasure.Balance(slide.Subtitle!, subtitle.Size, grid.W(9), context.Theme.BodyFont);
                writer.Text("Subtitle", new PowerPointLayoutRect(grid.X(1), header.SubtitleY, grid.W(9),
                    header.SubtitleLines * subtitle.Size * subtitle.LineHeight + 6D), context.Paragraph(text, PowerPointTextRole.Subtitle));
            }

            return header;
        }

        /// <summary>
        ///     The header on a template's layout: the title placeholder keeps the template's position, font, size and
        ///     colour; the eyebrow sits above it when there is room; content starts below the template title.
        /// </summary>
        private static HeaderGeometry RenderTemplateHeader(PowerPointLayoutContext context, PowerPointLayoutWriter writer,
            PowerPointLayoutSlide slide, PowerPointLayoutBox titleBox) {
            PowerPointLayoutGrid grid = context.Grid;
            var title = new PowerPointLayoutRect(titleBox.LeftPoints, titleBox.TopPoints, titleBox.WidthPoints, titleBox.HeightPoints);
            string? eyebrow = Eyebrow(context, slide, null);
            PowerPointTextStyleSpec eyebrowStyle = context.Style(PowerPointTextRole.Eyebrow);
            if (!string.IsNullOrWhiteSpace(eyebrow) && title.Y >= eyebrowStyle.Size * 2.4D) {
                writer.Text("Eyebrow", new PowerPointLayoutRect(title.X, title.Y - eyebrowStyle.Size * 1.8D, title.Width, eyebrowStyle.Size * 1.4D),
                    context.Paragraph(eyebrow!, PowerPointTextRole.Eyebrow));
            }

            string text = (slide.Title ?? string.Empty) + (slide.ContinuedSuffix ?? string.Empty);
            if (!string.IsNullOrWhiteSpace(text)) {
                PowerPointTextBox shape = writer.Text("Title", title, context.Paragraph(text, PowerPointTextRole.Title),
                    placeholder: "<p:ph type=\"title\"/>");
                // Inherit geometry and character formatting from the template's title placeholder.
                var element = (DocumentFormat.OpenXml.Presentation.Shape)shape.Element;
                element.ShapeProperties?.RemoveAllChildren();
                foreach (DocumentFormat.OpenXml.Drawing.RunProperties run in element.Descendants<DocumentFormat.OpenXml.Drawing.RunProperties>()) {
                    run.RemoveAllChildren();
                    run.FontSize = null;
                    run.Bold = null;
                    run.Spacing = null;
                }

                DocumentFormat.OpenXml.Drawing.BodyProperties? body = element.TextBody?.BodyProperties;
                if (body != null) {
                    body.RemoveAllChildren();
                    body.RemoveAttribute("lIns", string.Empty);
                    body.RemoveAttribute("tIns", string.Empty);
                    body.RemoveAttribute("rIns", string.Empty);
                    body.RemoveAttribute("bIns", string.Empty);
                    body.RemoveAttribute("anchor", string.Empty);
                    body.RemoveAttribute("wrap", string.Empty);
                }

                element.TextBody?.Descendants<DocumentFormat.OpenXml.Drawing.ParagraphProperties>().ToList().ForEach(p => p.Remove());
            }

            double bottom = title.Bottom;
            if (!string.IsNullOrWhiteSpace(slide.Subtitle)) {
                PowerPointTextStyleSpec subtitle = context.Style(PowerPointTextRole.Subtitle);
                double height = context.Height(slide.Subtitle!, subtitle, title.Width) + 4D;
                writer.Text("Subtitle", new PowerPointLayoutRect(title.X, bottom + 6D, title.Width, height), context.Paragraph(slide.Subtitle!, PowerPointTextRole.Subtitle));
                bottom += 6D + height;
            }

            return new HeaderGeometry {
                TitleY = title.Y,
                TitleSize = context.Style(PowerPointTextRole.Title).Size,
                TitleLines = 1,
                Bottom = bottom,
                ContentRow = Math.Max(1, Math.Min(PowerPointLayoutGrid.Rows - 3, grid.FirstRowBelow(bottom, 12D)))
            };
        }

        /// <summary>Footer text and slide number in native footer placeholders.</summary>
        private static void RenderFooter(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointLayoutSlide slide) {
            bool text, number;
            string color = context.Theme.TextMuted;
            double alpha = 1D;
            switch (slide.Layout) {
                case PowerPointLayout.Title:
                case PowerPointLayout.Section:
                case PowerPointLayout.Closing:
                    return;
                case PowerPointLayout.Quote:
                    text = false;
                    number = true;
                    break;
                case PowerPointLayout.ImageFull:
                    text = number = true;
                    color = "FFFFFF";
                    alpha = 0.8D;
                    break;
                default:
                    text = number = true;
                    break;
            }

            if (slide.NoFooter) text = false;
            if (slide.NoSlideNumber || !context.Deck.ShowSlideNumber) number = false;
            PowerPointLayoutGrid grid = context.Grid;
            PowerPointTextStyleSpec footer = context.Style(PowerPointTextRole.Footer);
            string? footerText = context.Deck.EffectiveFooterText;
            if (text && !string.IsNullOrWhiteSpace(footerText)) {
                string fitted = Ellipsize(footerText!, footer.Size, grid.W(8), context.Theme.BodyFont);
                PowerPointLayoutParagraph p = context.Paragraph(fitted, PowerPointTextRole.Footer, color);
                p.Runs[0].Alpha = alpha;
                PowerPointTextBox footerShape = writer.Text("Footer", new PowerPointLayoutRect(grid.X(1), grid.FooterY, grid.W(8), PowerPointLayoutGrid.FooterHeight), p,
                    anchor: "b", placeholder: "<p:ph type=\"ftr\" sz=\"quarter\" idx=\"" + PowerPointLayoutMasterWriter.FooterIndex + "\"/>");
                InheritTemplatePlaceholder(context, writer, footerShape, PowerPointPlaceholderType.Footer);
            }

            if (number) {
                int index = writer.Slide.OwnerPresentation?.Slides.Count ?? 1;
                var run = new PowerPointLayoutRun(index.ToString(CultureInfo.InvariantCulture), footer, color: color) { Field = true, Alpha = alpha };
                var p = new PowerPointLayoutParagraph(run) { Align = "r", LineHeight = footer.LineHeight };
                PowerPointTextBox numberShape = writer.Text("Slide Number", new PowerPointLayoutRect(grid.X(9), grid.FooterY, grid.W(4), PowerPointLayoutGrid.FooterHeight), p,
                    anchor: "b", placeholder: "<p:ph type=\"sldNum\" sz=\"quarter\" idx=\"" + PowerPointLayoutMasterWriter.SlideNumberIndex + "\"/>");
                InheritTemplatePlaceholder(context, writer, numberShape, PowerPointPlaceholderType.SlideNumber);
            }
        }

        /// <summary>On a template's layout, a footer placeholder keeps the template's own position.</summary>
        private static void InheritTemplatePlaceholder(PowerPointLayoutContext context, PowerPointLayoutWriter writer, PowerPointTextBox shape,
            PowerPointPlaceholderType type) {
            if (context.Document.OwnsLayoutTheme || writer.Slide.GetLayoutPlaceholder(type)?.Bounds == null) return;
            ((DocumentFormat.OpenXml.Presentation.Shape)shape.Element).ShapeProperties?.RemoveAllChildren();
        }

        private static string Ellipsize(string text, double size, double width, string font) {
            if (PowerPointTextMeasure.Width(text, size, font) <= width) return text;
            string value = text;
            while (value.Length > 1 && PowerPointTextMeasure.Width(value + "…", size, font) > width) value = value.Substring(0, value.Length - 1);
            return value.TrimEnd() + "…";
        }

        /// <summary>One item of a bottom-anchored text stack (title slides, section, closing, image).</summary>
        internal sealed class StackItem {
            internal StackItem(string name, PowerPointLayoutParagraph paragraph, double height, double gapAfter, string? placeholder = null) {
                Name = name;
                Paragraph = paragraph;
                Height = height;
                GapAfter = gapAfter;
                Placeholder = placeholder;
            }

            internal string Name { get; }
            internal PowerPointLayoutParagraph Paragraph { get; }
            internal double Height { get; }
            internal double GapAfter { get; }
            internal string? Placeholder { get; }
        }

        internal static StackItem Stacked(PowerPointLayoutContext context, string name, string text, PowerPointTextRole role,
            double width, double gapAfter, string? color = null, double? size = null, string? placeholder = null, bool balance = true,
            int? maxLines = null, double? minimum = null) {
            PowerPointTextStyleSpec style = context.Style(role);
            double resolved = size ?? style.Size;
            if (maxLines.HasValue) {
                double floor = minimum ?? style.Minimum;
                while (resolved - 2D >= floor && context.Lines(text, style, width, resolved) > maxLines.Value) resolved -= 2D;
            }

            string value = balance ? PowerPointTextMeasure.Balance(text, resolved, width, context.FontOf(style), style.Bold, style.Tracking) : text;
            int lines = context.Lines(text, style, width / 1.06D, resolved);
            PowerPointLayoutParagraph paragraph = context.Paragraph(value, role, color, resolved);
            // Tight display leading (1.0) still gets a frame of normal leading so no renderer clips the last line.
            return new StackItem(name, paragraph, lines * resolved * Math.Max(1.2D, style.LineHeight) + 4D, gapAfter, placeholder);
        }

        /// <summary>Writes stack items bottom-anchored (or top-anchored) inside <paramref name="box" />.</summary>
        internal static void RenderStack(PowerPointLayoutWriter writer, PowerPointLayoutRect box, IList<StackItem> items, bool bottom = true,
            Func<StackItem, double>? alphaFor = null) {
            double total = items.Sum(i => i.Height) + items.Take(items.Count - 1).Sum(i => i.GapAfter);
            double y = bottom ? box.Bottom - total : box.Y;
            foreach (StackItem item in items) {
                writer.Text(item.Name, new PowerPointLayoutRect(box.X, y, box.Width, item.Height), item.Paragraph,
                    anchor: "t", placeholder: item.Placeholder);
                y += item.Height + item.GapAfter;
            }
        }
    }
}
