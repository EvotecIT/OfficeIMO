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
                slide.TitleSize = context.UnambiguousSize(title, style, width, style.Size, style.Minimum, 4D);
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

            slide.TitleSize = context.UnambiguousSize(title, style, width, size, style.Minimum, 4D);
        }

        internal static HeaderGeometry MeasureHeader(PowerPointLayoutContext context, PowerPointLayoutSlide slide) {
            if (TemplateTitleBounds(context, slide) is PowerPointLayoutBox titleBox) return MeasureTemplateHeader(context, slide, titleBox);
            FitTitle(context, slide);
            PowerPointLayoutGrid grid = context.Grid;
            PowerPointTextStyleSpec title = context.Style(PowerPointTextRole.Title);
            PowerPointTextStyleSpec subtitle = context.Style(PowerPointTextRole.Subtitle);
            double size = slide.TitleSize ?? title.Size;
            string text = slide.Title ?? string.Empty;
            List<string> lines = PowerPointTextMeasure.Wrap(text, size, grid.W(10) / PowerPointLayoutContext.SafetyFactor, context.Theme.HeadingFont, true, title.Tracking);
            var header = new HeaderGeometry { TitleY = grid.Y(2), TitleSize = size, TitleLines = string.IsNullOrWhiteSpace(text) ? 0 : lines.Count };
            if (!string.IsNullOrEmpty(slide.ContinuedSuffix) && lines.Count > 0) {
                double last = PowerPointTextMeasure.Width(lines[lines.Count - 1], size, context.Theme.HeadingFont, true, title.Tracking);
                double suffix = PowerPointTextMeasure.Width(slide.ContinuedSuffix!, subtitle.Size, context.Theme.BodyFont);
                header.SuffixOnEyebrow = (last + suffix) * 1.1D > grid.W(10);
            }

            double bottom = header.TitleY + header.TitleLines * size * title.LineHeight;
            if (!string.IsNullOrWhiteSpace(slide.Subtitle)) {
                header.SubtitleY = bottom + SubtitleGap;
                header.SubtitleLines = context.SafeLines(slide.Subtitle!, subtitle, grid.W(9));
                bottom = header.SubtitleY + header.SubtitleLines * subtitle.Size * subtitle.LineHeight;
            }

            header.Bottom = bottom;
            header.ContentRow = Math.Max(4, Math.Min(8, grid.FirstRowBelow(bottom, HeaderGap)));
            return header;
        }

        /// <summary>Reads the selected layout's authored title frame without creating a provisional slide.</summary>
        internal static PowerPointLayoutBox? TemplateTitleBounds(PowerPointLayoutContext context, PowerPointLayoutSlide slide) {
            if (context.Document.OwnsLayoutTheme && slide.TemplateLayout == null) return null;
            int master, layout;
            if (slide.TemplateLayout != null) {
                master = slide.TemplateLayout.MasterIndex;
                layout = slide.TemplateLayout.LayoutIndex;
            } else {
                PowerPointSlideLayoutInfo? selected = SelectNativeLayout(context, slide.Layout, reportMissing: false);
                if (!selected.HasValue) return null;
                master = selected.Value.MasterIndex;
                layout = selected.Value.LayoutIndex;
            }
            return TemplateTitle(context, context.Document.GetSlideLayoutPart(master, layout))?.Bounds;
        }

        private static PowerPointLayoutPlaceholderInfo? TemplateTitle(PowerPointLayoutContext context,
            DocumentFormat.OpenXml.Packaging.SlideLayoutPart? layout) {
            var tree = layout?.SlideLayout?.CommonSlideData?.ShapeTree;
            if (tree == null) return null;
            foreach (var shape in tree.Elements<DocumentFormat.OpenXml.Presentation.Shape>()) {
                var placeholder = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
                var type = placeholder?.Type?.Value;
                if (placeholder == null || type != DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title &&
                    type != DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle) continue;
                PowerPointLayoutBox? bounds = PowerPointSlide.GetExportPlaceholderBounds(layout, placeholder);
                if (!bounds.HasValue || !WithinSlide(context.Document, bounds.Value.Left.ToPoints(), bounds.Value.Top.ToPoints(),
                        bounds.Value.Width.ToPoints(), bounds.Value.Height.ToPoints())) continue;
                return new PowerPointLayoutPlaceholderInfo(shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty,
                    type.Value.ToOfficeEnum(), placeholder.Index?.Value, bounds);
            }
            return null;
        }

        internal static bool WithinSlide(PowerPointDocument document, double left, double top, double width, double height) =>
            left >= 0D && top >= 0D && width > 0D && height > 0D &&
            left + width <= document.SlideSize.Width.ToPoints() + 1D / 12700D &&
            top + height <= document.SlideSize.Height.ToPoints() + 1D / 12700D;

        private static HeaderGeometry MeasureTemplateHeader(PowerPointLayoutContext context, PowerPointLayoutSlide slide, PowerPointLayoutBox titleBox) {
            double bottom = titleBox.Bottom.ToPoints();
            double subtitleY = 0D;
            if (!string.IsNullOrWhiteSpace(slide.Subtitle)) {
                subtitleY = bottom + 6D;
                bottom = subtitleY + context.Height(slide.Subtitle!, context.Style(PowerPointTextRole.Subtitle), titleBox.Width.ToPoints()) + 4D;
            }
            return new HeaderGeometry {
                TitleY = titleBox.Top.ToPoints(),
                TitleSize = context.Style(PowerPointTextRole.Title).Size,
                TitleLines = 1,
                SubtitleY = subtitleY,
                Bottom = bottom,
                ContentRow = Math.Max(1, Math.Min(PowerPointLayoutGrid.Rows - 3, context.Grid.FirstRowBelow(bottom, 12D)))
            };
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
            PowerPointLayoutSlide slide, bool decorate = true) {
            PowerPointLayoutGrid grid = context.Grid;
            if ((!context.Document.OwnsLayoutTheme || slide.TemplateLayout != null) &&
                TemplateTitle(context, writer.Slide.SlidePart.SlideLayoutPart) is PowerPointLayoutPlaceholderInfo templateTitle) {
                return RenderTemplateHeader(context, writer, slide, templateTitle);
            }

            HeaderGeometry header = MeasureHeader(context, slide);
            if (decorate) writer.Rect("Accent Mark", new PowerPointLayoutRect(grid.X(1), grid.Y(1) - 12D, 28D, 3D), context.Theme.Accent1);
            string? eyebrow = Eyebrow(context, slide, header);
            if (decorate && !string.IsNullOrWhiteSpace(eyebrow)) {
                writer.Text("Eyebrow", grid.Box(1, 8, 1, 1), context.Paragraph(eyebrow!, PowerPointTextRole.Eyebrow));
            }

            PowerPointTextStyleSpec title = context.Style(PowerPointTextRole.Title);
            if (!string.IsNullOrWhiteSpace(slide.Title)) {
                // A continuation suffix shares the last line, so only an unsuffixed title narrows to balance.
                double width = string.IsNullOrEmpty(slide.ContinuedSuffix)
                    ? PowerPointTextMeasure.BalancedWidth(slide.Title!, header.TitleSize, grid.W(10), context.Theme.HeadingFont, true, title.Tracking)
                    : grid.W(10);
                PowerPointLayoutParagraph paragraph = context.Paragraph(slide.Title!, PowerPointTextRole.Title, size: header.TitleSize);
                if (!string.IsNullOrEmpty(slide.ContinuedSuffix) && !header.SuffixOnEyebrow) {
                    PowerPointTextStyleSpec subtitle = context.Style(PowerPointTextRole.Subtitle);
                    paragraph.Runs.Add(new PowerPointLayoutRun(slide.ContinuedSuffix!, subtitle, bold: false) { Tracking = 0D });
                }

                double height = header.TitleLines * header.TitleSize * title.LineHeight + 8D;
                writer.Text("Title", new PowerPointLayoutRect(grid.X(1), header.TitleY, width, height), paragraph,
                    placeholder: "<p:ph type=\"title\"/>");
            }

            if (!string.IsNullOrWhiteSpace(slide.Subtitle)) {
                PowerPointTextStyleSpec subtitle = context.Style(PowerPointTextRole.Subtitle);
                double width = PowerPointTextMeasure.BalancedWidth(slide.Subtitle!, subtitle.Size, grid.W(9), context.Theme.BodyFont);
                writer.Text("Subtitle", new PowerPointLayoutRect(grid.X(1), header.SubtitleY, width,
                    header.SubtitleLines * subtitle.Size * subtitle.LineHeight + 6D), context.Paragraph(slide.Subtitle!, PowerPointTextRole.Subtitle));
            }

            return header;
        }

        /// <summary>
        ///     The header on a template's layout: the title placeholder keeps the template's position, font, size and
        ///     colour; the eyebrow sits above it when there is room; content starts below the template title.
        /// </summary>
        private static HeaderGeometry RenderTemplateHeader(PowerPointLayoutContext context, PowerPointLayoutWriter writer,
            PowerPointLayoutSlide slide, PowerPointLayoutPlaceholderInfo placeholder) {
            PowerPointLayoutBox titleBox = placeholder.Bounds!.Value;
            HeaderGeometry header = MeasureTemplateHeader(context, slide, titleBox);
            var title = new PowerPointLayoutRect(titleBox.Left.ToPoints(), titleBox.Top.ToPoints(), titleBox.Width.ToPoints(), titleBox.Height.ToPoints());
            string? eyebrow = Eyebrow(context, slide, null);
            PowerPointTextStyleSpec eyebrowStyle = context.Style(PowerPointTextRole.Eyebrow);
            if (!string.IsNullOrWhiteSpace(eyebrow) && title.Y >= eyebrowStyle.Size * 2.4D) {
                writer.Text("Eyebrow", new PowerPointLayoutRect(title.X, title.Y - eyebrowStyle.Size * 1.8D, title.Width, eyebrowStyle.Size * 1.4D),
                    context.Paragraph(eyebrow!, PowerPointTextRole.Eyebrow));
            }

            string text = (slide.Title ?? string.Empty) + (slide.ContinuedSuffix ?? string.Empty);
            if (slide.Title != null || !string.IsNullOrWhiteSpace(text)) {
                string type = placeholder.PlaceholderType == PowerPointPlaceholderType.CenteredTitle ? "ctrTitle" : "title";
                string index = placeholder.PlaceholderIndex.HasValue
                    ? " idx=\"" + placeholder.PlaceholderIndex.Value.ToString(CultureInfo.InvariantCulture) + "\"" : string.Empty;
                PowerPointTextBox shape = writer.Text("Title", title, context.Paragraph(text, PowerPointTextRole.Title),
                    placeholder: "<p:ph type=\"" + type + "\"" + index + "/>");
                // Inherit geometry and character formatting from the template's title placeholder.
                var element = (DocumentFormat.OpenXml.Presentation.Shape)shape.Element;
                element.ShapeProperties?.RemoveAllChildren();
                foreach (var properties in element.Descendants().Where(item =>
                    item is DocumentFormat.OpenXml.Drawing.RunProperties ||
                    item is DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties).ToArray()) {
                    properties.RemoveAllChildren();
                    properties.ClearAllAttributes();
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

            if (!string.IsNullOrWhiteSpace(slide.Subtitle)) {
                writer.Text("Subtitle", new PowerPointLayoutRect(title.X, header.SubtitleY, title.Width, header.Bottom - header.SubtitleY),
                    context.Paragraph(slide.Subtitle!, PowerPointTextRole.Subtitle));
            }
            return header;
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

            if (slide.NoFooter || context.Deck.NoFooter) text = false;
            if (slide.NoSlideNumber || !context.Deck.ShowSlideNumber) number = false;
            PowerPointLayoutGrid grid = context.Grid;
            PowerPointTextStyleSpec footer = context.Style(PowerPointTextRole.Footer);
            bool rightText = text && !string.IsNullOrWhiteSpace(slide.FooterRightText);
            string? footerText = context.Deck.EffectiveFooterText;
            if (text && !string.IsNullOrWhiteSpace(footerText)) {
                string fitted = Ellipsize(footerText!, footer.Size, grid.W(8), context.Theme.BodyFont);
                PowerPointLayoutParagraph p = context.Paragraph(fitted, PowerPointTextRole.Footer, color);
                p.Runs[0].Alpha = alpha;
                PowerPointTextBox footerShape = writer.Text("Footer", new PowerPointLayoutRect(grid.X(1), grid.FooterY, grid.W(8), PowerPointLayoutGrid.FooterHeight), p,
                    anchor: "b", placeholder: "<p:ph type=\"ftr\" sz=\"quarter\" idx=\"" + PowerPointLayoutMasterWriter.FooterIndex + "\"/>");
                InheritTemplatePlaceholder(context, writer, footerShape, PowerPointPlaceholderType.Footer);
            }

            if (rightText) {
                double width = grid.W(number ? 3 : 4);
                string fitted = Ellipsize(slide.FooterRightText!, footer.Size, width, context.Theme.BodyFont);
                PowerPointLayoutParagraph p = context.Paragraph(fitted, PowerPointTextRole.Footer, color);
                p.Align = "r";
                p.Runs[0].Alpha = alpha;
                writer.Text("Footer Right", new PowerPointLayoutRect(grid.X(9), grid.FooterY, width,
                    PowerPointLayoutGrid.FooterHeight), p, anchor: "b");
            }

            if (number) {
                int index = writer.Slide.OwnerDocument?.Slides.Count ?? 1;
                var run = new PowerPointLayoutRun(index.ToString(CultureInfo.InvariantCulture), footer, color: color) { Field = true, Alpha = alpha };
                var p = new PowerPointLayoutParagraph(run) { Align = "r", LineHeight = footer.LineHeight };
                PowerPointTextBox numberShape = writer.Text("Slide Number", new PowerPointLayoutRect(grid.X(rightText ? 12 : 9), grid.FooterY, grid.W(rightText ? 1 : 4), PowerPointLayoutGrid.FooterHeight), p,
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
            internal StackItem(string name, PowerPointLayoutParagraph paragraph, double height, double gapAfter, string? placeholder = null,
                double? width = null) {
                Name = name;
                Paragraph = paragraph;
                Height = height;
                GapAfter = gapAfter;
                Placeholder = placeholder;
                Width = width;
            }

            /// <summary>Frame width when narrower than the stack (measured or balanced text).</summary>
            internal double? Width { get; }

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

            resolved = context.UnambiguousSize(text, style, width, resolved, minimum ?? style.Minimum);
            double frame = balance ? PowerPointTextMeasure.BalancedWidth(text, resolved, width, context.FontOf(style), style.Bold, style.Tracking) : width;
            int lines = context.SafeLines(text, style, width, resolved);
            PowerPointLayoutParagraph paragraph = context.Paragraph(text, role, color, resolved);
            // The frame holds exactly the measured lines at the written line pitch (plus the descender of the last
            // line for tight display leading), so stacked items keep their designed gaps in every renderer.
            double pitch = resolved * style.LineHeight;
            double height = lines * pitch + Math.Max(0D, resolved * 1.2D - pitch) + 2D;
            return new StackItem(name, paragraph, height, gapAfter, placeholder, frame);
        }

        /// <summary>Writes stack items bottom-anchored (or top-anchored) inside <paramref name="box" />.</summary>
        internal static void RenderStack(PowerPointLayoutWriter writer, PowerPointLayoutRect box, IList<StackItem> items, bool bottom = true,
            Func<StackItem, double>? alphaFor = null) {
            double total = items.Sum(i => i.Height) + items.Take(items.Count - 1).Sum(i => i.GapAfter);
            double y = bottom ? box.Bottom - total : box.Y;
            foreach (StackItem item in items) {
                writer.Text(item.Name, new PowerPointLayoutRect(box.X, y, Math.Min(box.Width, item.Width ?? box.Width), item.Height), item.Paragraph,
                    anchor: "t", placeholder: item.Placeholder);
                y += item.Height + item.GapAfter;
            }
        }
    }
}
