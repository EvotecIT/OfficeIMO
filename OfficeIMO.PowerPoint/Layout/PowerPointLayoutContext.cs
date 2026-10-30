using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>Everything a layout renderer needs: document, tokens, grid and deck state.</summary>
    internal sealed class PowerPointLayoutContext {
        internal PowerPointLayoutContext(PowerPointPresentation document, PowerPointLayoutTheme theme) {
            Document = document;
            Theme = theme.Scaled(SizeScale(document));
            Grid = (document.OwnsLayoutTheme ? null : TemplateGrid(document)) ?? PowerPointLayoutGrid.For(document, theme);
            Deck = document.LayoutDeck;
        }

        internal PowerPointPresentation Document { get; }

        /// <summary>
        ///     The slide size relative to the 960 x 540 pt reference (and 720 x 540 pt for 4:3). The type scale and
        ///     spacing are designed for those sizes; a smaller slide, such as a 10 inch 720 x 405 pt template, scales
        ///     them down so the same content keeps the same composition instead of overflowing.
        /// </summary>
        internal static double SizeScale(PowerPointPresentation document) {
            PowerPointSlideSize size = document.SlideSize;
            double width = PowerPointUnits.ToPoints(size.WidthEmus), height = PowerPointUnits.ToPoints(size.HeightEmus);
            if (width <= 0D || height <= 0D) return 1D;
            double referenceWidth = width / height < 1.5D ? 720D : 960D;
            return Math.Min(1D, Math.Min(width / referenceWidth, height / 540D));
        }

        /// <summary>Derives the grid from a loaded deck's or template's "Title and Content" body placeholder.</summary>
        /// <summary>
        ///     The first master's title placeholder. A layout title without its own geometry inherits it, so every
        ///     template slide gets the same header whether its layout repeats the position or not.
        /// </summary>
        internal PowerPointLayoutRect? MasterTitle { get; private set; }

        private PowerPointLayoutGrid? TemplateGrid(PowerPointPresentation document) {
            DocumentFormat.OpenXml.Packaging.SlideMasterPart? master = document.OpenXmlDocument.PresentationPart?.SlideMasterParts.FirstOrDefault();
            if (master == null) return null;
            DocumentFormat.OpenXml.Packaging.SlideLayoutPart? layout = master.SlideLayoutParts.FirstOrDefault(l =>
                    l.SlideLayout?.Type?.Value == DocumentFormat.OpenXml.Presentation.SlideLayoutValues.Object)
                ?? master.SlideLayoutParts.FirstOrDefault(l => l.SlideLayout?.Type?.Value == DocumentFormat.OpenXml.Presentation.SlideLayoutValues.Text);
            PowerPointLayoutRect? Bounds(DocumentFormat.OpenXml.OpenXmlPartRootElement? root, bool title) {
                if (root == null) return null;
                foreach (DocumentFormat.OpenXml.Presentation.Shape shape in root.Descendants<DocumentFormat.OpenXml.Presentation.Shape>()) {
                    DocumentFormat.OpenXml.Presentation.PlaceholderShape? ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
                    if (ph == null) continue;
                    bool isTitle = ph.Type?.Value == DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title ||
                                   ph.Type?.Value == DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle;
                    bool isBody = ph.Type == null || ph.Type.Value == DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body ||
                                  ph.Type.Value == DocumentFormat.OpenXml.Presentation.PlaceholderValues.Object;
                    if (title ? !isTitle : !isBody) continue;
                    DocumentFormat.OpenXml.Drawing.Transform2D? xfrm = shape.ShapeProperties?.Transform2D;
                    if (xfrm?.Offset == null || xfrm.Extents == null) continue;
                    return new PowerPointLayoutRect(xfrm.Offset.X!.Value / 12700D, xfrm.Offset.Y!.Value / 12700D,
                        xfrm.Extents.Cx!.Value / 12700D, xfrm.Extents.Cy!.Value / 12700D);
                }

                return null;
            }

            PowerPointLayoutRect? body = Bounds(layout?.SlideLayout, false) ?? Bounds(master.SlideMaster, false);
            PowerPointLayoutRect? titleBox = Bounds(layout?.SlideLayout, true) ?? Bounds(master.SlideMaster, true);
            MasterTitle = Bounds(master.SlideMaster, true);
            if (body == null || body.Value.Width < 100D || body.Value.Height < 100D) return null;
            PowerPointSlideSize size = document.SlideSize;
            return PowerPointLayoutGrid.ForTemplate(PowerPointUnits.ToPoints(size.WidthEmus), PowerPointUnits.ToPoints(size.HeightEmus),
                body.Value, Math.Max(12D, (titleBox?.Y ?? body.Value.Y - 100D) - 24D), SizeScale(document));
        }
        internal PowerPointLayoutTheme Theme { get; }
        internal PowerPointLayoutGrid Grid { get; }
        internal PowerPointLayoutDeck Deck { get; }

        internal PowerPointTextStyleSpec Style(PowerPointTextRole role) => Theme[role];

        internal string FontOf(PowerPointTextStyleSpec style) => style.HeadingFont ? Theme.HeadingFont : Theme.BodyFont;

        internal double Width(string text, PowerPointTextStyleSpec style, double? size = null) =>
            PowerPointTextMeasure.Width(style.Upper ? text.ToUpperInvariant() : text, size ?? style.Size, FontOf(style), style.Bold, style.Tracking);

        internal int Lines(string text, PowerPointTextStyleSpec style, double width, double? size = null) =>
            PowerPointTextMeasure.LineCount(style.Upper ? text.ToUpperInvariant() : text, size ?? style.Size, width, FontOf(style),
                style.Bold, style.Tracking);

        /// <summary>Height of <paramref name="text" /> wrapped at <paramref name="width" />.</summary>
        /// <remarks>Frame heights count lines at a slightly narrower measure, so a renderer whose glyphs run a few percent
        /// wider than the estimate still fits the text instead of clipping its last line.</remarks>
        internal double Height(string text, PowerPointTextStyleSpec style, double width, double? size = null) =>
            SafeLines(text, style, width, size) * (size ?? style.Size) * style.LineHeight;

        /// <summary>Line count at the frame-safety measure used for every frame height.</summary>
        internal int SafeLines(string text, PowerPointTextStyleSpec style, double width, double? size = null) =>
            Lines(text, style, width / SafetyFactor, size);

        /// <summary>Frames are measured this much narrower than drawn, so a wider renderer cannot add a line.</summary>
        internal const double SafetyFactor = 1.06D;

        /// <summary>
        ///     Steps a display size down (at most <paramref name="maxStep" /> points) while the safe line count is
        ///     larger than the drawn one. Text that sits on the edge of a line break otherwise gets a frame with a
        ///     phantom line: one renderer fills it, another leaves an empty band under the text.
        /// </summary>
        internal double UnambiguousSize(string text, PowerPointTextStyleSpec style, double width, double size, double floor,
            double maxStep = 6D) {
            double resolved = size;
            while (SafeLines(text, style, width, resolved) > Lines(text, style, width, resolved) &&
                   resolved - 2D >= floor && size - (resolved - 2D) <= maxStep) {
                resolved -= 2D;
            }

            return SafeLines(text, style, width, resolved) > Lines(text, style, width, resolved) ? size : resolved;
        }

        /// <summary>The style size, reduced until the longest word fits <paramref name="width" />, so words never break.</summary>
        internal double WordFitSize(string text, PowerPointTextStyleSpec style, double width, double? floor = null) {
            double size = style.Size, minimum = floor ?? Math.Min(style.Minimum, 12D);
            while (size - 1D >= minimum && PowerPointTextMeasure.LongestWord(style.Upper ? text.ToUpperInvariant() : text, size, FontOf(style), style.Bold) * 1.05D > width) {
                size -= 1D;
            }

            return size;
        }

        internal PowerPointLayoutParagraph Paragraph(string text, PowerPointTextRole role, string? color = null,
            double? size = null, string align = "l", bool? bold = null, OfficeColor? explicitColor = null) {
            PowerPointTextStyleSpec style = Style(role);
            return new PowerPointLayoutParagraph(new PowerPointLayoutRun(text, style, size, color, bold) { ExplicitColor = explicitColor }) {
                Align = align,
                LineHeight = style.LineHeight
            };
        }

        /// <summary>Colour for a role name or hex literal supplied by content (KPI change colour, card accent).</summary>
        internal string? ResolveColor(OfficeColor? value) {
            if (!value.HasValue) return null;
            OfficeColor color = value.Value;
            if (!color.IsThemeRole) return color.A == 255 ? color.ToRgbHex() : color.ToHex();
            return color.Role switch {
                OfficeColorRole.Success => Theme.Success,
                OfficeColorRole.Warning => Theme.Warning,
                OfficeColorRole.Danger => Theme.Danger,
                OfficeColorRole.Text => Theme.Text,
                OfficeColorRole.TextMuted => Theme.TextMuted,
                OfficeColorRole.Background => Theme.Background,
                OfficeColorRole.Accent1 => Theme.Accent1,
                OfficeColorRole.Accent2 => Theme.Accent2,
                OfficeColorRole.Accent3 => Theme.Accent3,
                OfficeColorRole.Accent4 => Theme.Accent4,
                OfficeColorRole.Accent5 => Theme.Accent5,
                OfficeColorRole.Accent6 => Theme.Accent6,
                _ => null
            };

        }
    }
}
