using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>Everything a layout renderer needs: document, tokens, grid and deck state.</summary>
    internal sealed class PowerPointLayoutContext {
        internal PowerPointLayoutContext(PowerPointPresentation document, PowerPointLayoutTheme theme) {
            Document = document;
            Theme = theme;
            Grid = (document.OwnsLayoutTheme ? null : TemplateGrid(document)) ?? PowerPointLayoutGrid.For(document, theme);
            Deck = document.LayoutDeck;
        }

        internal PowerPointPresentation Document { get; }

        /// <summary>Derives the grid from a loaded deck's or template's "Title and Content" body placeholder.</summary>
        private static PowerPointLayoutGrid? TemplateGrid(PowerPointPresentation document) {
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
            if (body == null || body.Value.Width < 100D || body.Value.Height < 100D) return null;
            PowerPointSlideSize size = document.SlideSize;
            return PowerPointLayoutGrid.ForTemplate(PowerPointUnits.ToPoints(size.WidthEmus), PowerPointUnits.ToPoints(size.HeightEmus),
                body.Value, Math.Max(12D, (titleBox?.Y ?? body.Value.Y - 100D) - 24D));
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
            Lines(text, style, width / 1.06D, size) * (size ?? style.Size) * style.LineHeight;

        /// <summary>The style size, reduced until the longest word fits <paramref name="width" />, so words never break.</summary>
        internal double WordFitSize(string text, PowerPointTextStyleSpec style, double width, double? floor = null) {
            double size = style.Size, minimum = floor ?? Math.Min(style.Minimum, 12D);
            while (size - 1D >= minimum && PowerPointTextMeasure.LongestWord(style.Upper ? text.ToUpperInvariant() : text, size, FontOf(style), style.Bold) * 1.05D > width) {
                size -= 1D;
            }

            return size;
        }

        internal PowerPointLayoutParagraph Paragraph(string text, PowerPointTextRole role, string? color = null,
            double? size = null, string align = "l", bool? bold = null) {
            PowerPointTextStyleSpec style = Style(role);
            return new PowerPointLayoutParagraph(new PowerPointLayoutRun(text, style, size, color, bold)) {
                Align = align,
                LineHeight = style.LineHeight
            };
        }

        /// <summary>Colour for a role name or hex literal supplied by content (KPI change colour, card accent).</summary>
        internal string? ResolveColor(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return null;
            string v = value!.Trim().TrimStart('#');
            switch (v.ToLowerInvariant()) {
                case "success": return Theme.Success;
                case "warning": return Theme.Warning;
                case "danger": return Theme.Danger;
                case "text": return Theme.Text;
                case "textmuted": return Theme.TextMuted;
                case "accent1": return Theme.Accent1;
                case "accent2": return Theme.Accent2;
                case "accent3": return Theme.Accent3;
                case "accent4": return Theme.Accent4;
                case "accent5": return Theme.Accent5;
                case "accent6": return Theme.Accent6;
            }

            if (v.Length == 6 && v.All(Uri.IsHexDigit)) return v.ToUpperInvariant();
            return null;
        }
    }
}
