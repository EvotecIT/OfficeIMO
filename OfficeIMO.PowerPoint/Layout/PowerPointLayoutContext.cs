using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;
using Ns = OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces;

namespace OfficeIMO.PowerPoint {
    /// <summary>Everything a layout renderer needs: document, tokens, grid and deck state.</summary>
    internal sealed class PowerPointLayoutContext {
        internal PowerPointLayoutContext(PowerPointDocument document, PowerPointLayoutTheme theme, PowerPointDesignIntent? intent = null) {
            Document = document;
            Intent = intent ?? document.LayoutDeck.Design?.BaseIntent;
            Theme = theme.WithIntent(Intent).Scaled(SizeScale(document));
            Grid = (document.OwnsLayoutTheme ? null : TemplateGrid(document)) ?? PowerPointLayoutGrid.For(document, theme);
            Deck = document.LayoutDeck;
        }

        internal PowerPointDocument Document { get; }
        internal PowerPointDesignIntent? Intent { get; }

        /// <summary>
        ///     The slide size relative to the 960 x 540 pt reference (and 720 x 540 pt for 4:3). The type scale and
        ///     spacing are designed for those sizes; a smaller slide, such as a 10 inch 720 x 405 pt template, scales
        ///     them down so the same content keeps the same composition instead of overflowing.
        /// </summary>
        internal static double SizeScale(PowerPointDocument document) {
            PowerPointSlideSize size = document.SlideSize;
            double width = size.Width.ToPoints(), height = size.Height.ToPoints();
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

        private PowerPointLayoutGrid? TemplateGrid(PowerPointDocument document) {
            PowerPointPartRef master = document.LayoutPresentationPartRef.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType);
            if (master.IsNone) return null;
            PowerPointPartRef[] layouts = master.RelatedParts(PowerPointPartRef.SlideLayoutRelationshipType).ToArray();
            PowerPointPartRef layout = layouts.FirstOrDefault(l => l.Root.Attribute("type") == "obj");
            if (layout.IsNone) layout = layouts.FirstOrDefault(l => l.Root.Attribute("type") == "tx");
            OfficeXmlNode masterRoot = master.Root;
            OfficeXmlNode layoutRoot = layout.IsNone ? default : layout.Root;
            PowerPointLayoutRect? Bounds(OfficeXmlNode root, bool title) {
                if (root.IsNone) return null;
                foreach (OfficeXmlNode shape in root.Descendants(Ns.PresentationML, "sp")) {
                    OfficeXmlNode ph = shape.Child(Ns.PresentationML, "nvSpPr").Child(Ns.PresentationML, "nvPr").Child(Ns.PresentationML, "ph");
                    if (ph.IsNone) continue;
                    string? type = ph.Attribute("type");
                    bool isTitle = type is "title" or "ctrTitle";
                    bool isBody = type is null or "body" or "obj";
                    if (title ? !isTitle : !isBody) continue;
                    OfficeXmlNode xfrm = shape.Child(Ns.PresentationML, "spPr").Child(Ns.DrawingML, "xfrm");
                    OfficeXmlNode offset = xfrm.Child(Ns.DrawingML, "off"), extents = xfrm.Child(Ns.DrawingML, "ext");
                    if (offset.IsNone || extents.IsNone) continue;
                    return new PowerPointLayoutRect((offset.Int64Attribute("x") ?? 0L) / 12700D, (offset.Int64Attribute("y") ?? 0L) / 12700D,
                        (extents.Int64Attribute("cx") ?? 0L) / 12700D, (extents.Int64Attribute("cy") ?? 0L) / 12700D);
                }

                return null;
            }

            PowerPointLayoutRect? body = Bounds(layoutRoot, false) ?? Bounds(masterRoot, false);
            PowerPointLayoutRect? titleBox = Bounds(layoutRoot, true) ?? Bounds(masterRoot, true);
            MasterTitle = Bounds(masterRoot, true);
            if (body == null || body.Value.Width < 100D || body.Value.Height < 100D) return null;
            if (!PowerPointLayoutEngine.WithinSlide(document, body.Value.X, body.Value.Y, body.Value.Width, body.Value.Height))
                return null;
            PowerPointSlideSize size = document.SlideSize;
            return PowerPointLayoutGrid.ForTemplate(PowerPointUnits.ToPoints(size.Width.ToEmus()), PowerPointUnits.ToPoints(size.Height.ToEmus()),
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
