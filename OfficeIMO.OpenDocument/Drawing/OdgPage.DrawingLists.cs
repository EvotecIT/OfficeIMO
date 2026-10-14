using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private sealed class ListLabelText : OdfTextContent {
        private readonly IReadOnlyList<OdfStyle> _styles;
        internal bool HasExplicitLevelForeground { get; }
        internal ListLabelText(OdfTextParagraph paragraph, XElement level, double firstCharacterSize, HashSet<string> losses)
            : base(paragraph.Document, paragraph.Element, paragraph.Graphic, OdfStyleFamily.Text) {
            var styles = new List<OdfStyle>();
            XElement? properties = level.Element(OdfNamespaces.Style + "text-properties");
            if (properties != null) styles.Add(new OdfStyle(Document, new XElement(OdfNamespaces.Style + "style",
                new XAttribute(OdfNamespaces.Style + "family", "text"), new XElement(properties)), PartPath, true));
            // Draw resets the marker's alpha when it applies the level's explicit color.
            // Named label styles and paragraph/graphic colors do not qualify this reset.
            HasExplicitLevelForeground = HasExplicitTextForeground(styles);
            string? name = (string?)level.Attribute(OdfNamespaces.Text + "style-name");
            if (name != null) {
                OdfStyle? named = Document.Styles.FindInPart(OdfStyleFamily.Text, name, PartPath);
                if (named == null) losses.Add("list-label-font-style");
                else styles.AddRange(Document.Styles.Resolve(named));
            }
            // Draw percentages use the first displayed character's size, including span
            // formatting that native resaves can move out of the paragraph style.
            styles.Add(new OdfStyle(Document, new XElement(OdfNamespaces.Style + "style",
                new XAttribute(OdfNamespaces.Style + "family", "text"),
                new XElement(OdfNamespaces.Style + "text-properties",
                    new XAttribute(OdfNamespaces.Fo + "font-size", OdfLength.Points(firstCharacterSize)))), PartPath, true));
            styles.AddRange(paragraph.Styles); _styles = styles;
        }
        internal override IEnumerable<OdfStyle> Styles => _styles;
    }

    private static OfficeTextParagraphLabel? ProjectListLabel(OdfDrawingListResolver.Entry? entry, OdfTextParagraph source, double firstCharacterSize,
        HashSet<string> losses, ref double left, ref OfficeTextParagraphIndent paragraphIndent, ref double paragraphTabOrigin) {
        if (entry == null) return null;
        double originalLeft = left, originalTabOrigin = paragraphTabOrigin; OfficeTextParagraphIndent originalIndent = paragraphIndent;
        try {
            XElement level = entry.Level;
            XElement? properties = level.Element(OdfNamespaces.Style + "list-level-properties");
            XElement? labelAlignment = properties?.Element(OdfNamespaces.Style + "list-level-label-alignment");
            string mode = (string?)properties?.Attribute(OdfNamespaces.Text + "list-level-position-and-space-mode") ?? "label-width-and-position";
            OfficeTextAlignment alignment = ((string?)properties?.Attribute(OdfNamespaces.Fo + "text-align")) switch {
                null or "start" or "left" => OfficeTextAlignment.Left,
                "end" or "right" => OfficeTextAlignment.Right,
                "center" => OfficeTextAlignment.Center,
                _ => throw new NotSupportedException("Unsupported list-label alignment.")
            };
            double anchor, width = 0, distance = 0;
            OfficeTextParagraphLabelFollowedBy followedBy = OfficeTextParagraphLabelFollowedBy.Position;
            double textPosition;
            if (mode == "label-width-and-position") {
                anchor = left + ListLength(properties, OdfNamespaces.Text + "space-before");
                width = ListLength(properties, OdfNamespaces.Text + "min-label-width");
                distance = ListLength(properties, OdfNamespaces.Text + "min-label-distance");
                textPosition = anchor + width; left = textPosition;
            } else if (mode == "label-alignment") {
                string? paragraphLeft = ExplicitParagraphValue(source, OdfNamespaces.Fo + "margin-left", OdfNamespaces.Fo + "margin");
                string? paragraphFirst = ExplicitParagraphValue(source, OdfNamespaces.Fo + "text-indent");
                left = paragraphLeft == null ? ListLength(labelAlignment, OdfNamespaces.Fo + "margin-left", true) : TextLength(OdfLength.Parse(paragraphLeft), true);
                // A modern list can supply the effective paragraph margin used by body tabs,
                // including subsequent paragraphs that have no visible list label.
                paragraphTabOrigin = left;
                double first = paragraphFirst == null ? ListLength(labelAlignment, OdfNamespaces.Fo + "text-indent", true) : TextLength(OdfLength.Parse(paragraphFirst), true);
                anchor = left + first;
                followedBy = ((string?)labelAlignment?.Attribute(OdfNamespaces.Text + "label-followed-by")) switch {
                    null or "listtab" => OfficeTextParagraphLabelFollowedBy.Position,
                    "space" => OfficeTextParagraphLabelFollowedBy.Space,
                    "nothing" => OfficeTextParagraphLabelFollowedBy.Nothing,
                    _ => throw new NotSupportedException("Unsupported list-label separator.")
                };
                textPosition = followedBy == OfficeTextParagraphLabelFollowedBy.Position ? ListLength(labelAlignment, OdfNamespaces.Text + "list-tab-stop-position") : left;
                if (followedBy == OfficeTextParagraphLabelFollowedBy.Position) losses.Add("list-tab-stops");
            } else throw new NotSupportedException("Unsupported list-label spacing mode.");
            if (left < 0 || anchor < 0) throw new NotSupportedException("List text outside the frame is not projected.");
            paragraphIndent = OfficeTextParagraphIndent.Empty;
            if (entry.Label == null || entry.Label.Length == 0) return null;
            var labelSource = new ListLabelText(source, level, firstCharacterSize, losses);
            OfficeRichTextRun run = CreateDrawingRun(entry.Label, labelSource, losses);
            string? relative = (string?)level.Attribute(OdfNamespaces.Text + "bullet-relative-size");
            if (relative != null) run = new OfficeRichTextRun(run.Text, ResolveTextLength(OdfLength.Parse(relative), OdfTextStyleResolver.ResolveFontSize(source.Styles, losses)), run.Color,
                run.Bold, run.Italic, run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle, run.Baseline);
            return mode == "label-width-and-position" ? OfficeTextParagraphLabel.InBox(run, anchor, width, alignment, distance) :
                OfficeTextParagraphLabel.AtPosition(run, anchor, alignment, followedBy, textPosition);
        } catch (Exception exception) when (exception is ArgumentException or FormatException or NotSupportedException or OverflowException) {
            left = originalLeft; paragraphIndent = originalIndent; paragraphTabOrigin = originalTabOrigin;
            losses.Add("list-label-layout"); return null;
        }
    }

    private static double ListLength(XElement? element, XName name, bool allowNegative = false) {
        string? value = (string?)element?.Attribute(name);
        return value == null ? 0 : TextLength(OdfLength.Parse(value), allowNegative);
    }
    private static string? ExplicitParagraphValue(OdfTextParagraph source, XName name, XName? shorthand = null) => source.Styles
        .Where(s => s.Family == OdfStyleFamily.Paragraph && s.Element.Name == OdfNamespaces.Style + "style")
        .Select(s => (string?)s.ParagraphProperties?.Attribute(name) ?? (shorthand == null ? null : (string?)s.ParagraphProperties?.Attribute(shorthand)))
        .FirstOrDefault(v => v != null);
}
