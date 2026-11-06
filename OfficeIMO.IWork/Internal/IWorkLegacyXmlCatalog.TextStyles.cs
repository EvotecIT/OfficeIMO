using System.Xml.Linq;

namespace OfficeIMO.IWork.Internal;

internal sealed partial class IWorkLegacyXmlCatalog {
    private static readonly IWorkTextStyle PlainTextStyle = new(null, null, null, null, null, null, null, null, null);
    private static readonly IWorkParagraphStyle PlainParagraphStyle =
        new(null, null, null, null, null, null, null, null, null, null, PlainTextStyle);
    private readonly Dictionary<XElement, IWorkParagraphStyle> _paragraphStyles = new();
    private readonly Dictionary<(XElement Style, IWorkTextStyle Inherited), IWorkTextStyle> _characterStyles = new();

    internal IWorkParagraphStyle ParagraphStyle(XElement? style) {
        if (style == null) return PlainParagraphStyle;
        if (_paragraphStyles.TryGetValue(style, out IWorkParagraphStyle? cached)) return cached;
        IWorkTextAlignment? alignment = NumberProperty(style, "alignment") switch {
            0 => IWorkTextAlignment.Left,
            1 => IWorkTextAlignment.Right,
            2 => IWorkTextAlignment.Center,
            3 => IWorkTextAlignment.Justified,
            4 => IWorkTextAlignment.Natural,
            _ => null
        };
        var result = new IWorkParagraphStyle((string?)style.Attribute(Sf + "name"), alignment,
            NumberProperty(style, "firstLineIndent"), NumberProperty(style, "leftIndent"),
            NumberProperty(style, "rightIndent"), NumberProperty(style, "spaceBefore"),
            NumberProperty(style, "spaceAfter"), BooleanProperty(style, "pageBreakBefore"),
            BooleanProperty(style, "keepWithNext"), BooleanProperty(style, "keepLinesTogether"),
            TextStyle(style, PlainTextStyle));
        _paragraphStyles.Add(style, result);
        return result;
    }

    private IWorkTextStyle TextStyle(XElement? style, IWorkTextStyle inherited) {
        if (style == null) return inherited;
        var key = (style, inherited);
        if (_characterStyles.TryGetValue(key, out IWorkTextStyle? cached)) return cached;
        XElement? fontName = Property(style, "fontName");
        XElement? fontSize = Property(style, "fontSize");
        XElement? foreground = Property(style, "fontColor");
        XElement? background = Property(style, "textBackground");
        double? points = fontSize == null ? inherited.FontSizePoints : NumberProperty(style, "fontSize");
        if (points is <= 0) {
            points = null;
            AddLoss("IWORK_LEGACY_FONT_SIZE_UNSUPPORTED", "A selected XML font size is outside the supported positive range.");
        }
        var result = new IWorkTextStyle((string?)style.Attribute(Sf + "name"),
            Property(style, "bold") == null ? inherited.Bold : BooleanProperty(style, "bold"),
            Property(style, "italic") == null ? inherited.Italic : BooleanProperty(style, "italic"),
            Property(style, "underline") == null ? inherited.Underline : BooleanProperty(style, "underline"),
            Property(style, "strikethru") == null ? inherited.Strikethrough : BooleanProperty(style, "strikethru"),
            points, fontName == null ? inherited.FontName : (string?)fontName.Element(Sf + "string")?.Attribute(Sfa + "string"),
            foreground == null ? inherited.Color : Color(foreground.Elements().FirstOrDefault(e => e.Name != Sf + "null")),
            background == null ? inherited.BackgroundColor : Color(background.Elements().FirstOrDefault(e => e.Name != Sf + "null")));
        _characterStyles.Add(key, result);
        return result;
    }

    private double? NumberProperty(XElement style, string name) {
        XElement? property = Property(style, name);
        XElement? number = property?.Element(Sf + "number");
        double? value = AttributeNumber(number, Sfa + "number");
        if (number != null && value == null)
            AddLoss("IWORK_LEGACY_STYLE_VALUE_UNSUPPORTED", "A selected XML style property has an unqualified value.");
        return value;
    }

    private bool? BooleanProperty(XElement style, string name) => NumberProperty(style, name) switch {
        0 => false,
        1 => true,
        _ => null
    };
}
