using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private static OfficeColor DrawingTextColor(OdfTextContent source, HashSet<string> losses) {
        OfficeColor color = ToColor(source.Color) ?? OfficeColor.Black;
        OdfStyle[] styles = source.Styles.ToArray();
        TextForegroundKind foreground = ResolveTextForeground(styles);
        if (foreground == TextForegroundKind.UnsupportedWindowColor) losses.Add("text-window-color");
        double? opacity = null;
        bool invalid = false;
        foreach (OdfStyle style in styles) {
            XElement? properties = style.TextProperties;
            foreach (XAttribute attribute in properties?.Attributes() ?? Enumerable.Empty<XAttribute>())
                if (attribute.Name.LocalName == "opacity" && attribute.Name != OdfNamespaces.LoExt + "opacity" &&
                    attribute.Name != OdfNamespaces.Draw + "opacity") losses.Add("text-opacity");
            if (opacity.HasValue || invalid) continue;
            try {
                opacity = style.TextOpacity;
            } catch (InvalidDataException) { invalid = true; }
        }
        if (invalid) { losses.Add("text-opacity"); return color; }
        // A separately styled leader compiles RGB/opacity overrides independently;
        // its missing properties inherit the active tab run during shared layout.
        if (source is TabLeaderText) return opacity.HasValue ? OfficeColorTransforms.WithAlpha(color, opacity.Value) : color;
        if (!opacity.HasValue || opacity.Value == 1) return color;
        // Explicit level colors keep native Draw markers opaque independently of text
        // opacity. Named/default label foregrounds retain the unqualified fallback.
        if (source is ListLabelText label) {
            if (!label.HasExplicitLevelForeground) losses.Add("text-opacity");
            return color;
        }
        if (foreground != TextForegroundKind.FixedColor) {
            losses.Add("text-opacity"); return color;
        }
        return OfficeColorTransforms.WithAlpha(color, opacity.Value);
    }

    private enum TextForegroundKind { Unqualified, FixedColor, UnsupportedWindowColor }

    private static bool HasExplicitTextForeground(IEnumerable<OdfStyle> styles) =>
        ResolveTextForeground(styles) == TextForegroundKind.FixedColor;

    // Preserve the existing nearest-foreground rule: a nearer color stops an
    // inherited window policy, while an explicit false/0 policy allows color inheritance.
    private static TextForegroundKind ResolveTextForeground(IEnumerable<OdfStyle> styles) {
        bool explicitWindowPolicy = false;
        foreach (OdfStyle style in styles) {
            XElement? properties = style.TextProperties;
            string? window = (string?)properties?.Attribute(OdfNamespaces.Style + "use-window-font-color");
            if (!explicitWindowPolicy && window != null) {
                try {
                    if (XmlConvert.ToBoolean(window)) return TextForegroundKind.UnsupportedWindowColor;
                } catch (FormatException) { return TextForegroundKind.UnsupportedWindowColor; }
                explicitWindowPolicy = true;
            }
            if (properties?.Attribute(OdfNamespaces.Fo + "color") != null) return TextForegroundKind.FixedColor;
        }
        return TextForegroundKind.Unqualified;
    }
}
