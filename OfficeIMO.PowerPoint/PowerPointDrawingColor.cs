using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint;

/// <summary>Encodes explicit DrawingML colors without resolving native theme slots to RGB.</summary>
internal static partial class PowerPointDrawingColor {
    /// <summary>Whether writing <paramref name="color"/> needs the consuming theme: a theme role without a slot.</summary>
    internal static bool NeedsTheme(OfficeColor? color) => color is { IsThemeRole: true, ThemeSlot: null };

    /// <summary>
    /// The theme content in <paramref name="part"/> consumes: its slide's authoring theme, or its layout's and master's.
    /// </summary>
    internal static OfficeTheme ThemeFor(PowerPointPartRef part, PowerPointSlide? slide) {
        if (slide != null) return slide.AuthoringTheme;
        if (part.IsNone || part.Document is not { } document) return OfficeTheme.Default;
        PowerPointPartRef owner = part.Kind == PowerPointPartKind.NotesSlide ? part.RelatedPart(PowerPointPartRef.SlideRelationshipType) : part;
        if (owner.Kind == PowerPointPartKind.Slide) {
            foreach (PowerPointSlide candidate in document.Slides) {
                if (candidate.PartRef.Equals(owner)) return candidate.AuthoringTheme;
            }
        }
        PowerPointThemeScope scope = PowerPointThemeScope.For(owner);
        return document.ReadLayoutTheme(scope.Master.IsNone ? document.MasterPartRef(0) : scope.Master, scope.ColorScheme, scope.FontScheme).Theme;
    }

    /// <summary>
    /// Appends the color element for <paramref name="color"/>: a theme slot as <c>a:schemeClr</c>, anything else as
    /// <c>a:srgbClr</c> with its opacity. <paramref name="theme"/> resolves a theme role without a slot.
    /// </summary>
    internal static OfficeXmlNode Append(OfficeXmlNode parent, OfficeColor color, OfficeTheme? theme) {
        if (color.ThemeSlot is string slot) {
            OfficeXmlNode scheme = parent.Append(parent.CreateElement(OfficeXmlNamespaces.DrawingML, "schemeClr"));
            scheme.SetAttribute("val", slot);
            return scheme;
        }
        if (color.IsThemeRole) color = color.Resolve(theme ?? OfficeTheme.Default);
        OfficeXmlNode rgb = parent.CreateElement(OfficeXmlNamespaces.DrawingML, "srgbClr");
        rgb.SetAttribute("val", color.ToRgbHex());
        if (color.A < 255) {
            rgb.Append(rgb.CreateElement(OfficeXmlNamespaces.DrawingML, "alpha"))
                .SetInt32Attribute("val", (int)Math.Round(color.A * 100000D / 255D));
        }
        return parent.Append(rgb);
    }

    /// <summary>Builds a detached <c>a:solidFill</c> of <paramref name="color"/> in the part of <paramref name="context"/>.</summary>
    internal static OfficeXmlNode CreateFill(OfficeXmlNode context, OfficeColor color, OfficeTheme? theme) {
        OfficeXmlNode fill = context.CreateElement(OfficeXmlNamespaces.DrawingML, "solidFill");
        Append(fill, color, theme);
        return fill;
    }

    // These getters expose local formatting. Inherited fills and transformed scheme colors
    // have no exact OfficeColor representation and remain absent rather than being flattened.
    internal static OfficeColor? Read(OfficeXmlNode container) {
        if (container.IsNone) return null;
        OfficeXmlNode scheme = container.Child(OfficeXmlNamespaces.DrawingML, "schemeClr");
        if (!scheme.IsNone) {
            if (!scheme.FirstChild.IsNone) return null;
            return scheme.Attribute("val") switch {
                "accent1" => OfficeColor.Accent1, "accent2" => OfficeColor.Accent2,
                "accent3" => OfficeColor.Accent3, "accent4" => OfficeColor.Accent4,
                "accent5" => OfficeColor.Accent5, "accent6" => OfficeColor.Accent6,
                "dk1" => OfficeColor.Text, "dk2" => OfficeColor.TextMuted,
                "lt1" => OfficeColor.Background, _ => null
            };
        }
        return ReadLiteral(FindColorElement(container));
    }

    // The color element itself, or the first color child in the order the shared resolver looks for them.
    private static OfficeXmlNode FindColorElement(OfficeXmlNode container) {
        if (container.NamespaceId == OfficeXmlNamespaces.DrawingML && IsColorElement(container.LocalName)) return container;
        foreach (string name in ColorElements) {
            OfficeXmlNode color = container.Child(OfficeXmlNamespaces.DrawingML, name);
            if (!color.IsNone) return color;
        }
        return default;
    }

    private static readonly string[] ColorElements = { "srgbClr", "scrgbClr", "hslClr", "schemeClr", "sysClr", "prstClr" };

    private static bool IsColorElement(string name) => Array.IndexOf(ColorElements, name) >= 0;

    // A color model without a theme: scheme colors have nothing to resolve against and read as null.
    private static OfficeColor? ReadLiteral(OfficeXmlNode element) {
        if (element.IsNone) return null;
        OfficeColor? color = element.LocalName switch {
            "srgbClr" => OfficeDrawingColorTransforms.ParseRgb(element.Attribute("val")),
            "scrgbClr" => OfficeDrawingColorTransforms.ParseScRgb(Percentage(element, "r"), Percentage(element, "g"), Percentage(element, "b")),
            "hslClr" => OfficeDrawingColorTransforms.ParseHsl(element.Int32Attribute("hue"), Percentage(element, "sat"), Percentage(element, "lum")),
            "sysClr" => OfficeDrawingColorTransforms.ParseRgb(element.Attribute("lastClr")),
            "prstClr" => OfficeDrawingColorTransforms.ParsePreset(element.Attribute("val")),
            _ => null
        };
        if (!color.HasValue) return null;
        var transforms = new OfficeDrawingColorTransforms(color.Value);
        foreach (OfficeXmlNode transform in element.Children()) {
            string name = transform.LocalName;
            transforms.Apply(name, OfficeDrawingColorTransforms.TakesValue(name) ? transform.Attribute("val") : null);
        }
        return transforms.Result;
    }

    private static int? Percentage(OfficeXmlNode element, string localName) => OfficeDrawingColorTransforms.ParsePercentage(element.Attribute(localName));
}
