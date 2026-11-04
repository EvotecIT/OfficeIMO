using DocumentFormat.OpenXml;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

/// <summary>Encodes explicit DrawingML colors without resolving native theme slots to RGB.</summary>
internal static class PowerPointDrawingColor {
    internal static OpenXmlElement Create(OfficeColor color, OfficeTheme? theme = null) {
        if (color.ThemeSlot is string slot) {
            var scheme = new A.SchemeColor();
            scheme.SetAttribute(new OpenXmlAttribute("val", string.Empty, slot));
            return scheme;
        }
        color = color.Resolve(theme ?? OfficeTheme.Default);
        var rgb = new A.RgbColorModelHex { Val = color.ToRgbHex() };
        if (color.A < 255) rgb.Append(new A.Alpha { Val = (int)System.Math.Round(color.A * 100000D / 255D) });
        return rgb;
    }

    internal static A.SolidFill Fill(OfficeColor? color, OfficeTheme? theme = null) =>
        new A.SolidFill(Create(color ?? throw new System.ArgumentNullException(nameof(color)), theme));

    // These getters expose local formatting. Inherited fills and transformed scheme colors
    // have no exact OfficeColor representation and remain absent rather than being flattened.
    internal static OfficeColor? Read(OpenXmlElement? container) {
        A.SchemeColor? scheme = container?.GetFirstChild<A.SchemeColor>();
        if (scheme != null) {
            if (scheme.HasChildren) return null;
            return scheme.Val?.InnerText switch {
                "accent1" => OfficeColor.Accent1, "accent2" => OfficeColor.Accent2,
                "accent3" => OfficeColor.Accent3, "accent4" => OfficeColor.Accent4,
                "accent5" => OfficeColor.Accent5, "accent6" => OfficeColor.Accent6,
                "dk1" => OfficeColor.Text, "dk2" => OfficeColor.TextMuted,
                "lt1" => OfficeColor.Background, _ => null
            };
        }
        return OfficeOpenXmlThemeColorResolver.ResolveColor(container, null);
    }
}
