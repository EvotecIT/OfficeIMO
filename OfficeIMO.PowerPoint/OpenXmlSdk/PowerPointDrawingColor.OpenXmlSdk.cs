using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

/// <summary>
/// The SDK forms of the color helpers, for code that still builds SDK elements; they go once that code works on
/// <see cref="OfficeXmlNode"/>.
/// </summary>
internal static partial class PowerPointDrawingColor {
    internal static OpenXmlElement Create(OfficeColor color, OfficeTheme theme) {
        if (color.ThemeSlot is string slot) {
            var scheme = new A.SchemeColor();
            scheme.SetAttribute(new OpenXmlAttribute("val", string.Empty, slot));
            return scheme;
        }
        color = color.Resolve(theme);
        var rgb = new A.RgbColorModelHex { Val = color.ToRgbHex() };
        if (color.A < 255) rgb.Append(new A.Alpha { Val = (int)System.Math.Round(color.A * 100000D / 255D) });
        return rgb;
    }

    internal static A.SolidFill Fill(OfficeColor? color, OfficeTheme theme) =>
        new A.SolidFill(Create(color ?? throw new System.ArgumentNullException(nameof(color)), theme));

    /// <summary>Finds the current consuming theme without persisting wrapper state in the native XML.</summary>
    internal static OfficeTheme ThemeFor(OpenXmlPartContainer? owner) =>
        owner is OpenXmlPart part ? ThemeFor(OpenXml.PowerPointPartRef.FromSdk(part), null) : OfficeTheme.Default;

    internal static OfficeColor? Read(OpenXmlElement? container) => Read(OfficeXmlNode.FromSdk(container));
}
