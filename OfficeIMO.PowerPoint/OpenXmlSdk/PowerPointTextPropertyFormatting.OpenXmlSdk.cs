using OfficeIMO.PowerPoint.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

/// <summary>The SDK forms of the text property helpers, for code that still holds SDK elements.</summary>
internal static partial class PowerPointTextPropertyFormatting {
    internal static void SetFontFamily(A.TextCharacterPropertiesType properties, string? value) =>
        SetFontFamily(PowerPointXmlNode.FromSdk(properties), value);

    internal static void SetColor(A.TextCharacterPropertiesType properties, OfficeColor? value, OfficeTheme theme) =>
        SetColor(PowerPointXmlNode.FromSdk(properties), value, theme);

    internal static void SetBackgroundColor(A.TextCharacterPropertiesType properties, OfficeColor? value, OfficeTheme theme) =>
        SetBackgroundColor(PowerPointXmlNode.FromSdk(properties), value, theme);
}
