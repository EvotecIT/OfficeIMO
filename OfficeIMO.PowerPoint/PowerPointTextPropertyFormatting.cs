using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

/// <summary>Writes run fonts, fills and highlights into DrawingML text character properties.</summary>
internal static partial class PowerPointTextPropertyFormatting {
    private static readonly string[] OtherFills = { "noFill", "gradFill", "blipFill", "pattFill", "grpFill" };

    internal static void SetFontFamily(OfficeXmlNode properties, string? value) {
        properties.RemoveChildren(OfficeXmlNamespaces.DrawingML, "latin");
        if (value != null) properties.AddOrdered(OfficeXmlNamespaces.DrawingML, "latin", ATextCharacterPropertiesView.Order).SetAttribute("typeface", value);
    }

    /// <summary>Writes a text fill, retaining native theme slots and opacity.</summary>
    internal static void SetColor(OfficeXmlNode properties, OfficeColor? value, OfficeTheme? theme) {
        properties.RemoveChildren(OfficeXmlNamespaces.DrawingML, "solidFill");
        if (!value.HasValue) return;
        foreach (string fill in OtherFills) properties.RemoveChildren(OfficeXmlNamespaces.DrawingML, fill);
        properties.InsertOrdered(PowerPointDrawingColor.CreateFill(properties, value.Value, theme), ATextCharacterPropertiesView.Order);
        PowerPointTextHyperlinkColor.UseExplicitTextColor(properties);
    }

    /// <summary>Writes the run highlight color, retaining native theme slots and opacity.</summary>
    internal static void SetBackgroundColor(OfficeXmlNode properties, OfficeColor? value, OfficeTheme? theme) {
        properties.RemoveChildren(OfficeXmlNamespaces.DrawingML, "highlight");
        if (!value.HasValue) return;
        OfficeXmlNode highlight = properties.CreateElement(OfficeXmlNamespaces.DrawingML, "highlight");
        PowerPointDrawingColor.Append(highlight, value.Value, theme);
        properties.InsertOrdered(highlight, ATextCharacterPropertiesView.Order);
    }
}
