using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordChart {
    private static readonly string[] FillNames = { "noFill", "solidFill", "gradFill", "pattFill", "blipFill", "grpFill" };

    /// <summary>Replaces the fill of shape, line or text properties with a solid <paramref name="color"/>.</summary>
    private void SetChartFill(OfficeXmlNode properties, OfficeColor color) {
        foreach (string fill in FillNames) properties.RemoveChildren(NsA, fill);
        OfficeChartXml.SetOrdered(properties, CreateChartFill(properties, color), DrawingOrderOf(properties));
    }

    /// <summary>A detached <c>a:solidFill</c> of <paramref name="color"/>, made for <paramref name="context"/>.</summary>
    private OfficeXmlNode CreateChartFill(OfficeXmlNode context, OfficeColor color) {
        OfficeXmlNode fill = context.CreateElement(NsA, "solidFill");
        if (color.ThemeSlot is string slot) {
            AddA(fill, "schemeClr").SetAttribute("val", slot);
            return fill;
        }
        color = color.Resolve(_document?.EffectiveCreationTheme ?? OfficeTheme.Default);
        OfficeXmlNode rgb = AddA(fill, "srgbClr");
        rgb.SetAttribute("val", color.ToRgbHex());
        if (color.A < 255) rgb.AppendVal(NsA, "alpha", OfficeChartXml.Text((int)System.Math.Round(color.A * 100000D / 255D)));
        return fill;
    }

    /// <summary>The schema order of DrawingML shape, line or text-run properties.</summary>
    private static OfficeXmlChildOrder DrawingOrderOf(OfficeXmlNode properties) => properties.LocalName switch {
        "ln" => DmlLinePropertiesView.Order,
        "defRPr" or "rPr" or "endParaRPr" => DmlTextCharacterPropertiesView.Order,
        _ => DmlShapePropertiesView.Order
    };
}
