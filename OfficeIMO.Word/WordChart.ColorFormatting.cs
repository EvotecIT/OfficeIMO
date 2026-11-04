using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word;

public partial class WordChart {
    private void SetChartFill(OpenXmlCompositeElement properties, OfficeColor color) {
        properties.RemoveAllChildren<A.NoFill>();
        properties.RemoveAllChildren<A.SolidFill>();
        properties.RemoveAllChildren<A.GradientFill>();
        properties.RemoveAllChildren<A.PatternFill>();
        properties.RemoveAllChildren<A.BlipFill>();
        properties.RemoveAllChildren<A.GroupFill>();
        properties.AddChild(CreateChartFill(color), true);
    }

    private A.SolidFill CreateChartFill(OfficeColor color) {
        if (color.ThemeSlot is string slot) {
            var scheme = new A.SchemeColor();
            scheme.SetAttribute(new OpenXmlAttribute("val", string.Empty, slot));
            return new A.SolidFill(scheme);
        }
        color = color.Resolve(_document?.EffectiveCreationTheme ?? OfficeTheme.Default);
        var rgb = new A.RgbColorModelHex { Val = color.ToRgbHex() };
        if (color.A < 255) rgb.Append(new A.Alpha { Val = (int)System.Math.Round(color.A * 100000D / 255D) });
        return new A.SolidFill(rgb);
    }
}
