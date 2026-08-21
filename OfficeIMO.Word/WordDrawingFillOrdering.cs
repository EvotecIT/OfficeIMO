using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word;

internal static class WordDrawingFillOrdering {
    internal static void InsertAfterGeometryOrBeforeFormatting(
        DocumentFormat.OpenXml.OpenXmlCompositeElement properties,
        DocumentFormat.OpenXml.OpenXmlElement fill) {
        DocumentFormat.OpenXml.OpenXmlElement? geometry =
            (DocumentFormat.OpenXml.OpenXmlElement?)properties.GetFirstChild<A.CustomGeometry>()
            ?? properties.GetFirstChild<A.PresetGeometry>();
        if (geometry != null) {
            properties.InsertAfter(fill, geometry);
            return;
        }

        DocumentFormat.OpenXml.OpenXmlElement? laterFormatting = properties.ChildElements.FirstOrDefault(child =>
            child is A.Outline
            || child is A.EffectList
            || child is A.EffectDag
            || child is A.ShapePropertiesExtensionList
            || child.LocalName.Equals("scene3d", StringComparison.Ordinal)
            || child.LocalName.Equals("sp3d", StringComparison.Ordinal));
        if (laterFormatting != null) properties.InsertBefore(fill, laterFormatting);
        else properties.Append(fill);
    }
}