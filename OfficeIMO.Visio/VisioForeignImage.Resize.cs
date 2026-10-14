using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

internal static partial class VisioForeignImage {
    /// <summary>Materializes the effective image rectangle on a detached resized instance, retaining its crop fractions.</summary>
    internal static void ScalePlacement(VisioShape source, VisioShape target, double x, double y) {
        if (!IsForeign(source)) return;
        foreach (string name in new[] { "ImgOffsetX", "ImgOffsetY", "ImgWidth", "ImgHeight" }) {
            bool horizontal = name == "ImgOffsetX" || name == "ImgWidth";
            double fallback = name == "ImgWidth" ? source.Width : name == "ImgHeight" ? source.Height : 0;
            double value = ReadCell(source, name, fallback, null, null) * (horizontal ? x : y);
            XElement? original = source.PreservedCellElements.FirstOrDefault(c => (string?)c.Attribute("N") == name);
            if (original == null || (string?)original.Attribute("F") == "Inh")
                original = MasterShape(source)?.PreservedCellElements.FirstOrDefault(c => (string?)c.Attribute("N") == name) ?? original;
            XElement cell = original == null
                ? new XElement(XName.Get("Cell", "http://schemas.microsoft.com/office/visio/2012/main"), new XAttribute("N", name), new XAttribute("U", "IN"))
                : new XElement(original);
            VisioGeometryScaling.MaterializeCell(cell, value, target);
            XElement? previous = target.PreservedCellElements.FirstOrDefault(c => (string?)c.Attribute("N") == name);
            if (previous == null) target.PreservedCellElements.Add(cell);
            else target.PreservedCellElements[target.PreservedCellElements.IndexOf(previous)] = cell;
        }
    }
}
