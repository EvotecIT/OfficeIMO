using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointTableCell {
    private static readonly string[] LineFills = { "noFill", "solidFill", "gradFill", "pattFill" };

    internal static void ValidateBorderStyle(PowerPointTableCellBorders borders, OfficeColor color, OfficeLength? width) {
        if ((borders & ~(PowerPointTableCellBorders.All | PowerPointTableCellBorders.Diagonals)) != 0)
            throw new ArgumentOutOfRangeException(nameof(borders));
        if (width.HasValue) PowerPointShape.ToDrawingLineWidth(width.Value, nameof(width));
    }

    /// <summary>Explicitly hides the selected borders, overriding inherited table styling.
    /// <see cref="ClearBorders"/> removes the local declarations and allows style inheritance instead.</summary>
    public void HideBorders(PowerPointTableCellBorders borders) {
        if ((borders & ~(PowerPointTableCellBorders.All | PowerPointTableCellBorders.Diagonals)) != 0)
            throw new ArgumentOutOfRangeException(nameof(borders));
        if (borders == PowerPointTableCellBorders.None) return;
        OfficeXmlNode properties = EnsureProperties();
        foreach ((PowerPointTableCellBorders side, string name) in BorderElements) {
            if (!borders.HasFlag(side)) continue;
            OfficeXmlNode line = properties.GetOrAddChild(D, name, ATableCellPropertiesView.Order);
            SetBorderLineFill(line, line.CreateElement(D, "noFill"));
        }
        MarkEdited();
    }

    // Replaces the line's fill with another, first in the line as the schema places it.
    private static void SetBorderLineFill(OfficeXmlNode line, OfficeXmlNode fill) {
        foreach (string name in LineFills) line.RemoveChildren(D, name);
        line.InsertBefore(fill, line.FirstChild);
    }
}
