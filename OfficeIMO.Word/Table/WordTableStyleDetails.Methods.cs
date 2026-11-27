using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

// Manages table style settings and borders.
public partial class WordTableStyleDetails {
    // Element names and matching cell edges, in the order of WordTableBorderSide.
    private static readonly string[] SideNames = { "top", "bottom", "left", "right", "insideH", "insideV" };
    private static readonly WordTableCellBorder.Edge[] SideEdges = {
        WordTableCellBorder.Edge.Top, WordTableCellBorder.Edge.Bottom, WordTableCellBorder.Edge.Left,
        WordTableCellBorder.Edge.Right, WordTableCellBorder.Edge.InsideHorizontal, WordTableCellBorder.Edge.InsideVertical
    };

    // Word writes w:left before w:start and w:right before w:end when both are present.
    private static readonly OfficeXmlChildOrder BorderOrder = new(
        new[] { W, W, W, W, W, W, W, W },
        new[] { "top", "left", "start", "bottom", "right", "end", "insideH", "insideV" });

    // Only the table's own w:tblBorders counts, not one inside extension markup.
    private OfficeXmlNode Borders => _properties.Child(W, "tblBorders");

    /// <summary>
    /// Creates and sets a TableBorders object with the specified border settings
    /// </summary>
    /// <param name="style">Border style for all sides</param>
    /// <param name="size">Nonnegative thickness, rounded to eighths of a point within the native unsigned 32-bit range.</param>
    /// <param name="color">Border color for all sides</param>
    public void SetBordersForAllSides(WordBorderStyle style, global::OfficeIMO.OfficeLength size, OfficeIMO.OfficeColor color) {
        SetBordersOutsideInside(style, size, color, style, size, color);
    }

    /// <summary>
    /// Creates a TableBorders object with different settings for outside and inside borders
    /// </summary>
    /// <param name="outsideStyle">Style for outside borders</param>
    /// <param name="outsideSize">Nonnegative outside thickness, rounded to eighths of a point.</param>
    /// <param name="outsideColor">Color for outside borders</param>
    /// <param name="insideStyle">Style for inside borders</param>
    /// <param name="insideSize">Nonnegative inside thickness, rounded to eighths of a point.</param>
    /// <param name="insideColor">Color for inside borders</param>
    public void SetBordersOutsideInside(
        WordBorderStyle outsideStyle, global::OfficeIMO.OfficeLength outsideSize, OfficeIMO.OfficeColor outsideColor,
        WordBorderStyle insideStyle, global::OfficeIMO.OfficeLength insideSize, OfficeIMO.OfficeColor insideColor) {
        uint nativeOutsideSize = ToBorderSize(outsideSize, nameof(outsideSize));
        uint nativeInsideSize = ToBorderSize(insideSize, nameof(insideSize));
        string outsideToken = outsideStyle.ToOpenXmlToken();
        string insideToken = insideStyle.ToOpenXmlToken();
        string outsideRgb = outsideColor.ToRgbHex();
        string insideRgb = insideColor.ToRgbHex();

        // The new borders replace any the table had.
        _table.CheckTableProperties();
        Borders.Remove();
        OfficeXmlNode borders = _properties.AddOrdered(W, "tblBorders", WTblPrView.Order);
        foreach (string side in new[] { "top", "left", "bottom", "right" }) {
            WriteBorder(borders.AddOrdered(W, side, BorderOrder), outsideToken, nativeOutsideSize, outsideRgb);
        }
        WriteBorder(borders.AddOrdered(W, "insideH", BorderOrder), insideToken, nativeInsideSize, insideRgb);
        WriteBorder(borders.AddOrdered(W, "insideV", BorderOrder), insideToken, nativeInsideSize, insideRgb);
    }

    /// <summary>
    /// Updates only the supplied border settings. Null arguments preserve existing values.
    /// Thicknesses are nonnegative and round to eighths of a point within the native unsigned 32-bit range.
    /// All thicknesses are validated before any border is changed.
    /// </summary>
    /// <param name="topStyle">Optional border style.</param>
    /// <param name="topSize">Optional physical thickness; zero remains explicit.</param>
    /// <param name="topColor">Optional border color.</param>
    /// <param name="bottomStyle">Optional border style.</param>
    /// <param name="bottomSize">Optional physical thickness; zero remains explicit.</param>
    /// <param name="bottomColor">Optional border color.</param>
    /// <param name="leftStyle">Optional border style.</param>
    /// <param name="leftSize">Optional physical thickness; zero remains explicit.</param>
    /// <param name="leftColor">Optional border color.</param>
    /// <param name="rightStyle">Optional border style.</param>
    /// <param name="rightSize">Optional physical thickness; zero remains explicit.</param>
    /// <param name="rightColor">Optional border color.</param>
    /// <param name="insideHStyle">Optional border style.</param>
    /// <param name="insideHSize">Optional physical thickness; zero remains explicit.</param>
    /// <param name="insideHColor">Optional border color.</param>
    /// <param name="insideVStyle">Optional border style.</param>
    /// <param name="insideVSize">Optional physical thickness; zero remains explicit.</param>
    /// <param name="insideVColor">Optional border color.</param>
    public void SetCustomBorders(
        WordBorderStyle? topStyle = null, global::OfficeIMO.OfficeLength? topSize = null, OfficeIMO.OfficeColor? topColor = null,
        WordBorderStyle? bottomStyle = null, global::OfficeIMO.OfficeLength? bottomSize = null, OfficeIMO.OfficeColor? bottomColor = null,
        WordBorderStyle? leftStyle = null, global::OfficeIMO.OfficeLength? leftSize = null, OfficeIMO.OfficeColor? leftColor = null,
        WordBorderStyle? rightStyle = null, global::OfficeIMO.OfficeLength? rightSize = null, OfficeIMO.OfficeColor? rightColor = null,
        WordBorderStyle? insideHStyle = null, global::OfficeIMO.OfficeLength? insideHSize = null, OfficeIMO.OfficeColor? insideHColor = null,
        WordBorderStyle? insideVStyle = null, global::OfficeIMO.OfficeLength? insideVSize = null, OfficeIMO.OfficeColor? insideVColor = null) {
        // Validate every value before changing anything, in the order of WordTableBorderSide.
        WordBorderStyle?[] styles = { topStyle, bottomStyle, leftStyle, rightStyle, insideHStyle, insideVStyle };
        OfficeIMO.OfficeColor?[] colors = { topColor, bottomColor, leftColor, rightColor, insideHColor, insideVColor };
        global::OfficeIMO.OfficeLength?[] lengths = { topSize, bottomSize, leftSize, rightSize, insideHSize, insideVSize };
        string[] sizeNames = { nameof(topSize), nameof(bottomSize), nameof(leftSize), nameof(rightSize), nameof(insideHSize), nameof(insideVSize) };
        uint?[] sizes = new uint?[SideNames.Length];
        string?[] tokens = new string?[SideNames.Length];
        for (int side = 0; side < SideNames.Length; side++) {
            sizes[side] = lengths[side].HasValue ? ToBorderSize(lengths[side]!.Value, sizeNames[side]) : null;
            tokens[side] = styles[side]?.ToOpenXmlToken();
        }

        _table.CheckTableProperties();
        OfficeXmlNode borders = _properties.GetOrAddChild(W, "tblBorders", WTblPrView.Order);
        for (int side = 0; side < SideNames.Length; side++) {
            if (tokens[side] == null && sizes[side] == null && colors[side] == null) continue;
            OfficeXmlNode border = borders.GetOrAddChild(W, SideNames[side], BorderOrder);
            if (tokens[side] != null) border.SetAttribute(W, "val", tokens[side]);
            if (sizes[side] != null) border.SetAttribute(W, "sz", OfficeXmlViewValues.Format(sizes[side]));
            if (colors[side] != null) border.SetAttribute(W, "color", colors[side]!.Value.ToRgbHex());
        }
    }

    /// <summary>
    /// Gets border properties for the specified side
    /// </summary>
    /// <param name="side">The border side to get properties for</param>
    /// <returns>The direct border style, physical thickness and RGB color; absent or automatic values remain null.</returns>
    public (WordBorderStyle? Style, global::OfficeIMO.OfficeLength? Size, global::OfficeIMO.OfficeColor? Color) GetBorderProperties(WordTableBorderSide side) {
        if ((uint)side >= (uint)SideNames.Length) return (null, null, null);
        OfficeXmlNode border = Borders.Child(W, SideNames[(int)side]);
        return (ReadStyle(border), ReadBorderSize(OfficeXmlViewValues.ParseUInt32(border.Attribute(W, "sz"))), Helpers.ParseNativeColor(border.Attribute(W, "color")));
    }

    private static void WriteBorder(OfficeXmlNode border, string token, uint size, string color) {
        border.SetAttribute(W, "val", token);
        border.SetAttribute(W, "color", color);
        border.SetAttribute(W, "sz", OfficeXmlViewValues.Format(size));
    }

    private static WordBorderStyle? ReadStyle(OfficeXmlNode border) =>
        border.Attribute(W, "val") is string token && WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordBorderStyle style) ? style : null;

    private static uint ToBorderSize(global::OfficeIMO.OfficeLength value, string parameter) {
        decimal native = Math.Round(value.ToEmus() * 8m / 12700m, 0, MidpointRounding.AwayFromZero);
        if (value < global::OfficeIMO.OfficeLength.Zero || native > uint.MaxValue)
            throw new ArgumentOutOfRangeException(parameter, "Border thickness must fit Word's unsigned 32-bit eighth-point range.");
        return (uint)native;
    }

    private static global::OfficeIMO.OfficeLength? ReadBorderSize(uint? value) => value.HasValue
        ? global::OfficeIMO.OfficeLength.Points(value.Value / 8D) : (global::OfficeIMO.OfficeLength?)null;

    /// <summary>
    /// Apply the current table's border settings to all cells in the table
    /// </summary>
    public void ApplyBordersToAllCells() {
        OfficeXmlNode borders = Borders;
        if (borders.IsNone) {
            return;
        }

        // Read the table's borders once; every cell gets the attributes each side states.
        var sides = new List<(WordTableCellBorder.Edge Edge, WordBorderStyle? Style, global::OfficeIMO.OfficeLength? Size, string? Color)>();
        for (int side = 0; side < SideNames.Length; side++) {
            OfficeXmlNode border = borders.Child(W, SideNames[side]);
            if (border.IsNone) continue;
            sides.Add((SideEdges[side], ReadStyle(border), WordBorderLengths.FromSize(OfficeXmlViewValues.ParseUInt32(border.Attribute(W, "sz"))), border.Attribute(W, "color")));
        }

        foreach (var cell in _table.Cells) {
            WordTableCellBorder cellBorders = cell.Borders;
            foreach (var side in sides) {
                if (side.Style.HasValue) cellBorders.SetStyle(side.Edge, side.Style);
                if (side.Size.HasValue) cellBorders.SetSize(side.Edge, side.Size);
                if (side.Color != null) cellBorders.SetColorValue(side.Edge, side.Color);
            }
        }
    }
}
