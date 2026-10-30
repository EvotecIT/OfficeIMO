using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

// Manages table style settings and borders.
public partial class WordTableStyleDetails {
    /// <summary>
    /// Creates and sets a TableBorders object with the specified border settings
    /// </summary>
    /// <param name="style">Border style for all sides</param>
    /// <param name="size">Nonnegative thickness, rounded to eighths of a point within the native unsigned 32-bit range.</param>
    /// <param name="color">Border color for all sides</param>
    public void SetBordersForAllSides(WordBorderStyle style, global::OfficeIMO.OfficeLength size, OfficeIMO.OfficeColor color) {
        uint nativeSize = ToBorderSize(size, nameof(size));
        string colorHex = color.ToRgbHex();

        TableBorders borders = new TableBorders(
            new TopBorder() { Val = style.ToOpenXml(), Size = nativeSize, Color = colorHex },
            new LeftBorder() { Val = style.ToOpenXml(), Size = nativeSize, Color = colorHex },
            new BottomBorder() { Val = style.ToOpenXml(), Size = nativeSize, Color = colorHex },
            new RightBorder() { Val = style.ToOpenXml(), Size = nativeSize, Color = colorHex },
            new InsideHorizontalBorder() { Val = style.ToOpenXml(), Size = nativeSize, Color = colorHex },
            new InsideVerticalBorder() { Val = style.ToOpenXml(), Size = nativeSize, Color = colorHex }
        );

        TableBorders = borders;
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
        string outsideColorHex = outsideColor.ToRgbHex();
        string insideColorHex = insideColor.ToRgbHex();

        TableBorders borders = new TableBorders(
            new TopBorder() { Val = outsideStyle.ToOpenXml(), Size = nativeOutsideSize, Color = outsideColorHex },
            new LeftBorder() { Val = outsideStyle.ToOpenXml(), Size = nativeOutsideSize, Color = outsideColorHex },
            new BottomBorder() { Val = outsideStyle.ToOpenXml(), Size = nativeOutsideSize, Color = outsideColorHex },
            new RightBorder() { Val = outsideStyle.ToOpenXml(), Size = nativeOutsideSize, Color = outsideColorHex },
            new InsideHorizontalBorder() { Val = insideStyle.ToOpenXml(), Size = nativeInsideSize, Color = insideColorHex },
            new InsideVerticalBorder() { Val = insideStyle.ToOpenXml(), Size = nativeInsideSize, Color = insideColorHex }
        );

        TableBorders = borders;
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
        uint? nativeTopSize = topSize.HasValue ? ToBorderSize(topSize.Value, nameof(topSize)) : (uint?)null;
        uint? nativeBottomSize = bottomSize.HasValue ? ToBorderSize(bottomSize.Value, nameof(bottomSize)) : (uint?)null;
        uint? nativeLeftSize = leftSize.HasValue ? ToBorderSize(leftSize.Value, nameof(leftSize)) : (uint?)null;
        uint? nativeRightSize = rightSize.HasValue ? ToBorderSize(rightSize.Value, nameof(rightSize)) : (uint?)null;
        uint? nativeInsideHSize = insideHSize.HasValue ? ToBorderSize(insideHSize.Value, nameof(insideHSize)) : (uint?)null;
        uint? nativeInsideVSize = insideVSize.HasValue ? ToBorderSize(insideVSize.Value, nameof(insideVSize)) : (uint?)null;

        // Get existing borders or create new
        TableBorders borders = TableBorders ?? new TableBorders();

        // Update top border if any parameter is set
        if (topStyle != null || topSize != null || topColor != null) {
            var topBorder = borders.TopBorder ?? new TopBorder();
            if (topStyle != null) topBorder.Val = topStyle.Value.ToOpenXml();
            if (topSize != null) topBorder.Size = nativeTopSize;
            if (topColor != null) {
                topBorder.Color = topColor.Value.ToRgbHex();
            }
            borders.TopBorder = topBorder;
        }

        // Update bottom border if any parameter is set
        if (bottomStyle != null || bottomSize != null || bottomColor != null) {
            var bottomBorder = borders.BottomBorder ?? new BottomBorder();
            if (bottomStyle != null) bottomBorder.Val = bottomStyle.Value.ToOpenXml();
            if (bottomSize != null) bottomBorder.Size = nativeBottomSize;
            if (bottomColor != null) {
                bottomBorder.Color = bottomColor.Value.ToRgbHex();
            }
            borders.BottomBorder = bottomBorder;
        }

        // Update left border if any parameter is set
        if (leftStyle != null || leftSize != null || leftColor != null) {
            var leftBorder = borders.LeftBorder ?? new LeftBorder();
            if (leftStyle != null) leftBorder.Val = leftStyle.Value.ToOpenXml();
            if (leftSize != null) leftBorder.Size = nativeLeftSize;
            if (leftColor != null) {
                leftBorder.Color = leftColor.Value.ToRgbHex();
            }
            borders.LeftBorder = leftBorder;
        }

        // Update right border if any parameter is set
        if (rightStyle != null || rightSize != null || rightColor != null) {
            var rightBorder = borders.RightBorder ?? new RightBorder();
            if (rightStyle != null) rightBorder.Val = rightStyle.Value.ToOpenXml();
            if (rightSize != null) rightBorder.Size = nativeRightSize;
            if (rightColor != null) {
                rightBorder.Color = rightColor.Value.ToRgbHex();
            }
            borders.RightBorder = rightBorder;
        }

        // Update inside horizontal border if any parameter is set
        if (insideHStyle != null || insideHSize != null || insideHColor != null) {
            var insideHBorder = borders.InsideHorizontalBorder ?? new InsideHorizontalBorder();
            if (insideHStyle != null) insideHBorder.Val = insideHStyle.Value.ToOpenXml();
            if (insideHSize != null) insideHBorder.Size = nativeInsideHSize;
            if (insideHColor != null) {
                insideHBorder.Color = insideHColor.Value.ToRgbHex();
            }
            borders.InsideHorizontalBorder = insideHBorder;
        }

        // Update inside vertical border if any parameter is set
        if (insideVStyle != null || insideVSize != null || insideVColor != null) {
            var insideVBorder = borders.InsideVerticalBorder ?? new InsideVerticalBorder();
            if (insideVStyle != null) insideVBorder.Val = insideVStyle.Value.ToOpenXml();
            if (insideVSize != null) insideVBorder.Size = nativeInsideVSize;
            if (insideVColor != null) {
                insideVBorder.Color = insideVColor.Value.ToRgbHex();
            }
            borders.InsideVerticalBorder = insideVBorder;
        }

        TableBorders = borders;
    }

    /// <summary>
    /// Gets border properties for the specified side
    /// </summary>
    /// <param name="side">The border side to get properties for</param>
    /// <returns>The direct border style, physical thickness and RGB hex color; absent values remain null.</returns>
    public (WordBorderStyle? Style, global::OfficeIMO.OfficeLength? Size, string? ColorHex) GetBorderProperties(WordTableBorderSide side) {
        if (TableBorders == null) {
            return (null, null, null);
        }

        switch (side) {
            case WordTableBorderSide.Top:
                return (
                    TableBorders.TopBorder?.Val?.Value.ToOfficeEnum(),
                    ReadBorderSize(TableBorders.TopBorder?.Size?.Value),
                    TableBorders.TopBorder?.Color?.Value
                );
            case WordTableBorderSide.Bottom:
                return (
                    TableBorders.BottomBorder?.Val?.Value.ToOfficeEnum(),
                    ReadBorderSize(TableBorders.BottomBorder?.Size?.Value),
                    TableBorders.BottomBorder?.Color?.Value
                );
            case WordTableBorderSide.Left:
                return (
                    TableBorders.LeftBorder?.Val?.Value.ToOfficeEnum(),
                    ReadBorderSize(TableBorders.LeftBorder?.Size?.Value),
                    TableBorders.LeftBorder?.Color?.Value
                );
            case WordTableBorderSide.Right:
                return (
                    TableBorders.RightBorder?.Val?.Value.ToOfficeEnum(),
                    ReadBorderSize(TableBorders.RightBorder?.Size?.Value),
                    TableBorders.RightBorder?.Color?.Value
                );
            case WordTableBorderSide.InsideHorizontal:
                return (
                    TableBorders.InsideHorizontalBorder?.Val?.Value.ToOfficeEnum(),
                    ReadBorderSize(TableBorders.InsideHorizontalBorder?.Size?.Value),
                    TableBorders.InsideHorizontalBorder?.Color?.Value
                );
            case WordTableBorderSide.InsideVertical:
                return (
                    TableBorders.InsideVerticalBorder?.Val?.Value.ToOfficeEnum(),
                    ReadBorderSize(TableBorders.InsideVerticalBorder?.Size?.Value),
                    TableBorders.InsideVerticalBorder?.Color?.Value
                );
            default:
                return (null, null, null);
        }
    }

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
        if (TableBorders == null) {
            return;
        }

        foreach (var cell in _table.Cells) {
            // Top border
            if (TableBorders.TopBorder != null) {
                if (TableBorders.TopBorder.Val != null) {
                    cell.Borders.TopStyle = TableBorders.TopBorder.Val!.Value.ToOfficeEnum();
                }
                if (TableBorders.TopBorder.Size != null)
                    cell.Borders.TopSize = WordBorderLengths.FromSize(TableBorders.TopBorder.Size.Value);
                if (TableBorders.TopBorder.Color != null)
                    cell.Borders.TopColorHex = TableBorders.TopBorder.Color!;
            }

            // Bottom border
            if (TableBorders.BottomBorder != null) {
                if (TableBorders.BottomBorder.Val != null) {
                    cell.Borders.BottomStyle = TableBorders.BottomBorder.Val!.Value.ToOfficeEnum();
                }
                if (TableBorders.BottomBorder.Size != null)
                    cell.Borders.BottomSize = WordBorderLengths.FromSize(TableBorders.BottomBorder.Size.Value);
                if (TableBorders.BottomBorder.Color != null)
                    cell.Borders.BottomColorHex = TableBorders.BottomBorder.Color!;
            }

            // Left border
            if (TableBorders.LeftBorder != null) {
                if (TableBorders.LeftBorder.Val != null) {
                    cell.Borders.LeftStyle = TableBorders.LeftBorder.Val!.Value.ToOfficeEnum();
                }
                if (TableBorders.LeftBorder.Size != null)
                    cell.Borders.LeftSize = WordBorderLengths.FromSize(TableBorders.LeftBorder.Size.Value);
                if (TableBorders.LeftBorder.Color != null)
                    cell.Borders.LeftColorHex = TableBorders.LeftBorder.Color!;
            }

            // Right border
            if (TableBorders.RightBorder != null) {
                if (TableBorders.RightBorder.Val != null) {
                    cell.Borders.RightStyle = TableBorders.RightBorder.Val!.Value.ToOfficeEnum();
                }
                if (TableBorders.RightBorder.Size != null)
                    cell.Borders.RightSize = WordBorderLengths.FromSize(TableBorders.RightBorder.Size.Value);
                if (TableBorders.RightBorder.Color != null)
                    cell.Borders.RightColorHex = TableBorders.RightBorder.Color!;
            }

            // Inside horizontal border
            if (TableBorders.InsideHorizontalBorder != null) {
                if (TableBorders.InsideHorizontalBorder.Val != null) {
                    cell.Borders.InsideHorizontalStyle = TableBorders.InsideHorizontalBorder.Val!.Value.ToOfficeEnum();
                }
                if (TableBorders.InsideHorizontalBorder.Size != null)
                    cell.Borders.InsideHorizontalSize = WordBorderLengths.FromSize(TableBorders.InsideHorizontalBorder.Size.Value);
                if (TableBorders.InsideHorizontalBorder.Color != null)
                    cell.Borders.InsideHorizontalColorHex = TableBorders.InsideHorizontalBorder.Color!;
            }

            // Inside vertical border
            if (TableBorders.InsideVerticalBorder != null) {
                if (TableBorders.InsideVerticalBorder.Val != null) {
                    cell.Borders.InsideVerticalStyle = TableBorders.InsideVerticalBorder.Val!.Value.ToOfficeEnum();
                }
                if (TableBorders.InsideVerticalBorder.Size != null)
                    cell.Borders.InsideVerticalSize = WordBorderLengths.FromSize(TableBorders.InsideVerticalBorder.Size.Value);
                if (TableBorders.InsideVerticalBorder.Color != null)
                    cell.Borders.InsideVerticalColorHex = TableBorders.InsideVerticalBorder.Color!;
            }
        }
    }
}
