using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Manages table style settings and borders.
/// </summary>
public partial class WordTableStyleDetails {
    /// <summary>
    /// Creates and sets a TableBorders object with the specified border settings
    /// </summary>
    /// <param name="style">Border style for all sides</param>
    /// <param name="size">Border size for all sides</param>
    /// <param name="color">Border color for all sides</param>
    public void SetBordersForAllSides(BorderValues style, UInt32Value size, SixLabors.ImageSharp.Color color) {
        _table.CheckTableProperties();

        string colorHex = color.ToHexColor();

        TableBorders borders = new TableBorders(
            new TopBorder() { Val = style, Size = size, Color = colorHex },
            new BottomBorder() { Val = style, Size = size, Color = colorHex },
            new LeftBorder() { Val = style, Size = size, Color = colorHex },
            new RightBorder() { Val = style, Size = size, Color = colorHex },
            new InsideHorizontalBorder() { Val = style, Size = size, Color = colorHex },
            new InsideVerticalBorder() { Val = style, Size = size, Color = colorHex }
        );

        TableBorders = borders;
    }

    /// <summary>
    /// Creates a TableBorders object with different settings for outside and inside borders
    /// </summary>
    /// <param name="outsideStyle">Style for outside borders</param>
    /// <param name="outsideSize">Size for outside borders</param>
    /// <param name="outsideColor">Color for outside borders</param>
    /// <param name="insideStyle">Style for inside borders</param>
    /// <param name="insideSize">Size for inside borders</param>
    /// <param name="insideColor">Color for inside borders</param>
    public void SetBordersOutsideInside(
        BorderValues outsideStyle, UInt32Value outsideSize, SixLabors.ImageSharp.Color outsideColor,
        BorderValues insideStyle, UInt32Value insideSize, SixLabors.ImageSharp.Color insideColor) {
        _table.CheckTableProperties();

        string outsideColorHex = outsideColor.ToHexColor();
        string insideColorHex = insideColor.ToHexColor();

        TableBorders borders = new TableBorders(
            new TopBorder() { Val = outsideStyle, Size = outsideSize, Color = outsideColorHex },
            new BottomBorder() { Val = outsideStyle, Size = outsideSize, Color = outsideColorHex },
            new LeftBorder() { Val = outsideStyle, Size = outsideSize, Color = outsideColorHex },
            new RightBorder() { Val = outsideStyle, Size = outsideSize, Color = outsideColorHex },
            new InsideHorizontalBorder() { Val = insideStyle, Size = insideSize, Color = insideColorHex },
            new InsideVerticalBorder() { Val = insideStyle, Size = insideSize, Color = insideColorHex }
        );

        TableBorders = borders;
    }

    /// <summary>
    /// Creates a TableBorders object with custom settings for each side
    /// </summary>
    public void SetCustomBorders(
        BorderValues? topStyle = null, UInt32Value? topSize = null, SixLabors.ImageSharp.Color? topColor = null,
        BorderValues? bottomStyle = null, UInt32Value? bottomSize = null, SixLabors.ImageSharp.Color? bottomColor = null,
        BorderValues? leftStyle = null, UInt32Value? leftSize = null, SixLabors.ImageSharp.Color? leftColor = null,
        BorderValues? rightStyle = null, UInt32Value? rightSize = null, SixLabors.ImageSharp.Color? rightColor = null,
        BorderValues? insideHStyle = null, UInt32Value? insideHSize = null, SixLabors.ImageSharp.Color? insideHColor = null,
        BorderValues? insideVStyle = null, UInt32Value? insideVSize = null, SixLabors.ImageSharp.Color? insideVColor = null) {
        _table.CheckTableProperties();

        // Get existing borders or create new
        TableBorders borders = TableBorders ?? new TableBorders();

        // Update top border if any parameter is set
        if (topStyle != null || topSize != null || topColor != null) {
            var topBorder = borders.TopBorder ?? new TopBorder();
            if (topStyle != null) topBorder.Val = topStyle;
            if (topSize != null) topBorder.Size = topSize;
            if (topColor != null) {
                topBorder.Color = topColor.Value.ToHexColor();
            }
            borders.TopBorder = topBorder;
        }

        // Update bottom border if any parameter is set
        if (bottomStyle != null || bottomSize != null || bottomColor != null) {
            var bottomBorder = borders.BottomBorder ?? new BottomBorder();
            if (bottomStyle != null) bottomBorder.Val = bottomStyle;
            if (bottomSize != null) bottomBorder.Size = bottomSize;
            if (bottomColor != null) {
                bottomBorder.Color = bottomColor.Value.ToHexColor();
            }
            borders.BottomBorder = bottomBorder;
        }

        // Update left border if any parameter is set
        if (leftStyle != null || leftSize != null || leftColor != null) {
            var leftBorder = borders.LeftBorder ?? new LeftBorder();
            if (leftStyle != null) leftBorder.Val = leftStyle;
            if (leftSize != null) leftBorder.Size = leftSize;
            if (leftColor != null) {
                leftBorder.Color = leftColor.Value.ToHexColor();
            }
            borders.LeftBorder = leftBorder;
        }

        // Update right border if any parameter is set
        if (rightStyle != null || rightSize != null || rightColor != null) {
            var rightBorder = borders.RightBorder ?? new RightBorder();
            if (rightStyle != null) rightBorder.Val = rightStyle;
            if (rightSize != null) rightBorder.Size = rightSize;
            if (rightColor != null) {
                rightBorder.Color = rightColor.Value.ToHexColor();
            }
            borders.RightBorder = rightBorder;
        }

        // Update inside horizontal border if any parameter is set
        if (insideHStyle != null || insideHSize != null || insideHColor != null) {
            var insideHBorder = borders.InsideHorizontalBorder ?? new InsideHorizontalBorder();
            if (insideHStyle != null) insideHBorder.Val = insideHStyle;
            if (insideHSize != null) insideHBorder.Size = insideHSize;
            if (insideHColor != null) {
                insideHBorder.Color = insideHColor.Value.ToHexColor();
            }
            borders.InsideHorizontalBorder = insideHBorder;
        }

        // Update inside vertical border if any parameter is set
        if (insideVStyle != null || insideVSize != null || insideVColor != null) {
            var insideVBorder = borders.InsideVerticalBorder ?? new InsideVerticalBorder();
            if (insideVStyle != null) insideVBorder.Val = insideVStyle;
            if (insideVSize != null) insideVBorder.Size = insideVSize;
            if (insideVColor != null) {
                insideVBorder.Color = insideVColor.Value.ToHexColor();
            }
            borders.InsideVerticalBorder = insideVBorder;
        }

        TableBorders = borders;
    }

    /// <summary>
    /// Gets border properties for the specified side
    /// </summary>
    /// <param name="side">The border side to get properties for</param>
    /// <returns>A tuple with style, size, and color</returns>
    public (BorderValues? Style, UInt32Value? Size, string? ColorHex) GetBorderProperties(WordTableBorderSide side) {
        if (TableBorders == null) {
            return (null, null, null);
        }

        switch (side) {
            case WordTableBorderSide.Top:
                return (
                    TableBorders.TopBorder?.Val?.Value,
                    TableBorders.TopBorder?.Size,
                    TableBorders.TopBorder?.Color?.Value
                );
            case WordTableBorderSide.Bottom:
                return (
                    TableBorders.BottomBorder?.Val?.Value,
                    TableBorders.BottomBorder?.Size,
                    TableBorders.BottomBorder?.Color?.Value
                );
            case WordTableBorderSide.Left:
                return (
                    TableBorders.LeftBorder?.Val?.Value,
                    TableBorders.LeftBorder?.Size,
                    TableBorders.LeftBorder?.Color?.Value
                );
            case WordTableBorderSide.Right:
                return (
                    TableBorders.RightBorder?.Val?.Value,
                    TableBorders.RightBorder?.Size,
                    TableBorders.RightBorder?.Color?.Value
                );
            case WordTableBorderSide.InsideHorizontal:
                return (
                    TableBorders.InsideHorizontalBorder?.Val?.Value,
                    TableBorders.InsideHorizontalBorder?.Size,
                    TableBorders.InsideHorizontalBorder?.Color?.Value
                );
            case WordTableBorderSide.InsideVertical:
                return (
                    TableBorders.InsideVerticalBorder?.Val?.Value,
                    TableBorders.InsideVerticalBorder?.Size,
                    TableBorders.InsideVerticalBorder?.Color?.Value
                );
            default:
                return (null, null, null);
        }
    }

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
                    cell.Borders.TopStyle = TableBorders.TopBorder.Val.Value;
                }
                if (TableBorders.TopBorder.Size != null)
                    cell.Borders.TopSize = TableBorders.TopBorder.Size;
                if (TableBorders.TopBorder.Color != null)
                    cell.Borders.TopColorHex = TableBorders.TopBorder.Color.Value;
            }

            // Bottom border
            if (TableBorders.BottomBorder != null) {
                if (TableBorders.BottomBorder.Val != null) {
                    cell.Borders.BottomStyle = TableBorders.BottomBorder.Val.Value;
                }
                if (TableBorders.BottomBorder.Size != null)
                    cell.Borders.BottomSize = TableBorders.BottomBorder.Size;
                if (TableBorders.BottomBorder.Color != null)
                    cell.Borders.BottomColorHex = TableBorders.BottomBorder.Color.Value;
            }

            // Left border
            if (TableBorders.LeftBorder != null) {
                if (TableBorders.LeftBorder.Val != null) {
                    cell.Borders.LeftStyle = TableBorders.LeftBorder.Val.Value;
                }
                if (TableBorders.LeftBorder.Size != null)
                    cell.Borders.LeftSize = TableBorders.LeftBorder.Size;
                if (TableBorders.LeftBorder.Color != null)
                    cell.Borders.LeftColorHex = TableBorders.LeftBorder.Color.Value;
            }

            // Right border
            if (TableBorders.RightBorder != null) {
                if (TableBorders.RightBorder.Val != null) {
                    cell.Borders.RightStyle = TableBorders.RightBorder.Val.Value;
                }
                if (TableBorders.RightBorder.Size != null)
                    cell.Borders.RightSize = TableBorders.RightBorder.Size;
                if (TableBorders.RightBorder.Color != null)
                    cell.Borders.RightColorHex = TableBorders.RightBorder.Color.Value;
            }

            // Inside horizontal border
            if (TableBorders.InsideHorizontalBorder != null) {
                if (TableBorders.InsideHorizontalBorder.Val != null) {
                    cell.Borders.InsideHorizontalStyle = TableBorders.InsideHorizontalBorder.Val.Value;
                }
                if (TableBorders.InsideHorizontalBorder.Size != null)
                    cell.Borders.InsideHorizontalSize = TableBorders.InsideHorizontalBorder.Size;
                if (TableBorders.InsideHorizontalBorder.Color != null)
                    cell.Borders.InsideHorizontalColorHex = TableBorders.InsideHorizontalBorder.Color.Value;
            }

            // Inside vertical border
            if (TableBorders.InsideVerticalBorder != null) {
                if (TableBorders.InsideVerticalBorder.Val != null) {
                    cell.Borders.InsideVerticalStyle = TableBorders.InsideVerticalBorder.Val.Value;
                }
                if (TableBorders.InsideVerticalBorder.Size != null)
                    cell.Borders.InsideVerticalSize = TableBorders.InsideVerticalBorder.Size;
                if (TableBorders.InsideVerticalBorder.Color != null)
                    cell.Borders.InsideVerticalColorHex = TableBorders.InsideVerticalBorder.Color.Value;
            }
        }
    }
}
