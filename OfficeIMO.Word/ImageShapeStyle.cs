namespace OfficeIMO.Word;

/// <summary>
/// Represents style information for positioned images.
/// </summary>
public class ImageShapeStyle {
    /// <summary>Gets or sets the CSS position property.</summary>
    public string Position { get; set; } = string.Empty;
    /// <summary>Gets or sets the left margin.</summary>
    public string MarginLeft { get; set; } = string.Empty;
    /// <summary>Gets or sets the top margin.</summary>
    public string MarginTop { get; set; } = string.Empty;
    /// <summary>Gets or sets the width.</summary>
    public string Width { get; set; } = string.Empty;
    /// <summary>Gets or sets the height.</summary>
    public string Height { get; set; } = string.Empty;
    /// <summary>Gets or sets the Z-index value.</summary>
    public string ZIndex { get; set; } = string.Empty;
    /// <summary>Gets or sets the horizontal position mode.</summary>
    public string MsoPositionHorizontal { get; set; } = string.Empty;
    /// <summary>Gets or sets the horizontal position relative to.</summary>
    public string MsoPositionHorizontalRelative { get; set; } = string.Empty;
    /// <summary>Gets or sets the vertical position mode.</summary>
    public string MsoPositionVertical { get; set; } = string.Empty;
    /// <summary>Gets or sets the vertical position relative to.</summary>
    public string MsoPositionVerticalRelative { get; set; } = string.Empty;

    /// <summary>
    /// Parses a semicolon delimited style string into an <see cref="ImageShapeStyle"/> instance.
    /// </summary>
    /// <param name="styleString">The style string.</param>
    /// <returns>A populated <see cref="ImageShapeStyle"/>.</returns>
    public static ImageShapeStyle FromString(string styleString) {
        var styleParts = styleString.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Split(new[] { ':' }, 2))
            .Where(split => split.Length == 2)
            .ToDictionary(split => split[0].Trim().ToLowerInvariant(), split => split[1].Trim());

        var shapeStyle = new ImageShapeStyle();

        static string Key(string name) => name.ToLowerInvariant();
        string Require(string propertyName) {
            var key = Key(propertyName);
            if (!styleParts.TryGetValue(key, out var value)) {
                throw new ArgumentException($"Missing required style property: {propertyName}");
            }
            return value;
        }

        shapeStyle.Position = Require(nameof(Position));
        shapeStyle.MarginLeft = Require(nameof(MarginLeft));
        shapeStyle.MarginTop = Require(nameof(MarginTop));
        shapeStyle.Width = Require(nameof(Width));
        shapeStyle.Height = Require(nameof(Height));
        shapeStyle.ZIndex = Require(nameof(ZIndex));
        shapeStyle.MsoPositionHorizontal = Require(nameof(MsoPositionHorizontal));
        shapeStyle.MsoPositionHorizontalRelative = Require(nameof(MsoPositionHorizontalRelative));
        shapeStyle.MsoPositionVertical = Require(nameof(MsoPositionVertical));
        shapeStyle.MsoPositionVerticalRelative = Require(nameof(MsoPositionVerticalRelative));

        return shapeStyle;
    }

    /// <summary>
    /// Serializes the instance to a style string.
    /// </summary>
    /// <returns>A semicolon delimited style string.</returns>
    public override string ToString() {
        var styleParts = new List<string> {
            $"{nameof(Position).ToLowerInvariant()}:{Position}",
            $"{nameof(MarginLeft).ToLowerInvariant()}:{MarginLeft}",
            $"{nameof(MarginTop).ToLowerInvariant()}:{MarginTop}",
            $"{nameof(Width).ToLowerInvariant()}:{Width}",
            $"{nameof(Height).ToLowerInvariant()}:{Height}",
            $"{nameof(ZIndex).ToLowerInvariant()}:{ZIndex}",
            $"{nameof(MsoPositionHorizontal).ToLowerInvariant()}:{MsoPositionHorizontal}",
            $"{nameof(MsoPositionHorizontalRelative).ToLowerInvariant()}:{MsoPositionHorizontalRelative}",
            $"{nameof(MsoPositionVertical).ToLowerInvariant()}:{MsoPositionVertical}",
            $"{nameof(MsoPositionVerticalRelative).ToLowerInvariant()}:{MsoPositionVerticalRelative}"
        };

        return string.Join(";", styleParts);
    }
}

