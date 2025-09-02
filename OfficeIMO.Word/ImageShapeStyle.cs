using System;
using System.Collections.Generic;
using System.Linq;

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
        var styleParts = styleString.Split(';')
            .Select(part => part.Split(':'))
            .ToDictionary(split => split[0], split => split[1]);

        var shapeStyle = new ImageShapeStyle();

        var properties = typeof(ImageShapeStyle).GetProperties();
        foreach (var property in properties) {
            if (styleParts.ContainsKey(property.Name.ToLower())) {
                property.SetValue(shapeStyle, styleParts[property.Name.ToLower()]);
            } else {
                throw new ArgumentException($"Missing required style property: {property.Name}");
            }
        }

        return shapeStyle;
    }

    /// <summary>
    /// Serializes the instance to a style string.
    /// </summary>
    /// <returns>A semicolon delimited style string.</returns>
    public override string ToString() {
        var properties = typeof(ImageShapeStyle).GetProperties();
        var styleParts = new List<string>();

        foreach (var property in properties) {
            var value = property.GetValue(this) as string;
            if (value != null) {
                styleParts.Add($"{property.Name.ToLower()}:{value}");
            }
        }

        return string.Join(";", styleParts);
    }
}

