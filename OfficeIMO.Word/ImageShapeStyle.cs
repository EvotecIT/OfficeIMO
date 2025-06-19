using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Word;

public class ImageShapeStyle {
    /// <summary>
    /// Gets or sets the Position.
    /// </summary>
    public string Position { get; set; }
    /// <summary>
    /// Gets or sets the MarginLeft.
    /// </summary>
    public string MarginLeft { get; set; }
    /// <summary>
    /// Gets or sets the MarginTop.
    /// </summary>
    public string MarginTop { get; set; }
    /// <summary>
    /// Gets or sets the Width.
    /// </summary>
    public string Width { get; set; }
    /// <summary>
    /// Gets or sets the Height.
    /// </summary>
    public string Height { get; set; }
    /// <summary>
    /// Gets or sets the ZIndex.
    /// </summary>
    public string ZIndex { get; set; }
    /// <summary>
    /// Gets or sets the MsoPositionHorizontal.
    /// </summary>
    public string MsoPositionHorizontal { get; set; }
    /// <summary>
    /// Gets or sets the MsoPositionHorizontalRelative.
    /// </summary>
    public string MsoPositionHorizontalRelative { get; set; }
    /// <summary>
    /// Gets or sets the MsoPositionVertical.
    /// </summary>
    public string MsoPositionVertical { get; set; }
    /// <summary>
    /// Gets or sets the MsoPositionVerticalRelative.
    /// </summary>
    public string MsoPositionVerticalRelative { get; set; }

    /// <summary>
    /// Parses a semicolon-separated list of property assignments and produces a
    /// new <see cref="ImageShapeStyle"/> instance initialized with those values.
    /// </summary>
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
    /// Converts the current style settings into a semicolon-separated property
    /// string that can be persisted or parsed back later.
    /// </summary>
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
