using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Word;

public class ImageShapeStyle {
    public string Position { get; set; }
    public string MarginLeft { get; set; }
    public string MarginTop { get; set; }
    public string Width { get; set; }
    public string Height { get; set; }
    public string ZIndex { get; set; }
    public string MsoPositionHorizontal { get; set; }
    public string MsoPositionHorizontalRelative { get; set; }
    public string MsoPositionVertical { get; set; }
    public string MsoPositionVerticalRelative { get; set; }

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
