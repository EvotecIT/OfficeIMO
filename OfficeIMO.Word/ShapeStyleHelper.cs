using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    /// <summary>
    /// Helper class for getting and setting the style of a V.Shape object.
    /// </summary>
    /// <example>
    /// This is a sample of how to use this class.
    /// <code>
    /// // Get the style
    /// var style = ShapeStyleHelper.GetStyle(shape1);
    ///
    /// // Modify the style
    /// style["position"] = "absolute";
    /// style["margin-left"] = "0";
    ///
    /// // Set the style
    /// ShapeStyleHelper.SetStyle(shape1, style);
    /// </code>
    /// </example>
    public static class ShapeStyleHelper {
        public static Dictionary<string, string> GetStyle(Shape shape) {
            return shape.Style.Value.Split(';')
                .Select(part => part.Split(':'))
                .ToDictionary(split => split[0], split => split[1]);
        }

        public static void SetStyle(Shape shape, Dictionary<string, string> style) {
            shape.Style.Value = string.Join(";", style.Select(kvp => $"{kvp.Key}:{kvp.Value}"));
        }
    }

    public class ShapeStyle {
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

        public static ShapeStyle FromString(string styleString) {
            var styleParts = styleString.Split(';')
                .Select(part => part.Split(':'))
                .ToDictionary(split => split[0], split => split[1]);

            var shapeStyle = new ShapeStyle();

            var properties = typeof(ShapeStyle).GetProperties();
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
            var properties = typeof(ShapeStyle).GetProperties();
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
}
