using System.Collections.Generic;
using System.Reflection;
using SixLabors.ImageSharp;

namespace OfficeIMO.Word {
    public class ColorNameResolver {
        private static readonly Dictionary<Color, string> colorNameMap = new Dictionary<Color, string>();

        static ColorNameResolver() {
            // Use reflection to get all named colors from the Color structure
            foreach (var prop in typeof(Color).GetFields(BindingFlags.Static | BindingFlags.Public)) {
                if (prop.FieldType == typeof(Color)) {
                    Color color = (Color)prop.GetValue(null);
                    colorNameMap[color] = prop.Name;
                }
            }
        }

        public static string GetColorName(Color color) {
            if (colorNameMap.TryGetValue(color, out string colorName)) {
                return colorName;
            }

            // If the color is not found in the map, return the RGBA value
            return color.ToString();
        }
    }
}
