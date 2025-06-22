using System.Collections.Generic;
using System.Reflection;
using SixLabors.ImageSharp;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides utilities for resolving <see cref="Color"/> instances to their
    /// named equivalents.
    /// </summary>
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

        /// <summary>
        /// Returns the standard color name for the specified <see cref="Color"/>,
        /// or the RGBA value if no name exists.
        /// </summary>
        /// <param name="color">The color to resolve.</param>
        /// <returns>The color name or RGBA string.</returns>
        public static string GetColorName(Color color) {
            if (colorNameMap.TryGetValue(color, out string colorName)) {
                return colorName;
            }

            // If the color is not found in the map, return the RGBA value
            return color.ToString();
        }
    }
}
