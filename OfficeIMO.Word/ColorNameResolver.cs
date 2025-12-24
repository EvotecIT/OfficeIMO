using SixLabors.ImageSharp;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides utilities for resolving <see cref="Color"/> instances to their
    /// named equivalents.
    /// </summary>
    public class ColorNameResolver {
        private static readonly Dictionary<Color, string> colorNameMap = new Dictionary<Color, string>();
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicFields | DynamicallyAccessedMemberTypes.PublicProperties)]
        private static readonly Type ColorType = typeof(Color);

        static ColorNameResolver() {
            // Use reflection to get all named colors from the Color structure
            foreach (var field in ColorType.GetFields(BindingFlags.Static | BindingFlags.Public)) {
                if (field.FieldType == typeof(Color) && field.GetValue(null) is Color color) {
                    colorNameMap[color] = field.Name;
                }
            }
            foreach (var prop in ColorType.GetProperties(BindingFlags.Static | BindingFlags.Public)) {
                if (prop.PropertyType == typeof(Color) && prop.GetValue(null) is Color color) {
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
            if (colorNameMap.TryGetValue(color, out var colorName)) {
                return colorName;
            }

            // If the color is not found in the map, return the RGBA value
            return color.ToString();
        }
    }
}
