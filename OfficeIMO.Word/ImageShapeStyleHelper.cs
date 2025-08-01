using System.Collections.Generic;
using System.Linq;
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
    public static class ImageShapeStyleHelper {
        /// <summary>
        /// Retrieves the style attributes of a <see cref="Shape"/> as a dictionary.
        /// </summary>
        /// <param name="shape">The VML <see cref="Shape"/> whose style should be parsed.</param>
        /// <returns>A dictionary containing style names and values.</returns>
        public static Dictionary<string, string> GetStyle(Shape shape) {
            return shape.Style.Value.Split(';')
                .Select(part => part.Split(':'))
                .ToDictionary(split => split[0], split => split[1]);
        }

        /// <summary>
        /// Applies the provided style dictionary to the given <see cref="Shape"/>.
        /// </summary>
        /// <param name="shape">The VML <see cref="Shape"/> to update.</param>
        /// <param name="style">The style dictionary to serialize and assign.</param>
        public static void SetStyle(Shape shape, Dictionary<string, string> style) {
            shape.Style.Value = string.Join(";", style.Select(kvp => $"{kvp.Key}:{kvp.Value}"));
        }
    }
}
