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
            if (shape == null) throw new ArgumentNullException(nameof(shape));

            string? styleValue = shape.Style?.Value;
            if (string.IsNullOrWhiteSpace(styleValue)) {
                return new Dictionary<string, string>();
            }

            // Compatible with net472/netstandard2.0: avoid newer Split overloads and TrimEntries flag
            var pairs = new List<KeyValuePair<string, string>>();
            var segments = styleValue!.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var segment in segments) {
                var parts = segment.Split(new[] { ':' }, 2);
                if (parts.Length < 2) continue;
                var key = parts[0].Trim();
                if (key.Length == 0) continue;
                var value = parts[1].Trim();
                pairs.Add(new KeyValuePair<string, string>(key, value));
            }
            return pairs.ToDictionary(kv => kv.Key, kv => kv.Value);
        }

        /// <summary>
        /// Applies the provided style dictionary to the given <see cref="Shape"/>.
        /// </summary>
        /// <param name="shape">The VML <see cref="Shape"/> to update.</param>
        /// <param name="style">The style dictionary to serialize and assign.</param>
        public static void SetStyle(Shape shape, Dictionary<string, string> style) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (style == null) throw new ArgumentNullException(nameof(style));

            if (style.Count == 0) {
                shape.Style = null;
                return;
            }

            if (shape.Style == null) shape.Style = new StringValue();
            shape.Style.Value = string.Join(";", style.Select(kvp => $"{kvp.Key}:{kvp.Value}"));
        }
    }
}
