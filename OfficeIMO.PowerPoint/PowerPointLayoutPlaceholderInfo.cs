using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a placeholder defined by a slide layout.
    /// </summary>
    public readonly struct PowerPointLayoutPlaceholderInfo {
        /// <summary>
        ///     Creates a layout placeholder info entry.
        /// </summary>
        public PowerPointLayoutPlaceholderInfo(string name, PlaceholderValues? placeholderType, uint? placeholderIndex, PowerPointLayoutBox? bounds) {
            Name = name ?? string.Empty;
            PlaceholderType = placeholderType;
            PlaceholderIndex = placeholderIndex;
            Bounds = bounds;
        }

        /// <summary>
        ///     Placeholder name from the layout.
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///     Placeholder type (may be null for custom placeholders).
        /// </summary>
        public PlaceholderValues? PlaceholderType { get; }

        /// <summary>
        ///     Placeholder index (may be null).
        /// </summary>
        public uint? PlaceholderIndex { get; }

        /// <summary>
        ///     Placeholder bounds if present in the layout.
        /// </summary>
        public PowerPointLayoutBox? Bounds { get; }

        /// <summary>
        ///     Returns a display-friendly string.
        /// </summary>
        public override string ToString() {
            string typeText = PlaceholderType?.ToString() ?? "Placeholder";
            if (PlaceholderIndex != null) {
                typeText = $"{typeText} {PlaceholderIndex.Value}";
            }

            if (!string.IsNullOrWhiteSpace(Name)) {
                return $"{Name} ({typeText})";
            }

            return typeText;
        }
    }
}
