using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a slide layout entry in a presentation.
    /// </summary>
    public readonly struct PowerPointSlideLayoutInfo {
        /// <summary>
        ///     Creates a layout info entry.
        /// </summary>
        public PowerPointSlideLayoutInfo(int masterIndex, int layoutIndex, string name, SlideLayoutValues? type, string? relationshipId) {
            MasterIndex = masterIndex;
            LayoutIndex = layoutIndex;
            Name = name;
            Type = type;
            RelationshipId = relationshipId;
        }

        /// <summary>
        ///     Index of the slide master that owns the layout.
        /// </summary>
        public int MasterIndex { get; }

        /// <summary>
        ///     Index of the layout within its master.
        /// </summary>
        public int LayoutIndex { get; }

        /// <summary>
        ///     Layout display name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///     Layout type (may be null for custom layouts).
        /// </summary>
        public SlideLayoutValues? Type { get; }

        /// <summary>
        ///     Relationship ID for the layout part.
        /// </summary>
        public string? RelationshipId { get; }

        /// <summary>
        ///     Returns the display name or type.
        /// </summary>
        public override string ToString() {
            if (!string.IsNullOrWhiteSpace(Name)) {
                return Name;
            }
            return Type?.ToString() ?? $"Layout {LayoutIndex}";
        }
    }
}
