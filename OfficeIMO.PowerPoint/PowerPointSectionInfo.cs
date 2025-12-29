using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a slide section within a presentation.
    /// </summary>
    public readonly struct PowerPointSectionInfo {
        /// <summary>
        ///     Creates a section info entry.
        /// </summary>
        public PowerPointSectionInfo(string name, string id, IReadOnlyList<int> slideIndices) {
            Name = name ?? string.Empty;
            Id = id ?? string.Empty;
            SlideIndices = slideIndices ?? new List<int>();
        }

        /// <summary>
        ///     Section display name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///     Section identifier (GUID string).
        /// </summary>
        public string Id { get; }

        /// <summary>
        ///     Zero-based slide indices included in the section.
        /// </summary>
        public IReadOnlyList<int> SlideIndices { get; }

        /// <summary>
        ///     Returns a display-friendly string.
        /// </summary>
        public override string ToString() {
            return $"{Name} ({SlideIndices.Count} slides)";
        }
    }
}
