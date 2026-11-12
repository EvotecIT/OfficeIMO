using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint {
    /// <summary>Inspectable accessibility metadata for one slide shape.</summary>
    public sealed class PowerPointAccessibilityShapeInfo {
        internal PowerPointAccessibilityShapeInfo(int readingOrder, uint? shapeId, string? name, string? title,
            string? description, bool decorative, string? language, PowerPointShapeContentType contentType) {
            ReadingOrder = readingOrder;
            ShapeId = shapeId;
            Name = name;
            Title = title;
            Description = description;
            Decorative = decorative;
            Language = language;
            ContentType = contentType;
        }

        /// <summary>Zero-based shape-tree reading order.</summary>
        public int ReadingOrder { get; }
        /// <summary>OOXML shape identifier.</summary>
        public uint? ShapeId { get; }
        /// <summary>Authored shape name.</summary>
        public string? Name { get; }
        /// <summary>Accessibility title.</summary>
        public string? Title { get; }
        /// <summary>Accessibility description.</summary>
        public string? Description { get; }
        /// <summary>Whether the shape is decorative.</summary>
        public bool Decorative { get; }
        /// <summary>First explicit text language.</summary>
        public string? Language { get; }
        /// <summary>Detected shape content type.</summary>
        public PowerPointShapeContentType ContentType { get; }
    }

    /// <summary>Inspectable accessibility metadata for one slide.</summary>
    public sealed class PowerPointAccessibilitySlideInfo {
        internal PowerPointAccessibilitySlideInfo(int slideIndex, string? title,
            IList<PowerPointAccessibilityShapeInfo> shapes) {
            SlideIndex = slideIndex;
            Title = title;
            Shapes = new ReadOnlyCollection<PowerPointAccessibilityShapeInfo>(
                new List<PowerPointAccessibilityShapeInfo>(shapes));
        }

        /// <summary>Zero-based slide index.</summary>
        public int SlideIndex { get; }
        /// <summary>Detected visible slide title.</summary>
        public string? Title { get; }
        /// <summary>Shapes in inspectable reading order.</summary>
        public IReadOnlyList<PowerPointAccessibilityShapeInfo> Shapes { get; }
    }
}
