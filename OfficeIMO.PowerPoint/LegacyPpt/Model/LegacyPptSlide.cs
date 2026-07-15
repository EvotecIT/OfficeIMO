namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents one presentation slide decoded from a binary PowerPoint file.</summary>
    public sealed class LegacyPptSlide {
        private readonly List<LegacyPptShape> _shapes = new();

        internal LegacyPptSlide(uint slideId, uint persistId) {
            SlideId = slideId;
            PersistId = persistId;
        }

        /// <summary>Gets the legacy slide identifier.</summary>
        public uint SlideId { get; }

        /// <summary>Gets the persist object id that owns this slide container.</summary>
        public uint PersistId { get; }

        /// <summary>Gets the slide name, when present.</summary>
        public string? Name { get; internal set; }

        /// <summary>Gets the projected shapes in drawing order.</summary>
        public IReadOnlyList<LegacyPptShape> Shapes => _shapes;

        /// <summary>Gets speaker notes flattened to plain text.</summary>
        public string NotesText { get; internal set; } = string.Empty;

        internal void AddShape(LegacyPptShape shape) => _shapes.Add(shape);
    }
}
