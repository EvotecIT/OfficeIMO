namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents one presentation slide decoded from a binary PowerPoint file.</summary>
    public sealed class LegacyPptSlide {
        private readonly List<LegacyPptShape> _shapes = new();
        private readonly List<LegacyPptConnectorRule> _connectorRules = new();

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

        /// <summary>Gets whether the slide is hidden during a slide show.</summary>
        public bool Hidden { get; internal set; }

        /// <summary>Gets the legacy master identifier referenced by this slide.</summary>
        public uint MasterId { get; internal set; }

        /// <summary>Gets the legacy layout hint stored on this slide.</summary>
        public uint LayoutType { get; internal set; }

        /// <summary>Gets whether the slide inherits shapes from its master.</summary>
        public bool FollowsMasterObjects { get; internal set; }

        /// <summary>Gets whether the slide inherits its master's color scheme.</summary>
        public bool FollowsMasterColorScheme { get; internal set; }

        /// <summary>Gets whether the slide inherits its master's background.</summary>
        public bool FollowsMasterBackground { get; internal set; }

        /// <summary>Gets the color scheme stored on this slide.</summary>
        public LegacyPptColorScheme? ColorScheme { get; internal set; }

        /// <summary>Gets the projected shapes in drawing order.</summary>
        public IReadOnlyList<LegacyPptShape> Shapes => _shapes;

        /// <summary>Gets OfficeArt connector attachment rules in solver order.</summary>
        public IReadOnlyList<LegacyPptConnectorRule> ConnectorRules => _connectorRules;

        /// <summary>Gets speaker notes flattened to plain text.</summary>
        public string NotesText { get; internal set; } = string.Empty;

        internal void AddShape(LegacyPptShape shape) => _shapes.Add(shape);

        internal void AddConnectorRule(LegacyPptConnectorRule rule) => _connectorRules.Add(rule);
    }
}
