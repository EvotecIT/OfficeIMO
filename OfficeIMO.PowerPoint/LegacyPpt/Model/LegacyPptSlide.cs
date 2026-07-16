namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents one presentation slide decoded from a binary PowerPoint file.</summary>
    public sealed class LegacyPptSlide {
        private readonly List<LegacyPptShape> _shapes = new();
        private readonly List<LegacyPptConnectorRule> _connectorRules = new();
        private readonly List<LegacyPptComment> _comments = new();
        private IReadOnlyList<LegacyPptPlaceholderKind> _layoutPlaceholderTypes =
            Array.Empty<LegacyPptPlaceholderKind>();

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

        /// <summary>Gets the typed legacy layout hint, or null when the source value is undefined.</summary>
        public LegacyPptSlideLayoutType? Layout { get; internal set; }

        /// <summary>Gets the eight placeholder-kind slots stored in the SlideAtom layout signature.</summary>
        public IReadOnlyList<LegacyPptPlaceholderKind> LayoutPlaceholderTypes =>
            _layoutPlaceholderTypes;

        /// <summary>Gets the legacy notes-slide identifier referenced by this slide.</summary>
        public uint NotesId { get; internal set; }

        /// <summary>Gets whether the slide inherits shapes from its master.</summary>
        public bool FollowsMasterObjects { get; internal set; }

        /// <summary>Gets whether the slide inherits its master's color scheme.</summary>
        public bool FollowsMasterColorScheme { get; internal set; }

        /// <summary>Gets whether the slide inherits its master's background.</summary>
        public bool FollowsMasterBackground { get; internal set; }

        /// <summary>Gets the color scheme stored on this slide.</summary>
        public LegacyPptColorScheme? ColorScheme { get; internal set; }

        /// <summary>Gets the DrawingML theme override stored in PowerPoint 2007+ round-trip records.</summary>
        public LegacyPptRoundTripTheme? RoundTripTheme { get; internal set; }

        /// <summary>Gets the explicit OfficeArt background shape stored on this slide.</summary>
        public LegacyPptBackground? Background { get; internal set; }

        /// <summary>Gets this slide's explicit header/footer override, when present.</summary>
        public LegacyPptHeaderFooterSettings? HeaderFooter { get; internal set; }

        /// <summary>Gets this slide's transition and advance settings, when present.</summary>
        public LegacyPptTransition? Transition { get; internal set; }

        /// <summary>Gets the legacy review comments in record order.</summary>
        public IReadOnlyList<LegacyPptComment> Comments => _comments;

        /// <summary>Gets the projected shapes in drawing order.</summary>
        public IReadOnlyList<LegacyPptShape> Shapes => _shapes;

        /// <summary>Gets OfficeArt connector attachment rules in solver order.</summary>
        public IReadOnlyList<LegacyPptConnectorRule> ConnectorRules => _connectorRules;

        /// <summary>Gets the associated notes page, when the slide references one.</summary>
        public LegacyPptNotesPage? NotesPage { get; internal set; }

        /// <summary>Gets speaker notes flattened to plain text.</summary>
        public string NotesText { get; internal set; } = string.Empty;

        internal void AddShape(LegacyPptShape shape) => _shapes.Add(shape);

        internal void AddConnectorRule(LegacyPptConnectorRule rule) => _connectorRules.Add(rule);

        internal void AddComment(LegacyPptComment comment) => _comments.Add(comment);

        internal void SetLayoutPlaceholderTypes(
            IReadOnlyList<LegacyPptPlaceholderKind> placeholderTypes) =>
            _layoutPlaceholderTypes = placeholderTypes?.ToArray()
                ?? throw new ArgumentNullException(nameof(placeholderTypes));
    }
}
