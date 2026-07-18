namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents a main master or title master decoded from a binary PowerPoint file.</summary>
    public sealed class LegacyPptMaster {
        private readonly List<LegacyPptShape> _shapes = new();
        private readonly List<LegacyPptConnectorRule> _connectorRules = new();
        private readonly List<LegacyPptTextMasterStyle> _textMasterStyles = new();
        private IReadOnlyList<LegacyPptPlaceholderKind> _layoutPlaceholderTypes =
            Array.Empty<LegacyPptPlaceholderKind>();

        internal LegacyPptMaster(uint masterId, uint persistId, bool isMainMaster, uint parentMasterId) {
            MasterId = masterId;
            PersistId = persistId;
            IsMainMaster = isMainMaster;
            ParentMasterId = parentMasterId;
        }

        /// <summary>Gets the legacy master identifier.</summary>
        public uint MasterId { get; }

        /// <summary>Gets the persist object id that owns this master.</summary>
        public uint PersistId { get; }

        /// <summary>Gets whether this is a main master rather than a title master.</summary>
        public bool IsMainMaster { get; }

        /// <summary>Gets the raw legacy layout hint stored on the master.</summary>
        public uint LayoutType { get; internal set; }

        /// <summary>Gets the typed legacy layout hint, or null when the source value is undefined.</summary>
        public LegacyPptSlideLayoutType? Layout { get; internal set; }

        /// <summary>Gets the eight placeholder-kind slots stored in the master's SlideAtom.</summary>
        public IReadOnlyList<LegacyPptPlaceholderKind> LayoutPlaceholderTypes =>
            _layoutPlaceholderTypes;

        /// <summary>Gets the main master identifier inherited by a title master, or zero when absent.</summary>
        public uint ParentMasterId { get; }

        /// <summary>Gets whether this title master inherits its parent's color scheme.</summary>
        public bool FollowsMasterColorScheme { get; internal set; }

        /// <summary>Gets whether this title master inherits its parent's shapes.</summary>
        public bool FollowsMasterObjects { get; internal set; }

        /// <summary>Gets whether this title master inherits its parent's background.</summary>
        public bool FollowsMasterBackground { get; internal set; }

        /// <summary>Gets the color scheme stored on this master.</summary>
        public LegacyPptColorScheme? ColorScheme { get; internal set; }

        /// <summary>Gets the DrawingML theme stored in PowerPoint 2007+ round-trip records.</summary>
        public LegacyPptRoundTripTheme? RoundTripTheme { get; internal set; }

        /// <summary>Gets the explicit OfficeArt background shape stored on this master.</summary>
        public LegacyPptBackground? Background { get; internal set; }

        /// <summary>Gets this master's explicit header/footer options, when present.</summary>
        public LegacyPptHeaderFooterSettings? HeaderFooter { get; internal set; }

        /// <summary>Gets the projected shapes in drawing order.</summary>
        public IReadOnlyList<LegacyPptShape> Shapes => _shapes;

        /// <summary>Gets OfficeArt connector attachment rules in solver order.</summary>
        public IReadOnlyList<LegacyPptConnectorRule> ConnectorRules => _connectorRules;

        /// <summary>Gets the decoded title, body, notes, and other master text styles.</summary>
        public IReadOnlyList<LegacyPptTextMasterStyle> TextMasterStyles => _textMasterStyles;

        internal void AddShape(LegacyPptShape shape) => _shapes.Add(shape);

        internal void AddConnectorRule(LegacyPptConnectorRule rule) => _connectorRules.Add(rule);

        internal void AddTextMasterStyle(LegacyPptTextMasterStyle style) => _textMasterStyles.Add(style);

        internal void SetLayoutPlaceholderTypes(
            IReadOnlyList<LegacyPptPlaceholderKind> placeholderTypes) =>
            _layoutPlaceholderTypes = placeholderTypes?.ToArray()
                ?? throw new ArgumentNullException(nameof(placeholderTypes));
    }
}
