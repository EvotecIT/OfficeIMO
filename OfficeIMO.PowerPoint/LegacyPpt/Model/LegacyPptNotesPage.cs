namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents one notes page associated with a binary presentation slide.</summary>
    public sealed class LegacyPptNotesPage {
        private readonly List<LegacyPptShape> _shapes = new();
        private readonly List<LegacyPptConnectorRule> _connectorRules = new();

        internal LegacyPptNotesPage(uint notesId, uint persistId, uint slideId) {
            NotesId = notesId;
            PersistId = persistId;
            SlideId = slideId;
        }

        /// <summary>Gets the notes identifier referenced by the owning SlideAtom.</summary>
        public uint NotesId { get; }

        /// <summary>Gets the persist object identifier for the NotesContainer.</summary>
        public uint PersistId { get; }

        /// <summary>Gets the presentation-slide identifier referenced by the NotesAtom.</summary>
        public uint SlideId { get; }

        /// <summary>Gets whether the notes page inherits shapes from the notes master.</summary>
        public bool FollowsMasterObjects { get; internal set; }

        /// <summary>Gets whether the notes page inherits the notes-master color scheme.</summary>
        public bool FollowsMasterColorScheme { get; internal set; }

        /// <summary>Gets whether the notes page inherits the notes-master background.</summary>
        public bool FollowsMasterBackground { get; internal set; }

        /// <summary>Gets the color scheme stored on this notes page.</summary>
        public LegacyPptColorScheme? ColorScheme { get; internal set; }

        /// <summary>Gets the DrawingML theme override stored in PowerPoint 2007+ round-trip records.</summary>
        public LegacyPptRoundTripTheme? RoundTripTheme { get; internal set; }

        /// <summary>Gets the explicit OfficeArt background shape stored on this notes page.</summary>
        public LegacyPptBackground? Background { get; internal set; }

        /// <summary>Gets the decoded notes-page shapes in drawing order.</summary>
        public IReadOnlyList<LegacyPptShape> Shapes => _shapes;

        /// <summary>Gets OfficeArt connector attachment rules in solver order.</summary>
        public IReadOnlyList<LegacyPptConnectorRule> ConnectorRules => _connectorRules;

        /// <summary>Gets the flattened editable speaker-note body text.</summary>
        public string Text { get; internal set; } = string.Empty;

        internal void AddShape(LegacyPptShape shape) => _shapes.Add(shape);

        internal void AddConnectorRule(LegacyPptConnectorRule rule) =>
            _connectorRules.Add(rule);
    }
}
