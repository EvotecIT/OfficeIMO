namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Identifies a notes or handout master decoded from a binary PowerPoint file.</summary>
    public enum LegacyPptSpecialMasterKind {
        /// <summary>The master used by notes pages.</summary>
        Notes,

        /// <summary>The master used by printed handouts.</summary>
        Handout
    }

    /// <summary>Represents a notes or handout master referenced by the binary document atom.</summary>
    public sealed class LegacyPptSpecialMaster {
        private readonly List<LegacyPptShape> _shapes = new();
        private readonly List<LegacyPptConnectorRule> _connectorRules = new();

        internal LegacyPptSpecialMaster(LegacyPptSpecialMasterKind kind, uint persistId) {
            Kind = kind;
            PersistId = persistId;
        }

        /// <summary>Gets whether this is the notes or handout master.</summary>
        public LegacyPptSpecialMasterKind Kind { get; }

        /// <summary>Gets the persist object id referenced by the document atom.</summary>
        public uint PersistId { get; }

        /// <summary>Gets the color scheme stored on this master.</summary>
        public LegacyPptColorScheme? ColorScheme { get; internal set; }

        /// <summary>Gets the DrawingML theme stored in PowerPoint 2007+ round-trip records.</summary>
        public LegacyPptRoundTripTheme? RoundTripTheme { get; internal set; }

        /// <summary>Gets the explicit OfficeArt background shape stored on this master.</summary>
        public LegacyPptBackground? Background { get; internal set; }

        /// <summary>Gets the projected shapes in drawing order.</summary>
        public IReadOnlyList<LegacyPptShape> Shapes => _shapes;

        /// <summary>Gets OfficeArt connector attachment rules in solver order.</summary>
        public IReadOnlyList<LegacyPptConnectorRule> ConnectorRules => _connectorRules;

        internal void AddShape(LegacyPptShape shape) => _shapes.Add(shape);

        internal void AddConnectorRule(LegacyPptConnectorRule rule) => _connectorRules.Add(rule);
    }
}
