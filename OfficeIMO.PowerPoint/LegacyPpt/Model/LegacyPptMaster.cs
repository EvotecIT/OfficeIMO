namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents a main master or title master decoded from a binary PowerPoint file.</summary>
    public sealed class LegacyPptMaster {
        private readonly List<LegacyPptShape> _shapes = new();

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

        /// <summary>Gets the projected shapes in drawing order.</summary>
        public IReadOnlyList<LegacyPptShape> Shapes => _shapes;

        internal void AddShape(LegacyPptShape shape) => _shapes.Add(shape);
    }
}
