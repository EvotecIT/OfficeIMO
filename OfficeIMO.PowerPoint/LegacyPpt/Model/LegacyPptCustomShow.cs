using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents one named sequence of presentation slides.</summary>
    public sealed class LegacyPptCustomShow {
        private bool _hasStructuralLoss;

        internal LegacyPptCustomShow(string name, IReadOnlyList<uint> slideIds,
            long recordOffset, bool hasStructuralLoss) {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            SlideIds = new ReadOnlyCollection<uint>(slideIds.ToArray());
            RecordOffset = recordOffset;
            _hasStructuralLoss = hasStructuralLoss;
        }

        /// <summary>Gets the custom-show name.</summary>
        public string Name { get; }

        /// <summary>Gets binary slide identifiers in custom-show order.</summary>
        public IReadOnlyList<uint> SlideIds { get; }

        /// <summary>Gets whether one or more referenced slides are absent.</summary>
        public bool HasUnresolvedSlides { get; private set; }

        internal long RecordOffset { get; }

        internal bool IsEditable => !_hasStructuralLoss && !HasUnresolvedSlides
            && Name.Length > 0;

        internal void MarkStructurallyLossy() => _hasStructuralLoss = true;

        internal void MarkUnresolvedSlides() => HasUnresolvedSlides = true;
    }
}
