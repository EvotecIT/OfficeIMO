using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Represents one UserEditAtom and the persist-directory entries written by that edit.</summary>
    internal sealed class LegacyPptUserEdit {
        internal LegacyPptUserEdit(uint streamOffset, uint previousEditOffset, uint persistDirectoryOffset,
            uint documentPersistId, uint persistIdSeed, IReadOnlyDictionary<uint, uint> persistObjectOffsets) {
            StreamOffset = streamOffset;
            PreviousEditOffset = previousEditOffset;
            PersistDirectoryOffset = persistDirectoryOffset;
            DocumentPersistId = documentPersistId;
            PersistIdSeed = persistIdSeed;
            PersistObjectOffsets = new ReadOnlyDictionary<uint, uint>(
                persistObjectOffsets.ToDictionary(pair => pair.Key, pair => pair.Value));
        }

        internal uint StreamOffset { get; }

        internal uint PreviousEditOffset { get; }

        internal uint PersistDirectoryOffset { get; }

        internal uint DocumentPersistId { get; }

        internal uint PersistIdSeed { get; }

        internal IReadOnlyDictionary<uint, uint> PersistObjectOffsets { get; }
    }
}
