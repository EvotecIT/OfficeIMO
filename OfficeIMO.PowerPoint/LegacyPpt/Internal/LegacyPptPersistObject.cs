namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Retains one live persist object and its exact top-level record bytes.</summary>
    internal sealed class LegacyPptPersistObject {
        internal LegacyPptPersistObject(uint persistId, uint streamOffset, ushort recordType, byte[] recordBytes) {
            PersistId = persistId;
            StreamOffset = streamOffset;
            RecordType = recordType;
            RecordBytes = recordBytes ?? throw new ArgumentNullException(nameof(recordBytes));
        }

        internal uint PersistId { get; }

        internal uint StreamOffset { get; }

        internal ushort RecordType { get; }

        internal byte[] RecordBytes { get; }
    }
}
