namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes document-wide OfficeArtFDGGBlock drawing metadata discovered in an XLS drawing group.
    /// </summary>
    public sealed class LegacyXlsDrawingGroupBlock {
        /// <summary>
        /// Creates preserve-only metadata for an OfficeArtFDGGBlock record.
        /// </summary>
        public LegacyXlsDrawingGroupBlock(uint maxShapeId, uint declaredIdentifierClusterCount, uint savedShapeCount, uint savedDrawingCount, IReadOnlyList<LegacyXlsDrawingIdentifierCluster>? identifierClusters = null) {
            MaxShapeId = maxShapeId;
            DeclaredIdentifierClusterCount = declaredIdentifierClusterCount;
            SavedShapeCount = savedShapeCount;
            SavedDrawingCount = savedDrawingCount;
            IdentifierClusters = identifierClusters?.ToArray() ?? Array.Empty<LegacyXlsDrawingIdentifierCluster>();
        }

        /// <summary>Gets the current maximum shape identifier used in any drawing.</summary>
        public uint MaxShapeId { get; }

        /// <summary>Gets the saved OfficeArtIDCL cluster count plus one, as stored by OfficeArtFDGG.</summary>
        public uint DeclaredIdentifierClusterCount { get; }

        /// <summary>Gets the total number of saved shapes across drawings.</summary>
        public uint SavedShapeCount { get; }

        /// <summary>Gets the total number of saved drawings in the file.</summary>
        public uint SavedDrawingCount { get; }

        /// <summary>Gets identifier clusters decoded after the OfficeArtFDGG header.</summary>
        public IReadOnlyList<LegacyXlsDrawingIdentifierCluster> IdentifierClusters { get; }
    }
}
