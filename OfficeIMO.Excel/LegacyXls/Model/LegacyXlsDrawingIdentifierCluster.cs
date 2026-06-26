namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes an OfficeArtIDCL shape identifier cluster inside document-wide drawing metadata.
    /// </summary>
    public sealed class LegacyXlsDrawingIdentifierCluster {
        /// <summary>
        /// Creates preserve-only metadata for an OfficeArtIDCL record.
        /// </summary>
        public LegacyXlsDrawingIdentifierCluster(uint drawingId, uint currentShapeId) {
            DrawingId = drawingId;
            CurrentShapeId = currentShapeId;
        }

        /// <summary>Gets the drawing identifier that owns this identifier cluster.</summary>
        public uint DrawingId { get; }

        /// <summary>Gets the largest shape identifier currently assigned in this cluster.</summary>
        public uint CurrentShapeId { get; }
    }
}
