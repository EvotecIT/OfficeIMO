namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes per-drawing OfficeArtFDG metadata discovered in an XLS drawing.
    /// </summary>
    public sealed class LegacyXlsDrawingGroupInfo {
        /// <summary>
        /// Creates preserve-only metadata for an OfficeArtFDG record.
        /// </summary>
        public LegacyXlsDrawingGroupInfo(ushort drawingId, uint shapeCount, uint lastShapeId) {
            DrawingId = drawingId;
            ShapeCount = shapeCount;
            LastShapeId = lastShapeId;
        }

        /// <summary>Gets the drawing identifier from the OfficeArtFDG record header instance field.</summary>
        public ushort DrawingId { get; }

        /// <summary>Gets the number of shapes in this drawing.</summary>
        public uint ShapeCount { get; }

        /// <summary>Gets the shape identifier of the last shape in this drawing.</summary>
        public uint LastShapeId { get; }
    }
}
