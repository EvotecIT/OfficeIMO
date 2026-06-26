namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes OfficeArt fill-property metadata decoded from a legacy XLS chart GelFrame record.
    /// </summary>
    public sealed class LegacyXlsChartGelFrame {
        internal LegacyXlsChartGelFrame(
            IReadOnlyList<LegacyXlsDrawingOfficeArtRecord> officeArtRecords,
            IReadOnlyList<LegacyXlsDrawingShapeProperty> shapeProperties) {
            OfficeArtRecords = officeArtRecords?.ToArray() ?? Array.Empty<LegacyXlsDrawingOfficeArtRecord>();
            ShapeProperties = shapeProperties?.ToArray() ?? Array.Empty<LegacyXlsDrawingShapeProperty>();
        }

        /// <summary>Gets nested OfficeArt record headers discovered in the GelFrame payload.</summary>
        public IReadOnlyList<LegacyXlsDrawingOfficeArtRecord> OfficeArtRecords { get; }

        /// <summary>Gets OfficeArtFOPT properties discovered in the GelFrame payload.</summary>
        public IReadOnlyList<LegacyXlsDrawingShapeProperty> ShapeProperties { get; }

        /// <summary>Gets the number of nested OfficeArt record headers discovered in the GelFrame payload.</summary>
        public int OfficeArtRecordCount => OfficeArtRecords.Count;

        /// <summary>Gets the number of OfficeArtFOPT properties discovered in the GelFrame payload.</summary>
        public int ShapePropertyCount => ShapeProperties.Count;

        /// <summary>Gets whether the GelFrame payload exposed at least one OfficeArtFOPT property.</summary>
        public bool HasShapeProperties => ShapeProperties.Count > 0;
    }
}
