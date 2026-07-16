using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes an OfficeArtFOPT property entry discovered in an XLS MsoDrawing payload.
    /// </summary>
    public sealed class LegacyXlsDrawingShapeProperty {
        private readonly OfficeArtProperty _property;

        /// <summary>
        /// Creates preserve-only metadata for an OfficeArtFOPTE property entry.
        /// </summary>
        public LegacyXlsDrawingShapeProperty(int index, ushort rawOperationId, uint value, int? availableComplexDataLength = null, string? complexText = null) {
            _property = new OfficeArtProperty(index, rawOperationId, value,
                availableComplexDataLength, complexText);
        }

        internal LegacyXlsDrawingShapeProperty(OfficeArtProperty property) {
            _property = property ?? throw new ArgumentNullException(nameof(property));
        }

        /// <summary>Gets the zero-based index of this property entry inside its FOPT record.</summary>
        public int Index => _property.Index;

        /// <summary>Gets the raw OfficeArtFOPTEOPID bitfield.</summary>
        public ushort RawOperationId => _property.RawOperationId;

        /// <summary>Gets the low 14-bit OfficeArt property identifier.</summary>
        public ushort PropertyId => _property.PropertyId;

        /// <summary>Gets a stable property identifier display key.</summary>
        public string PropertyIdKey => _property.PropertyIdKey;

        /// <summary>Gets the decoded OfficeArt property name, or a stable raw identifier for unknown property identifiers.</summary>
        public string PropertyName => _property.PropertyName;

        /// <summary>Gets the decoded OfficeArt property family, or <c>Unknown</c> when the identifier is not mapped yet.</summary>
        public string PropertyGroupName => _property.PropertyGroupName;

        /// <summary>Gets whether the property value references BLIP data.</summary>
        public bool IsBlipId => _property.IsBlipId;

        /// <summary>Gets whether the property value declares a following complex data payload length.</summary>
        public bool IsComplex => _property.IsComplex;

        /// <summary>Gets the raw 32-bit property value.</summary>
        public uint Value => _property.Value;

        /// <summary>Gets the declared complex data length when this is a complex property.</summary>
        public uint? DeclaredComplexDataLength => _property.DeclaredComplexDataLength;

        /// <summary>Gets the complex data bytes available in the containing record, when this is a complex property.</summary>
        public int? AvailableComplexDataLength => _property.AvailableComplexDataLength;

        /// <summary>Gets decoded complex text for text-bearing OfficeArt properties when available.</summary>
        public string? ComplexText => _property.ComplexText;
    }
}
