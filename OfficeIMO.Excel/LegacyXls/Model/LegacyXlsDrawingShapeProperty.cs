namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes an OfficeArtFOPT property entry discovered in an XLS MsoDrawing payload.
    /// </summary>
    public sealed class LegacyXlsDrawingShapeProperty {
        /// <summary>
        /// Creates preserve-only metadata for an OfficeArtFOPTE property entry.
        /// </summary>
        public LegacyXlsDrawingShapeProperty(int index, ushort rawOperationId, uint value, int? availableComplexDataLength = null) {
            if (index < 0) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            Index = index;
            RawOperationId = rawOperationId;
            PropertyId = checked((ushort)(rawOperationId & 0x3fff));
            IsBlipId = (rawOperationId & 0x4000) != 0;
            IsComplex = (rawOperationId & 0x8000) != 0;
            Value = value;
            AvailableComplexDataLength = availableComplexDataLength;
        }

        /// <summary>Gets the zero-based index of this property entry inside its FOPT record.</summary>
        public int Index { get; }

        /// <summary>Gets the raw OfficeArtFOPTEOPID bitfield.</summary>
        public ushort RawOperationId { get; }

        /// <summary>Gets the low 14-bit OfficeArt property identifier.</summary>
        public ushort PropertyId { get; }

        /// <summary>Gets a stable property identifier display key.</summary>
        public string PropertyIdKey => $"PropertyId:0x{PropertyId:X4}";

        /// <summary>Gets whether the property value references BLIP data.</summary>
        public bool IsBlipId { get; }

        /// <summary>Gets whether the property value declares a following complex data payload length.</summary>
        public bool IsComplex { get; }

        /// <summary>Gets the raw 32-bit property value.</summary>
        public uint Value { get; }

        /// <summary>Gets the declared complex data length when this is a complex property.</summary>
        public uint? DeclaredComplexDataLength => IsComplex ? Value : null;

        /// <summary>Gets the complex data bytes available in the containing record, when this is a complex property.</summary>
        public int? AvailableComplexDataLength { get; }
    }
}
