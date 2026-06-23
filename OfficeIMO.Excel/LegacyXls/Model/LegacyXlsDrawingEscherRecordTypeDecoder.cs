namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Maps OfficeArt record type identifiers to stable preserve-only model names.
    /// </summary>
    internal static class LegacyXlsDrawingEscherRecordTypeDecoder {
        /// <summary>
        /// Gets the known OfficeArt record type for a raw identifier.
        /// </summary>
        internal static LegacyXlsDrawingEscherRecordType? TryGetKind(ushort? recordType) {
            if (!recordType.HasValue) {
                return null;
            }

            return recordType.Value switch {
                0xF000 => LegacyXlsDrawingEscherRecordType.OfficeArtDggContainer,
                0xF001 => LegacyXlsDrawingEscherRecordType.OfficeArtBStoreContainer,
                0xF002 => LegacyXlsDrawingEscherRecordType.OfficeArtDgContainer,
                0xF003 => LegacyXlsDrawingEscherRecordType.OfficeArtSpgrContainer,
                0xF004 => LegacyXlsDrawingEscherRecordType.OfficeArtSpContainer,
                0xF005 => LegacyXlsDrawingEscherRecordType.OfficeArtSolverContainer,
                0xF006 => LegacyXlsDrawingEscherRecordType.OfficeArtFDGGBlock,
                0xF007 => LegacyXlsDrawingEscherRecordType.OfficeArtFBSE,
                0xF008 => LegacyXlsDrawingEscherRecordType.OfficeArtFDG,
                0xF009 => LegacyXlsDrawingEscherRecordType.OfficeArtFSPGR,
                0xF00A => LegacyXlsDrawingEscherRecordType.OfficeArtFSP,
                0xF00B => LegacyXlsDrawingEscherRecordType.OfficeArtFOPT,
                0xF00D => LegacyXlsDrawingEscherRecordType.OfficeArtFClientTextbox,
                0xF00F => LegacyXlsDrawingEscherRecordType.OfficeArtChildAnchor,
                0xF010 => LegacyXlsDrawingEscherRecordType.OfficeArtFClientAnchor,
                0xF011 => LegacyXlsDrawingEscherRecordType.OfficeArtFClientData,
                0xF11E => LegacyXlsDrawingEscherRecordType.OfficeArtSplitMenuColorContainer,
                _ => null
            };
        }

        /// <summary>
        /// Gets a stable OfficeArt record type name, using a hexadecimal fallback for unknown identifiers.
        /// </summary>
        internal static string GetName(ushort recordType) {
            return TryGetKind(recordType)?.ToString() ?? $"EscherRecordType:0x{recordType:X4}";
        }
    }
}
