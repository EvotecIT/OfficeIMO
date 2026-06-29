namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a BIFF rich-text formatting run inside a cell string.
    /// </summary>
    public sealed class LegacyXlsTextFormattingRun {
        /// <summary>
        /// Creates rich-text run metadata.
        /// </summary>
        /// <param name="startCharacter">Zero-based character offset where this formatting run starts.</param>
        /// <param name="fontIndex">BIFF font index used by the run.</param>
        public LegacyXlsTextFormattingRun(ushort startCharacter, ushort fontIndex) {
            StartCharacter = startCharacter;
            FontIndex = fontIndex;
        }

        /// <summary>
        /// Gets the zero-based character offset where this formatting run starts.
        /// </summary>
        public ushort StartCharacter { get; }

        /// <summary>
        /// Gets the BIFF font index used by the run.
        /// </summary>
        public ushort FontIndex { get; }
    }
}
