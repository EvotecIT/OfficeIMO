namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a parsed formatting run from a legacy XLS TxO comment.
    /// </summary>
    public sealed class LegacyXlsCommentFormattingRun {
        /// <summary>
        /// Creates parsed legacy XLS comment formatting run metadata.
        /// </summary>
        /// <param name="startCharacter">Zero-based index of the first character in this run.</param>
        /// <param name="fontIndex">Legacy FontIndex referenced by the run.</param>
        public LegacyXlsCommentFormattingRun(ushort startCharacter, ushort fontIndex) {
            StartCharacter = startCharacter;
            FontIndex = fontIndex;
        }

        /// <summary>
        /// Gets the zero-based index of the first character in this run.
        /// </summary>
        public ushort StartCharacter { get; }

        /// <summary>
        /// Gets the legacy FontIndex referenced by the run.
        /// </summary>
        public ushort FontIndex { get; }
    }
}
