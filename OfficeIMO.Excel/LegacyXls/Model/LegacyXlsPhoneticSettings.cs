namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents worksheet-level phonetic display defaults decoded from a BIFF PhoneticInfo record.
    /// </summary>
    public sealed class LegacyXlsPhoneticSettings {
        /// <summary>
        /// Initializes worksheet-level phonetic display defaults.
        /// </summary>
        /// <param name="fontId">Zero-based font index used to display phonetic text.</param>
        /// <param name="type">Character conversion type for phonetic text.</param>
        /// <param name="alignment">Alignment for phonetic text.</param>
        /// <param name="ranges">Optional BIFF range list attached to the PhoneticInfo record.</param>
        public LegacyXlsPhoneticSettings(
            ushort fontId,
            LegacyXlsPhoneticType type,
            LegacyXlsPhoneticAlignment alignment,
            IReadOnlyList<string> ranges) {
            FontId = fontId;
            Type = type;
            Alignment = alignment;
            Ranges = ranges ?? Array.Empty<string>();
        }

        /// <summary>Gets the zero-based font index used to display phonetic text.</summary>
        public ushort FontId { get; }

        /// <summary>Gets the character conversion type for phonetic text.</summary>
        public LegacyXlsPhoneticType Type { get; }

        /// <summary>Gets the alignment for phonetic text.</summary>
        public LegacyXlsPhoneticAlignment Alignment { get; }

        /// <summary>Gets BIFF cell ranges attached to the PhoneticInfo record.</summary>
        public IReadOnlyList<string> Ranges { get; }

        /// <summary>Gets whether the record carries range-scoped phonetic guide metadata.</summary>
        public bool HasRanges => Ranges.Count > 0;
    }
}
