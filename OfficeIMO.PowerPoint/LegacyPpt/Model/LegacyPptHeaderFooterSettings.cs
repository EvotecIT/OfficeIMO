namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents the header and footer options stored by a binary PowerPoint container.</summary>
    public sealed class LegacyPptHeaderFooterSettings {
        internal LegacyPptHeaderFooterSettings(short dateTimeFormatId, ushort rawFlags,
            string userDateText, string headerText, string footerText) {
            DateTimeFormatId = dateTimeFormatId;
            RawFlags = rawFlags;
            UserDateText = userDateText ?? string.Empty;
            HeaderText = headerText ?? string.Empty;
            FooterText = footerText ?? string.Empty;
        }

        /// <summary>Gets the legacy date/time format identifier, from 0 through 13.</summary>
        public short DateTimeFormatId { get; }

        /// <summary>Gets the complete 16-bit options field, including ignored reserved bits.</summary>
        public ushort RawFlags { get; }

        /// <summary>Gets whether a date is displayed in the footer area.</summary>
        public bool ShowDate => (RawFlags & 0x0001) != 0;

        /// <summary>Gets whether the current date/time supplies the displayed date.</summary>
        public bool UseAutomaticDateTime => (RawFlags & 0x0002) != 0;

        /// <summary>Gets whether the stored user date supplies the displayed date.</summary>
        public bool UseUserDate => (RawFlags & 0x0004) != 0;

        /// <summary>Gets whether the slide or page number is displayed.</summary>
        public bool ShowSlideNumber => (RawFlags & 0x0008) != 0;

        /// <summary>Gets whether the header text is displayed.</summary>
        public bool ShowHeader => (RawFlags & 0x0010) != 0;

        /// <summary>Gets whether the footer text is displayed.</summary>
        public bool ShowFooter => (RawFlags & 0x0020) != 0;

        /// <summary>Gets the fixed user-date text, when present.</summary>
        public string UserDateText { get; }

        /// <summary>Gets the header text, when present.</summary>
        public string HeaderText { get; }

        /// <summary>Gets the footer text, when present.</summary>
        public string FooterText { get; }

        internal string CreateLayoutKey() =>
            $"{DateTimeFormatId}:{RawFlags & 0x003F:X2}:{UserDateText}:{HeaderText}:{FooterText}";
    }
}
