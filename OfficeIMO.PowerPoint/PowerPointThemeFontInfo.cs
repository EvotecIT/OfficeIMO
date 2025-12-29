namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents the major/minor Latin fonts in a theme.
    /// </summary>
    public readonly struct PowerPointThemeFontInfo {
        /// <summary>
        ///     Creates a theme font info entry.
        /// </summary>
        public PowerPointThemeFontInfo(string? majorLatin, string? minorLatin) {
            MajorLatin = majorLatin;
            MinorLatin = minorLatin;
        }

        /// <summary>
        ///     Major font Latin typeface.
        /// </summary>
        public string? MajorLatin { get; }

        /// <summary>
        ///     Minor font Latin typeface.
        /// </summary>
        public string? MinorLatin { get; }

        /// <summary>
        ///     Returns a display-friendly string.
        /// </summary>
        public override string ToString() {
            return $"{MajorLatin ?? "?"} / {MinorLatin ?? "?"}";
        }
    }
}
