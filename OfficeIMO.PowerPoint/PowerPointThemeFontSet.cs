namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents the major/minor fonts used by a theme.
    /// </summary>
    public readonly struct PowerPointThemeFontSet {
        /// <summary>
        ///     Creates a theme font set.
        /// </summary>
        public PowerPointThemeFontSet(string? majorLatin, string? minorLatin,
            string? majorEastAsian, string? minorEastAsian,
            string? majorComplexScript, string? minorComplexScript) {
            MajorLatin = majorLatin;
            MinorLatin = minorLatin;
            MajorEastAsian = majorEastAsian;
            MinorEastAsian = minorEastAsian;
            MajorComplexScript = majorComplexScript;
            MinorComplexScript = minorComplexScript;
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
        ///     Major font East Asian typeface.
        /// </summary>
        public string? MajorEastAsian { get; }

        /// <summary>
        ///     Minor font East Asian typeface.
        /// </summary>
        public string? MinorEastAsian { get; }

        /// <summary>
        ///     Major font complex script typeface.
        /// </summary>
        public string? MajorComplexScript { get; }

        /// <summary>
        ///     Minor font complex script typeface.
        /// </summary>
        public string? MinorComplexScript { get; }

        /// <summary>
        ///     Returns a display-friendly string.
        /// </summary>
        public override string ToString() {
            return $"{MajorLatin ?? "?"} / {MinorLatin ?? "?"}";
        }
    }
}
