namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Lightweight description of a generated deck design alternative.
    /// </summary>
    public sealed class PowerPointDeckDesignSummary {
        internal PowerPointDeckDesignSummary(int index, PowerPointDeckDesign design) {
            Index = index;
            ThemeName = design.Theme.Name;
            DirectionName = design.Direction.Name;
            Mood = design.BaseIntent.Mood;
            Density = design.BaseIntent.Density;
            VisualStyle = design.BaseIntent.VisualStyle;
            LayoutStrategy = design.BaseIntent.LayoutStrategy;
            HeadingFontName = design.Theme.HeadingFontName;
            BodyFontName = design.Theme.BodyFontName;
            AccentColor = design.Theme.AccentColor;
            Accent2Color = design.Theme.Accent2Color;
            Accent3Color = design.Theme.Accent3Color;
            WarningColor = design.Theme.WarningColor;
            PaletteStyle = design.Theme.PaletteStyle;
            TypographyStyle = design.Theme.TypographyStyle;
            ShowsDirectionMotif = design.ShowDirectionMotif;
        }

        /// <summary>
        ///     Zero-based alternative index.
        /// </summary>
        public int Index { get; }

        /// <summary>
        ///     Generated theme name.
        /// </summary>
        public string ThemeName { get; }

        /// <summary>
        ///     Creative direction name.
        /// </summary>
        public string DirectionName { get; }

        /// <summary>
        ///     Broad deck mood.
        /// </summary>
        public PowerPointDesignMood Mood { get; }

        /// <summary>
        ///     Preferred content density.
        /// </summary>
        public PowerPointSlideDensity Density { get; }

        /// <summary>
        ///     Preferred visual style.
        /// </summary>
        public PowerPointVisualStyle VisualStyle { get; }

        /// <summary>
        ///     Strategy used by Auto slide variants.
        /// </summary>
        public PowerPointAutoLayoutStrategy LayoutStrategy { get; }

        /// <summary>
        ///     Heading font used by this alternative.
        /// </summary>
        public string HeadingFontName { get; }

        /// <summary>
        ///     Body font used by this alternative.
        /// </summary>
        public string BodyFontName { get; }

        /// <summary>
        ///     Primary brand accent.
        /// </summary>
        public string AccentColor { get; }

        /// <summary>
        ///     Secondary accent color.
        /// </summary>
        public string Accent2Color { get; }

        /// <summary>
        ///     Tertiary accent color.
        /// </summary>
        public string Accent3Color { get; }

        /// <summary>
        ///     Warm accent used for markers and highlights.
        /// </summary>
        public string WarningColor { get; }

        /// <summary>
        ///     Supporting palette strategy used by this alternative.
        /// </summary>
        public PowerPointPaletteStyle PaletteStyle { get; }

        /// <summary>
        ///     Typography strategy used by this alternative.
        /// </summary>
        public PowerPointTypographyStyle TypographyStyle { get; }

        /// <summary>
        ///     Whether this alternative uses repeated direction markers by default.
        /// </summary>
        public bool ShowsDirectionMotif { get; }

        /// <inheritdoc />
        public override string ToString() {
            return Index + ": " + DirectionName + " (" + Mood + ", " + VisualStyle + ", " +
                   HeadingFontName + "/" + BodyFontName + ")";
        }
    }
}
