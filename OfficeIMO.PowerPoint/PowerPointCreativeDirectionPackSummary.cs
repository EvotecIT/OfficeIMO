using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Lightweight description of a creative direction pack that can be shown before applying it to a design brief.
    /// </summary>
    public sealed class PowerPointCreativeDirectionPackSummary {
        internal PowerPointCreativeDirectionPackSummary(int index, PowerPointCreativeDirectionPack pack,
            string name, string description, PowerPointDesignRecipe? recipe,
            PowerPointPaletteStyle? paletteStyle, PowerPointAutoLayoutStrategy? layoutStrategy,
            PowerPointDesignVariety variety, IReadOnlyList<PowerPointDesignMood> preferredMoods,
            IReadOnlyList<PowerPointSlideDensity> preferredDensities,
            IReadOnlyList<PowerPointVisualStyle> preferredVisualStyles) {
            Index = index;
            Pack = pack;
            Name = name;
            Description = description;
            Recipe = recipe;
            RecipeName = recipe?.Name;
            PaletteStyle = paletteStyle;
            LayoutStrategy = layoutStrategy;
            Variety = variety;
            PreferredMoods = preferredMoods;
            PreferredDensities = preferredDensities;
            PreferredVisualStyles = preferredVisualStyles;
        }

        /// <summary>
        ///     Zero-based pack index within the built-in pack list.
        /// </summary>
        public int Index { get; }

        /// <summary>
        ///     Creative direction pack value.
        /// </summary>
        public PowerPointCreativeDirectionPack Pack { get; }

        /// <summary>
        ///     Display name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///     Short explanation of where the pack fits.
        /// </summary>
        public string Description { get; }

        /// <summary>
        ///     Name of the recipe applied by this pack, when the pack sets one.
        /// </summary>
        public string? RecipeName { get; }

        /// <summary>
        ///     Supporting palette strategy applied by this pack, when the pack sets one.
        /// </summary>
        public PowerPointPaletteStyle? PaletteStyle { get; }

        /// <summary>
        ///     Auto layout strategy applied by this pack, when the pack sets one.
        /// </summary>
        public PowerPointAutoLayoutStrategy? LayoutStrategy { get; }

        /// <summary>
        ///     Alternative variety applied by this pack.
        /// </summary>
        public PowerPointDesignVariety Variety { get; }

        /// <summary>
        ///     Preferred moods used to rank recipe directions.
        /// </summary>
        public IReadOnlyList<PowerPointDesignMood> PreferredMoods { get; }

        /// <summary>
        ///     Preferred densities used to rank recipe directions.
        /// </summary>
        public IReadOnlyList<PowerPointSlideDensity> PreferredDensities { get; }

        /// <summary>
        ///     Preferred visual styles used to rank recipe directions.
        /// </summary>
        public IReadOnlyList<PowerPointVisualStyle> PreferredVisualStyles { get; }

        internal PowerPointDesignRecipe? Recipe { get; }

        /// <inheritdoc />
        public override string ToString() {
            string recipeName = RecipeName ?? "Purpose matched";
            string layout = LayoutStrategy?.ToString() ?? "Default layout";
            return Index + ": " + Name + " (" + recipeName + ", " + layout + ")";
        }
    }
}
