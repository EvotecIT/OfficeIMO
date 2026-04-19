using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Lightweight description of a design recipe that can be shown before creating deck alternatives.
    /// </summary>
    public sealed class PowerPointDesignRecipeSummary {
        internal PowerPointDesignRecipeSummary(int index, PowerPointDesignRecipe recipe) {
            Index = index;
            Name = recipe.Name;
            Description = recipe.Description;
            DefaultEyebrow = recipe.DefaultEyebrow;
            Keywords = recipe.Keywords;

            PowerPointDesignDirectionSummary[] directions =
                new PowerPointDesignDirectionSummary[recipe.Directions.Count];
            for (int i = 0; i < recipe.Directions.Count; i++) {
                directions[i] = new PowerPointDesignDirectionSummary(i, recipe.Directions[i]);
            }

            Directions = directions;
        }

        /// <summary>
        ///     Zero-based recipe index within the built-in recipe list.
        /// </summary>
        public int Index { get; }

        /// <summary>
        ///     Recipe display name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///     Optional short explanation of where the recipe fits.
        /// </summary>
        public string? Description { get; }

        /// <summary>
        ///     Optional eyebrow applied when the caller does not supply one.
        /// </summary>
        public string? DefaultEyebrow { get; }

        /// <summary>
        ///     Purpose keywords used for recipe matching.
        /// </summary>
        public IReadOnlyList<string> Keywords { get; }

        /// <summary>
        ///     Creative directions included in the recipe.
        /// </summary>
        public IReadOnlyList<PowerPointDesignDirectionSummary> Directions { get; }

        /// <summary>
        ///     Number of creative directions included in the recipe.
        /// </summary>
        public int DirectionCount => Directions.Count;

        /// <inheritdoc />
        public override string ToString() {
            return Index + ": " + Name + " (" + DirectionCount + " directions)";
        }
    }

    /// <summary>
    ///     Lightweight description of one creative direction inside a design recipe.
    /// </summary>
    public sealed class PowerPointDesignDirectionSummary {
        internal PowerPointDesignDirectionSummary(int index, PowerPointDesignDirection direction) {
            Index = index;
            Name = direction.Name;
            Mood = direction.Mood;
            Density = direction.Density;
            VisualStyle = direction.VisualStyle;
            HeadingFontName = direction.HeadingFontName;
            BodyFontName = direction.BodyFontName;
            ShowsDirectionMotif = direction.ShowDirectionMotif;
        }

        /// <summary>
        ///     Zero-based direction index within the recipe.
        /// </summary>
        public int Index { get; }

        /// <summary>
        ///     Creative direction name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///     Broad deck mood.
        /// </summary>
        public PowerPointDesignMood Mood { get; }

        /// <summary>
        ///     Preferred content density.
        /// </summary>
        public PowerPointSlideDensity Density { get; }

        /// <summary>
        ///     Preferred primitive visual style.
        /// </summary>
        public PowerPointVisualStyle VisualStyle { get; }

        /// <summary>
        ///     Heading font used by this direction.
        /// </summary>
        public string HeadingFontName { get; }

        /// <summary>
        ///     Body font used by this direction.
        /// </summary>
        public string BodyFontName { get; }

        /// <summary>
        ///     Whether this direction uses repeated direction markers by default.
        /// </summary>
        public bool ShowsDirectionMotif { get; }

        /// <inheritdoc />
        public override string ToString() {
            return Index + ": " + Name + " (" + Mood + ", " + VisualStyle + ", " +
                   HeadingFontName + "/" + BodyFontName + ")";
        }
    }
}
