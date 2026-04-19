using System;
using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Small deck design brief that can resolve a brand, purpose, identity, and optional custom directions
    ///     into one or more deterministic deck designs.
    /// </summary>
    public sealed class PowerPointDesignBrief {
        private readonly List<PowerPointDesignDirection> _directions = new();

        /// <summary>
        ///     Creates a brief from a brand accent and stable seed.
        /// </summary>
        public static PowerPointDesignBrief FromBrand(string accentColor, string seed, string? purpose = null) {
            PowerPointDesignBrief brief = new(accentColor, seed);
            if (!string.IsNullOrWhiteSpace(purpose)) {
                brief.ForPurpose(purpose!);
            }

            return brief;
        }

        /// <summary>
        ///     Creates a brief from a brand accent and stable seed.
        /// </summary>
        public PowerPointDesignBrief(string accentColor, string seed) {
            if (string.IsNullOrWhiteSpace(accentColor)) {
                throw new ArgumentException("Brand accent cannot be null or empty.", nameof(accentColor));
            }
            if (string.IsNullOrWhiteSpace(seed)) {
                throw new ArgumentException("Deck design seed cannot be null or empty.", nameof(seed));
            }

            AccentColor = accentColor;
            Seed = seed;
        }

        /// <summary>
        ///     Brand accent color used to derive the deck palette.
        /// </summary>
        public string AccentColor { get; }

        /// <summary>
        ///     Stable seed used to make generated alternatives repeatable.
        /// </summary>
        public string Seed { get; }

        /// <summary>
        ///     Plain-language purpose used to select a built-in recipe when no explicit recipe is set.
        /// </summary>
        public string? Purpose { get; private set; }

        /// <summary>
        ///     Explicit recipe. When omitted, the purpose text is matched against built-in recipes.
        /// </summary>
        public PowerPointDesignRecipe? Recipe { get; private set; }

        /// <summary>
        ///     Optional display name passed into generated deck themes.
        /// </summary>
        public string? Name { get; private set; }

        /// <summary>
        ///     Optional default eyebrow for generated slides.
        /// </summary>
        public string? Eyebrow { get; private set; }

        /// <summary>
        ///     Optional left footer for generated slides.
        /// </summary>
        public string? FooterLeft { get; private set; }

        /// <summary>
        ///     Optional right footer for generated slides.
        /// </summary>
        public string? FooterRight { get; private set; }

        /// <summary>
        ///     Optional heading font override applied to every generated alternative.
        /// </summary>
        public string? HeadingFontName { get; private set; }

        /// <summary>
        ///     Optional body font override applied to every generated alternative.
        /// </summary>
        public string? BodyFontName { get; private set; }

        /// <summary>
        ///     Caller-supplied creative directions. When present, these take precedence over recipes.
        /// </summary>
        public IReadOnlyList<PowerPointDesignDirection> Directions => _directions;

        /// <summary>
        ///     Sets the plain-language deck purpose used for recipe matching.
        /// </summary>
        public PowerPointDesignBrief ForPurpose(string purpose) {
            if (string.IsNullOrWhiteSpace(purpose)) {
                throw new ArgumentException("Deck purpose cannot be null or empty.", nameof(purpose));
            }

            Purpose = purpose;
            return this;
        }

        /// <summary>
        ///     Uses an explicit recipe instead of matching one from the purpose text.
        /// </summary>
        public PowerPointDesignBrief WithRecipe(PowerPointDesignRecipe recipe) {
            Recipe = recipe ?? throw new ArgumentNullException(nameof(recipe));
            return this;
        }

        /// <summary>
        ///     Sets deck identity chrome shared by generated slides.
        /// </summary>
        public PowerPointDesignBrief WithIdentity(string? name = null, string? eyebrow = null,
            string? footerLeft = null, string? footerRight = null) {
            Name = name;
            Eyebrow = eyebrow;
            FooterLeft = footerLeft;
            FooterRight = footerRight;
            return this;
        }

        /// <summary>
        ///     Applies font overrides while preserving the rest of the selected recipe or direction.
        /// </summary>
        public PowerPointDesignBrief WithFonts(string? headingFontName = null, string? bodyFontName = null) {
            HeadingFontName = headingFontName;
            BodyFontName = bodyFontName;
            return this;
        }

        /// <summary>
        ///     Replaces recipe-based alternatives with caller-supplied creative directions.
        /// </summary>
        public PowerPointDesignBrief WithDirections(IEnumerable<PowerPointDesignDirection> directions) {
            if (directions == null) {
                throw new ArgumentNullException(nameof(directions));
            }

            _directions.Clear();
            foreach (PowerPointDesignDirection direction in directions) {
                AddDirection(direction);
            }

            if (_directions.Count == 0) {
                throw new ArgumentException("At least one design direction is required.", nameof(directions));
            }

            return this;
        }

        /// <summary>
        ///     Adds one caller-supplied creative direction.
        /// </summary>
        public PowerPointDesignBrief AddDirection(PowerPointDesignDirection direction) {
            if (direction == null) {
                throw new ArgumentException("Design direction cannot be null.", nameof(direction));
            }

            _directions.Add(direction);
            return this;
        }

        /// <summary>
        ///     Creates deterministic deck design alternatives from this brief.
        /// </summary>
        public IReadOnlyList<PowerPointDeckDesign> CreateAlternatives(int count = 0) {
            if (count < 0) {
                throw new ArgumentOutOfRangeException(nameof(count), "Design alternative count cannot be negative.");
            }

            if (_directions.Count > 0) {
                return CreateDirectionAlternatives(count);
            }

            PowerPointDesignRecipe recipe = Recipe
                ?? (!string.IsNullOrWhiteSpace(Purpose) ? PowerPointDesignRecipe.FindBuiltIn(Purpose!) : null)
                ?? PowerPointDesignRecipe.ConsultingPortfolio;

            return recipe.CreateAlternativesFromBrand(AccentColor, Seed, count, Name, Eyebrow, FooterLeft,
                FooterRight, HeadingFontName, BodyFontName);
        }

        /// <summary>
        ///     Creates one deterministic deck design from this brief.
        /// </summary>
        public PowerPointDeckDesign CreateDesign(int alternativeIndex = 0) {
            if (alternativeIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(alternativeIndex),
                    "Design alternative index cannot be negative.");
            }

            IReadOnlyList<PowerPointDeckDesign> alternatives = CreateAlternatives(alternativeIndex + 1);
            return alternatives[alternativeIndex];
        }

        private IReadOnlyList<PowerPointDeckDesign> CreateDirectionAlternatives(int count) {
            int designCount = count == 0 ? _directions.Count : count;
            List<PowerPointDesignDirection> selectedDirections = new(designCount);
            for (int i = 0; i < designCount; i++) {
                selectedDirections.Add(_directions[i % _directions.Count]);
            }

            return PowerPointDeckDesign.CreateAlternativesFromBrand(AccentColor, Seed, selectedDirections, Name,
                Eyebrow, FooterLeft, FooterRight, HeadingFontName, BodyFontName);
        }
    }
}
