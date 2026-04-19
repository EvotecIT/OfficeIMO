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
        ///     Optional secondary accent override applied after recipe variations.
        /// </summary>
        public string? SecondaryAccentColor { get; private set; }

        /// <summary>
        ///     Optional tertiary accent override applied after recipe variations.
        /// </summary>
        public string? TertiaryAccentColor { get; private set; }

        /// <summary>
        ///     Optional warm accent override applied after recipe variations.
        /// </summary>
        public string? WarmAccentColor { get; private set; }

        /// <summary>
        ///     Optional surface color override applied after recipe variations.
        /// </summary>
        public string? SurfaceColor { get; private set; }

        /// <summary>
        ///     Optional panel border color override applied after recipe variations.
        /// </summary>
        public string? PanelBorderColor { get; private set; }

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
        ///     Applies optional supporting palette overrides while preserving the primary brand accent.
        /// </summary>
        public PowerPointDesignBrief WithPalette(string? secondaryAccentColor = null,
            string? tertiaryAccentColor = null, string? warmAccentColor = null, string? surfaceColor = null,
            string? panelBorderColor = null) {
            SecondaryAccentColor = NormalizeOptionalColor(secondaryAccentColor, nameof(secondaryAccentColor));
            TertiaryAccentColor = NormalizeOptionalColor(tertiaryAccentColor, nameof(tertiaryAccentColor));
            WarmAccentColor = NormalizeOptionalColor(warmAccentColor, nameof(warmAccentColor));
            SurfaceColor = NormalizeOptionalColor(surfaceColor, nameof(surfaceColor));
            PanelBorderColor = NormalizeOptionalColor(panelBorderColor, nameof(panelBorderColor));
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
                return ApplyPaletteOverrides(CreateDirectionAlternatives(count));
            }

            PowerPointDesignRecipe recipe = Recipe
                ?? (!string.IsNullOrWhiteSpace(Purpose) ? PowerPointDesignRecipe.FindBuiltIn(Purpose!) : null)
                ?? PowerPointDesignRecipe.ConsultingPortfolio;

            return ApplyPaletteOverrides(recipe.CreateAlternativesFromBrand(AccentColor, Seed, count, Name,
                Eyebrow, FooterLeft, FooterRight, HeadingFontName, BodyFontName));
        }

        /// <summary>
        ///     Creates lightweight descriptions of generated alternatives without requiring callers to inspect the deck objects.
        /// </summary>
        public IReadOnlyList<PowerPointDeckDesignSummary> DescribeAlternatives(int count = 0) {
            IReadOnlyList<PowerPointDeckDesign> alternatives = CreateAlternatives(count);
            PowerPointDeckDesignSummary[] summaries = new PowerPointDeckDesignSummary[alternatives.Count];
            for (int i = 0; i < alternatives.Count; i++) {
                summaries[i] = alternatives[i].Describe(i);
            }

            return summaries;
        }

        /// <summary>
        ///     Creates lightweight descriptions of how a deck plan would resolve under one design alternative.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> DescribeDeckPlan(
            PowerPointDeckPlan plan, int alternativeIndex = 0) {
            if (plan == null) {
                throw new ArgumentNullException(nameof(plan));
            }

            return plan.DescribeSlides(CreateDesign(alternativeIndex));
        }

        /// <summary>
        ///     Creates lightweight descriptions of how a deck plan would resolve across several design alternatives.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanAlternativeSummary> DescribeDeckPlanAlternatives(
            PowerPointDeckPlan plan, int count = 0) {
            if (plan == null) {
                throw new ArgumentNullException(nameof(plan));
            }

            IReadOnlyList<PowerPointDeckDesign> alternatives = CreateAlternatives(count);
            IReadOnlyList<PowerPointDeckPlanDiagnostic> diagnostics = plan.ValidateSlides();
            PowerPointDeckPlanAlternativeSummary[] summaries =
                new PowerPointDeckPlanAlternativeSummary[alternatives.Count];

            for (int i = 0; i < alternatives.Count; i++) {
                PowerPointDeckDesign design = alternatives[i];
                summaries[i] = new PowerPointDeckPlanAlternativeSummary(
                    i,
                    design.Describe(i),
                    plan.DescribeSlides(design),
                    diagnostics);
            }

            return summaries;
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

        private IReadOnlyList<PowerPointDeckDesign> ApplyPaletteOverrides(
            IReadOnlyList<PowerPointDeckDesign> alternatives) {
            if (SecondaryAccentColor == null && TertiaryAccentColor == null && WarmAccentColor == null &&
                SurfaceColor == null && PanelBorderColor == null) {
                return alternatives;
            }

            foreach (PowerPointDeckDesign design in alternatives) {
                if (SecondaryAccentColor != null) {
                    design.Theme.Accent2Color = SecondaryAccentColor;
                }
                if (TertiaryAccentColor != null) {
                    design.Theme.Accent3Color = TertiaryAccentColor;
                }
                if (WarmAccentColor != null) {
                    design.Theme.WarningColor = WarmAccentColor;
                }
                if (SurfaceColor != null) {
                    design.Theme.SurfaceColor = SurfaceColor;
                }
                if (PanelBorderColor != null) {
                    design.Theme.PanelBorderColor = PanelBorderColor;
                }

                design.Theme.Validate();
            }

            return alternatives;
        }

        private static string? NormalizeOptionalColor(string? value, string name) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string color = value!.Trim();
            if (color.StartsWith("#", StringComparison.Ordinal)) {
                color = color.Substring(1);
            }
            if (color.Length != 6) {
                throw new ArgumentException("Color must be a six-character RGB hex value.", name);
            }

            for (int i = 0; i < color.Length; i++) {
                char c = color[i];
                bool valid = c is >= '0' and <= '9' or >= 'A' and <= 'F' or >= 'a' and <= 'f';
                if (!valid) {
                    throw new ArgumentException("Color must be a six-character RGB hex value.", name);
                }
            }

            return color.ToUpperInvariant();
        }
    }
}
