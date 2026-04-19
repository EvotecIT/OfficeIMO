using System;
using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Deck-level design facade for applying a consistent but variable visual direction to many slides.
    /// </summary>
    public sealed class PowerPointDeckDesign {
        private PowerPointDeckDesign(PowerPointDesignTheme theme, PowerPointDesignIntent baseIntent,
            PowerPointDesignDirection direction, string seed, string? eyebrow, string? footerLeft,
            string? footerRight, bool showDirectionMotif) {
            Theme = theme;
            BaseIntent = baseIntent;
            Direction = direction;
            Seed = seed;
            Eyebrow = eyebrow;
            FooterLeft = footerLeft;
            FooterRight = footerRight;
            ShowDirectionMotif = showDirectionMotif;
        }

        /// <summary>
        ///     Creates a deck design from a brand accent, deck seed, and mood.
        /// </summary>
        public static PowerPointDeckDesign FromBrand(string accentColor, string seed,
            PowerPointDesignMood mood = PowerPointDesignMood.Corporate, string? name = null,
            string? eyebrow = null, string? footerLeft = null, string? footerRight = null,
            string headingFontName = "Poppins", string bodyFontName = "Lato") {
            if (string.IsNullOrWhiteSpace(seed)) {
                throw new ArgumentException("Deck design seed cannot be null or empty.", nameof(seed));
            }

            PowerPointDesignTheme theme = PowerPointDesignTheme
                .FromBrand(accentColor, name, headingFontName, bodyFontName)
                .WithVariation(seed)
                .WithMood(mood);

            PowerPointDesignIntent intent = PowerPointDesignIntent.FromMood(mood, seed);
            PowerPointDesignDirection direction = new(mood.ToString(), mood, intent.Density, intent.VisualStyle,
                theme.HeadingFontName, theme.BodyFontName, mood != PowerPointDesignMood.Minimal);

            return new PowerPointDeckDesign(theme, intent, direction, seed, eyebrow, footerLeft, footerRight,
                direction.ShowDirectionMotif);
        }

        /// <summary>
        ///     Creates a deck design from a brand accent, deck seed, and named creative direction.
        /// </summary>
        public static PowerPointDeckDesign FromBrand(string accentColor, string seed,
            PowerPointDesignDirection direction, string? name = null, string? eyebrow = null,
            string? footerLeft = null, string? footerRight = null) {
            if (direction == null) {
                throw new ArgumentNullException(nameof(direction));
            }
            if (string.IsNullOrWhiteSpace(seed)) {
                throw new ArgumentException("Deck design seed cannot be null or empty.", nameof(seed));
            }

            PowerPointDesignTheme theme = PowerPointDesignTheme
                .FromBrand(accentColor, name, direction.HeadingFontName, direction.BodyFontName)
                .WithVariation(seed)
                .WithMood(direction.Mood);
            theme.HeadingFontName = direction.HeadingFontName;
            theme.BodyFontName = direction.BodyFontName;
            theme.Validate();

            PowerPointDesignIntent intent = PowerPointDesignIntent.FromMood(direction.Mood, seed);
            intent.Density = direction.Density;
            intent.VisualStyle = direction.VisualStyle;

            return new PowerPointDeckDesign(theme, intent, direction, seed, eyebrow, footerLeft, footerRight,
                direction.ShowDirectionMotif);
        }

        /// <summary>
        ///     Creates several deterministic design directions from the same brand accent and content seed.
        /// </summary>
        public static IReadOnlyList<PowerPointDeckDesign> CreateAlternativesFromBrand(string accentColor, string seed,
            int count = 3, string? name = null, string? eyebrow = null, string? footerLeft = null,
            string? footerRight = null, string headingFontName = "Poppins", string bodyFontName = "Lato") {
            if (string.IsNullOrWhiteSpace(seed)) {
                throw new ArgumentException("Deck design seed cannot be null or empty.", nameof(seed));
            }
            if (count <= 0) {
                throw new ArgumentOutOfRangeException(nameof(count), "At least one design alternative is required.");
            }

            PowerPointDeckDesign[] designs = new PowerPointDeckDesign[count];
            IReadOnlyList<PowerPointDesignDirection> directions = PowerPointDesignDirection.BuiltIn;
            for (int i = 0; i < count; i++) {
                PowerPointDesignDirection direction = directions[i % directions.Count];
                string candidateSeed = seed + "/direction-" + (i + 1);
                string candidateName = string.IsNullOrWhiteSpace(name)
                    ? direction.Name + " Direction"
                    : name + " " + direction.Name;
                string candidateHeadingFont = headingFontName == "Poppins" && bodyFontName == "Lato"
                    ? direction.HeadingFontName
                    : headingFontName;
                string candidateBodyFont = headingFontName == "Poppins" && bodyFontName == "Lato"
                    ? direction.BodyFontName
                    : bodyFontName;
                PowerPointDesignDirection candidateDirection = new(direction.Name, direction.Mood,
                    direction.Density, direction.VisualStyle, candidateHeadingFont, candidateBodyFont,
                    direction.ShowDirectionMotif);
                designs[i] = FromBrand(accentColor, candidateSeed, candidateDirection, candidateName,
                    eyebrow, footerLeft, footerRight);
            }

            return designs;
        }

        /// <summary>
        ///     Creates deterministic design alternatives from a curated scenario recipe.
        /// </summary>
        public static IReadOnlyList<PowerPointDeckDesign> CreateAlternativesFromBrand(string accentColor, string seed,
            PowerPointDesignRecipe recipe, int count = 0, string? name = null, string? eyebrow = null,
            string? footerLeft = null, string? footerRight = null, string? headingFontName = null,
            string? bodyFontName = null) {
            if (string.IsNullOrWhiteSpace(seed)) {
                throw new ArgumentException("Deck design seed cannot be null or empty.", nameof(seed));
            }
            if (recipe == null) {
                throw new ArgumentNullException(nameof(recipe));
            }
            if (count < 0) {
                throw new ArgumentOutOfRangeException(nameof(count), "Design alternative count cannot be negative.");
            }

            int designCount = count == 0 ? recipe.Directions.Count : count;
            PowerPointDeckDesign[] designs = new PowerPointDeckDesign[designCount];
            string recipeSeed = NormalizeSeedPart(recipe.Name);
            string? resolvedEyebrow = string.IsNullOrWhiteSpace(eyebrow) ? recipe.DefaultEyebrow : eyebrow;

            for (int i = 0; i < designCount; i++) {
                PowerPointDesignDirection direction = recipe.DirectionAt(i);
                string candidateSeed = seed + "/" + recipeSeed + "-" + NormalizeSeedPart(direction.Name) + "-" + (i + 1);
                string candidateName = string.IsNullOrWhiteSpace(name)
                    ? recipe.Name + " " + direction.Name
                    : name + " " + direction.Name;
                string candidateHeadingFont = string.IsNullOrWhiteSpace(headingFontName)
                    ? direction.HeadingFontName
                    : headingFontName!;
                string candidateBodyFont = string.IsNullOrWhiteSpace(bodyFontName)
                    ? direction.BodyFontName
                    : bodyFontName!;
                PowerPointDesignDirection candidateDirection = new(direction.Name, direction.Mood,
                    direction.Density, direction.VisualStyle, candidateHeadingFont, candidateBodyFont,
                    direction.ShowDirectionMotif);
                designs[i] = FromBrand(accentColor, candidateSeed, candidateDirection, candidateName,
                    resolvedEyebrow, footerLeft, footerRight);
            }

            return designs;
        }

        /// <summary>
        ///     Creates deterministic design alternatives from caller-supplied creative directions.
        /// </summary>
        public static IReadOnlyList<PowerPointDeckDesign> CreateAlternativesFromBrand(string accentColor, string seed,
            IEnumerable<PowerPointDesignDirection> directions, string? name = null, string? eyebrow = null,
            string? footerLeft = null, string? footerRight = null, string? headingFontName = null,
            string? bodyFontName = null) {
            if (string.IsNullOrWhiteSpace(seed)) {
                throw new ArgumentException("Deck design seed cannot be null or empty.", nameof(seed));
            }
            if (directions == null) {
                throw new ArgumentNullException(nameof(directions));
            }

            List<PowerPointDesignDirection> directionList = new();
            foreach (PowerPointDesignDirection direction in directions) {
                if (direction == null) {
                    throw new ArgumentException("Design direction list cannot contain null entries.", nameof(directions));
                }
                directionList.Add(direction);
            }

            if (directionList.Count == 0) {
                throw new ArgumentException("At least one design direction is required.", nameof(directions));
            }

            PowerPointDeckDesign[] designs = new PowerPointDeckDesign[directionList.Count];
            for (int i = 0; i < directionList.Count; i++) {
                PowerPointDesignDirection direction = directionList[i];
                string candidateSeed = seed + "/" + NormalizeSeedPart(direction.Name) + "-" + (i + 1);
                string candidateName = string.IsNullOrWhiteSpace(name)
                    ? direction.Name + " Direction"
                    : name + " " + direction.Name;
                string candidateHeadingFont = string.IsNullOrWhiteSpace(headingFontName)
                    ? direction.HeadingFontName
                    : headingFontName!;
                string candidateBodyFont = string.IsNullOrWhiteSpace(bodyFontName)
                    ? direction.BodyFontName
                    : bodyFontName!;
                PowerPointDesignDirection candidateDirection = new(direction.Name, direction.Mood,
                    direction.Density, direction.VisualStyle, candidateHeadingFont, candidateBodyFont,
                    direction.ShowDirectionMotif);
                designs[i] = FromBrand(accentColor, candidateSeed, candidateDirection, candidateName,
                    eyebrow, footerLeft, footerRight);
            }

            return designs;
        }

        /// <summary>
        ///     Theme used by the deck.
        /// </summary>
        public PowerPointDesignTheme Theme { get; }

        /// <summary>
        ///     Base intent used to derive per-slide deterministic variants.
        /// </summary>
        public PowerPointDesignIntent BaseIntent { get; }

        /// <summary>
        ///     Creative direction used by this deck design.
        /// </summary>
        public PowerPointDesignDirection Direction { get; }

        /// <summary>
        ///     Stable deck seed.
        /// </summary>
        public string Seed { get; }

        /// <summary>
        ///     Optional default eyebrow text for generated slides.
        /// </summary>
        public string? Eyebrow { get; set; }

        /// <summary>
        ///     Optional default left footer text.
        /// </summary>
        public string? FooterLeft { get; set; }

        /// <summary>
        ///     Optional default right footer text.
        /// </summary>
        public string? FooterRight { get; set; }

        /// <summary>
        ///     Default direction motif behavior for generated slides.
        /// </summary>
        public bool ShowDirectionMotif { get; set; }

        /// <summary>
        ///     Applies the deck theme to the presentation.
        /// </summary>
        public PowerPointPresentation ApplyTo(PowerPointPresentation presentation) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }

            return presentation.ApplyDesignerTheme(Theme);
        }

        /// <summary>
        ///     Creates a per-slide intent while preserving the deck's mood, density, and visual style.
        /// </summary>
        public PowerPointDesignIntent IntentFor(string slideSeed) {
            if (string.IsNullOrWhiteSpace(slideSeed)) {
                throw new ArgumentException("Slide seed cannot be null or empty.", nameof(slideSeed));
            }

            PowerPointDesignIntent intent = BaseIntent.Clone();
            intent.Seed = Seed + "/" + slideSeed;
            return intent;
        }

        /// <summary>
        ///     Applies deck chrome and a per-slide intent to an options object.
        /// </summary>
        public T Configure<T>(T options, string slideSeed) where T : PowerPointDesignerSlideOptions {
            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            options.DesignIntent = IntentFor(slideSeed);
            options.ShowDirectionMotif = options.ShowDirectionMotif && ShowDirectionMotif;

            if (string.IsNullOrWhiteSpace(options.Eyebrow)) {
                options.Eyebrow = Eyebrow;
            }
            if (string.IsNullOrWhiteSpace(options.FooterLeft) || options.FooterLeft == "OfficeIMO") {
                options.FooterLeft = FooterLeft;
            }
            if (string.IsNullOrWhiteSpace(options.FooterRight)) {
                options.FooterRight = FooterRight;
            }

            return options;
        }

        /// <summary>
        ///     Creates default designer slide options for the supplied slide seed.
        /// </summary>
        public PowerPointDesignerSlideOptions Options(string slideSeed) {
            return Configure(new PowerPointDesignerSlideOptions(), slideSeed);
        }

        private static string NormalizeSeedPart(string value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return "direction";
            }

            char[] chars = value.Trim().ToLowerInvariant().ToCharArray();
            for (int i = 0; i < chars.Length; i++) {
                char c = chars[i];
                if (!char.IsLetterOrDigit(c)) {
                    chars[i] = '-';
                }
            }

            string normalized = new(chars);
            while (normalized.IndexOf("--", StringComparison.Ordinal) >= 0) {
                normalized = normalized.Replace("--", "-");
            }

            return normalized.Trim('-');
        }
    }
}
