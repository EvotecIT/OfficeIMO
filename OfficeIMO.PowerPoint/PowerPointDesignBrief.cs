using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Small deck design brief that can resolve a brand, purpose, identity, and optional custom directions
    ///     into one or more deterministic deck designs.
    /// </summary>
    public sealed class PowerPointDesignBrief {
        private readonly List<PowerPointDesignDirection> _directions = new();
        private readonly List<PowerPointDesignMood> _preferredMoods = new();
        private readonly List<PowerPointSlideDensity> _preferredDensities = new();
        private readonly List<PowerPointVisualStyle> _preferredVisualStyles = new();
        private static readonly PowerPointCreativeDirectionPack[] BuiltInCreativeDirectionPacks = {
            PowerPointCreativeDirectionPack.Boardroom,
            PowerPointCreativeDirectionPack.FieldProof,
            PowerPointCreativeDirectionPack.EditorialCaseStudy,
            PowerPointCreativeDirectionPack.TechnicalMap,
            PowerPointCreativeDirectionPack.QuietAppendix
        };

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
        ///     Describes the built-in creative direction packs that can be applied to a design brief.
        /// </summary>
        public static IReadOnlyList<PowerPointCreativeDirectionPackSummary> DescribeCreativeDirectionPacks() {
            PowerPointCreativeDirectionPackSummary[] summaries =
                new PowerPointCreativeDirectionPackSummary[BuiltInCreativeDirectionPacks.Length];
            for (int i = 0; i < BuiltInCreativeDirectionPacks.Length; i++) {
                summaries[i] = DescribeCreativeDirectionPack(BuiltInCreativeDirectionPacks[i], i);
            }

            return summaries;
        }

        /// <summary>
        ///     Describes one creative direction pack before applying it to a design brief.
        /// </summary>
        public static PowerPointCreativeDirectionPackSummary DescribeCreativeDirectionPack(
            PowerPointCreativeDirectionPack pack) {
            int index = pack == PowerPointCreativeDirectionPack.Auto
                ? -1
                : Array.IndexOf(BuiltInCreativeDirectionPacks, pack);

            return DescribeCreativeDirectionPack(pack, index);
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
        ///     Optional supporting palette strategy applied before explicit palette color overrides.
        /// </summary>
        public PowerPointPaletteStyle? PaletteStyle { get; private set; }

        /// <summary>
        ///     Optional typography strategy applied before explicit font overrides.
        /// </summary>
        public PowerPointTypographyStyle? TypographyStyle { get; private set; }

        /// <summary>
        ///     Optional Auto layout strategy applied to generated design alternatives.
        /// </summary>
        public PowerPointAutoLayoutStrategy? LayoutStrategy { get; private set; }

        /// <summary>
        ///     Optional high-level creative pack used to configure recipe, palette, layout strategy, and ranking preferences.
        /// </summary>
        public PowerPointCreativeDirectionPack CreativeDirectionPack { get; private set; } =
            PowerPointCreativeDirectionPack.Auto;

        /// <summary>
        ///     Controls how far generated alternatives should move from the selected recipe or preferred direction.
        /// </summary>
        public PowerPointDesignVariety Variety { get; private set; } = PowerPointDesignVariety.Balanced;

        /// <summary>
        ///     Caller-supplied creative directions. When present, these take precedence over recipes.
        /// </summary>
        public IReadOnlyList<PowerPointDesignDirection> Directions => _directions;

        /// <summary>
        ///     Preferred moods used to rank recipe or custom directions before alternatives are created.
        /// </summary>
        public IReadOnlyList<PowerPointDesignMood> PreferredMoods => _preferredMoods;

        /// <summary>
        ///     Preferred slide densities used to rank recipe or custom directions before alternatives are created.
        /// </summary>
        public IReadOnlyList<PowerPointSlideDensity> PreferredDensities => _preferredDensities;

        /// <summary>
        ///     Preferred visual styles used to rank recipe or custom directions before alternatives are created.
        /// </summary>
        public IReadOnlyList<PowerPointVisualStyle> PreferredVisualStyles => _preferredVisualStyles;

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
        ///     Chooses a supporting palette strategy while preserving the primary brand accent.
        /// </summary>
        public PowerPointDesignBrief WithPaletteStyle(PowerPointPaletteStyle paletteStyle) {
            PaletteStyle = paletteStyle;
            return this;
        }

        /// <summary>
        ///     Chooses a typography strategy while preserving the selected recipe or creative direction.
        /// </summary>
        public PowerPointDesignBrief WithTypographyStyle(PowerPointTypographyStyle typographyStyle) {
            TypographyStyle = typographyStyle;
            return this;
        }

        /// <summary>
        ///     Chooses how Auto slide variants should balance content fit, design variety, and compactness.
        /// </summary>
        public PowerPointDesignBrief WithLayoutStrategy(PowerPointAutoLayoutStrategy layoutStrategy) {
            LayoutStrategy = layoutStrategy;
            return this;
        }

        /// <summary>
        ///     Applies a curated creative starting point while preserving seed-based variation and any later explicit overrides.
        /// </summary>
        public PowerPointDesignBrief WithCreativeDirectionPack(PowerPointCreativeDirectionPack pack) {
            CreativeDirectionPack = pack;

            if (pack == PowerPointCreativeDirectionPack.Auto) {
                Recipe = null;
                PaletteStyle = null;
                TypographyStyle = null;
                LayoutStrategy = null;
                Variety = PowerPointDesignVariety.Balanced;
                ClearDesignPreferences();
                return this;
            }

            PowerPointCreativeDirectionPackSummary summary = DescribeCreativeDirectionPack(pack);
            Recipe = summary.Recipe;
            PaletteStyle = summary.PaletteStyle;
            LayoutStrategy = summary.LayoutStrategy;
            Variety = summary.Variety;
            SetDirectionPreferences(summary.PreferredMoods, summary.PreferredDensities,
                summary.PreferredVisualStyles);

            return this;
        }

        /// <summary>
        ///     Sets how broad generated alternatives should be.
        /// </summary>
        public PowerPointDesignBrief WithVariety(PowerPointDesignVariety variety) {
            Variety = variety;
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
        ///     Prefers one or more moods when ordering recipe or custom directions.
        /// </summary>
        public PowerPointDesignBrief WithPreferredMoods(params PowerPointDesignMood[] moods) {
            ReplacePreferences(_preferredMoods, moods, nameof(moods));
            return this;
        }

        /// <summary>
        ///     Prefers one or more slide densities when ordering recipe or custom directions.
        /// </summary>
        public PowerPointDesignBrief WithPreferredDensities(params PowerPointSlideDensity[] densities) {
            ReplacePreferences(_preferredDensities, densities, nameof(densities));
            return this;
        }

        /// <summary>
        ///     Prefers one or more visual styles when ordering recipe or custom directions.
        /// </summary>
        public PowerPointDesignBrief WithPreferredVisualStyles(params PowerPointVisualStyle[] visualStyles) {
            ReplacePreferences(_preferredVisualStyles, visualStyles, nameof(visualStyles));
            return this;
        }

        /// <summary>
        ///     Clears direction ordering preferences.
        /// </summary>
        public PowerPointDesignBrief ClearDesignPreferences() {
            _preferredMoods.Clear();
            _preferredDensities.Clear();
            _preferredVisualStyles.Clear();
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
                return ApplyBriefOverrides(CreateDirectionAlternatives(count));
            }

            PowerPointDesignRecipe recipe = Recipe
                ?? (!string.IsNullOrWhiteSpace(Purpose) ? PowerPointDesignRecipe.FindBuiltIn(Purpose!) : null)
                ?? PowerPointDesignRecipe.ConsultingPortfolio;
            if (HasDirectionPreferences || Variety != PowerPointDesignVariety.Balanced) {
                recipe = new PowerPointDesignRecipe(recipe.Name, ResolveDirectionSet(recipe.Directions, true),
                    recipe.DefaultEyebrow, recipe.Description, recipe.Keywords);
            }

            return ApplyBriefOverrides(recipe.CreateAlternativesFromBrand(AccentColor, Seed, count, Name,
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
        ///     Creates lightweight recommendations explaining why generated alternatives fit this brief.
        /// </summary>
        public IReadOnlyList<PowerPointDeckDesignRecommendation> RecommendAlternatives(int count = 0) {
            IReadOnlyList<PowerPointDeckDesign> alternatives = CreateAlternatives(count);
            PowerPointDeckDesignRecommendation[] recommendations =
                new PowerPointDeckDesignRecommendation[alternatives.Count];

            for (int i = 0; i < alternatives.Count; i++) {
                recommendations[i] = RecommendAlternative(alternatives[i].Describe(i));
            }

            return recommendations;
        }

        /// <summary>
        ///     Creates lightweight descriptions of how a deck plan would resolve under one design alternative.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> DescribeDeckPlan(
            PowerPointDeckPlan plan, int alternativeIndex = 0) {
            return DescribeDeckPlan(plan, alternativeIndex, slideIndexOffset: 0);
        }

        /// <summary>
        ///     Creates lightweight descriptions of how a deck plan would resolve under one design alternative,
        ///     using an existing composer slide count for fallback seed generation.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> DescribeDeckPlan(
            PowerPointDeckPlan plan, int alternativeIndex, int slideIndexOffset) {
            if (plan == null) {
                throw new ArgumentNullException(nameof(plan));
            }

            return plan.DescribeSlides(CreateDesign(alternativeIndex), slideIndexOffset);
        }

        /// <summary>
        ///     Creates lightweight descriptions of how a deck plan would resolve across several design alternatives.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanAlternativeSummary> DescribeDeckPlanAlternatives(
            PowerPointDeckPlan plan, int count = 0) {
            return DescribeDeckPlanAlternatives(plan, count, slideIndexOffset: 0);
        }

        /// <summary>
        ///     Creates lightweight descriptions of how a deck plan would resolve across several design alternatives,
        ///     using an existing composer slide count for fallback seed generation.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanAlternativeSummary> DescribeDeckPlanAlternatives(
            PowerPointDeckPlan plan, int count, int slideIndexOffset) {
            if (plan == null) {
                throw new ArgumentNullException(nameof(plan));
            }

            IReadOnlyList<PowerPointDeckDesign> alternatives = CreateAlternatives(count);
            IReadOnlyList<PowerPointDeckPlanDiagnostic> diagnostics = plan.ValidateSlides();
            PowerPointDeckPlanAlternativeSummary[] summaries =
                new PowerPointDeckPlanAlternativeSummary[alternatives.Count];

            for (int i = 0; i < alternatives.Count; i++) {
                PowerPointDeckDesign design = alternatives[i];
                PowerPointDeckDesignSummary designSummary = design.Describe(i);
                IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> slides =
                    plan.DescribeSlides(design, slideIndexOffset);
                PowerPointDeckPlanContentFit fit = ScoreDeckPlanAlternative(designSummary, slides, diagnostics);
                summaries[i] = new PowerPointDeckPlanAlternativeSummary(
                    i,
                    Variety,
                    designSummary,
                    slides,
                    diagnostics,
                    fit.Score,
                    fit.Reasons);
            }

            return summaries;
        }

        /// <summary>
        ///     Creates content-fit recommendations for a deck plan, ordered with the strongest match first.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanAlternativeSummary> RecommendDeckPlanAlternatives(
            PowerPointDeckPlan plan, int count = 0) {
            return RecommendDeckPlanAlternatives(plan, count, slideIndexOffset: 0);
        }

        /// <summary>
        ///     Creates content-fit recommendations for a deck plan, ordered with the strongest match first,
        ///     using an existing composer slide count for fallback seed generation.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanAlternativeSummary> RecommendDeckPlanAlternatives(
            PowerPointDeckPlan plan, int count, int slideIndexOffset) {
            return OrderDeckPlanAlternatives(DescribeDeckPlanAlternatives(plan, count, slideIndexOffset));
        }

        /// <summary>
        ///     Selects the strongest content-fit recommendation for a deck plan.
        /// </summary>
        public PowerPointDeckPlanAlternativeSummary RecommendDeckPlanAlternative(PowerPointDeckPlan plan,
            int count = 0) {
            return RecommendDeckPlanAlternative(plan, count, slideIndexOffset: 0);
        }

        /// <summary>
        ///     Selects the strongest content-fit recommendation for a deck plan,
        ///     using an existing composer slide count for fallback seed generation.
        /// </summary>
        public PowerPointDeckPlanAlternativeSummary RecommendDeckPlanAlternative(PowerPointDeckPlan plan,
            int count, int slideIndexOffset) {
            IReadOnlyList<PowerPointDeckPlanAlternativeSummary> recommendations =
                RecommendDeckPlanAlternatives(plan, count, slideIndexOffset);
            return recommendations[0];
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
            IReadOnlyList<PowerPointDesignDirection> resolvedDirections = ResolveDirectionSet(_directions, false);
            int designCount = count == 0 ? resolvedDirections.Count : count;
            List<PowerPointDesignDirection> selectedDirections = new(designCount);
            for (int i = 0; i < designCount; i++) {
                selectedDirections.Add(resolvedDirections[i % resolvedDirections.Count]);
            }

            return PowerPointDeckDesign.CreateAlternativesFromBrand(AccentColor, Seed, selectedDirections, Name,
                Eyebrow, FooterLeft, FooterRight, HeadingFontName, BodyFontName);
        }

        private bool HasDirectionPreferences =>
            _preferredMoods.Count > 0 || _preferredDensities.Count > 0 || _preferredVisualStyles.Count > 0;

        private IReadOnlyList<PowerPointDesignDirection> ResolveDirectionSet(
            IEnumerable<PowerPointDesignDirection> directions, bool allowExploratoryExpansion) {
            IReadOnlyList<PowerPointDesignDirection> ranked = RankDirections(directions);
            if (Variety == PowerPointDesignVariety.Focused && ranked.Count > 0) {
                return new[] { ranked[0] };
            }

            if (Variety != PowerPointDesignVariety.Exploratory || !allowExploratoryExpansion) {
                return ranked;
            }

            List<PowerPointDesignDirection> expanded = new(ranked);
            HashSet<string> directionNames = new(StringComparer.OrdinalIgnoreCase);
            foreach (PowerPointDesignDirection direction in expanded) {
                directionNames.Add(direction.Name);
            }

            foreach (PowerPointDesignDirection direction in PowerPointDesignDirection.BuiltIn) {
                if (directionNames.Add(direction.Name)) {
                    expanded.Add(direction);
                }
            }

            return expanded.AsReadOnly();
        }

        private IReadOnlyList<PowerPointDesignDirection> RankDirections(
            IEnumerable<PowerPointDesignDirection> directions) {
            List<PowerPointDesignDirection> source = directions.ToList();

            if (!HasDirectionPreferences) {
                return source.AsReadOnly();
            }

            List<PowerPointDesignDirection> ranked = source.Where(MatchesDirectionPreferences).ToList();
            ranked.AddRange(source.Where(direction => !MatchesDirectionPreferences(direction)));

            return ranked.AsReadOnly();
        }

        private bool MatchesDirectionPreferences(PowerPointDesignDirection direction) {
            if (_preferredMoods.Count > 0 && !_preferredMoods.Contains(direction.Mood)) {
                return false;
            }
            if (_preferredDensities.Count > 0 && !_preferredDensities.Contains(direction.Density)) {
                return false;
            }
            if (_preferredVisualStyles.Count > 0 && !_preferredVisualStyles.Contains(direction.VisualStyle)) {
                return false;
            }

            return true;
        }

        private PowerPointDeckDesignRecommendation RecommendAlternative(PowerPointDeckDesignSummary design) {
            List<string> reasons = new();
            int score = 0;

            if (_preferredMoods.Contains(design.Mood)) {
                score += 3;
                reasons.Add("Matches preferred mood: " + design.Mood + ".");
            }
            if (_preferredVisualStyles.Contains(design.VisualStyle)) {
                score += 2;
                reasons.Add("Matches preferred visual style: " + design.VisualStyle + ".");
            }
            if (_preferredDensities.Contains(design.Density)) {
                score += 1;
                reasons.Add("Matches preferred density: " + design.Density + ".");
            }

            if (Variety == PowerPointDesignVariety.Focused) {
                reasons.Add("Focused variety keeps this option close to the strongest matching direction.");
            } else if (Variety == PowerPointDesignVariety.Exploratory) {
                reasons.Add("Exploratory variety allows broader directions when more visual distance is useful.");
            } else {
                reasons.Add("Balanced variety keeps the selected recipe breadth.");
            }

            if (design.ShowsDirectionMotif) {
                reasons.Add("Uses direction motifs for stronger visual rhythm.");
            } else {
                reasons.Add("Avoids repeated direction motifs for a quieter deck.");
            }

            return new PowerPointDeckDesignRecommendation(design, score, reasons.AsReadOnly());
        }

        private static PowerPointDeckPlanContentFit ScoreDeckPlanAlternative(
            PowerPointDeckDesignSummary design,
            IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> slides,
            IReadOnlyList<PowerPointDeckPlanDiagnostic> diagnostics) {
            List<string> reasons = new();
            int score = 0;

            bool hasCaseStudy = HasSlideKind(slides, PowerPointDeckPlanSlideKind.CaseStudy);
            bool hasProcess = HasSlideKind(slides, PowerPointDeckPlanSlideKind.Process);
            bool hasCardGrid = HasSlideKind(slides, PowerPointDeckPlanSlideKind.CardGrid);
            bool hasLogoWall = HasSlideKind(slides, PowerPointDeckPlanSlideKind.LogoWall);
            bool hasCoverage = HasSlideKind(slides, PowerPointDeckPlanSlideKind.Coverage);
            bool hasCapability = HasSlideKind(slides, PowerPointDeckPlanSlideKind.Capability);
            bool hasCustom = HasSlideKind(slides, PowerPointDeckPlanSlideKind.Custom);
            bool hasDenseSlide = HasDenseSlide(slides);
            bool hasWarning = HasDiagnosticSeverity(diagnostics, PowerPointDeckPlanDiagnosticSeverity.Warning);

            if ((hasProcess || hasCoverage) && design.VisualStyle == PowerPointVisualStyle.Geometric) {
                score += 2;
                reasons.Add("Geometric visual style supports process, timeline, and coverage slides.");
            }

            if ((hasCaseStudy || hasCapability) && design.Density != PowerPointSlideDensity.Compact) {
                score += 2;
                reasons.Add("Balanced or relaxed density gives narrative sections more breathing room.");
            }

            if ((hasCardGrid || hasLogoWall) && design.VisualStyle != PowerPointVisualStyle.Minimal) {
                score += 1;
                reasons.Add("Decorative visual styles give reusable cards and proof walls clearer rhythm.");
            }

            if (hasDenseSlide && design.Density == PowerPointSlideDensity.Compact) {
                score += 2;
                reasons.Add("Compact density fits denser planned slides without manual placement.");
            } else if (!hasDenseSlide && slides.Count > 0 && design.Density == PowerPointSlideDensity.Relaxed) {
                score += 1;
                reasons.Add("Relaxed density suits lighter slide plans with more whitespace.");
            }

            if (hasWarning && design.Density == PowerPointSlideDensity.Compact) {
                score += 1;
                reasons.Add("Compact density can help warning-level dense content stay inspectable before rendering.");
            }

            if (hasCustom && !design.ShowsDirectionMotif) {
                score += 1;
                reasons.Add("A quieter motif leaves raw-composition slides more neutral.");
            }

            if (reasons.Count == 0) {
                reasons.Add("Uses the selected design direction without a strong content-fit signal.");
            }

            return new PowerPointDeckPlanContentFit(score, reasons.AsReadOnly());
        }

        private static IReadOnlyList<PowerPointDeckPlanAlternativeSummary> OrderDeckPlanAlternatives(
            IEnumerable<PowerPointDeckPlanAlternativeSummary> alternatives) {
            return alternatives
                .OrderBy(alternative => alternative.HasErrors)
                .ThenByDescending(alternative => alternative.ContentFitScore)
                .ThenBy(alternative => alternative.HasWarnings)
                .ThenBy(alternative => alternative.Index)
                .ToList()
                .AsReadOnly();
        }

        private IReadOnlyList<PowerPointDeckDesign> ApplyBriefOverrides(
            IReadOnlyList<PowerPointDeckDesign> alternatives) {
            if (LayoutStrategy == null && PaletteStyle == null && TypographyStyle == null &&
                SecondaryAccentColor == null &&
                TertiaryAccentColor == null && WarmAccentColor == null && SurfaceColor == null &&
                PanelBorderColor == null && HeadingFontName == null && BodyFontName == null) {
                return alternatives;
            }

            foreach (PowerPointDeckDesign design in alternatives) {
                if (LayoutStrategy != null) {
                    design.BaseIntent.LayoutStrategy = LayoutStrategy.Value;
                }
                if (PaletteStyle != null) {
                    design.Theme.ApplyPaletteStyle(PaletteStyle.Value, design.Seed);
                }
                if (TypographyStyle != null) {
                    design.Theme.ApplyTypographyStyle(TypographyStyle.Value, design.Seed);
                }
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
                if (!string.IsNullOrWhiteSpace(HeadingFontName)) {
                    design.Theme.HeadingFontName = HeadingFontName!;
                }
                if (!string.IsNullOrWhiteSpace(BodyFontName)) {
                    design.Theme.BodyFontName = BodyFontName!;
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

        private static bool HasSlideKind(IEnumerable<PowerPointDeckPlanSlideRenderSummary> slides,
            PowerPointDeckPlanSlideKind kind) {
            return slides.Any(slide => slide.Kind == kind);
        }

        private static bool HasDenseSlide(IEnumerable<PowerPointDeckPlanSlideRenderSummary> slides) {
            return slides.Any(slide => slide.ContentItemCount > PowerPointDeckPlanLimits.DenseProcessSteps);
        }

        private static bool HasDiagnosticSeverity(IEnumerable<PowerPointDeckPlanDiagnostic> diagnostics,
            PowerPointDeckPlanDiagnosticSeverity severity) {
            return diagnostics.Any(diagnostic => diagnostic.Severity == severity);
        }

        private static void ReplacePreferences<T>(List<T> target, IEnumerable<T> values, string name) {
            if (values == null) {
                throw new ArgumentNullException(name);
            }

            target.Clear();
            target.AddRange(values.Distinct());
        }

        private static PowerPointCreativeDirectionPackSummary DescribeCreativeDirectionPack(
            PowerPointCreativeDirectionPack pack, int index) {
            return pack switch {
                PowerPointCreativeDirectionPack.Auto => new PowerPointCreativeDirectionPackSummary(
                    index,
                    PowerPointCreativeDirectionPack.Auto,
                    "Auto",
                    "Clear curated pack settings and let purpose matching or later explicit brief settings choose the design.",
                    null,
                    null,
                    null,
                    PowerPointDesignVariety.Balanced,
                    Array.Empty<PowerPointDesignMood>(),
                    Array.Empty<PowerPointSlideDensity>(),
                    Array.Empty<PowerPointVisualStyle>()),
                PowerPointCreativeDirectionPack.Boardroom => new PowerPointCreativeDirectionPackSummary(
                    index,
                    PowerPointCreativeDirectionPack.Boardroom,
                    "Boardroom",
                    "Restrained, board-ready hierarchy for executive and decision decks.",
                    PowerPointDesignRecipe.ExecutiveBrief,
                    PowerPointPaletteStyle.CoolNeutral,
                    PowerPointAutoLayoutStrategy.ContentFirst,
                    PowerPointDesignVariety.Balanced,
                    new[] { PowerPointDesignMood.Corporate },
                    new[] { PowerPointSlideDensity.Balanced, PowerPointSlideDensity.Relaxed },
                    new[] { PowerPointVisualStyle.Soft }),
                PowerPointCreativeDirectionPack.FieldProof => new PowerPointCreativeDirectionPackSummary(
                    index,
                    PowerPointCreativeDirectionPack.FieldProof,
                    "Field Proof",
                    "Visual proof and stronger contrast for case studies, portfolios, and service stories.",
                    PowerPointDesignRecipe.ConsultingPortfolio,
                    PowerPointPaletteStyle.SplitComplementary,
                    PowerPointAutoLayoutStrategy.VisualFirst,
                    PowerPointDesignVariety.Exploratory,
                    new[] { PowerPointDesignMood.Energetic },
                    new[] { PowerPointSlideDensity.Balanced },
                    new[] { PowerPointVisualStyle.Geometric }),
                PowerPointCreativeDirectionPack.EditorialCaseStudy => new PowerPointCreativeDirectionPackSummary(
                    index,
                    PowerPointCreativeDirectionPack.EditorialCaseStudy,
                    "Editorial Case Study",
                    "Editorial spacing and softer surfaces for narrative-heavy customer stories.",
                    PowerPointDesignRecipe.ConsultingPortfolio,
                    PowerPointPaletteStyle.WarmNeutral,
                    PowerPointAutoLayoutStrategy.VisualFirst,
                    PowerPointDesignVariety.Balanced,
                    new[] { PowerPointDesignMood.Corporate, PowerPointDesignMood.Editorial },
                    new[] { PowerPointSlideDensity.Relaxed },
                    new[] { PowerPointVisualStyle.Soft }),
                PowerPointCreativeDirectionPack.TechnicalMap => new PowerPointCreativeDirectionPackSummary(
                    index,
                    PowerPointCreativeDirectionPack.TechnicalMap,
                    "Technical Map",
                    "Compact geometric structure for architecture, rollout, and operational decks.",
                    PowerPointDesignRecipe.TechnicalProposal,
                    PowerPointPaletteStyle.Complementary,
                    PowerPointAutoLayoutStrategy.Compact,
                    PowerPointDesignVariety.Exploratory,
                    new[] { PowerPointDesignMood.Corporate, PowerPointDesignMood.Energetic },
                    new[] { PowerPointSlideDensity.Balanced, PowerPointSlideDensity.Compact },
                    new[] { PowerPointVisualStyle.Geometric }),
                PowerPointCreativeDirectionPack.QuietAppendix => new PowerPointCreativeDirectionPackSummary(
                    index,
                    PowerPointCreativeDirectionPack.QuietAppendix,
                    "Quiet Appendix",
                    "Quiet, dense appendix treatment for supporting detail and reference slides.",
                    PowerPointDesignRecipe.TechnicalProposal,
                    PowerPointPaletteStyle.Monochrome,
                    PowerPointAutoLayoutStrategy.ContentFirst,
                    PowerPointDesignVariety.Focused,
                    new[] { PowerPointDesignMood.Minimal },
                    new[] { PowerPointSlideDensity.Compact },
                    new[] { PowerPointVisualStyle.Minimal }),
                _ => throw new ArgumentOutOfRangeException(nameof(pack), pack, "Unknown creative direction pack.")
            };
        }

        private void SetDirectionPreferences(IEnumerable<PowerPointDesignMood> moods,
            IEnumerable<PowerPointSlideDensity> densities, IEnumerable<PowerPointVisualStyle> visualStyles) {
            ReplacePreferences(_preferredMoods, moods, nameof(moods));
            ReplacePreferences(_preferredDensities, densities, nameof(densities));
            ReplacePreferences(_preferredVisualStyles, visualStyles, nameof(visualStyles));
        }

        private sealed class PowerPointDeckPlanContentFit {
            internal PowerPointDeckPlanContentFit(int score, IReadOnlyList<string> reasons) {
                Score = score;
                Reasons = reasons;
            }

            internal int Score { get; }

            internal IReadOnlyList<string> Reasons { get; }
        }
    }
}
