using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Controls how a semantic <see cref="PowerPointDeckPlan" /> is composed into a presentation.
    /// </summary>
    public sealed class PowerPointCompositionOptions {
        private PowerPointCompositionOptions(PowerPointDeckDesign? design, PowerPointDesignBrief? brief) {
            Design = design;
            Brief = brief;
        }

        /// <summary>Creates composition options for an already resolved deck design.</summary>
        public static PowerPointCompositionOptions FromDesign(PowerPointDeckDesign design) {
            if (design == null) throw new ArgumentNullException(nameof(design));
            return new PowerPointCompositionOptions(design, null);
        }

        /// <summary>Creates composition options that resolve a design from a reusable brief.</summary>
        public static PowerPointCompositionOptions FromBrief(PowerPointDesignBrief brief) {
            if (brief == null) throw new ArgumentNullException(nameof(brief));
            return new PowerPointCompositionOptions(null, brief);
        }

        /// <summary>Resolved design to apply. Set when options were created with <see cref="FromDesign" />.</summary>
        public PowerPointDeckDesign? Design { get; }

        /// <summary>Design brief to resolve. Set when options were created with <see cref="FromBrief" />.</summary>
        public PowerPointDesignBrief? Brief { get; }

        /// <summary>
        ///     Zero-based brief alternative to use when <see cref="SelectBestAlternative" /> is false.
        /// </summary>
        public int AlternativeIndex { get; set; }

        /// <summary>
        ///     Selects the brief alternative that best fits the supplied plan. This is the default for brief-based composition.
        /// </summary>
        public bool SelectBestAlternative { get; set; } = true;

        /// <summary>
        ///     Number of alternatives considered by best-fit selection. Zero uses the brief's natural alternative count.
        /// </summary>
        public int AlternativeCount { get; set; }

        /// <summary>Applies the resolved design theme before slides are composed.</summary>
        public bool ApplyTheme { get; set; } = true;

        /// <summary>Validates semantic slide contracts before composition.</summary>
        public bool ValidatePlan { get; set; } = true;

        /// <summary>Expands dense semantic content into continuation slides before composition.</summary>
        public bool ExpandContinuations { get; set; } = true;

        /// <summary>Optional continuation policy used when <see cref="ExpandContinuations" /> is enabled.</summary>
        public PowerPointDeckContinuationOptions? Continuation { get; set; }

        /// <summary>Optional preflight policy for the composed result.</summary>
        public PowerPointDeckPreflightOptions? Preflight { get; set; }

        /// <summary>
        ///     Optional mapping from semantic slide kinds to layouts owned by a copied PowerPoint template.
        /// </summary>
        public PowerPointTemplateLayoutMap? TemplateLayouts { get; set; }

        internal PowerPointDeckDesign ResolveDesign(PowerPointDeckPlan plan) {
            if (Design != null) return Design;
            if (Brief == null) {
                throw new InvalidOperationException("Composition options require a deck design or design brief.");
            }
            if (AlternativeIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(AlternativeIndex),
                    "Design alternative index cannot be negative.");
            }
            if (AlternativeCount < 0) {
                throw new ArgumentOutOfRangeException(nameof(AlternativeCount),
                    "Design alternative count cannot be negative.");
            }

            if (!SelectBestAlternative) return Brief.CreateDesign(AlternativeIndex);

            PowerPointDeckPlanAlternativeSummary recommendation =
                Brief.RecommendDeckPlanAlternative(plan, AlternativeCount);
            return Brief.CreateDesign(recommendation.Index);
        }
    }
}
