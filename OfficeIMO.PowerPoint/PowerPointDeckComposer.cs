using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Presentation-bound designer facade that applies one deck design across many semantic slides.
    /// </summary>
    public sealed class PowerPointDeckComposer {
        private readonly PowerPointPresentation _presentation;
        private readonly PowerPointDeckDesign _design;
        private int _slideIndex;

        internal PowerPointDeckComposer(PowerPointPresentation presentation, PowerPointDeckDesign design,
            bool applyTheme) {
            _presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
            _design = design ?? throw new ArgumentNullException(nameof(design));

            if (applyTheme) {
                _design.ApplyTo(_presentation);
            }
        }

        /// <summary>
        ///     Underlying presentation.
        /// </summary>
        public PowerPointPresentation Presentation => _presentation;

        /// <summary>
        ///     Active deck design.
        /// </summary>
        public PowerPointDeckDesign Design => _design;

        /// <summary>
        ///     Adds a section/title slide using the active deck design.
        /// </summary>
        public PowerPointSlide AddSectionSlide(string title, string? subtitle = null, string? seed = null,
            Action<PowerPointDesignerSlideOptions>? configure = null) {
            PowerPointDesignerSlideOptions options = Configure(new PowerPointDesignerSlideOptions(),
                seed ?? title, configure);
            return _presentation.AddDesignerSectionSlide(title, subtitle, _design.Theme, options);
        }

        /// <summary>
        ///     Adds a case-study slide using the active deck design.
        /// </summary>
        public PowerPointSlide AddCaseStudySlide(string clientTitle, IEnumerable<PowerPointCaseStudySection> sections,
            IEnumerable<PowerPointMetric>? metrics = null, string? seed = null,
            Action<PowerPointCaseStudySlideOptions>? configure = null) {
            PowerPointCaseStudySlideOptions options = Configure(new PowerPointCaseStudySlideOptions(),
                seed ?? clientTitle, configure);
            return _presentation.AddDesignerCaseStudySlide(clientTitle, sections, metrics, _design.Theme, options);
        }

        /// <summary>
        ///     Adds a process/timeline slide using the active deck design.
        /// </summary>
        public PowerPointSlide AddProcessSlide(string title, string? subtitle, IEnumerable<PowerPointProcessStep> steps,
            string? seed = null, Action<PowerPointProcessSlideOptions>? configure = null) {
            PowerPointProcessSlideOptions options = Configure(new PowerPointProcessSlideOptions(),
                seed ?? title, configure);
            return _presentation.AddDesignerProcessSlide(title, subtitle, steps, _design.Theme, options);
        }

        /// <summary>
        ///     Adds a card-grid slide using the active deck design.
        /// </summary>
        public PowerPointSlide AddCardGridSlide(string title, string? subtitle, IEnumerable<PowerPointCardContent> cards,
            string? seed = null, Action<PowerPointCardGridSlideOptions>? configure = null) {
            PowerPointCardGridSlideOptions options = Configure(new PowerPointCardGridSlideOptions(),
                seed ?? title, configure);
            return _presentation.AddDesignerCardGridSlide(title, subtitle, cards, _design.Theme, options);
        }

        /// <summary>
        ///     Adds a logo/proof wall slide using the active deck design.
        /// </summary>
        public PowerPointSlide AddLogoWallSlide(string title, string? subtitle, IEnumerable<PowerPointLogoItem> logos,
            string? seed = null, Action<PowerPointLogoWallSlideOptions>? configure = null) {
            PowerPointLogoWallSlideOptions options = Configure(new PowerPointLogoWallSlideOptions(),
                seed ?? title, configure);
            return _presentation.AddDesignerLogoWallSlide(title, subtitle, logos, _design.Theme, options);
        }

        /// <summary>
        ///     Adds a coverage/location slide using the active deck design.
        /// </summary>
        public PowerPointSlide AddCoverageSlide(string title, string? subtitle,
            IEnumerable<PowerPointCoverageLocation> locations, string? seed = null,
            Action<PowerPointCoverageSlideOptions>? configure = null) {
            PowerPointCoverageSlideOptions options = Configure(new PowerPointCoverageSlideOptions(),
                seed ?? title, configure);
            return _presentation.AddDesignerCoverageSlide(title, subtitle, locations, _design.Theme, options);
        }

        /// <summary>
        ///     Adds a capability/content slide using the active deck design.
        /// </summary>
        public PowerPointSlide AddCapabilitySlide(string title, string? subtitle,
            IEnumerable<PowerPointCapabilitySection> sections, string? seed = null,
            Action<PowerPointCapabilitySlideOptions>? configure = null) {
            PowerPointCapabilitySlideOptions options = Configure(new PowerPointCapabilitySlideOptions(),
                seed ?? title, configure);
            return _presentation.AddDesignerCapabilitySlide(title, subtitle, sections, _design.Theme, options);
        }

        /// <summary>
        ///     Adds a custom designer slide with raw composition primitives and active deck chrome.
        /// </summary>
        public PowerPointSlide ComposeSlide(Action<PowerPointSlideComposer> compose, string? seed = null,
            Action<PowerPointDesignerSlideOptions>? configure = null, bool dark = false) {
            PowerPointDesignerSlideOptions options = Configure(new PowerPointDesignerSlideOptions(),
                seed ?? "custom", configure);
            return _presentation.ComposeDesignerSlide(compose, _design.Theme, options, dark);
        }

        /// <summary>
        ///     Adds all slides described by a semantic deck plan using the active deck design.
        /// </summary>
        public IReadOnlyList<PowerPointSlide> AddSlides(PowerPointDeckPlan plan) {
            return AddSlides(plan, validate: true);
        }

        /// <summary>
        ///     Adds all slides described by a semantic deck plan using the active deck design.
        /// </summary>
        public IReadOnlyList<PowerPointSlide> AddSlides(PowerPointDeckPlan plan, bool validate) {
            if (plan == null) {
                throw new ArgumentNullException(nameof(plan));
            }

            if (validate) {
                IReadOnlyList<PowerPointDeckPlanDiagnostic> diagnostics = plan.ValidateSlides();
                if (diagnostics.Any(diagnostic =>
                        diagnostic.Severity == PowerPointDeckPlanDiagnosticSeverity.Error)) {
                    throw new PowerPointDeckPlanValidationException(diagnostics);
                }
            }

            List<PowerPointSlide> slides = new();
            foreach (PowerPointDeckPlanSlide slide in plan.Slides) {
                slides.Add(slide.AddTo(this));
            }

            return slides.AsReadOnly();
        }

        /// <summary>
        ///     Previews how a semantic deck plan resolves against the active deck design from the composer's
        ///     current slide position.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> DescribeSlides(PowerPointDeckPlan plan) {
            if (plan == null) {
                throw new ArgumentNullException(nameof(plan));
            }

            return plan.DescribeSlides(_design, _slideIndex);
        }

        private T Configure<T>(T options, string seed, Action<T>? configure)
            where T : PowerPointDesignerSlideOptions {
            string resolvedSeed = ResolveSeed(seed);
            _design.Configure(options, resolvedSeed);
            configure?.Invoke(options);
            return options;
        }

        private string ResolveSeed(string seed) {
            _slideIndex++;
            return ResolveSeed(seed, _slideIndex);
        }

        internal static string ResolveSeed(string? seed, int slideIndex) {
            if (string.IsNullOrWhiteSpace(seed)) {
                return "slide-" + slideIndex;
            }

            return seed!.Trim();
        }
    }

    public static partial class PowerPointDesignExtensions {
        /// <summary>
        ///     Creates a presentation-bound designer facade and optionally applies the deck theme immediately.
        /// </summary>
        public static PowerPointDeckComposer UseDesigner(this PowerPointPresentation presentation,
            PowerPointDeckDesign design, bool applyTheme = true) {
            return new PowerPointDeckComposer(presentation, design, applyTheme);
        }

        /// <summary>
        ///     Creates a designer facade from a reusable design brief.
        /// </summary>
        public static PowerPointDeckComposer UseDesigner(this PowerPointPresentation presentation,
            PowerPointDesignBrief brief, int alternativeIndex = 0, bool applyTheme = true) {
            if (brief == null) {
                throw new ArgumentNullException(nameof(brief));
            }

            return new PowerPointDeckComposer(presentation, brief.CreateDesign(alternativeIndex), applyTheme);
        }

        /// <summary>
        ///     Creates a designer facade directly from a brand accent and scenario recipe.
        /// </summary>
        public static PowerPointDeckComposer UseDesigner(this PowerPointPresentation presentation,
            string accentColor, string seed, PowerPointDesignRecipe recipe, int alternativeIndex = 0,
            string? name = null, string? eyebrow = null, string? footerLeft = null, string? footerRight = null,
            bool applyTheme = true) {
            if (recipe == null) {
                throw new ArgumentNullException(nameof(recipe));
            }
            if (alternativeIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(alternativeIndex),
                    "Design alternative index cannot be negative.");
            }

            IReadOnlyList<PowerPointDeckDesign> alternatives = recipe.CreateAlternativesFromBrand(accentColor, seed,
                count: alternativeIndex + 1, name: name, eyebrow: eyebrow, footerLeft: footerLeft,
                footerRight: footerRight);
            return new PowerPointDeckComposer(presentation, alternatives[alternativeIndex], applyTheme);
        }

        /// <summary>
        ///     Creates a designer facade from a brand accent and plain-language deck purpose.
        /// </summary>
        public static PowerPointDeckComposer UseDesigner(this PowerPointPresentation presentation,
            string accentColor, string seed, string purpose, int alternativeIndex = 0,
            string? name = null, string? eyebrow = null, string? footerLeft = null, string? footerRight = null,
            bool applyTheme = true) {
            PowerPointDesignRecipe recipe = PowerPointDesignRecipe.FindBuiltIn(purpose)
                ?? PowerPointDesignRecipe.ConsultingPortfolio;
            return presentation.UseDesigner(accentColor, seed, recipe, alternativeIndex, name, eyebrow,
                footerLeft, footerRight, applyTheme);
        }
    }
}
