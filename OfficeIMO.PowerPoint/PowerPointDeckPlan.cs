using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Shared content limits used by semantic PowerPoint deck plans.
    /// </summary>
    public static class PowerPointDeckPlanLimits {
        /// <summary>Maximum narrative sections rendered by a case-study slide.</summary>
        public const int MaxCaseStudySections = 4;
        /// <summary>Maximum metrics visibly rendered by a case-study slide.</summary>
        public const int MaxCaseStudyMetrics = 3;
        /// <summary>Maximum steps supported by a process slide.</summary>
        public const int MaxProcessSteps = 8;
        /// <summary>Process steps above this count are considered dense.</summary>
        public const int DenseProcessSteps = 5;
        /// <summary>Card counts above this value favor compact grid layouts.</summary>
        public const int ComfortableCardGridCards = 4;
        /// <summary>Maximum items supported by a logo/proof wall slide.</summary>
        public const int MaxLogoWallItems = 24;
        /// <summary>Logo/proof wall items above this count are considered dense.</summary>
        public const int DenseLogoWallItems = 12;
        /// <summary>Maximum locations supported by a coverage slide.</summary>
        public const int MaxCoverageLocations = 24;
        /// <summary>Maximum pins shown by map-like coverage variants before text-only overflow.</summary>
        public const int VisibleCoveragePins = 18;
        /// <summary>Maximum sections supported by a capability slide.</summary>
        public const int MaxCapabilitySections = 6;
        /// <summary>Capability sections above this count are considered dense.</summary>
        public const int DenseCapabilitySections = 4;
        /// <summary>Maximum rows rendered by one appendix-table slide.</summary>
        public const int MaxAppendixTableRows = 14;
    }

    /// <summary>
    ///     Semantic kind of a planned designer slide.
    /// </summary>
    public enum PowerPointDeckPlanSlideKind {
        /// <summary>Section or title slide.</summary>
        Section,
        /// <summary>Case-study summary slide.</summary>
        CaseStudy,
        /// <summary>Process or timeline slide.</summary>
        Process,
        /// <summary>Card-grid slide.</summary>
        CardGrid,
        /// <summary>Logo, partner, or proof wall slide.</summary>
        LogoWall,
        /// <summary>Coverage or location slide.</summary>
        Coverage,
        /// <summary>Capability or content slide.</summary>
        Capability,
        /// <summary>Executive-summary slide.</summary>
        ExecutiveSummary,
        /// <summary>Editable chart with narrative context.</summary>
        ChartStory,
        /// <summary>Side-by-side or matrix comparison.</summary>
        Comparison,
        /// <summary>Screenshot with semantic crop, metadata, and annotations.</summary>
        ScreenshotStory,
        /// <summary>Paginated editable appendix table.</summary>
        AppendixTable,
        /// <summary>Editable architecture diagram.</summary>
        Architecture,
        /// <summary>Closing statement or action slide.</summary>
        Closing,
        /// <summary>Custom raw-composition slide.</summary>
        Custom
    }

    /// <summary>
    ///     Severity of a planned slide diagnostic.
    /// </summary>
    public enum PowerPointDeckPlanDiagnosticSeverity {
        /// <summary>The plan can render, but content may be dense, hidden, or better split across slides.</summary>
        Warning,
        /// <summary>The plan contains content that the semantic renderer rejects.</summary>
        Error
    }

    /// <summary>
    ///     Semantic sequence of designer slides that can be applied to a deck composer.
    /// </summary>
    public sealed partial class PowerPointDeckPlan {
        private readonly List<PowerPointDeckPlanSlide> _slides = new();

        /// <summary>
        ///     Slides requested by this plan.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanSlide> Slides => _slides;

        /// <summary>
        ///     Adds a prebuilt semantic slide request.
        /// </summary>
        public PowerPointDeckPlan Add(PowerPointDeckPlanSlide slide) {
            if (slide == null) {
                throw new ArgumentNullException(nameof(slide));
            }

            _slides.Add(slide);
            return this;
        }

        /// <summary>
        ///     Adds a section/title slide request.
        /// </summary>
        public PowerPointDeckPlan AddSection(string title, string? subtitle = null, string? seed = null,
            Action<PowerPointDesignerSlideOptions>? configure = null) {
            return Add(new PowerPointSectionPlanSlide(title, subtitle, seed, configure));
        }

        /// <summary>
        ///     Adds a case-study slide request.
        /// </summary>
        public PowerPointDeckPlan AddCaseStudy(string clientTitle, IEnumerable<PowerPointCaseStudySection> sections,
            IEnumerable<PowerPointMetric>? metrics = null, string? seed = null,
            Action<PowerPointCaseStudySlideOptions>? configure = null) {
            return Add(new PowerPointCaseStudyPlanSlide(clientTitle, sections, metrics, seed, configure));
        }

        /// <summary>
        ///     Adds a process/timeline slide request.
        /// </summary>
        public PowerPointDeckPlan AddProcess(string title, string? subtitle,
            IEnumerable<PowerPointProcessStep> steps, string? seed = null,
            Action<PowerPointProcessSlideOptions>? configure = null) {
            return Add(new PowerPointProcessPlanSlide(title, subtitle, steps, seed, configure));
        }

        /// <summary>
        ///     Adds a card-grid slide request.
        /// </summary>
        public PowerPointDeckPlan AddCardGrid(string title, string? subtitle,
            IEnumerable<PowerPointCardContent> cards, string? seed = null,
            Action<PowerPointCardGridSlideOptions>? configure = null) {
            return Add(new PowerPointCardGridPlanSlide(title, subtitle, cards, seed, configure));
        }

        /// <summary>
        ///     Adds a logo/proof wall slide request.
        /// </summary>
        public PowerPointDeckPlan AddLogoWall(string title, string? subtitle,
            IEnumerable<PowerPointLogoItem> logos, string? seed = null,
            Action<PowerPointLogoWallSlideOptions>? configure = null) {
            return Add(new PowerPointLogoWallPlanSlide(title, subtitle, logos, seed, configure));
        }

        /// <summary>
        ///     Adds a coverage/location slide request.
        /// </summary>
        public PowerPointDeckPlan AddCoverage(string title, string? subtitle,
            IEnumerable<PowerPointCoverageLocation> locations, string? seed = null,
            Action<PowerPointCoverageSlideOptions>? configure = null) {
            return Add(new PowerPointCoveragePlanSlide(title, subtitle, locations, seed, configure));
        }

        /// <summary>
        ///     Adds a capability/content slide request.
        /// </summary>
        public PowerPointDeckPlan AddCapability(string title, string? subtitle,
            IEnumerable<PowerPointCapabilitySection> sections, string? seed = null,
            Action<PowerPointCapabilitySlideOptions>? configure = null) {
            return Add(new PowerPointCapabilityPlanSlide(title, subtitle, sections, seed, configure));
        }

        /// <summary>
        ///     Adds a custom slide request that can use raw composition primitives inside the same semantic plan.
        /// </summary>
        public PowerPointDeckPlan AddCustom(string title, Action<PowerPointSlideComposer> compose,
            string? seed = null, Action<PowerPointDesignerSlideOptions>? configure = null, bool dark = false) {
            return Add(new PowerPointCustomPlanSlide(title, compose, seed, configure, dark));
        }

        /// <summary>
        ///     Creates lightweight descriptions of the planned slide sequence before rendering.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanSlideSummary> DescribeSlides() {
            PowerPointDeckPlanSlideSummary[] summaries = new PowerPointDeckPlanSlideSummary[_slides.Count];
            for (int i = 0; i < _slides.Count; i++) {
                summaries[i] = _slides[i].Describe(i);
            }

            return summaries;
        }

        /// <summary>
        ///     Creates lightweight descriptions of how the planned slide sequence resolves under a deck design.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> DescribeSlides(PowerPointDeckDesign design) {
            return DescribeSlides(design, slideIndexOffset: 0);
        }

        /// <summary>
        ///     Creates lightweight descriptions of how the planned slide sequence resolves under a deck design,
        ///     using an existing composer slide count for fallback seed generation.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> DescribeSlides(PowerPointDeckDesign design,
            int slideIndexOffset) {
            if (design == null) {
                throw new ArgumentNullException(nameof(design));
            }
            if (slideIndexOffset < 0) {
                throw new ArgumentOutOfRangeException(nameof(slideIndexOffset),
                    "Slide index offset cannot be negative.");
            }

            PowerPointDeckPlanSlideRenderSummary[] summaries =
                new PowerPointDeckPlanSlideRenderSummary[_slides.Count];
            for (int i = 0; i < _slides.Count; i++) {
                summaries[i] = _slides[i].DescribeRender(i, design, slideIndexOffset);
            }

            return summaries;
        }

        /// <summary>
        ///     Returns warnings and errors for the planned slide sequence before rendering.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanDiagnostic> ValidateSlides() {
            List<PowerPointDeckPlanDiagnostic> diagnostics = new();
            for (int i = 0; i < _slides.Count; i++) {
                _slides[i].Validate(i, diagnostics);
            }

            return diagnostics.AsReadOnly();
        }

        /// <summary>
        ///     Expands dense semantic requests into deterministic continuation slides without dropping content.
        /// </summary>
        public PowerPointDeckPlan WithContinuations(PowerPointDeckContinuationOptions? options = null) {
            PowerPointDeckContinuationOptions resolved = options ?? new PowerPointDeckContinuationOptions();
            var expanded = new PowerPointDeckPlan();
            for (int slideIndex = 0; slideIndex < _slides.Count; slideIndex++) {
                foreach (PowerPointDeckPlanSlide page in _slides[slideIndex].ExpandContinuations(resolved)) {
                    expanded.Add(page);
                }
            }

            return expanded;
        }
    }
}
