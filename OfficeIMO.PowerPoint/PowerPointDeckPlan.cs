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
    public sealed class PowerPointDeckPlan {
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
    }

    /// <summary>
    ///     Base type for one semantic designer slide request.
    /// </summary>
    public abstract class PowerPointDeckPlanSlide {
        private protected PowerPointDeckPlanSlide(string title, string? subtitle, string? seed) {
            if (string.IsNullOrWhiteSpace(title)) {
                throw new ArgumentException("Plan slide title cannot be null or empty.", nameof(title));
            }

            Title = title;
            Subtitle = subtitle;
            Seed = seed;
        }

        /// <summary>
        ///     Slide title or primary label.
        /// </summary>
        public string Title { get; }

        /// <summary>
        ///     Optional slide subtitle.
        /// </summary>
        public string? Subtitle { get; }

        /// <summary>
        ///     Optional stable seed used for this slide.
        /// </summary>
        public string? Seed { get; }

        /// <summary>
        ///     Semantic slide kind.
        /// </summary>
        public abstract PowerPointDeckPlanSlideKind Kind { get; }

        internal abstract PowerPointSlide AddTo(PowerPointDeckComposer deck);

        internal virtual int ContentItemCount => 0;

        internal PowerPointDeckPlanSlideSummary Describe(int index) {
            return new PowerPointDeckPlanSlideSummary(index, Kind, Title, Subtitle, Seed, ContentItemCount);
        }

        internal PowerPointDeckPlanSlideRenderSummary DescribeRender(int index, PowerPointDeckDesign design,
            int slideIndexOffset = 0) {
            string slideSeed = ResolveSeed(index, slideIndexOffset);
            string? layoutVariant = ResolveLayoutVariant(design, slideSeed);
            return new PowerPointDeckPlanSlideRenderSummary(
                index,
                Kind,
                Title,
                Subtitle,
                Seed,
                slideSeed,
                design.Seed + "/" + slideSeed,
                ContentItemCount,
                layoutVariant,
                ResolveLayoutReasons(design, layoutVariant),
                design.Direction.Name,
                design.BaseIntent.Mood,
                design.BaseIntent.Density,
                design.BaseIntent.VisualStyle,
                design.BaseIntent.LayoutStrategy,
                design.Theme.HeadingFontName,
                design.Theme.BodyFontName);
        }

        internal virtual void Validate(int index, IList<PowerPointDeckPlanDiagnostic> diagnostics) {
        }

        private protected void AddDiagnostic(IList<PowerPointDeckPlanDiagnostic> diagnostics, int index,
            PowerPointDeckPlanDiagnosticSeverity severity, string code, string message) {
            diagnostics.Add(new PowerPointDeckPlanDiagnostic(index, Kind, Title, severity, code, message));
        }

        private protected virtual string? ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            return null;
        }

        private protected virtual IReadOnlyList<string> ResolveLayoutReasons(PowerPointDeckDesign design,
            string? layoutVariant) {
            List<string> reasons = new();
            if (!string.IsNullOrWhiteSpace(layoutVariant)) {
                reasons.Add("Uses " + layoutVariant + " for " + Kind + " content.");
            }
            if (design.BaseIntent.Density == PowerPointSlideDensity.Compact) {
                reasons.Add("Compact density keeps the slide content tighter.");
            } else if (design.BaseIntent.Density == PowerPointSlideDensity.Relaxed) {
                reasons.Add("Relaxed density leaves more whitespace around the slide content.");
            }

            return reasons.AsReadOnly();
        }

        private protected static T ConfigurePreview<T>(PowerPointDeckDesign design, string slideSeed,
            Action<T>? configure) where T : PowerPointDesignerSlideOptions, new() {
            T options = design.Configure(new T(), slideSeed);
            configure?.Invoke(options);
            return options;
        }

        private string ResolveSeed(int index, int slideIndexOffset) {
            string seed = Seed ?? Title;
            return PowerPointDeckComposer.ResolveSeed(seed, slideIndexOffset + index + 1);
        }

        private protected static IReadOnlyList<T> Materialize<T>(IEnumerable<T> values, string name) {
            if (values == null) {
                throw new ArgumentNullException(name);
            }

            List<T> list = values.Where(value => value != null).ToList();
            if (list.Count == 0) {
                throw new ArgumentException("Plan slide content cannot be empty.", name);
            }

            return list.AsReadOnly();
        }
    }

    /// <summary>
    ///     Section/title slide request.
    /// </summary>
    public sealed class PowerPointSectionPlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointDesignerSlideOptions>? _configure;

        /// <summary>
        ///     Creates a section/title slide request.
        /// </summary>
        public PowerPointSectionPlanSlide(string title, string? subtitle = null, string? seed = null,
            Action<PowerPointDesignerSlideOptions>? configure = null) : base(title, subtitle, seed) {
            _configure = configure;
        }

        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.Section;

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddSectionSlide(Title, Subtitle, Seed, _configure);
        }

        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointDesignerSlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveSectionVariant(options).ToString();
        }

        private protected override IReadOnlyList<string> ResolveLayoutReasons(PowerPointDeckDesign design,
            string? layoutVariant) {
            List<string> reasons = new() {
                "Section slides use " + layoutVariant + " to establish the deck rhythm before detailed content."
            };
            if (design.ShowDirectionMotif) {
                reasons.Add("Direction motifs are enabled for stronger opening-slide movement.");
            } else {
                reasons.Add("Direction motifs are disabled for a quieter opening slide.");
            }

            return reasons.AsReadOnly();
        }
    }

    /// <summary>
    ///     Case-study slide request.
    /// </summary>
    public sealed class PowerPointCaseStudyPlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointCaseStudySlideOptions>? _configure;

        /// <summary>
        ///     Creates a case-study slide request.
        /// </summary>
        public PowerPointCaseStudyPlanSlide(string clientTitle, IEnumerable<PowerPointCaseStudySection> sections,
            IEnumerable<PowerPointMetric>? metrics = null, string? seed = null,
            Action<PowerPointCaseStudySlideOptions>? configure = null) : base(clientTitle, null, seed) {
            Sections = Materialize(sections, nameof(sections));
            Metrics = (metrics ?? Enumerable.Empty<PowerPointMetric>()).Where(metric => metric != null)
                .ToList().AsReadOnly();
            _configure = configure;
        }

        /// <summary>
        ///     Case-study narrative sections.
        /// </summary>
        public IReadOnlyList<PowerPointCaseStudySection> Sections { get; }

        /// <summary>
        ///     Optional case-study metrics.
        /// </summary>
        public IReadOnlyList<PowerPointMetric> Metrics { get; }

        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.CaseStudy;

        internal override int ContentItemCount => Sections.Count + Metrics.Count;

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddCaseStudySlide(Title, Sections, Metrics, Seed, _configure);
        }

        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointCaseStudySlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveCaseStudyVariant(options, Sections, Metrics).ToString();
        }

        private protected override IReadOnlyList<string> ResolveLayoutReasons(PowerPointDeckDesign design,
            string? layoutVariant) {
            List<string> reasons = new();
            if (Sections.Count >= PowerPointDeckPlanLimits.MaxCaseStudySections) {
                reasons.Add("Four narrative sections favor an editorial split to keep each story block readable.");
            } else if (Metrics.Count > 0) {
                reasons.Add("Metrics are present, so the case study can reserve stronger visual emphasis.");
            } else {
                reasons.Add("Case-study content is balanced across narrative sections.");
            }
            if (design.BaseIntent.VisualStyle == PowerPointVisualStyle.Soft ||
                design.BaseIntent.VisualStyle == PowerPointVisualStyle.Minimal) {
                reasons.Add("Softer visual styles reduce decoration around narrative-heavy content.");
            }
            reasons.Add("Resolved case-study layout: " + layoutVariant + ".");
            return reasons.AsReadOnly();
        }

        internal override void Validate(int index, IList<PowerPointDeckPlanDiagnostic> diagnostics) {
            if (Sections.Count > PowerPointDeckPlanLimits.MaxCaseStudySections) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Error,
                    "CaseStudy.TooManySections", "Case-study slides support up to " +
                                                  PowerPointDeckPlanLimits.MaxCaseStudySections +
                                                  " narrative sections.");
            }
            if (Metrics.Count > PowerPointDeckPlanLimits.MaxCaseStudyMetrics) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Warning,
                    "CaseStudy.HiddenMetrics", "Case-study slides display up to " +
                                               PowerPointDeckPlanLimits.MaxCaseStudyMetrics +
                                               " metrics; extra metrics are ignored.");
            }
        }
    }

    /// <summary>
    ///     Process/timeline slide request.
    /// </summary>
    public sealed class PowerPointProcessPlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointProcessSlideOptions>? _configure;

        /// <summary>
        ///     Creates a process/timeline slide request.
        /// </summary>
        public PowerPointProcessPlanSlide(string title, string? subtitle, IEnumerable<PowerPointProcessStep> steps,
            string? seed = null, Action<PowerPointProcessSlideOptions>? configure = null)
            : base(title, subtitle, seed) {
            Steps = Materialize(steps, nameof(steps));
            _configure = configure;
        }

        /// <summary>
        ///     Process steps.
        /// </summary>
        public IReadOnlyList<PowerPointProcessStep> Steps { get; }

        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.Process;

        internal override int ContentItemCount => Steps.Count;

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddProcessSlide(Title, Subtitle, Steps, Seed, _configure);
        }

        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointProcessSlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveProcessVariant(options, Steps).ToString();
        }

        private protected override IReadOnlyList<string> ResolveLayoutReasons(PowerPointDeckDesign design,
            string? layoutVariant) {
            List<string> reasons = new();
            if (Steps.Count > PowerPointDeckPlanLimits.DenseProcessSteps) {
                reasons.Add("Six or more process steps use a rail so the sequence stays connected.");
            } else if (design.BaseIntent.Density == PowerPointSlideDensity.Compact) {
                reasons.Add("Compact density can use numbered columns for short step-by-step flows.");
            } else {
                reasons.Add("Shorter process flows can vary between rail and numbered columns.");
            }
            if (design.BaseIntent.VisualStyle == PowerPointVisualStyle.Minimal) {
                reasons.Add("Minimal style favors a rail over heavier process decoration.");
            }
            reasons.Add("Resolved process layout: " + layoutVariant + ".");
            return reasons.AsReadOnly();
        }

        internal override void Validate(int index, IList<PowerPointDeckPlanDiagnostic> diagnostics) {
            if (Steps.Count > PowerPointDeckPlanLimits.MaxProcessSteps) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Error,
                    "Process.TooManySteps", "Process slides support up to " +
                                            PowerPointDeckPlanLimits.MaxProcessSteps + " steps.");
            } else if (Steps.Count > PowerPointDeckPlanLimits.DenseProcessSteps) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Warning,
                    "Process.DenseSteps", "Process slides with more than " +
                                          PowerPointDeckPlanLimits.DenseProcessSteps +
                                          " steps are dense; consider splitting the flow.");
            }
        }
    }

    /// <summary>
    ///     Card-grid slide request.
    /// </summary>
    public sealed class PowerPointCardGridPlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointCardGridSlideOptions>? _configure;

        /// <summary>
        ///     Creates a card-grid slide request.
        /// </summary>
        public PowerPointCardGridPlanSlide(string title, string? subtitle, IEnumerable<PowerPointCardContent> cards,
            string? seed = null, Action<PowerPointCardGridSlideOptions>? configure = null)
            : base(title, subtitle, seed) {
            Cards = Materialize(cards, nameof(cards));
            _configure = configure;
        }

        /// <summary>
        ///     Cards displayed by the grid.
        /// </summary>
        public IReadOnlyList<PowerPointCardContent> Cards { get; }

        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.CardGrid;

        internal override int ContentItemCount => Cards.Count;

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddCardGridSlide(Title, Subtitle, Cards, Seed, _configure);
        }

        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointCardGridSlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveCardGridVariant(options, Cards).ToString();
        }

        private protected override IReadOnlyList<string> ResolveLayoutReasons(PowerPointDeckDesign design,
            string? layoutVariant) {
            List<string> reasons = new();
            if (Cards.Count > PowerPointDeckPlanLimits.ComfortableCardGridCards) {
                reasons.Add("More than four cards favor the accent-top grid for compact scanning.");
            } else if (design.BaseIntent.VisualStyle == PowerPointVisualStyle.Soft ||
                       design.BaseIntent.VisualStyle == PowerPointVisualStyle.Minimal) {
                reasons.Add("Softer visual styles favor quieter card tiles.");
            } else {
                reasons.Add("The card count leaves room for visual variation.");
            }
            reasons.Add("Resolved card-grid layout: " + layoutVariant + ".");
            return reasons.AsReadOnly();
        }
    }

    /// <summary>
    ///     Logo/proof wall slide request.
    /// </summary>
    public sealed class PowerPointLogoWallPlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointLogoWallSlideOptions>? _configure;

        /// <summary>
        ///     Creates a logo/proof wall slide request.
        /// </summary>
        public PowerPointLogoWallPlanSlide(string title, string? subtitle, IEnumerable<PowerPointLogoItem> logos,
            string? seed = null, Action<PowerPointLogoWallSlideOptions>? configure = null)
            : base(title, subtitle, seed) {
            Logos = Materialize(logos, nameof(logos));
            _configure = configure;
        }

        /// <summary>
        ///     Logo, partner, product, or certification items.
        /// </summary>
        public IReadOnlyList<PowerPointLogoItem> Logos { get; }

        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.LogoWall;

        internal override int ContentItemCount => Logos.Count;

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddLogoWallSlide(Title, Subtitle, Logos, Seed, _configure);
        }

        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointLogoWallSlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveLogoWallVariant(options, Logos).ToString();
        }

        private protected override IReadOnlyList<string> ResolveLayoutReasons(PowerPointDeckDesign design,
            string? layoutVariant) {
            List<string> reasons = new();
            if (Logos.Count > PowerPointDeckPlanLimits.DenseLogoWallItems) {
                reasons.Add("Large proof walls become compact, so the layout keeps logos in a readable system.");
            } else {
                reasons.Add("Logo-wall content can choose between proof mosaic and featured certificate layouts.");
            }
            reasons.Add("Resolved logo-wall layout: " + layoutVariant + ".");
            return reasons.AsReadOnly();
        }

        internal override void Validate(int index, IList<PowerPointDeckPlanDiagnostic> diagnostics) {
            if (Logos.Count > PowerPointDeckPlanLimits.MaxLogoWallItems) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Error,
                    "LogoWall.TooManyItems", "Logo wall slides support up to " +
                                             PowerPointDeckPlanLimits.MaxLogoWallItems + " items.");
            } else if (Logos.Count > PowerPointDeckPlanLimits.DenseLogoWallItems) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Warning,
                    "LogoWall.DenseItems", "Logo wall slides with more than " +
                                           PowerPointDeckPlanLimits.DenseLogoWallItems + " items become compact.");
            }
        }
    }

    /// <summary>
    ///     Coverage/location slide request.
    /// </summary>
    public sealed class PowerPointCoveragePlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointCoverageSlideOptions>? _configure;

        /// <summary>
        ///     Creates a coverage/location slide request.
        /// </summary>
        public PowerPointCoveragePlanSlide(string title, string? subtitle,
            IEnumerable<PowerPointCoverageLocation> locations, string? seed = null,
            Action<PowerPointCoverageSlideOptions>? configure = null) : base(title, subtitle, seed) {
            Locations = Materialize(locations, nameof(locations));
            _configure = configure;
        }

        /// <summary>
        ///     Coverage locations.
        /// </summary>
        public IReadOnlyList<PowerPointCoverageLocation> Locations { get; }

        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.Coverage;

        internal override int ContentItemCount => Locations.Count;

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddCoverageSlide(Title, Subtitle, Locations, Seed, _configure);
        }

        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointCoverageSlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveCoverageVariant(options, Locations).ToString();
        }

        private protected override IReadOnlyList<string> ResolveLayoutReasons(PowerPointDeckDesign design,
            string? layoutVariant) {
            List<string> reasons = new();
            if (Locations.Count > PowerPointDeckPlanLimits.VisibleCoveragePins) {
                reasons.Add("Many locations favor list support because map pins may become dense.");
            } else {
                reasons.Add("Coverage slides balance map-like visual proof with readable location labels.");
            }
            if (design.BaseIntent.VisualStyle == PowerPointVisualStyle.Geometric) {
                reasons.Add("Geometric style supports map and coverage structure.");
            }
            reasons.Add("Resolved coverage layout: " + layoutVariant + ".");
            return reasons.AsReadOnly();
        }

        internal override void Validate(int index, IList<PowerPointDeckPlanDiagnostic> diagnostics) {
            if (Locations.Count > PowerPointDeckPlanLimits.MaxCoverageLocations) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Error,
                    "Coverage.TooManyLocations", "Coverage slides support up to " +
                                                 PowerPointDeckPlanLimits.MaxCoverageLocations + " locations.");
            } else if (Locations.Count > PowerPointDeckPlanLimits.VisibleCoveragePins) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Warning,
                    "Coverage.HiddenPins", "Coverage map variants show up to " +
                                           PowerPointDeckPlanLimits.VisibleCoveragePins +
                                           " pins; extra locations may appear only in text.");
            }

            for (int i = 0; i < Locations.Count; i++) {
                PowerPointCoverageLocation location = Locations[i];
                if (location.X < 0 || location.X > 1 || location.Y < 0 || location.Y > 1) {
                    AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Error,
                        "Coverage.LocationOutOfBounds",
                        "Location '" + location.Name + "' must use X and Y values between 0 and 1.");
                }
            }
        }
    }

    /// <summary>
    ///     Capability/content slide request.
    /// </summary>
    public sealed class PowerPointCapabilityPlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointCapabilitySlideOptions>? _configure;

        /// <summary>
        ///     Creates a capability/content slide request.
        /// </summary>
        public PowerPointCapabilityPlanSlide(string title, string? subtitle,
            IEnumerable<PowerPointCapabilitySection> sections, string? seed = null,
            Action<PowerPointCapabilitySlideOptions>? configure = null) : base(title, subtitle, seed) {
            Sections = Materialize(sections, nameof(sections));
            _configure = configure;
        }

        /// <summary>
        ///     Capability or content sections.
        /// </summary>
        public IReadOnlyList<PowerPointCapabilitySection> Sections { get; }

        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.Capability;

        internal override int ContentItemCount => Sections.Count;

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddCapabilitySlide(Title, Subtitle, Sections, Seed, _configure);
        }

        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointCapabilitySlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveCapabilityVariant(options, Sections).ToString();
        }

        private protected override IReadOnlyList<string> ResolveLayoutReasons(PowerPointDeckDesign design,
            string? layoutVariant) {
            List<string> reasons = new();
            if (Sections.Count > PowerPointDeckPlanLimits.DenseCapabilitySections) {
                reasons.Add("Many capability sections favor stacked panels to avoid cramped columns.");
            } else if (design.BaseIntent.VisualStyle == PowerPointVisualStyle.Minimal) {
                reasons.Add("Minimal style keeps capability content quieter and more editorial.");
            } else {
                reasons.Add("Capability content can use a text-and-visual split.");
            }
            reasons.Add("Resolved capability layout: " + layoutVariant + ".");
            return reasons.AsReadOnly();
        }

        internal override void Validate(int index, IList<PowerPointDeckPlanDiagnostic> diagnostics) {
            if (Sections.Count > PowerPointDeckPlanLimits.MaxCapabilitySections) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Error,
                    "Capability.TooManySections", "Capability slides support up to " +
                                                  PowerPointDeckPlanLimits.MaxCapabilitySections + " sections.");
            } else if (Sections.Count > PowerPointDeckPlanLimits.DenseCapabilitySections) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Warning,
                    "Capability.DenseSections", "Capability slides with more than " +
                                                PowerPointDeckPlanLimits.DenseCapabilitySections +
                                                " sections are dense.");
            }
        }
    }

    /// <summary>
    ///     Custom raw-composition slide request.
    /// </summary>
    public sealed class PowerPointCustomPlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointSlideComposer> _compose;
        private readonly Action<PowerPointDesignerSlideOptions>? _configure;

        /// <summary>
        ///     Creates a custom slide request that can use slide composer primitives.
        /// </summary>
        public PowerPointCustomPlanSlide(string title, Action<PowerPointSlideComposer> compose,
            string? seed = null, Action<PowerPointDesignerSlideOptions>? configure = null, bool dark = false)
            : base(title, null, seed) {
            _compose = compose ?? throw new ArgumentNullException(nameof(compose));
            _configure = configure;
            Dark = dark;
        }

        /// <summary>
        ///     Whether the custom slide should use the dark designer surface.
        /// </summary>
        public bool Dark { get; }

        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.Custom;

        internal override int ContentItemCount => 1;

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.ComposeSlide(_compose, Seed ?? Title, _configure, Dark);
        }

        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            return Dark ? "CustomDark" : "Custom";
        }

        private protected override IReadOnlyList<string> ResolveLayoutReasons(PowerPointDeckDesign design,
            string? layoutVariant) {
            List<string> reasons = new() {
                "Custom slides keep raw composition control while still using deck identity and theme defaults."
            };
            if (Dark) {
                reasons.Add("The custom slide requests the dark designer surface.");
            } else {
                reasons.Add("The custom slide requests the light designer surface.");
            }

            return reasons.AsReadOnly();
        }
    }

    /// <summary>
    ///     Lightweight description of one planned slide.
    /// </summary>
    public sealed class PowerPointDeckPlanSlideSummary {
        internal PowerPointDeckPlanSlideSummary(int index, PowerPointDeckPlanSlideKind kind,
            string title, string? subtitle, string? seed, int contentItemCount) {
            Index = index;
            Kind = kind;
            Title = title;
            Subtitle = subtitle;
            Seed = seed;
            ContentItemCount = contentItemCount;
        }

        /// <summary>
        ///     Zero-based slide index within the plan.
        /// </summary>
        public int Index { get; }

        /// <summary>
        ///     Semantic slide kind.
        /// </summary>
        public PowerPointDeckPlanSlideKind Kind { get; }

        /// <summary>
        ///     Planned slide title.
        /// </summary>
        public string Title { get; }

        /// <summary>
        ///     Optional planned slide subtitle.
        /// </summary>
        public string? Subtitle { get; }

        /// <summary>
        ///     Optional stable seed used for the planned slide.
        /// </summary>
        public string? Seed { get; }

        /// <summary>
        ///     Count of primary content items such as sections, steps, cards, or locations.
        /// </summary>
        public int ContentItemCount { get; }

        /// <inheritdoc />
        public override string ToString() {
            return Index + ": " + Kind + " - " + Title;
        }
    }

    /// <summary>
    ///     Warning or error found while validating a planned designer slide.
    /// </summary>
    public sealed class PowerPointDeckPlanDiagnostic {
        internal PowerPointDeckPlanDiagnostic(int index, PowerPointDeckPlanSlideKind kind, string title,
            PowerPointDeckPlanDiagnosticSeverity severity, string code, string message) {
            Index = index;
            Kind = kind;
            Title = title;
            Severity = severity;
            Code = code;
            Message = message;
        }

        /// <summary>
        ///     Zero-based slide index within the plan.
        /// </summary>
        public int Index { get; }

        /// <summary>
        ///     Semantic slide kind.
        /// </summary>
        public PowerPointDeckPlanSlideKind Kind { get; }

        /// <summary>
        ///     Planned slide title.
        /// </summary>
        public string Title { get; }

        /// <summary>
        ///     Diagnostic severity.
        /// </summary>
        public PowerPointDeckPlanDiagnosticSeverity Severity { get; }

        /// <summary>
        ///     Stable machine-readable diagnostic code.
        /// </summary>
        public string Code { get; }

        /// <summary>
        ///     Human-readable diagnostic message.
        /// </summary>
        public string Message { get; }

        /// <inheritdoc />
        public override string ToString() {
            return Index + ": " + Severity + " " + Code + " - " + Message;
        }
    }

    /// <summary>
    ///     Raised when a semantic deck plan contains validation errors that would prevent rendering.
    /// </summary>
    public sealed class PowerPointDeckPlanValidationException : InvalidOperationException {
        internal PowerPointDeckPlanValidationException(IReadOnlyList<PowerPointDeckPlanDiagnostic> diagnostics)
            : base(CreateMessage(diagnostics)) {
            Diagnostics = diagnostics.ToList().AsReadOnly();
        }

        /// <summary>
        ///     Diagnostics collected from the deck plan.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanDiagnostic> Diagnostics { get; }

        private static string CreateMessage(IReadOnlyList<PowerPointDeckPlanDiagnostic> diagnostics) {
            IEnumerable<PowerPointDeckPlanDiagnostic> errors = diagnostics.Where(diagnostic =>
                diagnostic.Severity == PowerPointDeckPlanDiagnosticSeverity.Error);
            return "Deck plan contains rendering errors: " + string.Join("; ",
                errors.Select(diagnostic => diagnostic.Code + " on slide " + diagnostic.Index + " (" +
                                            diagnostic.Title + "): " + diagnostic.Message));
        }
    }

    /// <summary>
    ///     Lightweight description of one planned slide after resolving a deck design.
    /// </summary>
    public sealed class PowerPointDeckPlanSlideRenderSummary {
        internal PowerPointDeckPlanSlideRenderSummary(int index, PowerPointDeckPlanSlideKind kind,
            string title, string? subtitle, string? seed, string resolvedSeed, string designSeed,
            int contentItemCount, string? layoutVariant, IReadOnlyList<string> layoutReasons,
            string directionName, PowerPointDesignMood mood, PowerPointSlideDensity density,
            PowerPointVisualStyle visualStyle, PowerPointAutoLayoutStrategy layoutStrategy, string headingFontName,
            string bodyFontName) {
            Index = index;
            Kind = kind;
            Title = title;
            Subtitle = subtitle;
            Seed = seed;
            ResolvedSeed = resolvedSeed;
            DesignSeed = designSeed;
            ContentItemCount = contentItemCount;
            LayoutVariant = layoutVariant;
            LayoutReasons = layoutReasons;
            DirectionName = directionName;
            Mood = mood;
            Density = density;
            VisualStyle = visualStyle;
            LayoutStrategy = layoutStrategy;
            HeadingFontName = headingFontName;
            BodyFontName = bodyFontName;
        }

        /// <summary>
        ///     Zero-based slide index within the plan.
        /// </summary>
        public int Index { get; }

        /// <summary>
        ///     Semantic slide kind.
        /// </summary>
        public PowerPointDeckPlanSlideKind Kind { get; }

        /// <summary>
        ///     Planned slide title.
        /// </summary>
        public string Title { get; }

        /// <summary>
        ///     Optional planned slide subtitle.
        /// </summary>
        public string? Subtitle { get; }

        /// <summary>
        ///     Optional caller-supplied stable seed used for the planned slide.
        /// </summary>
        public string? Seed { get; }

        /// <summary>
        ///     Slide seed resolved the same way the deck composer resolves it before rendering.
        /// </summary>
        public string ResolvedSeed { get; }

        /// <summary>
        ///     Full deterministic design seed formed from the deck design seed and resolved slide seed.
        /// </summary>
        public string DesignSeed { get; }

        /// <summary>
        ///     Count of primary content items such as sections, steps, cards, or locations.
        /// </summary>
        public int ContentItemCount { get; }

        /// <summary>
        ///     Resolved layout variant name for semantic slides, or custom surface name for raw composition slides.
        /// </summary>
        public string? LayoutVariant { get; }

        /// <summary>
        ///     Short explanations for why this layout variant was selected for the planned slide.
        /// </summary>
        public IReadOnlyList<string> LayoutReasons { get; }

        /// <summary>
        ///     Creative direction name used by the deck design.
        /// </summary>
        public string DirectionName { get; }

        /// <summary>
        ///     Broad visual mood used by the deck design.
        /// </summary>
        public PowerPointDesignMood Mood { get; }

        /// <summary>
        ///     Preferred content density used by the deck design.
        /// </summary>
        public PowerPointSlideDensity Density { get; }

        /// <summary>
        ///     Preferred visual style used by the deck design.
        /// </summary>
        public PowerPointVisualStyle VisualStyle { get; }

        /// <summary>
        ///     Auto layout strategy used when resolving this planned slide.
        /// </summary>
        public PowerPointAutoLayoutStrategy LayoutStrategy { get; }

        /// <summary>
        ///     Heading font used by the deck design.
        /// </summary>
        public string HeadingFontName { get; }

        /// <summary>
        ///     Body font used by the deck design.
        /// </summary>
        public string BodyFontName { get; }

        /// <inheritdoc />
        public override string ToString() {
            return Index + ": " + Kind + " - " + Title + " [" + LayoutVariant + "]";
        }
    }
}
