using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
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

        internal abstract PowerPointSlide AddTo(PowerPointDeckComposer deck);

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

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddSectionSlide(Title, Subtitle, Seed, _configure);
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

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddCaseStudySlide(Title, Sections, Metrics, Seed, _configure);
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

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddProcessSlide(Title, Subtitle, Steps, Seed, _configure);
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

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddCardGridSlide(Title, Subtitle, Cards, Seed, _configure);
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

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddLogoWallSlide(Title, Subtitle, Logos, Seed, _configure);
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

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddCoverageSlide(Title, Subtitle, Locations, Seed, _configure);
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

        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) {
            return deck.AddCapabilitySlide(Title, Subtitle, Sections, Seed, _configure);
        }
    }
}
