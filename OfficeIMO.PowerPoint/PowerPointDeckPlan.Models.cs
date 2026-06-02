using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {

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
