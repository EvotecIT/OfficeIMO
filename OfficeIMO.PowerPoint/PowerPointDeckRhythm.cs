using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>Severity of a deck-rhythm finding.</summary>
    public enum PowerPointDeckRhythmSeverity {
        /// <summary>The deck can render, but its sequence may feel repetitive or poorly paced.</summary>
        Warning,
        /// <summary>The deck sequence has a structural problem that should be corrected.</summary>
        Error
    }

    /// <summary>Configures pre-render deck-rhythm inspection.</summary>
    public sealed class PowerPointDeckRhythmOptions {
        /// <summary>Maximum consecutive slides of the same semantic kind.</summary>
        public int MaximumRepeatedKind { get; set; } = 2;
        /// <summary>Maximum consecutive dense slides before a pacing warning.</summary>
        public int MaximumDenseStreak { get; set; } = 3;
        /// <summary>Maximum detailed slides between section or closing anchors.</summary>
        public int MaximumSlidesBetweenAnchors { get; set; } = 6;
        /// <summary>Deck length at which a closing slide is expected.</summary>
        public int ClosingExpectedAtSlideCount { get; set; } = 5;

        internal void Validate() {
            if (MaximumRepeatedKind < 1) throw new ArgumentOutOfRangeException(nameof(MaximumRepeatedKind));
            if (MaximumDenseStreak < 1) throw new ArgumentOutOfRangeException(nameof(MaximumDenseStreak));
            if (MaximumSlidesBetweenAnchors < 1) throw new ArgumentOutOfRangeException(nameof(MaximumSlidesBetweenAnchors));
            if (ClosingExpectedAtSlideCount < 1) throw new ArgumentOutOfRangeException(nameof(ClosingExpectedAtSlideCount));
        }
    }

    /// <summary>One stable, machine-readable deck-rhythm finding.</summary>
    public sealed class PowerPointDeckRhythmFinding {
        internal PowerPointDeckRhythmFinding(string code, PowerPointDeckRhythmSeverity severity,
            int slideIndex, string message) {
            Code = code;
            Severity = severity;
            SlideIndex = slideIndex;
            Message = message;
        }

        /// <summary>Stable finding code.</summary>
        public string Code { get; }
        /// <summary>Finding severity.</summary>
        public PowerPointDeckRhythmSeverity Severity { get; }
        /// <summary>Zero-based slide index associated with the finding, or -1 for a deck-level finding.</summary>
        public int SlideIndex { get; }
        /// <summary>Human-readable explanation.</summary>
        public string Message { get; }
        /// <inheritdoc />
        public override string ToString() => Code + ": " + Message;
    }

    /// <summary>Pre-render assessment of variety, density, anchors, and closing cadence.</summary>
    public sealed class PowerPointDeckRhythmReport {
        internal PowerPointDeckRhythmReport(IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> slides,
            IReadOnlyList<PowerPointDeckRhythmFinding> findings) {
            Slides = slides;
            Findings = findings;
            Score = Math.Max(0, 100 - findings.Count(finding =>
                finding.Severity == PowerPointDeckRhythmSeverity.Error) * 25 - findings.Count(finding =>
                finding.Severity == PowerPointDeckRhythmSeverity.Warning) * 8);
        }

        /// <summary>Resolved slide sequence inspected by the report.</summary>
        public IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> Slides { get; }
        /// <summary>Rhythm findings.</summary>
        public IReadOnlyList<PowerPointDeckRhythmFinding> Findings { get; }
        /// <summary>Simple zero-to-one-hundred rhythm score.</summary>
        public int Score { get; }
        /// <summary>Whether any warning or error was found.</summary>
        public bool HasFindings => Findings.Count > 0;
    }

    public sealed partial class PowerPointDeckPlan {
        /// <summary>Inspects the resolved semantic sequence before rendering.</summary>
        public PowerPointDeckRhythmReport InspectRhythm(PowerPointDeckDesign design,
            PowerPointDeckRhythmOptions? options = null) {
            if (design == null) throw new ArgumentNullException(nameof(design));
            PowerPointDeckRhythmOptions resolved = options ?? new PowerPointDeckRhythmOptions();
            resolved.Validate();
            IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> slides = DescribeSlides(design);
            var findings = new List<PowerPointDeckRhythmFinding>();
            InspectOpening(slides, findings);
            InspectRepeatedKinds(slides, resolved, findings);
            InspectRepeatedVariants(slides, resolved, findings);
            InspectDensity(slides, resolved, findings);
            InspectAnchors(slides, resolved, findings);
            InspectClosing(slides, resolved, findings);
            return new PowerPointDeckRhythmReport(slides,
                new ReadOnlyCollection<PowerPointDeckRhythmFinding>(findings));
        }

        private static void InspectOpening(IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> slides,
            ICollection<PowerPointDeckRhythmFinding> findings) {
            if (slides.Count == 0) {
                findings.Add(new PowerPointDeckRhythmFinding("Rhythm.EmptyDeck",
                    PowerPointDeckRhythmSeverity.Error, -1, "The plan contains no slides."));
                return;
            }
            PowerPointDeckPlanSlideKind kind = slides[0].Kind;
            if (kind != PowerPointDeckPlanSlideKind.Section &&
                kind != PowerPointDeckPlanSlideKind.ExecutiveSummary) {
                findings.Add(new PowerPointDeckRhythmFinding("Rhythm.WeakOpening",
                    PowerPointDeckRhythmSeverity.Warning, 0,
                    "Consider opening with a section or executive-summary slide before detailed content."));
            }
        }

        private static void InspectRepeatedKinds(IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> slides,
            PowerPointDeckRhythmOptions options, ICollection<PowerPointDeckRhythmFinding> findings) {
            int streak = 1;
            for (int index = 1; index < slides.Count; index++) {
                streak = slides[index].Kind == slides[index - 1].Kind ? streak + 1 : 1;
                if (streak == options.MaximumRepeatedKind + 1) {
                    findings.Add(new PowerPointDeckRhythmFinding("Rhythm.RepeatedKind",
                        PowerPointDeckRhythmSeverity.Warning, index,
                        streak + " consecutive " + slides[index].Kind + " slides reduce visual variety."));
                }
            }
        }

        private static void InspectRepeatedVariants(IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> slides,
            PowerPointDeckRhythmOptions options, ICollection<PowerPointDeckRhythmFinding> findings) {
            int streak = 1;
            for (int index = 1; index < slides.Count; index++) {
                bool same = slides[index].Kind == slides[index - 1].Kind &&
                    !string.IsNullOrWhiteSpace(slides[index].LayoutVariant) &&
                    string.Equals(slides[index].LayoutVariant, slides[index - 1].LayoutVariant,
                        StringComparison.Ordinal);
                streak = same ? streak + 1 : 1;
                if (streak == options.MaximumRepeatedKind + 1) {
                    findings.Add(new PowerPointDeckRhythmFinding("Rhythm.RepeatedVariant",
                        PowerPointDeckRhythmSeverity.Warning, index,
                        "The same " + slides[index].LayoutVariant + " composition repeats without a visual break."));
                }
            }
        }

        private static void InspectDensity(IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> slides,
            PowerPointDeckRhythmOptions options, ICollection<PowerPointDeckRhythmFinding> findings) {
            int streak = 0;
            for (int index = 0; index < slides.Count; index++) {
                streak = IsDense(slides[index]) ? streak + 1 : 0;
                if (streak == options.MaximumDenseStreak + 1) {
                    findings.Add(new PowerPointDeckRhythmFinding("Rhythm.DenseStreak",
                        PowerPointDeckRhythmSeverity.Warning, index,
                        "Several information-dense slides appear back to back; add a visual or section reset."));
                }
            }
        }

        private static bool IsDense(PowerPointDeckPlanSlideRenderSummary slide) {
            switch (slide.Kind) {
                case PowerPointDeckPlanSlideKind.Process: return slide.ContentItemCount > 5;
                case PowerPointDeckPlanSlideKind.CardGrid: return slide.ContentItemCount > 4;
                case PowerPointDeckPlanSlideKind.Comparison: return slide.ContentItemCount > 2;
                case PowerPointDeckPlanSlideKind.AppendixTable: return slide.ContentItemCount > 8;
                case PowerPointDeckPlanSlideKind.Architecture: return slide.ContentItemCount > 9;
                case PowerPointDeckPlanSlideKind.Capability: return slide.ContentItemCount > 4;
                default: return false;
            }
        }

        private static void InspectAnchors(IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> slides,
            PowerPointDeckRhythmOptions options, ICollection<PowerPointDeckRhythmFinding> findings) {
            int sinceAnchor = 0;
            for (int index = 0; index < slides.Count; index++) {
                if (IsAnchor(slides[index].Kind)) {
                    sinceAnchor = 0;
                    continue;
                }
                sinceAnchor++;
                if (sinceAnchor == options.MaximumSlidesBetweenAnchors + 1) {
                    findings.Add(new PowerPointDeckRhythmFinding("Rhythm.LongSection",
                        PowerPointDeckRhythmSeverity.Warning, index,
                        "The deck runs more than " + options.MaximumSlidesBetweenAnchors +
                        " detail slides without a section or closing anchor."));
                }
            }
        }

        private static bool IsAnchor(PowerPointDeckPlanSlideKind kind) =>
            kind == PowerPointDeckPlanSlideKind.Section || kind == PowerPointDeckPlanSlideKind.Closing;

        private static void InspectClosing(IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> slides,
            PowerPointDeckRhythmOptions options, ICollection<PowerPointDeckRhythmFinding> findings) {
            if (slides.Count >= options.ClosingExpectedAtSlideCount &&
                slides.All(slide => slide.Kind != PowerPointDeckPlanSlideKind.Closing)) {
                findings.Add(new PowerPointDeckRhythmFinding("Rhythm.MissingClosing",
                    PowerPointDeckRhythmSeverity.Warning, slides.Count - 1,
                    "A deck of this length should end with a takeaway, decision, or next action."));
            }
        }
    }
}
