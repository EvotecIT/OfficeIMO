using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint {
    /// <summary>Configures deterministic continuation of semantic deck-plan content.</summary>
    public sealed class PowerPointDeckContinuationOptions {
        private int _caseStudySectionsPerSlide = PowerPointDeckPlanLimits.MaxCaseStudySections;
        private int _caseStudyMetricsPerSlide = PowerPointDeckPlanLimits.MaxCaseStudyMetrics;
        private int _processStepsPerSlide = PowerPointDeckPlanLimits.DenseProcessSteps;
        private int _cardsPerSlide = 6;
        private int _logosPerSlide = PowerPointDeckPlanLimits.DenseLogoWallItems;
        private int _locationsPerSlide = PowerPointDeckPlanLimits.VisibleCoveragePins;
        private int _capabilitySectionsPerSlide = PowerPointDeckPlanLimits.DenseCapabilitySections;
        private int _appendixRowsPerSlide = PowerPointDeckPlanLimits.MaxAppendixTableRows;

        /// <summary>Maximum narrative case-study sections per generated page.</summary>
        public int CaseStudySectionsPerSlide {
            get => _caseStudySectionsPerSlide;
            set => _caseStudySectionsPerSlide = RequireRange(value, 1,
                PowerPointDeckPlanLimits.MaxCaseStudySections, nameof(CaseStudySectionsPerSlide));
        }

        /// <summary>Maximum case-study metrics per generated page.</summary>
        public int CaseStudyMetricsPerSlide {
            get => _caseStudyMetricsPerSlide;
            set => _caseStudyMetricsPerSlide = RequireRange(value, 1,
                PowerPointDeckPlanLimits.MaxCaseStudyMetrics, nameof(CaseStudyMetricsPerSlide));
        }

        /// <summary>Maximum process steps per generated page.</summary>
        public int ProcessStepsPerSlide {
            get => _processStepsPerSlide;
            set => _processStepsPerSlide = RequireRange(value, 1,
                PowerPointDeckPlanLimits.MaxProcessSteps, nameof(ProcessStepsPerSlide));
        }

        /// <summary>Maximum cards per generated page.</summary>
        public int CardsPerSlide {
            get => _cardsPerSlide;
            set => _cardsPerSlide = RequirePositive(value, nameof(CardsPerSlide));
        }

        /// <summary>Maximum logo/proof items per generated page.</summary>
        public int LogosPerSlide {
            get => _logosPerSlide;
            set => _logosPerSlide = RequireRange(value, 1,
                PowerPointDeckPlanLimits.MaxLogoWallItems, nameof(LogosPerSlide));
        }

        /// <summary>Maximum coverage locations per generated page.</summary>
        public int LocationsPerSlide {
            get => _locationsPerSlide;
            set => _locationsPerSlide = RequireRange(value, 1,
                PowerPointDeckPlanLimits.MaxCoverageLocations, nameof(LocationsPerSlide));
        }

        /// <summary>Maximum capability sections per generated page.</summary>
        public int CapabilitySectionsPerSlide {
            get => _capabilitySectionsPerSlide;
            set => _capabilitySectionsPerSlide = RequireRange(value, 1,
                PowerPointDeckPlanLimits.MaxCapabilitySections, nameof(CapabilitySectionsPerSlide));
        }

        /// <summary>Maximum appendix-table data rows per generated page.</summary>
        public int AppendixRowsPerSlide {
            get => _appendixRowsPerSlide;
            set => _appendixRowsPerSlide = RequireRange(value, 1,
                PowerPointDeckPlanLimits.MaxAppendixTableRows, nameof(AppendixRowsPerSlide));
        }

        /// <summary>
        ///     Formats continuation titles. Parameters are the original title, one-based page number, and page count.
        ///     The first page retains the original title.
        /// </summary>
        public string ContinuationTitleFormat { get; set; } = "{0} (continued {1}/{2})";

        internal string CreateTitle(string title, int pageIndex, int pageCount) {
            if (pageIndex == 0 || pageCount <= 1) return title;
            string format = string.IsNullOrWhiteSpace(ContinuationTitleFormat)
                ? "{0} (continued {1}/{2})"
                : ContinuationTitleFormat;
            return string.Format(System.Globalization.CultureInfo.InvariantCulture, format, title,
                pageIndex + 1, pageCount);
        }

        internal string? CreateSeed(string? seed, string title, int pageIndex) {
            if (pageIndex == 0) return seed;
            string value = string.IsNullOrWhiteSpace(seed) ? title : seed!;
            return value.Trim() + "-continuation-" + (pageIndex + 1);
        }

        internal static IReadOnlyList<IReadOnlyList<T>> Chunk<T>(IReadOnlyList<T> source, int size) {
            var chunks = new List<IReadOnlyList<T>>();
            for (int offset = 0; offset < source.Count; offset += size) {
                int count = Math.Min(size, source.Count - offset);
                var page = new List<T>(count);
                for (int index = 0; index < count; index++) page.Add(source[offset + index]);
                chunks.Add(new ReadOnlyCollection<T>(page));
            }
            return new ReadOnlyCollection<IReadOnlyList<T>>(chunks);
        }

        private static int RequirePositive(int value, string name) {
            if (value <= 0) throw new ArgumentOutOfRangeException(name, "Value must be positive.");
            return value;
        }

        private static int RequireRange(int value, int minimum, int maximum, string name) {
            if (value < minimum || value > maximum) {
                throw new ArgumentOutOfRangeException(name,
                    "Value must be between " + minimum + " and " + maximum + ".");
            }
            return value;
        }
    }
}
