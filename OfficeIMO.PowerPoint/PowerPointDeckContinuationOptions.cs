using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint {
    /// <summary>Configures deterministic continuation of semantic deck-plan content.</summary>
    public sealed class PowerPointDeckContinuationOptions {
        private int _caseStudySectionsPerSlide = PowerPointDeckPlanLimits.MaxCaseStudySections;
        private int _caseStudyMetricsPerSlide = PowerPointDeckPlanLimits.MaxCaseStudyMetrics;
        private int _processStepsPerSlide = PowerPointDeckPlanLimits.MaxProcessSteps;
        private int _cardsPerSlide = 8;
        private int _logosPerSlide = PowerPointDeckPlanLimits.DenseLogoWallItems;
        private int _locationsPerSlide = PowerPointDeckPlanLimits.VisibleCoveragePins;
        private int _capabilitySectionsPerSlide = PowerPointDeckPlanLimits.DenseCapabilitySections;
        private int _appendixRowsPerSlide = DefaultAppendixRowsPerSlide;
        internal const int DefaultAppendixRowsPerSlide = 8;

        /// <summary>
        ///     True when the caller chose a row limit. Otherwise appendix tables are left whole and the slide layout splits
        ///     them from the measured row height, so every part fills the slide evenly.
        /// </summary>
        internal bool AppendixRowsPerSlideSet { get; private set; }

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

        /// <summary>Maximum cards per generated page (default 8: two rows of four short cards). The layout still
        /// moves cards to another page when their text needs the room.</summary>
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
            set {
                _appendixRowsPerSlide = RequireRange(value, 1, PowerPointDeckPlanLimits.MaxAppendixTableRows, nameof(AppendixRowsPerSlide));
                AppendixRowsPerSlideSet = true;
            }
        }

        /// <summary>
        ///     Formats continuation titles. Parameters are the original title, one-based page number, and page count.
        ///     The first page retains the original title.
        /// </summary>
        public string? ContinuationTitleFormat { get; set; }

        internal string CreateTitle(string title, int pageIndex, int pageCount) {
            if (pageIndex == 0 || pageCount <= 1) return title;
            if (string.IsNullOrWhiteSpace(ContinuationTitleFormat)) {
                return pageCount == 2
                    ? title + " (continued)"
                    : title + " (continued " + (pageIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" +
                      pageCount.ToString(System.Globalization.CultureInfo.InvariantCulture) + ")";
            }

            return string.Format(System.Globalization.CultureInfo.InvariantCulture, ContinuationTitleFormat, title,
                pageIndex + 1, pageCount);
        }

        internal string? CreateSeed(string? seed, string title, int pageIndex) {
            if (pageIndex == 0) return seed;
            string value = string.IsNullOrWhiteSpace(seed) ? title : seed!;
            return value.Trim() + "-continuation-" + (pageIndex + 1);
        }

        /// <summary>Splits into the fewest pages of at most <paramref name="size" /> items, balanced (17 at 14 → 9 + 8).</summary>
        internal static IReadOnlyList<IReadOnlyList<T>> Chunk<T>(IReadOnlyList<T> source, int size) {
            var chunks = new List<IReadOnlyList<T>>();
            if (source.Count == 0) return new ReadOnlyCollection<IReadOnlyList<T>>(chunks);
            int pages = (source.Count + size - 1) / size;
            int baseSize = source.Count / pages, extra = source.Count % pages, offset = 0;
            for (int page = 0; page < pages; page++) {
                int count = baseSize + (page < extra ? 1 : 0);
                var items = new List<T>(count);
                for (int index = 0; index < count; index++) items.Add(source[offset + index]);
                chunks.Add(new ReadOnlyCollection<T>(items));
                offset += count;
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
