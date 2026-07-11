using System;
using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDeckPlan {
        /// <summary>Adds an executive-summary slide request.</summary>
        public PowerPointDeckPlan AddExecutiveSummary(string title, string? subtitle,
            PowerPointExecutiveSummaryContent content, string? seed = null,
            Action<PowerPointExecutiveSummarySlideOptions>? configure = null) =>
            Add(new PowerPointExecutiveSummaryPlanSlide(title, subtitle, content, seed, configure));

        /// <summary>Adds an editable chart-story slide request.</summary>
        public PowerPointDeckPlan AddChartStory(string title, string? subtitle,
            PowerPointChartStoryContent content, string? seed = null,
            Action<PowerPointChartStorySlideOptions>? configure = null) =>
            Add(new PowerPointChartStoryPlanSlide(title, subtitle, content, seed, configure));

        /// <summary>Adds a comparison slide request.</summary>
        public PowerPointDeckPlan AddComparison(string title, string? subtitle,
            IEnumerable<PowerPointComparisonItem> items, string? seed = null,
            Action<PowerPointComparisonSlideOptions>? configure = null) =>
            Add(new PowerPointComparisonPlanSlide(title, subtitle, items, seed, configure));

        /// <summary>Adds a semantic screenshot-story slide request.</summary>
        public PowerPointDeckPlan AddScreenshotStory(string title, string? subtitle, PowerPointImageAsset image,
            IEnumerable<string>? narrative = null, string? seed = null,
            Action<PowerPointScreenshotStorySlideOptions>? configure = null) =>
            Add(new PowerPointScreenshotStoryPlanSlide(title, subtitle, image, narrative, seed, configure));

        /// <summary>Adds a paginated editable appendix-table slide request.</summary>
        public PowerPointDeckPlan AddAppendixTable(string title, string? subtitle, PowerPointTableData data,
            string? seed = null, Action<PowerPointAppendixTableSlideOptions>? configure = null) =>
            Add(new PowerPointAppendixTablePlanSlide(title, subtitle, data, seed, configure));

        /// <summary>Adds an editable architecture slide request.</summary>
        public PowerPointDeckPlan AddArchitecture(string title, string? subtitle,
            PowerPointArchitectureContent content, string? seed = null,
            Action<PowerPointArchitectureSlideOptions>? configure = null) =>
            Add(new PowerPointArchitecturePlanSlide(title, subtitle, content, seed, configure));

        /// <summary>Adds a closing slide request.</summary>
        public PowerPointDeckPlan AddClosing(string title, PowerPointClosingContent content,
            string? seed = null, Action<PowerPointClosingSlideOptions>? configure = null) =>
            Add(new PowerPointClosingPlanSlide(title, content, seed, configure));
    }
}
