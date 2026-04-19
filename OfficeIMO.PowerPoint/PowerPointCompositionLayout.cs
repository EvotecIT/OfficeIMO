using System;
using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Named composition layouts for raw designer slides.
    /// </summary>
    public enum PowerPointCompositionPreset {
        /// <summary>Choose a stable preset from the active design intent.</summary>
        Auto,
        /// <summary>Two balanced columns with a full-width metric/support band.</summary>
        BalancedColumns,
        /// <summary>Primary narrative region paired with a larger visual region.</summary>
        VisualSplit,
        /// <summary>Large narrative/visual regions with a strong metric strip.</summary>
        MetricStory,
        /// <summary>Four equally weighted regions for dashboards and dense summaries.</summary>
        DashboardGrid
    }

    /// <summary>
    ///     Resolved layout regions for a raw designer composition preset.
    /// </summary>
    public sealed class PowerPointCompositionLayout {
        internal PowerPointCompositionLayout(PowerPointCompositionPreset preset, PowerPointLayoutBox content,
            PowerPointLayoutBox primary, PowerPointLayoutBox secondary, PowerPointLayoutBox visual,
            PowerPointLayoutBox metrics, PowerPointLayoutBox supporting, IReadOnlyList<PowerPointLayoutBox> grid) {
            Preset = preset;
            Content = content;
            Primary = primary;
            Secondary = secondary;
            Visual = visual;
            Metrics = metrics;
            Supporting = supporting;
            Grid = grid;
        }

        /// <summary>
        ///     Preset used after resolving Auto.
        /// </summary>
        public PowerPointCompositionPreset Preset { get; }

        /// <summary>
        ///     Full content region used to derive the preset.
        /// </summary>
        public PowerPointLayoutBox Content { get; }

        /// <summary>
        ///     Main narrative or primary content region.
        /// </summary>
        public PowerPointLayoutBox Primary { get; }

        /// <summary>
        ///     Secondary narrative or supporting content region.
        /// </summary>
        public PowerPointLayoutBox Secondary { get; }

        /// <summary>
        ///     Preferred region for a visual frame, map, proof wall, or custom illustration.
        /// </summary>
        public PowerPointLayoutBox Visual { get; }

        /// <summary>
        ///     Preferred region for metrics or compact quantified proof.
        /// </summary>
        public PowerPointLayoutBox Metrics { get; }

        /// <summary>
        ///     Preferred region for a callout, note, or supporting text.
        /// </summary>
        public PowerPointLayoutBox Supporting { get; }

        /// <summary>
        ///     Reusable grid regions supplied by the preset.
        /// </summary>
        public IReadOnlyList<PowerPointLayoutBox> Grid { get; }

        internal static PowerPointCompositionLayout Create(PowerPointCompositionPreset preset,
            PowerPointLayoutBox content, PowerPointSlideDensity density) {
            double gutter = GutterFor(density);

            switch (preset) {
                case PowerPointCompositionPreset.VisualSplit:
                    return CreateVisualSplit(content, gutter);
                case PowerPointCompositionPreset.MetricStory:
                    return CreateMetricStory(content, gutter, density);
                case PowerPointCompositionPreset.DashboardGrid:
                    return CreateDashboardGrid(content, gutter);
                default:
                    return CreateBalancedColumns(content, gutter, density);
            }
        }

        private static PowerPointCompositionLayout CreateBalancedColumns(PowerPointLayoutBox content, double gutter,
            PowerPointSlideDensity density) {
            double metricsHeight = MetricHeightFor(content, density);
            PowerPointLayoutBox main = Box(content.LeftCm, content.TopCm, content.WidthCm,
                content.HeightCm - metricsHeight - gutter);
            PowerPointLayoutBox[] columns = main.SplitColumnsCm(2, gutter);
            PowerPointLayoutBox metrics = Box(content.LeftCm, content.BottomCm - metricsHeight,
                content.WidthCm, metricsHeight);

            return new PowerPointCompositionLayout(PowerPointCompositionPreset.BalancedColumns, content,
                columns[0], columns[1], columns[1], metrics, metrics,
                new[] { columns[0], columns[1], metrics });
        }

        private static PowerPointCompositionLayout CreateVisualSplit(PowerPointLayoutBox content, double gutter) {
            double primaryWidth = content.WidthCm * 0.42;
            PowerPointLayoutBox primary = Box(content.LeftCm, content.TopCm, primaryWidth, content.HeightCm);
            PowerPointLayoutBox visual = Box(content.LeftCm + primaryWidth + gutter, content.TopCm,
                content.WidthCm - primaryWidth - gutter, content.HeightCm);
            PowerPointLayoutBox supporting = Box(primary.LeftCm, primary.BottomCm - Math.Min(1.45, primary.HeightCm * 0.28),
                primary.WidthCm, Math.Min(1.45, primary.HeightCm * 0.28));

            return new PowerPointCompositionLayout(PowerPointCompositionPreset.VisualSplit, content,
                primary, supporting, visual, supporting, supporting,
                new[] { primary, visual, supporting });
        }

        private static PowerPointCompositionLayout CreateMetricStory(PowerPointLayoutBox content, double gutter,
            PowerPointSlideDensity density) {
            double metricsHeight = MetricHeightFor(content, density) + 0.25;
            double storyHeight = content.HeightCm - metricsHeight - gutter;
            double visualWidth = content.WidthCm * 0.38;
            PowerPointLayoutBox primary = Box(content.LeftCm, content.TopCm,
                content.WidthCm - visualWidth - gutter, storyHeight);
            PowerPointLayoutBox visual = Box(primary.RightCm + gutter, content.TopCm, visualWidth, storyHeight);
            PowerPointLayoutBox metrics = Box(content.LeftCm, content.BottomCm - metricsHeight,
                content.WidthCm, metricsHeight);

            return new PowerPointCompositionLayout(PowerPointCompositionPreset.MetricStory, content,
                primary, visual, visual, metrics, metrics,
                new[] { primary, visual, metrics });
        }

        private static PowerPointCompositionLayout CreateDashboardGrid(PowerPointLayoutBox content, double gutter) {
            PowerPointLayoutBox[,] grid = content.SplitGridCm(2, 2, gutter, gutter);
            PowerPointLayoutBox[] flat = {
                grid[0, 0],
                grid[0, 1],
                grid[1, 0],
                grid[1, 1]
            };

            return new PowerPointCompositionLayout(PowerPointCompositionPreset.DashboardGrid, content,
                grid[0, 0], grid[0, 1], grid[1, 0], grid[1, 1], grid[1, 1], flat);
        }

        private static PowerPointLayoutBox Box(double left, double top, double width, double height) {
            if (width <= 0) {
                throw new InvalidOperationException("Composition region width is not positive.");
            }
            if (height <= 0) {
                throw new InvalidOperationException("Composition region height is not positive.");
            }

            return PowerPointLayoutBox.FromCentimeters(left, top, width, height);
        }

        private static double GutterFor(PowerPointSlideDensity density) {
            switch (density) {
                case PowerPointSlideDensity.Compact:
                    return 0.45;
                case PowerPointSlideDensity.Relaxed:
                    return 0.85;
                default:
                    return 0.65;
            }
        }

        private static double MetricHeightFor(PowerPointLayoutBox content, PowerPointSlideDensity density) {
            double preferred;
            switch (density) {
                case PowerPointSlideDensity.Compact:
                    preferred = 1.15;
                    break;
                case PowerPointSlideDensity.Relaxed:
                    preferred = 1.65;
                    break;
                default:
                    preferred = 1.4;
                    break;
            }

            double max = content.HeightCm * 0.30;
            return preferred > max ? max : preferred;
        }
    }
}
