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
    ///     Deterministic variants that adjust a composition preset without changing its semantic regions.
    /// </summary>
    public enum PowerPointCompositionVariant {
        /// <summary>Choose a stable variant from the active design intent.</summary>
        Auto,
        /// <summary>Use the preset's default region arrangement.</summary>
        Standard,
        /// <summary>Mirror the preset horizontally while preserving region roles.</summary>
        Mirrored,
        /// <summary>Give the visual region the leading side of the composition.</summary>
        VisualLead,
        /// <summary>Move metric or proof regions earlier in the reading order.</summary>
        EvidenceLead
    }

    /// <summary>
    ///     Resolved layout regions for a raw designer composition preset.
    /// </summary>
    public sealed class PowerPointCompositionLayout {
        internal PowerPointCompositionLayout(PowerPointCompositionPreset preset, PowerPointCompositionVariant variant,
            PowerPointLayoutBox content, PowerPointLayoutBox primary, PowerPointLayoutBox secondary,
            PowerPointLayoutBox visual, PowerPointLayoutBox metrics, PowerPointLayoutBox supporting,
            IReadOnlyList<PowerPointLayoutBox> grid) {
            Preset = preset;
            Variant = variant;
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
        ///     Variant used after resolving Auto.
        /// </summary>
        public PowerPointCompositionVariant Variant { get; }

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
            PowerPointCompositionVariant variant, PowerPointLayoutBox content, PowerPointSlideDensity density) {
            double gutter = GutterFor(density);
            PowerPointCompositionVariant resolvedVariant = variant == PowerPointCompositionVariant.Auto
                ? PowerPointCompositionVariant.Standard
                : variant;

            switch (preset) {
                case PowerPointCompositionPreset.VisualSplit:
                    return CreateVisualSplit(content, gutter, resolvedVariant);
                case PowerPointCompositionPreset.MetricStory:
                    return CreateMetricStory(content, gutter, density, resolvedVariant);
                case PowerPointCompositionPreset.DashboardGrid:
                    return CreateDashboardGrid(content, gutter, resolvedVariant);
                default:
                    return CreateBalancedColumns(content, gutter, density, resolvedVariant);
            }
        }

        private static PowerPointCompositionLayout CreateBalancedColumns(PowerPointLayoutBox content, double gutter,
            PowerPointSlideDensity density, PowerPointCompositionVariant variant) {
            double metricsHeight = MetricHeightFor(content, density);
            bool evidenceLead = variant == PowerPointCompositionVariant.EvidenceLead;
            PowerPointLayoutBox metrics = Box(content.LeftCm,
                evidenceLead ? content.TopCm : content.BottomCm - metricsHeight,
                content.WidthCm, metricsHeight);
            PowerPointLayoutBox main = Box(content.LeftCm,
                evidenceLead ? metrics.BottomCm + gutter : content.TopCm,
                content.WidthCm, content.HeightCm - metricsHeight - gutter);
            PowerPointLayoutBox[] columns = main.SplitColumnsCm(2, gutter);
            bool mirrored = IsMirrored(variant);
            PowerPointLayoutBox primary = mirrored ? columns[1] : columns[0];
            PowerPointLayoutBox secondary = mirrored ? columns[0] : columns[1];

            return new PowerPointCompositionLayout(PowerPointCompositionPreset.BalancedColumns, variant, content,
                primary, secondary, secondary, metrics, metrics,
                new[] { primary, secondary, metrics });
        }

        private static PowerPointCompositionLayout CreateVisualSplit(PowerPointLayoutBox content, double gutter,
            PowerPointCompositionVariant variant) {
            bool visualLead = IsMirrored(variant);
            double primaryWidth = variant == PowerPointCompositionVariant.VisualLead
                ? content.WidthCm * 0.36
                : content.WidthCm * 0.42;
            double visualWidth = content.WidthCm - primaryWidth - gutter;
            PowerPointLayoutBox primary = visualLead
                ? Box(content.LeftCm + visualWidth + gutter, content.TopCm, primaryWidth, content.HeightCm)
                : Box(content.LeftCm, content.TopCm, primaryWidth, content.HeightCm);
            PowerPointLayoutBox visual = visualLead
                ? Box(content.LeftCm, content.TopCm, visualWidth, content.HeightCm)
                : Box(content.LeftCm + primaryWidth + gutter, content.TopCm, visualWidth, content.HeightCm);
            PowerPointLayoutBox supporting = Box(primary.LeftCm, primary.BottomCm - Math.Min(1.45, primary.HeightCm * 0.28),
                primary.WidthCm, Math.Min(1.45, primary.HeightCm * 0.28));

            return new PowerPointCompositionLayout(PowerPointCompositionPreset.VisualSplit, variant, content,
                primary, supporting, visual, supporting, supporting,
                new[] { primary, visual, supporting });
        }

        private static PowerPointCompositionLayout CreateMetricStory(PowerPointLayoutBox content, double gutter,
            PowerPointSlideDensity density, PowerPointCompositionVariant variant) {
            double metricsHeight = MetricHeightFor(content, density) + 0.25;
            bool evidenceLead = variant == PowerPointCompositionVariant.EvidenceLead;
            bool visualLead = IsMirrored(variant);
            double storyTop = evidenceLead ? content.TopCm + metricsHeight + gutter : content.TopCm;
            double storyHeight = content.HeightCm - metricsHeight - gutter;
            double visualWidth = variant == PowerPointCompositionVariant.VisualLead
                ? content.WidthCm * 0.46
                : content.WidthCm * 0.38;
            double primaryWidth = content.WidthCm - visualWidth - gutter;
            PowerPointLayoutBox primary = visualLead
                ? Box(content.LeftCm + visualWidth + gutter, storyTop, primaryWidth, storyHeight)
                : Box(content.LeftCm, storyTop, primaryWidth, storyHeight);
            PowerPointLayoutBox visual = visualLead
                ? Box(content.LeftCm, storyTop, visualWidth, storyHeight)
                : Box(primary.RightCm + gutter, storyTop, visualWidth, storyHeight);
            PowerPointLayoutBox metrics = Box(content.LeftCm,
                evidenceLead ? content.TopCm : content.BottomCm - metricsHeight,
                content.WidthCm, metricsHeight);

            return new PowerPointCompositionLayout(PowerPointCompositionPreset.MetricStory, variant, content,
                primary, visual, visual, metrics, metrics,
                new[] { primary, visual, metrics });
        }

        private static PowerPointCompositionLayout CreateDashboardGrid(PowerPointLayoutBox content, double gutter,
            PowerPointCompositionVariant variant) {
            PowerPointLayoutBox[,] grid = content.SplitGridCm(2, 2, gutter, gutter);
            PowerPointLayoutBox primary;
            PowerPointLayoutBox secondary;
            PowerPointLayoutBox visual;
            PowerPointLayoutBox metrics;

            switch (variant) {
                case PowerPointCompositionVariant.Mirrored:
                    primary = grid[0, 1];
                    secondary = grid[0, 0];
                    visual = grid[1, 1];
                    metrics = grid[1, 0];
                    break;
                case PowerPointCompositionVariant.VisualLead:
                    primary = grid[0, 1];
                    secondary = grid[1, 1];
                    visual = grid[0, 0];
                    metrics = grid[1, 0];
                    break;
                case PowerPointCompositionVariant.EvidenceLead:
                    primary = grid[0, 1];
                    secondary = grid[1, 0];
                    visual = grid[1, 1];
                    metrics = grid[0, 0];
                    break;
                default:
                    primary = grid[0, 0];
                    secondary = grid[0, 1];
                    visual = grid[1, 0];
                    metrics = grid[1, 1];
                    break;
            }

            PowerPointLayoutBox[] flat = {
                primary,
                secondary,
                visual,
                metrics
            };

            return new PowerPointCompositionLayout(PowerPointCompositionPreset.DashboardGrid, variant, content,
                primary, secondary, visual, metrics, metrics, flat);
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

        private static bool IsMirrored(PowerPointCompositionVariant variant) {
            return variant == PowerPointCompositionVariant.Mirrored ||
                variant == PowerPointCompositionVariant.VisualLead;
        }
    }
}
