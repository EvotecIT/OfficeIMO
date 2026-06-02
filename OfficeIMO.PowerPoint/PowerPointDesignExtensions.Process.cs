using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public static partial class PowerPointDesignExtensions {
        internal static void AddProcessTimeline(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointProcessStep> steps, PowerPointProcessSlideOptions options,
            double slideWidthCm, double slideHeightCm) {
            PowerPointProcessLayoutVariant variant = ResolveProcessVariant(options, steps);
            if (variant == PowerPointProcessLayoutVariant.NumberedColumns) {
                AddProcessColumns(slide, theme, steps, options, slideWidthCm, slideHeightCm);
                return;
            }

            double left = options.DesignIntent.Density == PowerPointSlideDensity.Relaxed ? 2.35 : 2.1;
            double top = slideHeightCm * 0.47;
            double width = slideWidthCm - 4.2;
            double height = 4.7;
            AddProcessRailTimeline(slide, theme, steps, options,
                PowerPointLayoutBox.FromCentimeters(left, top, width, height));
        }

        internal static void AddProcessTimeline(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointProcessStep> steps, PowerPointProcessSlideOptions options,
            PowerPointLayoutBox bounds) {
            PowerPointProcessLayoutVariant variant = ResolveProcessVariant(options, steps);
            if (variant == PowerPointProcessLayoutVariant.NumberedColumns) {
                AddProcessColumns(slide, theme, steps, options, bounds);
                return;
            }

            AddProcessRailTimeline(slide, theme, steps, options, bounds);
        }

        private static void AddProcessRailTimeline(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointProcessStep> steps, PowerPointProcessSlideOptions options,
            PowerPointLayoutBox bounds) {
            int count = steps.Count;
            PowerPointLayoutBox[] boxes = PowerPointLayoutBox
                .FromCentimeters(bounds.LeftCm, bounds.TopCm, bounds.WidthCm, bounds.HeightCm)
                .SplitColumnsCm(count, count > 5 ? 0.45 : 0.75);

            double nodeSize = count > 5 ? 0.95 : 1.16;
            double railY = bounds.TopCm + nodeSize / 2;
            double railStart = boxes[0].LeftCm + nodeSize / 2;
            double railEnd = boxes[count - 1].LeftCm + nodeSize / 2;
            PowerPointProcessConnectorStyle connectorStyle = ResolveProcessConnectorStyle(options, steps);
            AddProcessConnectors(slide, theme, boxes, nodeSize, railY, railStart, railEnd, connectorStyle);

            for (int i = 0; i < count; i++) {
                PowerPointLayoutBox box = boxes[i];
                PowerPointProcessStep step = steps[i];
                string number = !string.IsNullOrWhiteSpace(step.Number) ? step.Number! : (i + 1) + ".";
                AddProcessNode(slide, theme, i, box.LeftCm, box.TopCm, nodeSize, number);
                AddText(slide, step.Title, box.LeftCm, box.TopCm + 1.55, box.WidthCm, 0.7, 13,
                    theme.AccentContrastColor, theme.HeadingFontName, bold: true);
                AddText(slide, step.Body, box.LeftCm, box.TopCm + 2.45, box.WidthCm, 1.7, 10,
                    theme.AccentLightColor, theme.BodyFontName);
            }
        }

        private static void AddProcessColumns(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointProcessStep> steps, PowerPointProcessSlideOptions options,
            double slideWidthCm, double slideHeightCm) {
            double left = 1.85;
            double top = slideHeightCm * 0.45;
            double width = slideWidthCm - 3.7;
            double height = 4.85;
            AddProcessColumns(slide, theme, steps, options, PowerPointLayoutBox.FromCentimeters(left, top, width, height));
        }

        private static void AddProcessColumns(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointProcessStep> steps, PowerPointProcessSlideOptions options,
            PowerPointLayoutBox bounds) {
            int count = steps.Count;
            double gutter = options.DesignIntent.Density == PowerPointSlideDensity.Compact ? 0.35 : 0.6;
            PowerPointLayoutBox[] boxes = PowerPointLayoutBox
                .FromCentimeters(bounds.LeftCm, bounds.TopCm, bounds.WidthCm, bounds.HeightCm)
                .SplitColumnsCm(count, gutter);

            for (int i = 0; i < count; i++) {
                PowerPointLayoutBox box = boxes[i];
                PowerPointProcessStep step = steps[i];
                string number = !string.IsNullOrWhiteSpace(step.Number) ? step.Number! : (i + 1).ToString("00");

                PowerPointAutoShape panel = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm, box.HeightCm,
                    "Process Column " + (i + 1));
                panel.FillColor = theme.AccentColor;
                panel.FillTransparency = 72;
                panel.OutlineColor = theme.AccentLightColor;
                panel.OutlineWidthPoints = 0.35;

                PowerPointAutoShape rule = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm * 0.34, 0.08,
                    "Process Column Rule " + (i + 1));
                rule.FillColor = GetAccent(theme, i);
                rule.OutlineColor = GetAccent(theme, i);

                AddText(slide, number.TrimEnd('.'), box.LeftCm + 0.26, box.TopCm + 0.45, box.WidthCm - 0.52, 0.85,
                    count > 5 ? 20 : 25, theme.AccentContrastColor, theme.HeadingFontName, bold: true);
                AddText(slide, step.Title, box.LeftCm + 0.26, box.TopCm + 1.65, box.WidthCm - 0.52, 0.72, 13,
                    theme.AccentContrastColor, theme.HeadingFontName, bold: true);
                AddText(slide, step.Body, box.LeftCm + 0.26, box.TopCm + 2.55, box.WidthCm - 0.52, 1.55, 10,
                    theme.AccentLightColor, theme.BodyFontName);
            }
        }
    }
}
