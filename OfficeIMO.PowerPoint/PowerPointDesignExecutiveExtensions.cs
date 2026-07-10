using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    public static partial class PowerPointDesignExtensions {
        /// <summary>Adds an executive-summary slide with metric-led or decision-brief composition.</summary>
        public static PowerPointSlide AddDesignerExecutiveSummarySlide(this PowerPointPresentation presentation,
            string title, string? subtitle, PowerPointExecutiveSummaryContent content,
            PowerPointDesignTheme? theme = null, PowerPointExecutiveSummarySlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("Title cannot be empty.", nameof(title));
            if (content == null) throw new ArgumentNullException(nameof(content));
            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointExecutiveSummarySlideOptions resolved = options ?? new PowerPointExecutiveSummarySlideOptions();
            PowerPointSlide slide = AddDesignerSlide(presentation, resolved);
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            PrepareLightStorySlide(slide, resolvedTheme, resolved, title, subtitle, width, height);
            PowerPointExecutiveSummaryLayoutVariant variant = ResolveExecutiveVariant(resolved, content);
            if (variant == PowerPointExecutiveSummaryLayoutVariant.DecisionBrief) {
                AddExecutiveDecisionBrief(slide, resolvedTheme, content, width, height, resolved);
            } else {
                AddExecutiveMetricLead(slide, resolvedTheme, content, width, height, resolved);
            }
            return FinalizeDesignerAccessibility(slide, title);
        }

        /// <summary>Adds a closing slide with statement or explicit action-panel composition.</summary>
        public static PowerPointSlide AddDesignerClosingSlide(this PowerPointPresentation presentation,
            string title, PowerPointClosingContent content, PowerPointDesignTheme? theme = null,
            PowerPointClosingSlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("Title cannot be empty.", nameof(title));
            if (content == null) throw new ArgumentNullException(nameof(content));
            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointClosingSlideOptions resolved = options ?? new PowerPointClosingSlideOptions();
            PowerPointSlide slide = AddDesignerSlide(presentation, resolved);
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            PowerPointClosingLayoutVariant variant = ResolveClosingVariant(resolved, content);
            if (variant == PowerPointClosingLayoutVariant.ActionPanel) {
                PrepareLightStorySlide(slide, resolvedTheme, resolved, title, null, width, height);
                AddClosingActionPanel(slide, resolvedTheme, content, width, height);
            } else {
                slide.BackgroundColor = resolvedTheme.AccentDarkColor;
                AddChrome(slide, resolvedTheme, width, height, dark: true, resolved);
                AddClosingStatement(slide, resolvedTheme, title, content, width, height);
            }
            return FinalizeDesignerAccessibility(slide, title);
        }

        private static void PrepareLightStorySlide(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointDesignerSlideOptions options, string title, string? subtitle, double width, double height) {
            slide.BackgroundColor = theme.BackgroundColor;
            AddSubtleLightBackground(slide, theme, width, height);
            AddChrome(slide, theme, width, height, dark: false, options);
            AddText(slide, title, 1.5, 1.42, width * 0.68, 1.02, 29,
                theme.PrimaryTextColor, theme.HeadingFontName, bold: true);
            if (!string.IsNullOrWhiteSpace(subtitle)) {
                AddText(slide, subtitle!, 1.55, 2.62, width * 0.66, 0.52, 12,
                    theme.SecondaryTextColor, theme.BodyFontName, bold: true);
            }
        }

        internal static PowerPointExecutiveSummaryLayoutVariant ResolveExecutiveVariant(
            PowerPointExecutiveSummarySlideOptions options, PowerPointExecutiveSummaryContent content) {
            if (options.Variant != PowerPointExecutiveSummaryLayoutVariant.Auto) return options.Variant;
            if (options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.Compact || content.Points.Count > 4)
                return PowerPointExecutiveSummaryLayoutVariant.DecisionBrief;
            return options.DesignIntent.Pick(2, "executive-summary") == 0
                ? PowerPointExecutiveSummaryLayoutVariant.MetricLead
                : PowerPointExecutiveSummaryLayoutVariant.DecisionBrief;
        }

        internal static PowerPointClosingLayoutVariant ResolveClosingVariant(PowerPointClosingSlideOptions options,
            PowerPointClosingContent content) {
            if (options.Variant != PowerPointClosingLayoutVariant.Auto) return options.Variant;
            if (!string.IsNullOrWhiteSpace(content.CallToAction) ||
                options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.Compact)
                return PowerPointClosingLayoutVariant.ActionPanel;
            return PowerPointClosingLayoutVariant.Statement;
        }

        private static void AddExecutiveMetricLead(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointExecutiveSummaryContent content, double width, double height,
            PowerPointExecutiveSummarySlideOptions options) {
            double top = 3.55;
            if (!string.IsNullOrWhiteSpace(content.Lead)) {
                AddText(slide, content.Lead!, 1.55, top, width - 3.1, 0.8, 18,
                    theme.AccentDarkColor, theme.HeadingFontName, bold: true);
                top += 1.05;
            }
            PowerPointLayoutBox metrics = PowerPointLayoutBox.FromCentimeters(1.5, top, width - 3, 2.15);
            AddMetrics(slide, theme, content.Metrics.Take(4).ToList(), metrics.LeftCm, metrics.TopCm,
                metrics.WidthCm, metrics.HeightCm);
            double cardTop = content.Metrics.Count == 0 ? top : top + 2.6;
            PowerPointLayoutBox cards = PowerPointLayoutBox.FromCentimeters(1.5, cardTop, width - 3,
                Math.Max(2.4, height - top - 4.1));
            var cardOptions = new PowerPointCardGridSlideOptions { MaxColumns = 4,
                Variant = PowerPointCardGridLayoutVariant.SoftTiles, DesignIntent = options.DesignIntent };
            if (content.Points.Count > 0) {
                AddCardGrid(slide, theme, content.Points.Take(4).ToList(), cardOptions,
                    PowerPointCardGridLayoutVariant.SoftTiles, cards);
            }
        }

        private static void AddExecutiveDecisionBrief(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointExecutiveSummaryContent content, double width, double height,
            PowerPointExecutiveSummarySlideOptions options) {
            PowerPointLayoutBox body = PowerPointLayoutBox.FromCentimeters(1.5, 3.65, width - 3, height - 5.3);
            PowerPointLayoutBox[] columns = body.SplitColumnsCm(2, 0.75);
            PowerPointAutoShape leadPanel = slide.AddRectangleCm(columns[0].LeftCm, columns[0].TopCm,
                columns[0].WidthCm, columns[0].HeightCm, "Executive Decision Panel");
            leadPanel.FillColor = theme.AccentDarkColor;
            leadPanel.OutlineColor = theme.AccentDarkColor;
            string lead = string.IsNullOrWhiteSpace(content.Lead)
                ? content.Points.FirstOrDefault()?.Title ?? "Decision required"
                : content.Lead!;
            AddText(slide, lead, columns[0].LeftCm + 0.65, columns[0].TopCm + 0.65,
                columns[0].WidthCm - 1.3, 2.0, 22, theme.AccentContrastColor, theme.HeadingFontName, bold: true);
            AddMetrics(slide, theme, content.Metrics.Take(3).ToList(), columns[0].LeftCm + 0.45,
                columns[0].BottomCm - 2.25, columns[0].WidthCm - 0.9, 1.7);
            var cardOptions = new PowerPointCardGridSlideOptions { MaxColumns = 1,
                Variant = PowerPointCardGridLayoutVariant.AccentTop, DesignIntent = options.DesignIntent };
            if (content.Points.Count > 0) {
                AddCardGrid(slide, theme, content.Points.Take(4).ToList(), cardOptions,
                    PowerPointCardGridLayoutVariant.AccentTop, columns[1]);
            }
        }

        private static void AddClosingStatement(PowerPointSlide slide, PowerPointDesignTheme theme, string title,
            PowerPointClosingContent content, double width, double height) {
            AddText(slide, title, 1.8, 1.5, width - 3.6, 0.65, 11,
                theme.AccentLightColor, theme.BodyFontName, bold: true);
            AddText(slide, content.Statement, 2.0, height * 0.31, width - 4.0, 2.8, 34,
                theme.AccentContrastColor, theme.HeadingFontName, bold: true);
            if (!string.IsNullOrWhiteSpace(content.CallToAction)) {
                AddText(slide, content.CallToAction!, 2.05, height * 0.62, width - 4.1, 0.8, 16,
                    theme.AccentLightColor, theme.BodyFontName, bold: true);
            }
            if (!string.IsNullOrWhiteSpace(content.Contact)) {
                AddText(slide, content.Contact!, 2.05, height * 0.72, width - 4.1, 0.55, 11,
                    theme.AccentLightColor, theme.BodyFontName);
            }
        }

        private static void AddClosingActionPanel(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointClosingContent content, double width, double height) {
            PowerPointLayoutBox body = PowerPointLayoutBox.FromCentimeters(1.5, 3.75, width - 3, height - 5.4);
            PowerPointLayoutBox[] columns = body.SplitColumnsCm(2, 0.85);
            AddText(slide, content.Statement, columns[0].LeftCm, columns[0].TopCm + 0.65,
                columns[0].WidthCm, columns[0].HeightCm - 1.3, 28,
                theme.PrimaryTextColor, theme.HeadingFontName, bold: true);
            PowerPointAutoShape panel = slide.AddRectangleCm(columns[1].LeftCm, columns[1].TopCm,
                columns[1].WidthCm, columns[1].HeightCm, "Closing Action Panel");
            panel.FillColor = theme.AccentColor;
            panel.OutlineColor = theme.AccentColor;
            string action = content.CallToAction ?? "Continue the conversation";
            int actionFontSize = action.Length > 42 ? 17 : action.Length > 28 ? 19 : 21;
            AddText(slide, action, columns[1].LeftCm + 0.7,
                columns[1].TopCm + 0.8, columns[1].WidthCm - 1.4, 2.65, actionFontSize,
                theme.AccentContrastColor, theme.HeadingFontName, bold: true);
            if (!string.IsNullOrWhiteSpace(content.Contact)) {
                AddText(slide, content.Contact!, columns[1].LeftCm + 0.7, columns[1].BottomCm - 1.55,
                    columns[1].WidthCm - 1.4, 0.7, 12, theme.AccentContrastColor, theme.BodyFontName);
            }
        }
    }
}
