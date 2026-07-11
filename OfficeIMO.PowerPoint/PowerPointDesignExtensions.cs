using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     High-level slide composition helpers for building polished decks without hand-placing every shape.
    /// </summary>
    internal static partial class PowerPointDesignExtensions {
        /// <summary>
        ///     Applies the designer theme colors and fonts to all slide masters.
        /// </summary>
        public static PowerPointPresentation ApplyDesignerTheme(this PowerPointPresentation presentation,
            PowerPointDesignTheme? theme = null) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            presentation.ThemeName = resolvedTheme.Name;
            presentation.SetThemeColorsForAllMasters(new Dictionary<PowerPointThemeColor, string> {
                [PowerPointThemeColor.Dark1] = resolvedTheme.PrimaryTextColor,
                [PowerPointThemeColor.Light1] = resolvedTheme.BackgroundColor,
                [PowerPointThemeColor.Dark2] = resolvedTheme.AccentDarkColor,
                [PowerPointThemeColor.Light2] = resolvedTheme.SurfaceColor,
                [PowerPointThemeColor.Accent1] = resolvedTheme.AccentColor,
                [PowerPointThemeColor.Accent2] = resolvedTheme.Accent2Color,
                [PowerPointThemeColor.Accent3] = resolvedTheme.Accent3Color,
                [PowerPointThemeColor.Accent4] = resolvedTheme.WarningColor,
                [PowerPointThemeColor.Accent5] = resolvedTheme.PanelBorderColor,
                [PowerPointThemeColor.Accent6] = resolvedTheme.MutedTextColor
            });
            presentation.SetThemeFontsForAllMasters(new PowerPointThemeFontSet(
                resolvedTheme.HeadingFontName,
                resolvedTheme.BodyFontName,
                null,
                null,
                null,
                null));

            return presentation;
        }

        /// <summary>
        ///     Adds a full-bleed section/title slide with diagonal planes and optional footer chrome.
        /// </summary>
        public static PowerPointSlide AddDesignerSectionSlide(this PowerPointPresentation presentation, string title,
            string? subtitle = null, PowerPointDesignTheme? theme = null,
            PowerPointDesignerSlideOptions? options = null) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointDesignerSlideOptions resolvedOptions = options ?? new PowerPointDesignerSlideOptions();
            PowerPointSlide slide = AddDesignerSlide(presentation, resolvedOptions);
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            PowerPointSectionLayoutVariant variant = ResolveSectionVariant(resolvedOptions);

            if (variant == PowerPointSectionLayoutVariant.EditorialRail) {
                AddSectionEditorialRail(slide, resolvedTheme, resolvedOptions, title, subtitle, width, height);
            } else if (variant == PowerPointSectionLayoutVariant.Poster) {
                AddSectionPoster(slide, resolvedTheme, resolvedOptions, title, subtitle, width, height);
            } else {
                AddSectionGeometricCover(slide, resolvedTheme, resolvedOptions, title, subtitle, width, height);
            }

            return FinalizeDesignerAccessibility(slide, title);
        }

        /// <summary>
        ///     Adds a case-study slide with summary columns, a strong visual band, metrics, and optional tags.
        /// </summary>
        public static PowerPointSlide AddDesignerCaseStudySlide(this PowerPointPresentation presentation,
            string clientTitle, IEnumerable<PowerPointCaseStudySection> sections,
            IEnumerable<PowerPointMetric>? metrics = null,
            PowerPointDesignTheme? theme = null,
            PowerPointCaseStudySlideOptions? options = null) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }
            if (clientTitle == null) {
                throw new ArgumentNullException(nameof(clientTitle));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointCaseStudySlideOptions resolvedOptions = options ?? new PowerPointCaseStudySlideOptions();
            List<PowerPointCaseStudySection> sectionList = NormalizeSections(sections, 4, nameof(sections));
            List<PowerPointMetric> metricList = (metrics ?? Enumerable.Empty<PowerPointMetric>()).Where(m => m != null).ToList();

            PowerPointSlide slide = AddDesignerSlide(presentation, resolvedOptions);
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            slide.BackgroundColor = resolvedTheme.BackgroundColor;

            PowerPointCaseStudyLayoutVariant variant = ResolveCaseStudyVariant(resolvedOptions, sectionList, metricList);
            if (variant == PowerPointCaseStudyLayoutVariant.EditorialSplit) {
                AddSubtleLightBackground(slide, resolvedTheme, width, height);
                AddChrome(slide, resolvedTheme, width, height, dark: false, resolvedOptions);
                AddCaseStudyEditorialSplit(slide, resolvedTheme, clientTitle, sectionList, resolvedOptions, metricList, width, height);
            } else if (variant == PowerPointCaseStudyLayoutVariant.VisualHero) {
                AddSubtleLightBackground(slide, resolvedTheme, width, height);
                AddChrome(slide, resolvedTheme, width, height, dark: false, resolvedOptions);
                AddCaseStudyVisualHero(slide, resolvedTheme, clientTitle, sectionList, resolvedOptions, metricList, width, height);
            } else {
                AddChrome(slide, resolvedTheme, width, height, dark: false, resolvedOptions);
                AddCaseStudyColumns(slide, resolvedTheme, clientTitle, sectionList, width);
                AddCaseStudyBand(slide, resolvedTheme, resolvedOptions, metricList, width, height);
            }

            return FinalizeDesignerAccessibility(slide, clientTitle);
        }

        /// <summary>
        ///     Adds a card grid slide that automatically chooses rows and columns for the supplied content.
        /// </summary>
        public static PowerPointSlide AddDesignerCardGridSlide(this PowerPointPresentation presentation,
            string title, string? subtitle, IEnumerable<PowerPointCardContent> cards,
            PowerPointDesignTheme? theme = null,
            PowerPointCardGridSlideOptions? options = null) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointCardGridSlideOptions resolvedOptions = options ?? new PowerPointCardGridSlideOptions();
            List<PowerPointCardContent> cardList = NormalizeCards(cards);

            PowerPointSlide slide = AddDesignerSlide(presentation, resolvedOptions);
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            slide.BackgroundColor = resolvedTheme.BackgroundColor;

            AddSubtleLightBackground(slide, resolvedTheme, width, height);
            AddChrome(slide, resolvedTheme, width, height, dark: false, resolvedOptions);
            AddText(slide, title, 1.5, 1.45, width * 0.6, 1.0, 29,
                resolvedTheme.PrimaryTextColor, resolvedTheme.HeadingFontName, bold: true);
            if (!string.IsNullOrWhiteSpace(subtitle)) {
                AddText(slide, subtitle!, 1.55, 2.7, width * 0.58, 0.5, 12,
                    resolvedTheme.SecondaryTextColor, resolvedTheme.BodyFontName, bold: true);
            }

            AddCardGrid(slide, resolvedTheme, cardList, resolvedOptions,
                ResolveCardGridVariant(resolvedOptions, cardList), width, height);

            return FinalizeDesignerAccessibility(slide, title);
        }

        /// <summary>
        ///     Adds a dark process slide with a readable timeline and automatic spacing.
        /// </summary>
        public static PowerPointSlide AddDesignerProcessSlide(this PowerPointPresentation presentation,
            string title, string? subtitle, IEnumerable<PowerPointProcessStep> steps,
            PowerPointDesignTheme? theme = null,
            PowerPointProcessSlideOptions? options = null) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointProcessSlideOptions resolvedOptions = options ?? new PowerPointProcessSlideOptions();
            List<PowerPointProcessStep> stepList = NormalizeSteps(steps);

            PowerPointSlide slide = AddDesignerSlide(presentation, resolvedOptions);
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            slide.BackgroundColor = resolvedTheme.AccentDarkColor;

            if (resolvedOptions.ShowDiagonalPlanes) {
                AddDiagonalPlanes(slide, resolvedTheme, width, height, dark: true);
            }
            AddChrome(slide, resolvedTheme, width, height, dark: true, resolvedOptions);
            AddText(slide, title, 1.85, 1.45, width * 0.52, 1.1, 33,
                resolvedTheme.AccentContrastColor, resolvedTheme.HeadingFontName, bold: true);
            if (!string.IsNullOrWhiteSpace(subtitle)) {
                AddText(slide, subtitle!, 1.9, 2.78, width * 0.58, 0.5, 13,
                    resolvedTheme.AccentLightColor, resolvedTheme.BodyFontName, bold: true);
            }

            AddProcessTimeline(slide, resolvedTheme, stepList, resolvedOptions, width, height);
            return FinalizeDesignerAccessibility(slide, title);
        }

    }
}
