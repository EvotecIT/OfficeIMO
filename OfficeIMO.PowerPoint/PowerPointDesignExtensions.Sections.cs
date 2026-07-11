using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointDesignExtensions {
        private static void AddSectionGeometricCover(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointDesignerSlideOptions options, string title, string? subtitle, double slideWidthCm,
            double slideHeightCm) {
            slide.BackgroundColor = theme.AccentDarkColor;
            AddDiagonalPlanes(slide, theme, slideWidthCm, slideHeightCm, dark: true);
            AddChrome(slide, theme, slideWidthCm, slideHeightCm, dark: true, options);

            PowerPointTitleAccentStyle titleAccent = ResolveTitleAccentStyle(options,
                PowerPointSectionLayoutVariant.GeometricCover);
            double titleLeft = 1.85;
            double titleTop = slideHeightCm * 0.47;
            double titleWidth = slideWidthCm * 0.58;
            AddText(slide, title, titleLeft, titleTop, titleWidth, 1.35, 40,
                theme.AccentContrastColor, theme.HeadingFontName, bold: true);
            AddSectionTitleAccent(slide, theme, titleAccent, titleLeft, titleTop, titleWidth, 1.35, dark: true);

            if (!string.IsNullOrWhiteSpace(subtitle)) {
                AddText(slide, subtitle!, 1.9, slideHeightCm * 0.59, slideWidthCm * 0.52, 0.8, 15,
                    theme.AccentLightColor, theme.BodyFontName);
            }

            if (ShouldShowDirectionMotif(options)) {
                AddDirectionMotif(slide, options, 1.95, slideHeightCm * 0.67, 11, 0.46, theme.WarningColor);
            }
        }

        private static void AddSectionEditorialRail(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointDesignerSlideOptions options, string title, string? subtitle, double slideWidthCm,
            double slideHeightCm) {
            slide.BackgroundColor = theme.BackgroundColor;
            AddSubtleLightBackground(slide, theme, slideWidthCm, slideHeightCm);
            AddChrome(slide, theme, slideWidthCm, slideHeightCm, dark: false, options);

            PowerPointAutoShape rail = slide.AddRectangleCm(1.45, 1.85, 0.18, slideHeightCm - 3.8,
                "Section Editorial Rail");
            rail.FillColor = theme.AccentColor;
            rail.OutlineColor = theme.AccentColor;

            PowerPointAutoShape block = slide.AddRectangleCm(1.9, 3.65, slideWidthCm * 0.42, 0.18,
                "Section Editorial Rule");
            block.FillColor = theme.WarningColor;
            block.OutlineColor = theme.WarningColor;

            PowerPointTitleAccentStyle titleAccent = ResolveTitleAccentStyle(options,
                PowerPointSectionLayoutVariant.EditorialRail);
            double titleLeft = 1.9;
            double titleTop = 2.15;
            double titleWidth = slideWidthCm * 0.55;
            AddText(slide, title, titleLeft, titleTop, titleWidth, 1.2, 38,
                theme.PrimaryTextColor, theme.HeadingFontName, bold: true);
            AddSectionTitleAccent(slide, theme, titleAccent, titleLeft, titleTop, titleWidth, 1.2, dark: false);

            if (!string.IsNullOrWhiteSpace(subtitle)) {
                AddText(slide, subtitle!, 1.95, 4.15, slideWidthCm * 0.47, 0.8, 14,
                    theme.SecondaryTextColor, theme.BodyFontName);
            }

            PowerPointAutoShape accentPanel = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram,
                slideWidthCm * 0.64, 0, slideWidthCm * 0.26, slideHeightCm, "Section Editorial Accent Plane");
            accentPanel.FillColor = theme.AccentLightColor;
            accentPanel.FillTransparency = 30;
            accentPanel.OutlineColor = theme.AccentLightColor;

            if (ShouldShowDirectionMotif(options)) {
                AddDirectionMotif(slide, options, slideWidthCm - 5.25, 2.05, 10, 0.36, theme.AccentColor,
                    flip: true);
            }
        }

        private static void AddSectionPoster(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointDesignerSlideOptions options, string title, string? subtitle, double slideWidthCm,
            double slideHeightCm) {
            slide.BackgroundColor = theme.AccentDarkColor;
            AddChrome(slide, theme, slideWidthCm, slideHeightCm, dark: true, options);

            PowerPointAutoShape wash = slide.AddRectangleCm(0, 0, slideWidthCm, slideHeightCm,
                "Section Poster Wash");
            wash.FillColor = theme.AccentColor;
            wash.FillTransparency = 50;
            wash.OutlineColor = theme.AccentColor;
            wash.SendToBack();

            PowerPointAutoShape frame = slide.AddRectangleCm(1.45, 1.55, slideWidthCm - 2.9, slideHeightCm - 3.2,
                "Section Poster Frame");
            frame.FillColor = theme.AccentDarkColor;
            frame.FillTransparency = 100;
            frame.OutlineColor = theme.AccentLightColor;
            frame.OutlineWidthPoints = 0.7;

            PowerPointTitleAccentStyle titleAccent = ResolveTitleAccentStyle(options,
                PowerPointSectionLayoutVariant.Poster);
            PowerPointTextBox titleBox = AddText(slide, title, 2.4, slideHeightCm * 0.42, slideWidthCm - 4.8, 1.4,
                42, theme.AccentContrastColor, theme.HeadingFontName, bold: true);
            CenterText(titleBox);
            AddSectionTitleAccent(slide, theme, titleAccent, 2.4, slideHeightCm * 0.42, slideWidthCm - 4.8, 1.4,
                dark: true, centered: true);

            if (!string.IsNullOrWhiteSpace(subtitle)) {
                PowerPointTextBox subtitleBox = AddText(slide, subtitle!, 4.1, slideHeightCm * 0.58,
                    slideWidthCm - 8.2, 0.65, 14, theme.AccentLightColor, theme.BodyFontName);
                CenterText(subtitleBox);
            }

            if (ShouldShowDirectionMotif(options)) {
                AddDirectionMotif(slide, options, slideWidthCm * 0.39, slideHeightCm * 0.68, 12, 0.4,
                    theme.WarningColor);
            }
        }

        private static void AddCaseStudyColumns(PowerPointSlide slide, PowerPointDesignTheme theme, string clientTitle,
            IReadOnlyList<PowerPointCaseStudySection> sections, double slideWidthCm) {
            double left = 1.45;
            double top = 1.75;
            double gutter = 0.85;
            double width = slideWidthCm - 2.9;
            PowerPointLayoutBox[] columns = PowerPointLayoutBox
                .FromCentimeters(left, top, width, 6.15)
                .SplitColumnsCm(sections.Count, gutter);

            for (int i = 0; i < sections.Count; i++) {
                PowerPointLayoutBox box = columns[i];
                PowerPointCaseStudySection section = sections[i];
                AddText(slide, section.Heading, box.LeftCm, box.TopCm, box.WidthCm, 0.55, i == 0 ? 19 : 11,
                    theme.PrimaryTextColor, theme.HeadingFontName, bold: true);

                if (i == 0) {
                    PowerPointAutoShape rule = slide.AddRectangleCm(box.LeftCm, box.TopCm + 0.78, box.WidthCm, 0.025,
                        "Case Study Client Rule");
                    rule.FillColor = theme.PanelBorderColor;
                    rule.OutlineColor = theme.PanelBorderColor;
                }

                string body = i == 0 ? clientTitle + Environment.NewLine + section.Body : section.Body;
                PowerPointTextBox bodyBox = AddText(slide, body, box.LeftCm, box.TopCm + 1.05, box.WidthCm, 4.5,
                    i == 0 ? 13 : 10, i == 0 ? theme.PrimaryTextColor : theme.SecondaryTextColor,
                    theme.BodyFontName, bold: i == 0);
                bodyBox.TextAutoFit = PowerPointTextAutoFit.Normal;
                bodyBox.TextAutoFitOptions = new PowerPointTextAutoFitOptions(fontScalePercent: 80, lineSpaceReductionPercent: 20);
            }
        }

        private static void AddCaseStudyEditorialSplit(PowerPointSlide slide, PowerPointDesignTheme theme,
            string clientTitle, IReadOnlyList<PowerPointCaseStudySection> sections, PowerPointCaseStudySlideOptions options,
            IReadOnlyList<PowerPointMetric> metrics, double slideWidthCm, double slideHeightCm) {
            AddText(slide, clientTitle, 1.45, 1.6, slideWidthCm * 0.58, 1.35, 22,
                theme.PrimaryTextColor, theme.HeadingFontName, bold: true);

            PowerPointAutoShape titleRule = slide.AddRectangleCm(1.47, 3.18, slideWidthCm * 0.24, 0.08,
                "Case Study Editorial Rule");
            titleRule.FillColor = theme.AccentColor;
            titleRule.OutlineColor = theme.AccentColor;

            int textCount = Math.Min(sections.Count, 4);
            PowerPointLayoutBox[,] boxes = PowerPointLayoutBox
                .FromCentimeters(1.45, 3.75, slideWidthCm * 0.54, 5.05)
                .SplitGridCm(textCount > 2 ? 2 : 1, textCount > 1 ? 2 : 1, 0.55, 0.55);

            for (int i = 0; i < textCount; i++) {
                int row = i / (textCount > 1 ? 2 : 1);
                int column = i % (textCount > 1 ? 2 : 1);
                PowerPointLayoutBox box = boxes[row, column];
                PowerPointCaseStudySection section = sections[i];

                PowerPointAutoShape panel = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm, box.HeightCm,
                    "Case Study Editorial Section " + (i + 1));
                panel.FillColor = theme.PanelColor;
                panel.OutlineColor = theme.PanelBorderColor;
                panel.OutlineWidthPoints = 0.45;
                panel.SetShadow("000000", blurPoints: 3, distancePoints: 0.8, angleDegrees: 90, transparencyPercent: 90);

                PowerPointAutoShape accent = slide.AddRectangleCm(box.LeftCm, box.TopCm, 0.12, box.HeightCm,
                    "Case Study Editorial Accent " + (i + 1));
                accent.FillColor = GetAccent(theme, i);
                accent.OutlineColor = GetAccent(theme, i);

                AddText(slide, section.Heading, box.LeftCm + 0.45, box.TopCm + 0.38, box.WidthCm - 0.75, 0.45, 11,
                    theme.PrimaryTextColor, theme.HeadingFontName, bold: true);
                PowerPointTextBox body = AddText(slide, section.Body, box.LeftCm + 0.45, box.TopCm + 1.0,
                    box.WidthCm - 0.75, box.HeightCm - 1.25, 9, theme.SecondaryTextColor, theme.BodyFontName);
                body.TextAutoFitOptions = new PowerPointTextAutoFitOptions(fontScalePercent: 82, lineSpaceReductionPercent: 18);
            }

            double visualLeft = slideWidthCm * 0.68;
            double visualTop = 2.15;
            double visualWidth = slideWidthCm - visualLeft - 1.45;
            double visualHeight = 4.75;
            AddVisualFrame(slide, theme, options.VisualImagePath, visualLeft, visualTop, visualWidth, visualHeight,
                options.VisualFrameVariant, options.DesignIntent);

            if (metrics.Count > 0) {
                PowerPointAutoShape metricBand = slide.AddRectangleCm(visualLeft, visualTop + visualHeight + 0.55,
                    visualWidth, 1.8, "Case Study Editorial Metric Band");
                metricBand.FillColor = theme.AccentColor;
                metricBand.OutlineColor = theme.AccentColor;
                metricBand.SetShadow("000000", blurPoints: 3, distancePoints: 0.8, angleDegrees: 90, transparencyPercent: 88);
                AddMetrics(slide, theme, metrics, visualLeft + 0.35, visualTop + visualHeight + 0.8,
                    visualWidth - 0.7, 1.35);
            }

            AddTags(slide, theme, options.Tags, 1.45, slideHeightCm - 2.05, slideWidthCm * 0.52, 0.55);
        }

        private static void AddCaseStudyBand(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointCaseStudySlideOptions options, IReadOnlyList<PowerPointMetric> metrics,
            double slideWidthCm, double slideHeightCm) {
            double bandTop = slideHeightCm * 0.55;
            double bandHeight = slideHeightCm * 0.36;
            PowerPointAutoShape band = slide.AddRectangleCm(1.2, bandTop, slideWidthCm - 2.4, bandHeight,
                "Case Study Visual Band");
            band.FillColor = theme.AccentColor;
            band.OutlineColor = theme.AccentColor;
            band.SetSoftEdges(1.1);

            PowerPointAutoShape wash = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, 8.0, bandTop + 0.25,
                7.2, bandHeight - 0.5, "Case Study Band Wash");
            wash.FillColor = theme.Accent2Color;
            wash.FillTransparency = 45;
            wash.OutlineColor = theme.Accent2Color;

            AddText(slide, options.BrandText ?? string.Empty, 2.0, bandTop + 0.85, 4.4, 0.7, 20,
                theme.AccentContrastColor, theme.HeadingFontName, bold: true);

            if (!string.IsNullOrWhiteSpace(options.BandLabel)) {
                AddText(slide, options.BandLabel!, 2.0, bandTop + bandHeight - 1.35, 6.8, 0.5, 17,
                    theme.AccentLightColor, theme.HeadingFontName, bold: true);
            }

            if (!string.IsNullOrWhiteSpace(options.PersonImagePath)) {
                AddPictureIfExists(slide, options.PersonImagePath!, 9.2, bandTop - 1.0, 5.1, bandHeight + 0.4, crop: true);
            }

            AddMetrics(slide, theme, metrics, slideWidthCm * 0.46, bandTop + 1.75, slideWidthCm * 0.22, 1.65);
            AddVisualFrame(slide, theme, options.VisualImagePath, slideWidthCm - 8.6, bandTop + 0.9, 6.8, bandHeight - 1.5,
                options.VisualFrameVariant, options.DesignIntent);
            AddTags(slide, theme, options.Tags, 9.4, bandTop + bandHeight - 1.15, slideWidthCm - 12.6, 0.7);
        }

    }
}
