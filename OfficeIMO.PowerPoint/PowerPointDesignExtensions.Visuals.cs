using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public static partial class PowerPointDesignExtensions {
        internal static void AddMetrics(PowerPointSlide slide, PowerPointDesignTheme theme, IReadOnlyList<PowerPointMetric> metrics,
            double leftCm, double topCm, double widthCm, double heightCm) {
            if (metrics.Count == 0) {
                return;
            }

            int count = Math.Min(metrics.Count, 3);
            PowerPointLayoutBox[] boxes = PowerPointLayoutBox
                .FromCentimeters(leftCm, topCm, widthCm, heightCm)
                .SplitColumnsCm(count, 0.45);
            double valueHeight = Math.Min(0.88, heightCm * 0.52);
            double labelTopOffset = valueHeight + Math.Min(0.14, heightCm * 0.08);
            double labelHeight = Math.Max(0.32, heightCm - labelTopOffset);
            int valueFontSize = heightCm < 1.6 ? 24 : 29;
            int labelFontSize = heightCm < 1.6 ? 8 : 9;
            for (int i = 0; i < count; i++) {
                PowerPointMetric metric = metrics[i];
                PowerPointLayoutBox box = boxes[i];
                int resolvedValueFontSize = ResolveMetricValueFontSize(metric.Value, box.WidthCm, valueFontSize);
                PowerPointTextBox value = AddText(slide, metric.Value, box.LeftCm, box.TopCm, box.WidthCm, valueHeight,
                    resolvedValueFontSize,
                    theme.AccentContrastColor, theme.HeadingFontName, bold: true);
                CenterText(value);
                int resolvedLabelFontSize = ResolveMetricLabelFontSize(metric.Label, box.WidthCm, labelFontSize);
                PowerPointTextBox label = AddText(slide, metric.Label, box.LeftCm, box.TopCm + labelTopOffset,
                    box.WidthCm, labelHeight, resolvedLabelFontSize,
                    theme.AccentContrastColor, theme.BodyFontName, bold: true);
                CenterText(label);
            }
        }

        internal static void AddVisualFrame(PowerPointSlide slide, PowerPointDesignTheme theme, string? imagePath,
            double leftCm, double topCm, double widthCm, double heightCm, PowerPointDesignIntent? intent = null) {
            AddVisualFrame(slide, theme, imagePath, leftCm, topCm, widthCm, heightCm,
                PowerPointVisualFrameVariant.Auto, intent);
        }

        internal static void AddVisualFrame(PowerPointSlide slide, PowerPointDesignTheme theme, string? imagePath,
            double leftCm, double topCm, double widthCm, double heightCm, PowerPointVisualFrameVariant variant,
            PowerPointDesignIntent? intent = null) {
            PowerPointAutoShape frame = slide.AddRectangleCm(leftCm, topCm, widthCm, heightCm, "Case Study Visual Frame");
            frame.FillColor = theme.AccentDarkColor;
            frame.OutlineColor = theme.AccentDarkColor;
            frame.OutlineWidthPoints = 0;
            frame.SetShadow("000000", blurPoints: 5, distancePoints: 1.5, angleDegrees: 90, transparencyPercent: 82);

            PowerPointVisualFrameVariant resolvedVariant = ResolveVisualPlaceholderVariant(variant, intent);
            if (!string.IsNullOrWhiteSpace(imagePath) && File.Exists(imagePath)) {
                AddVisualPicture(slide, theme, imagePath!, leftCm + 0.08, topCm + 0.08,
                    widthCm - 0.16, heightCm - 0.16, resolvedVariant);
                return;
            }

            AddVisualPlaceholder(slide, theme, leftCm + 0.08, topCm + 0.08, widthCm - 0.16, heightCm - 0.16,
                resolvedVariant, intent);
        }

        private static void AddVisualPicture(PowerPointSlide slide, PowerPointDesignTheme theme, string imagePath,
            double leftCm, double topCm, double widthCm, double heightCm, PowerPointVisualFrameVariant variant) {
            if (variant == PowerPointVisualFrameVariant.DeviceMockup) {
                AddVisualDeviceChrome(slide, theme, leftCm, topCm, widthCm, heightCm);
                GetVisualDeviceContentBounds(leftCm, topCm, widthCm, heightCm, out double contentLeft,
                    out double contentTop, out double contentWidth, out double contentHeight);
                AddPictureIfExists(slide, imagePath, contentLeft, contentTop, contentWidth, contentHeight,
                    crop: true);
                return;
            }

            if (variant == PowerPointVisualFrameVariant.ProofBoard) {
                AddVisualProofMat(slide, theme, leftCm, topCm, widthCm, heightCm);
                AddPictureIfExists(slide, imagePath, leftCm + widthCm * 0.10, topCm + heightCm * 0.14,
                    widthCm * 0.78, heightCm * 0.68, crop: true);
                return;
            }

            AddPictureIfExists(slide, imagePath, leftCm, topCm, widthCm, heightCm, crop: true);
        }

        private static void AddVisualPlaceholder(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm, PowerPointVisualFrameVariant variant,
            PowerPointDesignIntent? intent) {
            PowerPointAutoShape surface = slide.AddRectangleCm(leftCm, topCm, widthCm, heightCm,
                "Case Study Visual Surface");
            surface.FillColor = theme.AccentDarkColor;
            surface.OutlineColor = theme.AccentDarkColor;

            PowerPointVisualFrameVariant resolvedVariant = ResolveVisualPlaceholderVariant(variant, intent);
            if (resolvedVariant == PowerPointVisualFrameVariant.Collage) {
                AddVisualCollagePlaceholder(slide, theme, leftCm, topCm, widthCm, heightCm);
                return;
            }
            if (resolvedVariant == PowerPointVisualFrameVariant.Diagram) {
                AddVisualDiagramPlaceholder(slide, theme, leftCm, topCm, widthCm, heightCm);
                return;
            }
            if (resolvedVariant == PowerPointVisualFrameVariant.DeviceMockup) {
                AddVisualDeviceMockupPlaceholder(slide, theme, leftCm, topCm, widthCm, heightCm);
                return;
            }
            if (resolvedVariant == PowerPointVisualFrameVariant.ProofBoard) {
                AddVisualProofBoardPlaceholder(slide, theme, leftCm, topCm, widthCm, heightCm);
                return;
            }

            AddVisualDashboardPlaceholder(slide, theme, leftCm, topCm, widthCm, heightCm);
        }

        private static void AddVisualDashboardPlaceholder(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            PowerPointAutoShape glow = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, leftCm + widthCm * 0.08,
                topCm, widthCm * 0.42, heightCm, "Case Study Visual Wash");
            glow.FillColor = theme.Accent2Color;
            glow.FillTransparency = 68;
            glow.OutlineColor = theme.Accent2Color;

            double panelTop = topCm + heightCm * 0.24;
            double panelLeft = leftCm + widthCm * 0.17;
            double panelWidth = widthCm * 0.66;
            double panelHeight = heightCm * 0.42;
            PowerPointAutoShape panel = slide.AddRectangleCm(panelLeft, panelTop, panelWidth, panelHeight,
                "Case Study Visual Content Panel");
            panel.FillColor = theme.AccentColor;
            panel.FillTransparency = 35;
            panel.OutlineColor = theme.AccentColor;
            panel.OutlineWidthPoints = 0;

            PowerPointAutoShape imageBlock = slide.AddRectangleCm(panelLeft + panelWidth * 0.08, panelTop + panelHeight * 0.18,
                panelWidth * 0.34, panelHeight * 0.56, "Case Study Visual Image Block");
            imageBlock.FillColor = theme.AccentLightColor;
            imageBlock.FillTransparency = 10;
            imageBlock.OutlineColor = theme.AccentLightColor;

            for (int i = 0; i < 3; i++) {
                double barWidth = panelWidth * (i == 0 ? 0.36 : 0.28);
                PowerPointAutoShape bar = slide.AddRectangleCm(panelLeft + panelWidth * 0.49,
                    panelTop + panelHeight * (0.23 + i * 0.18), barWidth, 0.045,
                    "Case Study Visual Line " + (i + 1));
                bar.FillColor = theme.AccentLightColor;
                bar.FillTransparency = i == 0 ? 5 : 35;
                bar.OutlineColor = theme.AccentLightColor;
            }

            PowerPointAutoShape baseLine = slide.AddLineCm(leftCm + widthCm * 0.18, topCm + heightCm * 0.84,
                leftCm + widthCm * 0.82, topCm + heightCm * 0.84, "Case Study Visual Base Line");
            baseLine.OutlineColor = theme.AccentLightColor;
            baseLine.OutlineWidthPoints = 0.9;
        }

        private static void AddVisualCollagePlaceholder(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            PowerPointAutoShape wash = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, leftCm + widthCm * 0.58,
                topCm, widthCm * 0.24, heightCm, "Visual Collage Wash");
            wash.FillColor = theme.Accent2Color;
            wash.FillTransparency = 62;
            wash.OutlineColor = theme.Accent2Color;

            AddVisualTile(slide, theme, leftCm + widthCm * 0.10, topCm + heightCm * 0.17,
                widthCm * 0.38, heightCm * 0.42, "Visual Collage Tile 1", theme.AccentLightColor, 62);
            AddVisualTile(slide, theme, leftCm + widthCm * 0.43, topCm + heightCm * 0.09,
                widthCm * 0.43, heightCm * 0.28, "Visual Collage Tile 2", theme.AccentColor, 70);
            AddVisualTile(slide, theme, leftCm + widthCm * 0.36, topCm + heightCm * 0.55,
                widthCm * 0.48, heightCm * 0.25, "Visual Collage Tile 3", theme.Accent3Color, 72);

            for (int i = 0; i < 3; i++) {
                double dot = 0.18 - i * 0.025;
                PowerPointAutoShape node = slide.AddEllipseCm(leftCm + widthCm * (0.24 + i * 0.19),
                    topCm + heightCm * (0.77 - i * 0.08), dot, dot, "Visual Collage Marker " + (i + 1));
                node.FillColor = GetAccent(theme, i);
                node.OutlineColor = theme.AccentLightColor;
                node.OutlineWidthPoints = 0.55;
            }
        }

        private static void AddVisualDiagramPlaceholder(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            PowerPointAutoShape rail = slide.AddRectangleCm(leftCm + widthCm * 0.12, topCm + heightCm * 0.49,
                widthCm * 0.76, 0.045, "Visual Diagram Rail");
            rail.FillColor = theme.AccentLightColor;
            rail.FillTransparency = 15;
            rail.OutlineColor = theme.AccentLightColor;

            for (int i = 0; i < 4; i++) {
                double cx = leftCm + widthCm * (0.16 + i * 0.22);
                double cy = topCm + heightCm * (i % 2 == 0 ? 0.35 : 0.61);
                PowerPointAutoShape node = slide.AddEllipseCm(cx, cy, 0.58, 0.58, "Visual Diagram Node " + (i + 1));
                node.FillColor = GetAccent(theme, i);
                node.FillTransparency = i == 0 ? 0 : 22;
                node.OutlineColor = theme.AccentLightColor;
                node.OutlineWidthPoints = 0.55;

                PowerPointAutoShape label = slide.AddRectangleCm(cx + 0.78, cy + 0.22,
                    widthCm * (i == 0 ? 0.22 : 0.16), 0.04, "Visual Diagram Label " + (i + 1));
                label.FillColor = theme.AccentLightColor;
                label.FillTransparency = i == 0 ? 0 : 35;
                label.OutlineColor = theme.AccentLightColor;
            }

            PowerPointAutoShape plate = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, leftCm + widthCm * 0.20,
                topCm + heightCm * 0.14, widthCm * 0.50, heightCm * 0.72, "Visual Diagram Plate");
            plate.FillColor = theme.AccentColor;
            plate.FillTransparency = 82;
            plate.OutlineColor = theme.AccentLightColor;
            plate.OutlineWidthPoints = 0.45;
        }

        private static void AddVisualDeviceMockupPlaceholder(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            AddVisualDeviceChrome(slide, theme, leftCm, topCm, widthCm, heightCm);

            GetVisualDeviceContentBounds(leftCm, topCm, widthCm, heightCm, out double contentLeft,
                out double contentTop, out double contentWidth, out double contentHeight);

            PowerPointAutoShape hero = slide.AddRectangleCm(contentLeft, contentTop, contentWidth, contentHeight,
                "Visual Device Hero Area");
            hero.FillColor = theme.AccentColor;
            hero.FillTransparency = 34;
            hero.OutlineColor = theme.AccentColor;
            hero.OutlineWidthPoints = 0;

            for (int i = 0; i < 3; i++) {
                PowerPointAutoShape bar = slide.AddRectangleCm(contentLeft + contentWidth * 0.09,
                    contentTop + contentHeight * (0.22 + i * 0.18),
                    contentWidth * (i == 0 ? 0.38 : 0.25), 0.06, "Visual Device Content Line " + (i + 1));
                bar.FillColor = theme.AccentLightColor;
                bar.FillTransparency = i == 0 ? 12 : 42;
                bar.OutlineColor = theme.AccentLightColor;
                bar.OutlineWidthPoints = 0;
            }

            for (int i = 0; i < 3; i++) {
                PowerPointAutoShape tile = slide.AddRectangleCm(contentLeft + contentWidth * (0.58 + i * 0.11),
                    contentTop + contentHeight * 0.22, contentWidth * 0.08, contentHeight * 0.46,
                    "Visual Device Metric Tile " + (i + 1));
                tile.FillColor = GetAccent(theme, i);
                tile.FillTransparency = 18 + i * 10;
                tile.OutlineColor = theme.AccentLightColor;
                tile.OutlineWidthPoints = 0.35;
            }
        }

        private static void AddVisualDeviceChrome(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            PowerPointAutoShape screen = slide.AddRectangleCm(leftCm + widthCm * 0.06, topCm + heightCm * 0.08,
                widthCm * 0.88, heightCm * 0.76, "Visual Device Screen");
            screen.FillColor = theme.AccentDarkColor;
            screen.FillTransparency = 18;
            screen.OutlineColor = theme.AccentLightColor;
            screen.OutlineWidthPoints = 0.45;

            PowerPointAutoShape topBar = slide.AddRectangleCm(leftCm + widthCm * 0.06, topCm + heightCm * 0.08,
                widthCm * 0.88, 0.38, "Visual Device Chrome Bar");
            topBar.FillColor = theme.AccentLightColor;
            topBar.FillTransparency = 8;
            topBar.OutlineColor = theme.AccentLightColor;
            topBar.OutlineWidthPoints = 0;

            for (int i = 0; i < 3; i++) {
                PowerPointAutoShape dot = slide.AddEllipseCm(leftCm + widthCm * 0.10 + i * 0.18,
                    topCm + heightCm * 0.08 + 0.12, 0.09, 0.09, "Visual Device Chrome Dot " + (i + 1));
                dot.FillColor = GetAccent(theme, i);
                dot.FillTransparency = i == 0 ? 0 : 18;
                dot.OutlineColor = dot.FillColor;
                dot.OutlineWidthPoints = 0;
            }

            PowerPointAutoShape baseLine = slide.AddLineCm(leftCm + widthCm * 0.18, topCm + heightCm * 0.89,
                leftCm + widthCm * 0.82, topCm + heightCm * 0.89, "Visual Device Base");
            baseLine.OutlineColor = theme.AccentLightColor;
            baseLine.OutlineWidthPoints = 1.1;
        }

        private static void GetVisualDeviceContentBounds(double leftCm, double topCm, double widthCm,
            double heightCm, out double contentLeft, out double contentTop, out double contentWidth,
            out double contentHeight) {
            double screenLeft = leftCm + widthCm * 0.06;
            double screenTop = topCm + heightCm * 0.08;
            double screenWidth = widthCm * 0.88;
            double screenHeight = heightCm * 0.76;
            double topInset = Math.Min(0.46, Math.Max(0.12, screenHeight * 0.34));
            double bottomInset = Math.Min(0.08, Math.Max(0.04, screenHeight * 0.06));

            contentLeft = screenLeft + widthCm * 0.05;
            contentTop = screenTop + topInset;
            contentWidth = screenWidth - widthCm * 0.10;
            contentHeight = Math.Max(0.18, screenHeight - topInset - bottomInset);
        }

        private static void AddVisualProofBoardPlaceholder(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            AddVisualProofMat(slide, theme, leftCm, topCm, widthCm, heightCm);

            AddVisualProofPanel(slide, theme, leftCm + widthCm * 0.10, topCm + heightCm * 0.15,
                widthCm * 0.36, heightCm * 0.55, "Visual Proof Primary Panel", theme.PanelColor, 0);
            AddVisualProofPanel(slide, theme, leftCm + widthCm * 0.53, topCm + heightCm * 0.18,
                widthCm * 0.34, heightCm * 0.24, "Visual Proof Detail Panel 1", theme.AccentLightColor, 10);
            AddVisualProofPanel(slide, theme, leftCm + widthCm * 0.58, topCm + heightCm * 0.48,
                widthCm * 0.28, heightCm * 0.22, "Visual Proof Detail Panel 2", theme.AccentLightColor, 24);

            for (int i = 0; i < 3; i++) {
                PowerPointAutoShape rule = slide.AddRectangleCm(leftCm + widthCm * 0.15,
                    topCm + heightCm * (0.78 + i * 0.055), widthCm * (0.46 - i * 0.06), 0.045,
                    "Visual Proof Caption Line " + (i + 1));
                rule.FillColor = theme.AccentLightColor;
                rule.FillTransparency = 18 + i * 18;
                rule.OutlineColor = theme.AccentLightColor;
                rule.OutlineWidthPoints = 0;
            }
        }

        private static void AddVisualProofMat(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            PowerPointAutoShape mat = slide.AddRectangleCm(leftCm + widthCm * 0.05, topCm + heightCm * 0.08,
                widthCm * 0.90, heightCm * 0.78, "Visual Proof Mat");
            mat.FillColor = theme.AccentContrastColor;
            mat.FillTransparency = 8;
            mat.OutlineColor = theme.AccentLightColor;
            mat.OutlineWidthPoints = 0.4;

            PowerPointAutoShape accent = slide.AddRectangleCm(leftCm + widthCm * 0.05, topCm + heightCm * 0.08,
                widthCm * 0.90, 0.08, "Visual Proof Mat Accent");
            accent.FillColor = theme.Accent2Color;
            accent.OutlineColor = theme.Accent2Color;
            accent.OutlineWidthPoints = 0;
        }

        private static void AddVisualProofPanel(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm, string name, string fillColor,
            int fillTransparency) {
            PowerPointAutoShape panel = slide.AddRectangleCm(leftCm, topCm, widthCm, heightCm, name);
            panel.FillColor = fillColor;
            panel.FillTransparency = fillTransparency;
            panel.OutlineColor = theme.PanelBorderColor;
            panel.OutlineWidthPoints = 0.35;
        }

        private static void AddVisualTile(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm, string name, string fillColor,
            int fillTransparency) {
            PowerPointAutoShape tile = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, leftCm, topCm, widthCm, heightCm, name);
            tile.FillColor = fillColor;
            tile.FillTransparency = fillTransparency;
            tile.OutlineColor = theme.AccentLightColor;
            tile.OutlineWidthPoints = 0.45;
        }

        private static PowerPointVisualFrameVariant ResolveVisualPlaceholderVariant(PowerPointVisualFrameVariant variant,
            PowerPointDesignIntent? intent) {
            if (variant != PowerPointVisualFrameVariant.Auto) {
                return variant;
            }
            if (intent == null) {
                return PowerPointVisualFrameVariant.Dashboard;
            }
            if (string.IsNullOrWhiteSpace(intent.Seed) &&
                intent.Mood == PowerPointDesignMood.Corporate &&
                intent.Density == PowerPointSlideDensity.Balanced &&
                intent.VisualStyle == PowerPointVisualStyle.Geometric) {
                return PowerPointVisualFrameVariant.Dashboard;
            }
            if (intent.Mood == PowerPointDesignMood.Energetic) {
                return PowerPointVisualFrameVariant.DeviceMockup;
            }
            if (intent.Mood == PowerPointDesignMood.Editorial) {
                return PowerPointVisualFrameVariant.ProofBoard;
            }
            if (intent.VisualStyle == PowerPointVisualStyle.Soft) {
                return PowerPointVisualFrameVariant.Collage;
            }
            if (intent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointVisualFrameVariant.Diagram;
            }

            return intent.Pick(5, "visual-placeholder") switch {
                0 => PowerPointVisualFrameVariant.Dashboard,
                1 => PowerPointVisualFrameVariant.Collage,
                2 => PowerPointVisualFrameVariant.Diagram,
                3 => PowerPointVisualFrameVariant.DeviceMockup,
                _ => PowerPointVisualFrameVariant.ProofBoard
            };
        }

        private static void AddTags(PowerPointSlide slide, PowerPointDesignTheme theme, IList<string> tags,
            double leftCm, double topCm, double widthCm, double heightCm) {
            if (tags.Count == 0) {
                return;
            }

            int count = Math.Min(tags.Count, 7);
            PowerPointLayoutBox[] boxes = PowerPointLayoutBox
                .FromCentimeters(leftCm, topCm, widthCm, heightCm)
                .SplitColumnsCm(count, 0.28);
            for (int i = 0; i < count; i++) {
                PowerPointLayoutBox box = boxes[i];
                PowerPointAutoShape pill = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm, box.HeightCm,
                    "Case Study Tag " + (i + 1));
                pill.FillColor = i == count - 1 ? theme.Accent2Color : theme.AccentColor;
                pill.FillTransparency = i == count - 1 ? 0 : 40;
                pill.OutlineColor = theme.AccentLightColor;
                pill.OutlineWidthPoints = 0.6;

                PowerPointTextBox label = AddText(slide, tags[i], box.LeftCm + 0.08, box.TopCm + 0.16,
                    box.WidthCm - 0.16, box.HeightCm - 0.25, 8, theme.AccentContrastColor, theme.BodyFontName,
                    bold: i == count - 1);
                CenterText(label);
            }
        }
    }
}
