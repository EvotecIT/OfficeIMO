using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public static partial class PowerPointDesignExtensions {
        private static void AddSubtleLightBackground(PowerPointSlide slide, PowerPointDesignTheme theme,
            double slideWidthCm, double slideHeightCm) {
            PowerPointAutoShape diagonal = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, slideWidthCm * 0.28, 0,
                slideWidthCm * 0.22, slideHeightCm, "Designer Light Diagonal");
            diagonal.FillColor = theme.SurfaceColor;
            diagonal.FillTransparency = 35;
            diagonal.OutlineColor = theme.SurfaceColor;
            diagonal.SendToBack();
        }

        private static void AddDiagonalPlanes(PowerPointSlide slide, PowerPointDesignTheme theme, double slideWidthCm,
            double slideHeightCm, bool dark) {
            string baseColor = dark ? theme.AccentColor : theme.SurfaceColor;
            string second = dark ? theme.AccentDarkColor : theme.AccentLightColor;

            PowerPointAutoShape left = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, -1.0, 0,
                slideWidthCm * 0.48, slideHeightCm, "Designer Plane Left");
            left.FillColor = baseColor;
            left.FillTransparency = dark ? 18 : 60;
            left.OutlineColor = baseColor;

            PowerPointAutoShape middle = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, slideWidthCm * 0.46, 0,
                slideWidthCm * 0.27, slideHeightCm, "Designer Plane Middle");
            middle.FillColor = second;
            middle.FillTransparency = dark ? 35 : 72;
            middle.OutlineColor = second;
        }

        private static void AddSectionTitleAccent(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointTitleAccentStyle style, double titleLeftCm, double titleTopCm, double titleWidthCm,
            double titleHeightCm, bool dark, bool centered = false) {
            if (style == PowerPointTitleAccentStyle.None) {
                return;
            }

            string accent = dark ? theme.WarningColor : theme.AccentColor;
            string secondary = dark ? theme.AccentLightColor : theme.WarningColor;
            switch (style) {
                case PowerPointTitleAccentStyle.SideRule:
                    PowerPointAutoShape sideRule = slide.AddRectangleCm(titleLeftCm - 0.28, titleTopCm + 0.16,
                        0.08, Math.Max(0.82, titleHeightCm * 0.7), "Section Title Accent Side Rule");
                    sideRule.FillColor = accent;
                    sideRule.OutlineColor = accent;
                    sideRule.OutlineWidthPoints = 0;
                    break;
                case PowerPointTitleAccentStyle.KickerRule:
                    double kickerWidth = Math.Min(3.2, titleWidthCm * 0.28);
                    double kickerLeft = centered ? titleLeftCm + (titleWidthCm - kickerWidth) / 2d : titleLeftCm;
                    PowerPointAutoShape kicker = slide.AddLineCm(kickerLeft, Math.Max(0.65, titleTopCm - 0.22),
                        kickerLeft + kickerWidth, Math.Max(0.65, titleTopCm - 0.22),
                        "Section Title Accent Kicker Rule");
                    kicker.OutlineColor = secondary;
                    kicker.OutlineWidthPoints = 1.4;
                    break;
                case PowerPointTitleAccentStyle.Underline:
                    double underlineWidth = Math.Min(4.2, titleWidthCm * 0.36);
                    double underlineLeft = centered ? titleLeftCm + (titleWidthCm - underlineWidth) / 2d : titleLeftCm;
                    double underlineTop = titleTopCm + Math.Min(0.86, Math.Max(0.52, titleHeightCm * 0.62));
                    PowerPointAutoShape underline = slide.AddRectangleCm(underlineLeft, underlineTop, underlineWidth,
                        0.08, "Section Title Accent Underline");
                    underline.FillColor = accent;
                    underline.OutlineColor = accent;
                    underline.OutlineWidthPoints = 0;
                    break;
            }
        }

        private static void AddChrome(PowerPointSlide slide, PowerPointDesignTheme theme, double slideWidthCm,
            double slideHeightCm, bool dark, PowerPointDesignerSlideOptions options) {
            string text = dark ? theme.AccentLightColor : theme.MutedTextColor;
            string footer = dark ? theme.AccentContrastColor : theme.AccentDarkColor;

            if (!string.IsNullOrWhiteSpace(options.Eyebrow)) {
                AddText(slide, options.Eyebrow!, 1.8, 1.05, 8.0, 0.35, 8, text, theme.BodyFontName);
            }

            if (!string.IsNullOrWhiteSpace(options.FooterLeft)) {
                AddText(slide, options.FooterLeft!, 1.75, slideHeightCm - 1.35, 6.0, 0.45, 16, footer,
                    theme.HeadingFontName, bold: true);
            }

            if (!string.IsNullOrWhiteSpace(options.FooterRight)) {
                PowerPointTextBox right = AddText(slide, options.FooterRight!, slideWidthCm - 5.4,
                    slideHeightCm - 1.35, 4.1, 0.45, 12, footer, theme.HeadingFontName, bold: true);
                RightAlignText(right);
            }

            if (ShouldShowDirectionMotif(options) && !dark) {
                AddDirectionMotif(slide, options, slideWidthCm - 4.9, 1.48, 12, 0.35, theme.AccentColor,
                    flip: true);
            }
        }

        private static bool ShouldShowDirectionMotif(PowerPointDesignerSlideOptions options) {
            return options.ShowDirectionMotif &&
                   ResolveDirectionMotifStyle(options) != PowerPointDirectionMotifStyle.None;
        }

        private static void AddProcessRail(PowerPointSlide slide, PowerPointDesignTheme theme,
            double startXCm, double endXCm, double yCm) {
            PowerPointAutoShape rail = slide.AddLineCm(startXCm, yCm, endXCm, yCm, "Process Rail");
            rail.OutlineColor = theme.AccentLightColor;
            rail.OutlineWidthPoints = 1.1;
        }

        private static void AddProcessConnectors(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointLayoutBox> boxes, double nodeSize, double yCm, double railStartCm,
            double railEndCm, PowerPointProcessConnectorStyle style) {
            if (style == PowerPointProcessConnectorStyle.None) {
                return;
            }

            if (style == PowerPointProcessConnectorStyle.ContinuousRail) {
                AddProcessRail(slide, theme, railStartCm, railEndCm, yCm);
                return;
            }

            for (int i = 0; i < boxes.Count - 1; i++) {
                double start = boxes[i].LeftCm + nodeSize + 0.16;
                double end = boxes[i + 1].LeftCm - 0.16;
                if (end <= start) {
                    continue;
                }

                if (style == PowerPointProcessConnectorStyle.StepDots) {
                    AddProcessConnectorDots(slide, theme, i, start, end, yCm);
                } else {
                    AddProcessConnectorArrow(slide, theme, i, start, end, yCm);
                }
            }
        }

        private static void AddProcessConnectorArrow(PowerPointSlide slide, PowerPointDesignTheme theme, int index,
            double startXCm, double endXCm, double yCm) {
            PowerPointAutoShape connector = slide.AddLineCm(startXCm, yCm, endXCm, yCm,
                "Process Connector " + (index + 1));
            connector.OutlineColor = GetAccent(theme, index);
            connector.OutlineWidthPoints = 1.2;
            connector.SetLineEnds(null, A.LineEndValues.Triangle, A.LineEndWidthValues.Small,
                A.LineEndLengthValues.Small);
        }

        private static void AddProcessConnectorDots(PowerPointSlide slide, PowerPointDesignTheme theme, int index,
            double startXCm, double endXCm, double yCm) {
            const int dotCount = 4;
            double spacing = (endXCm - startXCm) / (dotCount + 1);
            for (int dot = 0; dot < dotCount; dot++) {
                double center = startXCm + spacing * (dot + 1);
                PowerPointAutoShape marker = slide.AddEllipseCm(center - 0.055, yCm - 0.055, 0.11, 0.11,
                    "Process Connector Dot " + (index + 1) + "-" + (dot + 1));
                marker.FillColor = GetAccent(theme, index);
                marker.FillTransparency = 12 + dot * 8;
                marker.OutlineColor = marker.FillColor;
                marker.OutlineWidthPoints = 0;
            }
        }

        private static void AddProcessNode(PowerPointSlide slide, PowerPointDesignTheme theme, int index,
            double leftCm, double topCm, double sizeCm, string number) {
            PowerPointAutoShape halo = slide.AddEllipseCm(leftCm - 0.08, topCm - 0.08,
                sizeCm + 0.16, sizeCm + 0.16, "Process Node Halo " + (index + 1));
            halo.FillColor = theme.AccentLightColor;
            halo.FillTransparency = 78;
            halo.OutlineColor = theme.AccentLightColor;
            halo.OutlineWidthPoints = 0.2;

            PowerPointAutoShape node = slide.AddEllipseCm(leftCm, topCm, sizeCm, sizeCm,
                "Process Node " + (index + 1));
            node.FillColor = theme.AccentDarkColor;
            node.FillTransparency = 8;
            node.OutlineColor = theme.AccentLightColor;
            node.OutlineWidthPoints = 1.2;

            PowerPointTextBox numberBox = AddText(slide, number.TrimEnd('.'), leftCm, topCm - 0.01, sizeCm, sizeCm,
                sizeCm < 1 ? 16 : 20, theme.AccentContrastColor, theme.HeadingFontName, bold: true);
            CenterText(numberBox);
        }

        private static PowerPointDirectionMotifStyle ResolveDirectionMotifStyle(
            PowerPointDesignerSlideOptions options) {
            if (options.DirectionMotifStyle != PowerPointDirectionMotifStyle.Auto) {
                return options.DirectionMotifStyle;
            }

            PowerPointDesignIntent intent = options.DesignIntent;
            if (intent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointDirectionMotifStyle.None;
            }
            if (string.IsNullOrWhiteSpace(intent.Seed)) {
                return PowerPointDirectionMotifStyle.Triangles;
            }
            if (intent.Mood == PowerPointDesignMood.Energetic) {
                return PowerPointDirectionMotifStyle.Chevrons;
            }
            if (intent.Mood == PowerPointDesignMood.Editorial) {
                return PowerPointDirectionMotifStyle.Bars;
            }
            if (intent.VisualStyle == PowerPointVisualStyle.Soft) {
                return PowerPointDirectionMotifStyle.Dots;
            }

            return intent.Pick(4, "direction-motif") switch {
                0 => PowerPointDirectionMotifStyle.Triangles,
                1 => PowerPointDirectionMotifStyle.Chevrons,
                2 => PowerPointDirectionMotifStyle.Dots,
                _ => PowerPointDirectionMotifStyle.Bars
            };
        }

        private static void AddDirectionMotif(PowerPointSlide slide, PowerPointDesignerSlideOptions options,
            double leftCm, double topCm, int count, double spacingCm, string color, bool flip = false) {
            PowerPointDirectionMotifStyle style = ResolveDirectionMotifStyle(options);
            if (style == PowerPointDirectionMotifStyle.None) {
                return;
            }

            for (int i = 0; i < count; i++) {
                double left = leftCm + i * spacingCm;
                int transparency = Math.Min(45, i * 3);
                switch (style) {
                    case PowerPointDirectionMotifStyle.Chevrons:
                        AddDirectionChevron(slide, left, topCm, i, color, transparency, flip);
                        break;
                    case PowerPointDirectionMotifStyle.Dots:
                        AddDirectionDot(slide, left, topCm, i, color, transparency);
                        break;
                    case PowerPointDirectionMotifStyle.Bars:
                        AddDirectionBar(slide, left, topCm, i, color, transparency);
                        break;
                    default:
                        AddDirectionTriangle(slide, left, topCm, i, color, transparency, flip);
                        break;
                }
            }
        }

        private static void AddDirectionTriangle(PowerPointSlide slide, double leftCm, double topCm, int index,
            string color, int transparency, bool flip) {
            PowerPointAutoShape arrow = slide.AddShapeCm(A.ShapeTypeValues.Triangle,
                leftCm, topCm, 0.22, 0.24, "Designer Direction " + (index + 1));
            arrow.FillColor = color;
            arrow.FillTransparency = transparency;
            arrow.OutlineColor = color;
            arrow.Rotation = flip ? 270 : 90;
        }

        private static void AddDirectionDot(PowerPointSlide slide, double leftCm, double topCm, int index,
            string color, int transparency) {
            PowerPointAutoShape dot = slide.AddEllipseCm(leftCm, topCm + 0.04, 0.16, 0.16,
                "Designer Direction " + (index + 1));
            dot.FillColor = color;
            dot.FillTransparency = transparency;
            dot.OutlineColor = color;
            dot.OutlineWidthPoints = 0;
        }

        private static void AddDirectionBar(PowerPointSlide slide, double leftCm, double topCm, int index,
            string color, int transparency) {
            PowerPointAutoShape bar = slide.AddRectangleCm(leftCm, topCm + 0.08, 0.24, 0.07,
                "Designer Direction " + (index + 1));
            bar.FillColor = color;
            bar.FillTransparency = transparency;
            bar.OutlineColor = color;
            bar.OutlineWidthPoints = 0;
        }

        private static void AddDirectionChevron(PowerPointSlide slide, double leftCm, double topCm, int index,
            string color, int transparency, bool flip) {
            double tip = flip ? leftCm : leftCm + 0.22;
            double back = flip ? leftCm + 0.22 : leftCm;
            double middleY = topCm + 0.12;

            PowerPointAutoShape upper = slide.AddLineCm(back, topCm + 0.02, tip, middleY,
                "Designer Direction " + (index + 1));
            upper.OutlineColor = color;
            upper.OutlineWidthPoints = Math.Max(0.55, 1.0 - transparency / 100d);

            PowerPointAutoShape lower = slide.AddLineCm(back, topCm + 0.22, tip, middleY,
                "Designer Direction Chevron " + (index + 1) + "B");
            lower.OutlineColor = color;
            lower.OutlineWidthPoints = Math.Max(0.55, 1.0 - transparency / 100d);
        }
    }
}
