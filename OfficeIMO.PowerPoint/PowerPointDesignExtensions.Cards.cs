using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointDesignExtensions {
        internal static void AddCardGrid(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointCardContent> cards, PowerPointCardGridSlideOptions options,
            PowerPointCardGridLayoutVariant variant, double slideWidthCm, double slideHeightCm) {
            double top = options.DesignIntent.Density == PowerPointSlideDensity.Relaxed ? 4.35 : 4.05;
            double height = string.IsNullOrWhiteSpace(options.SupportingText)
                ? slideHeightCm - 6.0
                : slideHeightCm - 8.7;
            PowerPointLayoutBox bounds = PowerPointLayoutBox.FromCentimeters(1.5, top, slideWidthCm - 3.0, height);
            AddCardGrid(slide, theme, cards, options, variant, bounds);

            if (!string.IsNullOrWhiteSpace(options.SupportingText)) {
                PowerPointAutoShape band = slide.AddRectangleCm(1.55, slideHeightCm - 3.25, slideWidthCm - 3.1, 1.8,
                    "Designer Supporting Band");
                band.FillColor = theme.PanelColor;
                band.OutlineColor = theme.PanelBorderColor;
                band.SetShadow("000000", blurPoints: 4, distancePoints: 1, angleDegrees: 90, transparencyPercent: 88);
                AddText(slide, options.SupportingText!, 2.15, slideHeightCm - 2.8, slideWidthCm - 4.3, 0.9, 13,
                    theme.SecondaryTextColor, theme.BodyFontName);
            }
        }

        internal static void AddCardGrid(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointCardContent> cards, PowerPointCardGridSlideOptions options,
            PowerPointCardGridLayoutVariant variant, PowerPointLayoutBox bounds) {
            int maxColumns = Math.Max(1, options.MaxColumns);
            int columns = Math.Min(maxColumns, cards.Count);
            int rows = (int)Math.Ceiling(cards.Count / (double)columns);
            double columnGap = variant == PowerPointCardGridLayoutVariant.SoftTiles ? 0.42 : 0.65;
            double rowGap = variant == PowerPointCardGridLayoutVariant.SoftTiles ? 0.42 : 0.55;
            PowerPointLayoutBox[,] grid = PowerPointLayoutBox
                .FromCentimeters(bounds.LeftCm, bounds.TopCm, bounds.WidthCm, bounds.HeightCm)
                .SplitGridCm(rows, columns, rowGap, columnGap);
            PowerPointCardSurfaceStyle surfaceStyle = ResolveCardSurfaceStyle(options, variant);

            for (int i = 0; i < cards.Count; i++) {
                int row = i / columns;
                int column = i % columns;
                AddDesignerCard(slide, theme, cards[i], grid[row, column], i, variant, surfaceStyle);
            }
        }

        private static void AddDesignerCard(PowerPointSlide slide, PowerPointDesignTheme theme, PowerPointCardContent card,
            PowerPointLayoutBox box, int index, PowerPointCardGridLayoutVariant variant,
            PowerPointCardSurfaceStyle surfaceStyle) {
            string accent = card.AccentColor ?? GetAccent(theme, index);
            PowerPointAutoShape panel = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm, box.HeightCm,
                "Designer Card " + (index + 1));
            ApplyDesignerCardSurface(panel, theme, accent, variant, surfaceStyle);

            if (variant == PowerPointCardGridLayoutVariant.SoftTiles) {
                PowerPointAutoShape accentStrip = slide.AddRectangleCm(box.LeftCm, box.TopCm, 0.13, box.HeightCm,
                    "Designer Card Accent " + (index + 1));
                accentStrip.FillColor = accent;
                accentStrip.OutlineColor = accent;
            } else {
                PowerPointAutoShape accentBar = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm, 0.18,
                    "Designer Card Accent " + (index + 1));
                accentBar.FillColor = accent;
                accentBar.OutlineColor = accent;
            }

            double titleLeft = variant == PowerPointCardGridLayoutVariant.SoftTiles ? box.LeftCm + 0.6 : box.LeftCm + 0.45;
            double titleWidth = Math.Max(0.5, box.WidthCm - 0.9);
            int titleFontSize = ResolveCardTitleFontSize(card.Title, titleWidth);
            double titleHeight = ResolveCardTitleHeight(card.Title, titleWidth, titleFontSize, box.HeightCm);
            AddText(slide, card.Title, titleLeft, box.TopCm + 0.65, titleWidth, titleHeight,
                titleFontSize, theme.PrimaryTextColor, theme.HeadingFontName, bold: true);

            double bodyTop = box.TopCm + 0.65 + titleHeight + 0.28;
            double bodyHeight = Math.Max(0.42, box.HeightCm - (bodyTop - box.TopCm) - 0.35);
            int bodyFontSize = ResolveCardBodyFontSize(card.Items, titleWidth, bodyHeight);
            PowerPointTextBox body = slide.AddTextBox("", PowerPointLayoutBox.FromCentimeters(
                titleLeft + 0.1, bodyTop, box.WidthCm - 1.05, bodyHeight));
            body.SetTextMarginsCm(0, 0, 0, 0);
            body.TextAutoFit = PowerPointTextAutoFit.Normal;

            if (card.Items.Count == 0) {
                body.SetParagraphs(new[] { " " });
                return;
            }

            int bulletSpaceAfter = bodyHeight < 1.15 || card.Items.Count > 3 ? 2 : 4;
            body.SetBullets(card.Items.Select(item => " " + item), configure: paragraph => {
                paragraph.SetFontName(theme.BodyFontName)
                    .SetFontSize(bodyFontSize)
                    .SetColor(theme.SecondaryTextColor)
                    .SetHangingPoints(bodyFontSize <= 8 ? 12 : 16)
                    .SetSpaceAfterPoints(bulletSpaceAfter)
                    .SetBulletSizePercent(70);
            });
        }

        private static void ApplyDesignerCardSurface(PowerPointAutoShape panel, PowerPointDesignTheme theme,
            string accent, PowerPointCardGridLayoutVariant variant, PowerPointCardSurfaceStyle surfaceStyle) {
            panel.FillColor = variant == PowerPointCardGridLayoutVariant.SoftTiles
                ? theme.SurfaceColor
                : theme.PanelColor;
            panel.FillTransparency = 0;
            panel.OutlineColor = theme.PanelBorderColor;
            panel.OutlineWidthPoints = variant == PowerPointCardGridLayoutVariant.SoftTiles ? 0.35 : 0.8;

            switch (surfaceStyle) {
                case PowerPointCardSurfaceStyle.Flat:
                    panel.FillColor = theme.SurfaceColor;
                    panel.OutlineWidthPoints = 0.25;
                    break;
                case PowerPointCardSurfaceStyle.Hairline:
                    panel.FillColor = theme.SurfaceColor;
                    panel.OutlineColor = theme.PanelBorderColor;
                    panel.OutlineWidthPoints = 0.35;
                    break;
                case PowerPointCardSurfaceStyle.AccentWash:
                    panel.FillColor = accent;
                    panel.FillTransparency = 88;
                    panel.OutlineColor = accent;
                    panel.OutlineWidthPoints = 0.25;
                    break;
                default:
                    panel.SetShadow("000000", blurPoints: variant == PowerPointCardGridLayoutVariant.SoftTiles ? 3 : 5,
                        distancePoints: variant == PowerPointCardGridLayoutVariant.SoftTiles ? 0.8 : 1.5,
                        angleDegrees: 90, transparencyPercent: 88);
                    break;
            }
        }

        private static int ResolveCardTitleFontSize(string title, double widthCm) {
            int length = string.IsNullOrWhiteSpace(title) ? 0 : title.Trim().Length;
            if (widthCm < 3.4 || length > 46) {
                return 12;
            }
            if (widthCm < 4.2 || length > 32) {
                return 13;
            }
            return 15;
        }

        private static double ResolveCardTitleHeight(string title, double widthCm, int fontSize, double cardHeightCm) {
            int lines = EstimateWrappedLines(title, widthCm, fontSize);
            double desiredHeight = 0.58 + Math.Max(0, lines - 1) * 0.34;
            double maxHeight = Math.Min(1.35, Math.Max(0.62, cardHeightCm * 0.36));
            return Math.Min(maxHeight, Math.Max(0.62, desiredHeight));
        }

        private static int ResolveCardBodyFontSize(IReadOnlyList<string> items, double widthCm, double bodyHeightCm) {
            if (items.Count == 0) {
                return 10;
            }

            int longest = items.Max(item => string.IsNullOrWhiteSpace(item) ? 0 : item.Trim().Length);
            int estimatedLines = items.Sum(item => EstimateWrappedLines(item, widthCm, 10));
            if (bodyHeightCm < 1.05 || items.Count > 5 || estimatedLines > 7 || longest > 60) {
                return 8;
            }
            if (bodyHeightCm < 1.35 || items.Count > 3 || estimatedLines > 5 || longest > 42) {
                return 9;
            }
            return 10;
        }

        private static int ResolveMetricValueFontSize(string value, double widthCm, int preferredFontSize) {
            int length = string.IsNullOrWhiteSpace(value) ? 0 : value.Trim().Length;
            if (length <= 3) {
                return preferredFontSize;
            }

            int estimate = (int)Math.Floor(widthCm * 7.5 / Math.Max(1, length));
            return Math.Max(12, Math.Min(preferredFontSize, estimate));
        }

        private static int ResolveMetricLabelFontSize(string label, double widthCm, int preferredFontSize) {
            int length = string.IsNullOrWhiteSpace(label) ? 0 : label.Trim().Length;
            if (length <= 12) {
                return preferredFontSize;
            }

            int estimate = (int)Math.Floor(widthCm * 5.5 / Math.Max(1, length));
            return Math.Max(7, Math.Min(preferredFontSize, estimate));
        }

        private static int EstimateWrappedLines(string? text, double widthCm, int fontSize) {
            string textValue = text == null ? string.Empty : text.Trim();
            if (textValue.Length == 0) {
                return 1;
            }

            double charsPerLine = Math.Max(8, widthCm * (fontSize <= 10 ? 5.1 : 4.3));
            return Math.Max(1, (int)Math.Ceiling(textValue.Length / charsPerLine));
        }
    }
}
