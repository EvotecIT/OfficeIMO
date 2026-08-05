using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static readonly Dictionary<HighlightColorValues, Color> _highlightColors = new() {
            { HighlightColorValues.Yellow, Color.Yellow },
            { HighlightColorValues.Green, Color.Lime },
            { HighlightColorValues.Cyan, Color.Cyan },
            { HighlightColorValues.Magenta, Color.Magenta },
            { HighlightColorValues.Blue, Color.Blue },
            { HighlightColorValues.Red, Color.Red },
            { HighlightColorValues.DarkBlue, Color.DarkBlue },
            { HighlightColorValues.DarkCyan, Color.DarkCyan },
            { HighlightColorValues.DarkGreen, Color.DarkGreen },
            { HighlightColorValues.DarkMagenta, Color.DarkMagenta },
            { HighlightColorValues.DarkRed, Color.DarkRed },
            { HighlightColorValues.DarkYellow, Color.Parse("#808000") },
            { HighlightColorValues.DarkGray, Color.DarkGray },
            { HighlightColorValues.LightGray, Color.LightGray },
            { HighlightColorValues.Black, Color.Black },
            { HighlightColorValues.White, Color.White }
        };

        private void ApplyTextBackground(
            WordParagraph run,
            TextFormatting formatting,
            HtmlToWordOptions options) {
            if (string.IsNullOrEmpty(formatting.BackgroundColorHex)) {
                return;
            }

            if (formatting.PreserveHighlightOverBackground && formatting.Highlight.HasValue) {
                return;
            }

            if (options.TextBackgroundMode == HtmlTextBackgroundMode.ExactShading) {
                run.Highlight = null;
                run.SetRunShadingFillColorHex(formatting.BackgroundColorHex!);
                return;
            }

            HighlightColorValues? highlight = MapColorToHighlight(formatting.BackgroundColorHex, out bool exact);
            if (!highlight.HasValue) {
                return;
            }

            run.SetHighlight(highlight.Value.ToOfficeEnum());
            if (!exact) {
                AddDiagnostic(
                    options,
                    "TextBackgroundColorApproximated",
                    "CSS text background color was approximated to the nearest Word highlight color.",
                    "background-color",
                    lossKind: HtmlConversionLossKind.Approximation);
            }
        }

        private static string ResolveOpaqueTextBackground(
            string foregroundHex,
            double alpha,
            string? inheritedBackgroundHex) {
            if (alpha >= 1d) {
                return foregroundHex;
            }

            Color foreground = Color.Parse("#" + foregroundHex);
            Color background;
            try {
                background = string.IsNullOrEmpty(inheritedBackgroundHex)
                    ? Color.White
                    : Color.Parse("#" + inheritedBackgroundHex);
            } catch {
                background = Color.White;
            }

            double foregroundRatio = Math.Max(0d, Math.Min(1d, alpha));
            double backgroundRatio = 1d - foregroundRatio;
            return Color.FromRgb(
                (byte)Math.Round(foreground.R * foregroundRatio + background.R * backgroundRatio),
                (byte)Math.Round(foreground.G * foregroundRatio + background.G * backgroundRatio),
                (byte)Math.Round(foreground.B * foregroundRatio + background.B * backgroundRatio))
                .ToRgbHex();
        }

        private static HighlightColorValues? MapColorToHighlight(string? hex, out bool exact) {
            exact = false;
            if (string.IsNullOrEmpty(hex)) {
                return null;
            }

            try {
                Color target = Color.Parse("#" + hex);
                HighlightColorValues? best = null;
                int bestDistance = int.MaxValue;
                foreach (KeyValuePair<HighlightColorValues, Color> pair in _highlightColors) {
                    Color candidate = pair.Value;
                    int distance = (candidate.R - target.R) * (candidate.R - target.R) +
                                   (candidate.G - target.G) * (candidate.G - target.G) +
                                   (candidate.B - target.B) * (candidate.B - target.B);
                    if (distance < bestDistance) {
                        bestDistance = distance;
                        best = pair.Key;
                    }
                }

                exact = bestDistance == 0;
                return best;
            } catch {
                return null;
            }
        }
    }
}
