using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Text measurement and line breaking for layout decisions. Widths use the deterministic Office measurer
    ///     scaled by a per-family advance factor (calibrated against the Aptos font files: the Office measurer reads about 6 % wide), so a layout sized here never overflows when
    ///     the font falls back to a narrower metric-compatible face.
    /// </summary>
    internal static class PowerPointTextMeasure {
        private static readonly OfficeTextMeasurer Measurer = OfficeTextMeasurer.Create(new OfficeFontInfo("Aptos", 10D));

        private static readonly Dictionary<string, double> FamilyFactors = new(StringComparer.OrdinalIgnoreCase) {
            ["Aptos"] = 0.94D,
            ["Aptos Display"] = 0.91D,
            ["Calibri"] = 0.87D,
            ["Carlito"] = 0.87D,
            ["Arial"] = 0.99D,
            ["Liberation Sans"] = 0.99D,
            ["Helvetica"] = 0.99D,
            ["Segoe UI"] = 0.98D,
            ["Lato"] = 0.94D,
            ["Poppins"] = 1.09D,
            ["Inter"] = 1.02D,
            ["Verdana"] = 1.13D,
            ["Georgia"] = 1.13D,
            ["Cambria"] = 0.99D,
            ["Times New Roman"] = 0.89D,
        };

        internal static double FamilyFactor(string family) =>
            FamilyFactors.TryGetValue(family ?? string.Empty, out double factor) ? factor : 1.0D;

        /// <summary>Advance width in points of <paramref name="text" /> at <paramref name="size" /> points.</summary>
        internal static double Width(string text, double size, string family, bool bold = false, double tracking = 0D) {
            if (string.IsNullOrEmpty(text)) return 0D;
            OfficeTextMeasurementStyle style = Measurer.CreateStyle(
                new OfficeFontInfo("Aptos", size, bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular), 72D);
            double width = Measurer.MeasureWidth(text, style) * FamilyFactor(family);
            return width + tracking * size * text.Length;
        }

        /// <summary>Greedy word wrap. Words are never broken; a word longer than the line takes its own line.</summary>
        internal static List<string> Wrap(string text, double size, double width, string family, bool bold = false,
            double tracking = 0D) {
            var lines = new List<string>();
            foreach (string paragraph in (text ?? string.Empty).Replace("\r\n", "\n").Split('\n')) {
                string current = string.Empty;
                foreach (string word in paragraph.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                    string candidate = current.Length == 0 ? word : current + " " + word;
                    if (current.Length == 0 || Width(candidate, size, family, bold, tracking) <= width) {
                        current = candidate;
                    } else {
                        lines.Add(current);
                        current = word;
                    }
                }

                lines.Add(current);
            }

            return lines;
        }

        internal static int LineCount(string text, double size, double width, string family, bool bold = false,
            double tracking = 0D) => Wrap(text, size, width, family, bold, tracking).Count;

        /// <summary>The widest single word, so a caller can keep words whole (logo names, labels).</summary>
        internal static double LongestWord(string text, double size, string family, bool bold = false) =>
            (text ?? string.Empty).Split(new[] { ' ', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(word => Width(word, size, family, bold)).DefaultIfEmpty(0D).Max();

        /// <summary>
        ///     Re-breaks a title or subtitle so the lines are balanced (no single-word last line) while keeping the
        ///     same line count. Returns the text with explicit line breaks, or the input when balancing changes nothing.
        /// </summary>
        internal static string Balance(string text, double size, double width, string family, bool bold = false,
            double tracking = 0D) {
            List<string> lines = Wrap(text, size, width, family, bold, tracking);
            if (lines.Count < 2 || text.IndexOf('\n') >= 0) return text;
            string[] words = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length < 3) return text;
            // Shrink the measure until the line count would grow; the narrowest measure that keeps the count is balanced.
            double low = width * 0.5D, high = width;
            for (int i = 0; i < 12; i++) {
                double mid = (low + high) / 2D;
                if (Wrap(text, size, mid, family, bold, tracking).Count > lines.Count) low = mid;
                else high = mid;
            }

            List<string> balanced = Wrap(text, size, high, family, bold, tracking);
            if (balanced.Count != lines.Count) return text;
            if (balanced.SequenceEqual(lines)) return text;
            return string.Join("\n", balanced);
        }
    }
}
