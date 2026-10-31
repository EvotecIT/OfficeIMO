using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Text measurement and line breaking for layout decisions. Widths use the deterministic Office measurer
    ///     scaled by calibrated family advances, with the shared drawing renderer as a lower bound
    ///     when an installed or fallback face is wider.
    /// </summary>
    internal static class PowerPointTextMeasure {
        private static readonly OfficeTextMeasurer Measurer = OfficeTextMeasurer.Create(new OfficeFontInfo("Aptos", 10D));

        private static readonly Dictionary<string, double> FamilyFactors = new(StringComparer.OrdinalIgnoreCase) {
            ["Aptos"] = 0.94D,
            // Aptos Display is downloaded on demand; where it is missing, the fallback runs wider than the face.
            ["Aptos Display"] = 0.96D,
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
            // A missing theme face can resolve to wider glyphs on another host. Keep the
            // native Office estimate, but never reserve less than the actual export face.
            var metrics = OfficeDrawingTextLayout.CreateMetrics((OfficeFontFaceCollection?)null);
            double exportWidth = metrics.MeasureText(text, size, family,
                bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular);
            return Math.Max(width + tracking * size * text.Length, exportWidth);
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
        ///     The frame width that balances a wrapped title or subtitle (no single-word last line) while keeping the
        ///     same line count. The text itself keeps soft wrapping: the frame narrows instead of the text gaining hard
        ///     line breaks, so the renderer still wraps with its own metrics and an edited title reflows. A small
        ///     allowance above the narrowest measure keeps a slightly wider renderer from adding a line.
        /// </summary>
        internal static double BalancedWidth(string text, double size, double width, string family, bool bold = false,
            double tracking = 0D) {
            List<string> lines = Wrap(text, size, width, family, bold, tracking);
            if (lines.Count < 2 || text.IndexOf('\n') >= 0) return width;
            string[] words = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length < 3) return width;
            // Shrink the measure until the line count would grow; the narrowest measure that keeps the count is balanced.
            double low = width * 0.5D, high = width;
            for (int i = 0; i < 12; i++) {
                double mid = (low + high) / 2D;
                if (Wrap(text, size, mid, family, bold, tracking).Count > lines.Count) low = mid;
                else high = mid;
            }

            if (Wrap(text, size, high, family, bold, tracking).Count != lines.Count) return width;
            // The longest word keeps a wider allowance than the line: a renderer that cannot fit a word breaks it.
            return Math.Min(width, Math.Max(high * 1.06D, LongestWord(text, size, family, bold) * 1.15D));
        }
    }
}
