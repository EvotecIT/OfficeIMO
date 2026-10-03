using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeMathRenderer {
    private sealed partial class LayoutEngine {
        private const int MaximumAssemblyGlyphs = 1024;
        private const int MaximumMathOutlinePoints = 1_000_000;

        /// <summary>Uses designed outlines without scaling their stroke weight along the stretch axis.</summary>
        private LayoutBox? StretchGlyph(string text, double scale, double target, bool horizontal = false,
            bool useAssembly = true, bool tightInk = false) {
            if (!_options.UseFontMathMetrics || !TryGlyph(text, scale, out var font, out var data, out int glyph)) return null;
            var constructions = horizontal ? data!.Horizontal : data!.Vertical;
            if (!constructions.TryGetValue(glyph, out var construction)) return null;
            double size = FontSize(scale), unit = size / font!.UnitsPerEm;
            var normal = GlyphContours(font, glyph, text, size);
            var bounds = Bounds(normal);
            double natural = horizontal ? GlyphAdvance(font, glyph, text, size) : bounds.Bottom - bounds.Top;
            if (natural >= target) return Text(text, scale, tightInk: tightInk);
            foreach (var variant in construction.Variants) {
                if (variant.Advance * unit < target) continue;
                return OutlineBox(font, data, variant.Glyph, text, size,
                    GlyphContours(font, variant.Glyph, text, size), tightInk: tightInk);
            }
            if (useAssembly && construction.Parts.Length > 0 &&
                TryAssembly(construction, data.MinimumOverlap, target / unit, out var parts, out double overlap)) {
                var contours = new List<List<OfficePoint>>();
                var cache = new Dictionary<int, List<List<OfficePoint>>>();
                double cursor = 0D, width = 0D;
                int points = 0;
                foreach (var part in parts) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    if (!cache.TryGetValue(part.Glyph, out var source)) {
                        source = GlyphContours(font, part.Glyph, text, size);
                        cache.Add(part.Glyph, source);
                    }
                    var ink = Bounds(source);
                    width = Math.Max(width, GlyphAdvance(font, part.Glyph, text, size));
                    foreach (var contour in source) {
                        if (contour.Count > MaximumMathOutlinePoints - points)
                            throw new InvalidOperationException("Math glyph assembly exceeds the outline point limit.");
                        points += contour.Count;
                        var translated = new List<OfficePoint>(contour.Count);
                        foreach (var p in contour) translated.Add(new OfficePoint(
                            p.X + (horizontal ? cursor * unit : 0D),
                            p.Y - (horizontal ? 0D : cursor * unit + ink.Bottom)));
                        contours.Add(translated);
                    }
                    cursor += part.Advance - overlap;
                }
                double extent = (cursor + overlap) * unit;
                var result = OutlineBox(font, data, glyph, text, size, contours,
                    horizontal ? extent : width, tightInk);
                result.ItalicCorrection = construction.ItalicCorrection * unit;
                // Assembly outlines no longer correspond to one glyph's per-corner kern table.
                result.GlyphData = null;
                return result;
            }
            if (construction.Variants.Length == 0) return Text(text, scale, tightInk: tightInk);
            var last = construction.Variants[construction.Variants.Length - 1];
            return OutlineBox(font, data, last.Glyph, text, size, GlyphContours(font, last.Glyph, text, size), tightInk: tightInk);
        }

        private bool TryGlyph(string text, double scale, out IOfficeFontProgram? font,
            out OfficeMathGlyphData? data, out int glyph) {
            font = null; data = null; glyph = 0;
            if (text.Length == 0 || text.Length > 2 || text.Length == 2 && !char.IsSurrogatePair(text, 0)
                || text.Length == 1 && char.IsSurrogate(text[0])) return false;
            if (!_options.Fonts.TryResolveFaceForText(text, _options.Font.FamilyName,
                    _options.Font.Style, FontSize(scale), out var face)) return false;
            font = face!.Program;
            data = (font as IOfficeMathGlyphProgram)?.MathGlyphData;
            return data != null && font.TryGetGlyphMetrics(char.ConvertToUtf32(text, 0), out glyph, out _);
        }

        private List<List<OfficePoint>> GlyphContours(IOfficeFontProgram font, int glyph, string text, double size) {
            _cancellationToken.ThrowIfCancellationRequested();
            var run = new OfficeTextShapingResult(new[] { new OfficeShapedGlyph(glyph, text, 0) });
            double y = -((font as IOfficeFontBaselineMetrics)?.BaselineOffset(size) ?? size);
            if (font is IOfficeBoundedFontProgram bounded)
                return bounded.GetShapedTextContoursBounded(text, run, 0D, y, size, MaximumMathOutlinePoints, _cancellationToken);
            // The glyph-data seam is internal and implemented only by the bounded first-party readers.
            throw new InvalidOperationException("Math glyph construction requires bounded outlines.");
        }

        private static double GlyphAdvance(IOfficeFontProgram font, int glyph, string text, double size) =>
            font.MeasureShapedText(text, new OfficeTextShapingResult(new[] { new OfficeShapedGlyph(glyph, text, 0) }), size);

        private LayoutBox OutlineBox(IOfficeFontProgram font, OfficeMathGlyphData data, int glyph,
            string text, double size, List<List<OfficePoint>> contours, double? advance = null, bool tightInk = false) {
            var bounds = Bounds(contours);
            double left = Math.Min(0D, bounds.Left), top = Math.Min(0D, bounds.Top);
            double right = Math.Max(advance ?? GlyphAdvance(font, glyph, text, size), bounds.Right);
            double bottom = tightInk ? bounds.Bottom : Math.Max(0D, bounds.Bottom);
            var box = new LayoutBox(Math.Max(.01D, right - left), Math.Max(.01D, bottom - top), -top);
            var commands = new List<OfficePathCommand>();
            foreach (var contour in contours) {
                if (contour.Count < 3) continue;
                commands.Add(OfficePathCommand.MoveTo(contour[0].X - left, contour[0].Y - top));
                for (int i = 1; i < contour.Count; i++) commands.Add(OfficePathCommand.LineTo(contour[i].X - left, contour[i].Y - top));
                commands.Add(OfficePathCommand.Close());
            }
            if (commands.Count == 0) return Text(text, size / FontSize(1D));
            box.Commands.Add(LayoutCommand.Outline(text, box.Width, box.Height, size, box.Baseline, commands));
            SetGlyphInfo(box, font, data, glyph, size, -left, advance ?? GlyphAdvance(font, glyph, text, size));
            return box;
        }

        private static (double Left, double Top, double Right, double Bottom) Bounds(List<List<OfficePoint>> contours) {
            double left = double.PositiveInfinity, top = double.PositiveInfinity,
                right = double.NegativeInfinity, bottom = double.NegativeInfinity;
            foreach (var contour in contours) foreach (var p in contour) {
                left = Math.Min(left, p.X); right = Math.Max(right, p.X);
                top = Math.Min(top, p.Y); bottom = Math.Max(bottom, p.Y);
            }
            return double.IsInfinity(left) ? (0D, 0D, 0D, 0D) : (left, top, right, bottom);
        }

        /// <summary>Repeats all extenders equally and chooses one legal connector overlap.</summary>
        private bool TryAssembly(OfficeMathGlyphConstruction construction, int minimum, double target,
            out List<OfficeMathGlyphPart> result, out double overlap) {
            result = new List<OfficeMathGlyphPart>(); overlap = 0D;
            if (double.IsNaN(target) || double.IsInfinity(target) || target < 0D) return false;
            int maximumRepetitions = construction.Parts.Any(p => p.Extender) ? MaximumAssemblyGlyphs : 0;
            for (int repetitions = 0; repetitions <= maximumRepetitions; repetitions++) {
                _cancellationToken.ThrowIfCancellationRequested(); result.Clear();
                foreach (var part in construction.Parts) {
                    int count = part.Extender ? repetitions : 1;
                    if (count > MaximumAssemblyGlyphs - result.Count) return false;
                    for (int i = 0; i < count; i++) result.Add(part);
                }
                if (result.Count == 0) continue;
                double sum = result.Sum(p => (double)p.Advance), maximum = double.PositiveInfinity;
                for (int i = 1; i < result.Count; i++)
                    maximum = Math.Min(maximum, Math.Min(result[i - 1].End, result[i].Start));
                if (result.Count == 1) { if (sum >= target) return true; continue; }
                if (maximum < minimum) continue;
                double largest = sum - minimum * (result.Count - 1);
                if (largest < target) continue;
                overlap = Math.Min(maximum, Math.Max(minimum, (sum - target) / (result.Count - 1)));
                return true;
            }
            return false;
        }

        private static void SetGlyphInfo(LayoutBox box, IOfficeFontProgram font, OfficeMathGlyphData data,
            int glyph, double size, double origin, double advance) {
            box.GlyphData = data; box.GlyphId = glyph; box.GlyphUnit = size / font.UnitsPerEm;
            box.GlyphOrigin = origin;
            box.GlyphAdvance = advance;
            box.ItalicCorrection = data.Italics.TryGetValue(glyph, out int italic) ? italic * box.GlyphUnit : 0D;
            box.AccentAttachment = data.Accents.TryGetValue(glyph, out int accent) ? origin + accent * box.GlyphUnit : (double?)null;
        }
    }
}
