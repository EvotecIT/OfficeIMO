using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Keeps a fitted line a hair inside its frame so floating-point round-off never wraps it.
    private const double TableCellFitTextTolerance = 0.05D;
    private const double TableCellFitTextMinimumScaling = 1D;

    /// <summary>
    /// Stretches or condenses each source line of a fit-text cell paragraph to the resolved line width.
    /// Long lines are condensed with horizontal glyph scaling; short lines are spread with character spacing,
    /// or widened with glyph scaling when they end in a connected script.
    /// Font height is unchanged.
    /// </summary>
    private static System.Collections.Generic.IReadOnlyList<PdfTextRun> FitTableCellParagraphText(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, double firstLineWidth, double followingLineWidth,
        PdfStandardFont baseFont, double fontSize, PdfCreateOptions? options) {
        if (runs.Count == 0) return runs;

        PdfCreateOptions measurementOptions = options ?? new PdfCreateOptions();
        var sourceLines = SplitTableCellFitTextLines(runs, out var separators);
        var fitted = new System.Collections.Generic.List<PdfTextRun>();
        for (int lineIndex = 0; lineIndex < sourceLines.Count; lineIndex++) {
            if (lineIndex > 0) fitted.Add(separators[lineIndex - 1]);
            // Tabs and inline objects position themselves, so their own lines keep the authored layout.
            if (sourceLines[lineIndex].Any(run => run.InlineElement != null || (run.Text?.IndexOf('\t') ?? -1) >= 0)) {
                fitted.AddRange(sourceLines[lineIndex]);
                continue;
            }

            // Fit Text replaces the authored character width and spacing on the lines it fits.
            System.Collections.Generic.IReadOnlyList<PdfTextRun> line = sourceLines[lineIndex]
                .Select(run => string.IsNullOrEmpty(run.Text) ? run : run.WithHorizontalTextScaling(100D).WithCharacterSpacing(OfficeIMO.OfficeLength.Zero)).ToArray();
            double target = (lineIndex == 0 ? firstLineWidth : followingLineWidth) - TableCellFitTextTolerance;
            double natural = MeasureTableCellFitTextLine(line, baseFont, fontSize, measurementOptions);
            if (natural <= 0D || !(target > 0D) || double.IsInfinity(target)) {
                fitted.AddRange(line);
            } else if (natural > target) {
                fitted.AddRange(CondenseTableCellFitTextLine(line, natural, target, baseFont, fontSize, measurementOptions));
            } else {
                // Spacing would leave a blank advance after a final joined glyph, and cannot widen a line with no
                // visible glyph before its last one; widen the glyphs instead.
                fitted.AddRange((EndsWithTableCellFitTextShapingContext(line)
                    ? null
                    : SpreadTableCellFitTextLine(line, natural, target, baseFont, fontSize, measurementOptions))
                    ?? WidenTableCellFitTextLine(line, natural, target, baseFont, fontSize, measurementOptions));
            }
        }

        return fitted.AsReadOnly();
    }

    /// <summary>Splits runs at authored line breaks while keeping empty metric-bearing runs such as paragraph-mark struts.</summary>
    private static System.Collections.Generic.List<System.Collections.Generic.List<PdfTextRun>> SplitTableCellFitTextLines(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, out System.Collections.Generic.List<PdfTextRun> separators) {
        // Authored breaks keep their own run formatting, which supplies the metrics of the line they end.
        separators = new System.Collections.Generic.List<PdfTextRun>();
        var lines = new System.Collections.Generic.List<System.Collections.Generic.List<PdfTextRun>>();
        var current = new System.Collections.Generic.List<PdfTextRun>();
        lines.Add(current);
        foreach (PdfTextRun run in runs) {
            string text = run.Text ?? string.Empty;
            if (text.IndexOf('\r') < 0 && text.IndexOf('\n') < 0) {
                current.Add(run);
                continue;
            }

            int segmentStart = 0;
            for (int index = 0; index < text.Length; index++) {
                char ch = text[index];
                if (ch != '\r' && ch != '\n') continue;
                if (index > segmentStart) current.Add(CreateStyledTextRun(text.Substring(segmentStart, index - segmentStart), run, run.StandardFont));
                int breakLength = ch == '\r' && index + 1 < text.Length && text[index + 1] == '\n' ? 2 : 1;
                separators.Add(text.Length == breakLength ? run : CreateStyledTextRun(text.Substring(index, breakLength), run, run.StandardFont));
                current = new System.Collections.Generic.List<PdfTextRun>();
                lines.Add(current);
                index += breakLength - 1;
                segmentStart = index + 1;
            }

            if (segmentStart < text.Length) current.Add(CreateStyledTextRun(text.Substring(segmentStart), run, run.StandardFont));
        }

        return lines;
    }

    private static double MeasureTableCellFitTextLine(System.Collections.Generic.IReadOnlyList<PdfTextRun> line,
        PdfStandardFont baseFont, double fontSize, PdfCreateOptions options) {
        if (!line.Any(run => !string.IsNullOrEmpty(run.Text))) return 0D;
        // Measure through the wrapper itself so font-unit advances, kerning and spacing match what wrapping will see.
        var layout = WrapRichRunsIntrinsic(line, double.MaxValue, fontSize, baseFont, fontSize * 1.2D, null,
            DefaultParagraphTabStopWidth, options);
        return layout.Lines.Count == 0 ? 0D : layout.Lines.Max(wrapped => MeasureRichLineWidth(wrapped, options));
    }

    private static System.Collections.Generic.IReadOnlyList<PdfTextRun> CondenseTableCellFitTextLine(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> line, double natural, double target,
        PdfStandardFont baseFont, double fontSize, PdfCreateOptions options) {
        // Fitted lines carry no character spacing, so the width scales with the glyphs. Kerning and
        // rounding can leave a small residue; refine against the measured result.
        System.Collections.Generic.IReadOnlyList<PdfTextRun> current = line;
        double width = natural;
        for (int attempt = 0; attempt < 3 && width > target; attempt++) {
            current = ScaleTableCellFitTextLine(current, target / width);
            width = MeasureTableCellFitTextLine(current, baseFont, fontSize, options);
        }

        return current;
    }
    private static PdfTextRun[] ScaleTableCellFitTextLine(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> line, double factor) =>
        line.Select(run => string.IsNullOrEmpty(run.Text)
            ? run
            : run.WithHorizontalTextScaling(Math.Max(TableCellFitTextMinimumScaling, run.HorizontalTextScaling * factor))).ToArray();

    private static System.Collections.Generic.IReadOnlyList<PdfTextRun> WidenTableCellFitTextLine(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> line, double natural, double target,
        PdfStandardFont baseFont, double fontSize, PdfCreateOptions options) {
        PdfTextRun[] widened = ScaleTableCellFitTextLine(line, target / natural);
        return CondenseTableCellFitTextLine(widened, MeasureTableCellFitTextLine(widened, baseFont, fontSize, options),
            target, baseFont, fontSize, options);
    }

    /// <summary>Returns <see langword="null"/> when no rendered glyph before the last visible one can take spacing.</summary>
    private static System.Collections.Generic.IReadOnlyList<PdfTextRun>? SpreadTableCellFitTextLine(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> line, double natural, double target,
        PdfStandardFont baseFont, double fontSize, PdfCreateOptions options) {
        // Character spacing follows every rendered glyph, so distribute the extra width over shaped glyphs rather
        // than source characters. The last visible grapheme is split off unspaced so the line ends at the cell
        // edge; trailing spaces after it are dropped by wrapping and stay unspaced too.
        int lastTextRun = -1;
        for (int index = line.Count - 1; index >= 0 && lastTextRun < 0; index--) {
            if (TrimTableCellFitTextTrailingSpaces(line[index].Text).Length > 0) lastTextRun = index;
        }
        if (lastTextRun < 0) return null;

        var pieces = new System.Collections.Generic.List<(PdfTextRun Run, bool Spaced)>();
        for (int index = 0; index < line.Count; index++) {
            PdfTextRun run = line[index];
            string text = run.Text ?? string.Empty;
            if (index != lastTextRun) {
                pieces.Add((run, index < lastTextRun && text.Length > 0));
                continue;
            }

            string visible = TrimTableCellFitTextTrailingSpaces(text);
            int[] boundaries = StringInfo.ParseCombiningCharacters(visible);
            int last = boundaries[boundaries.Length - 1];
            if (last > 0) pieces.Add((CreateStyledTextRun(visible.Substring(0, last), run, run.StandardFont), true));
            pieces.Add((CreateStyledTextRun(visible.Substring(last), run, run.StandardFont).WithHorizontalOffset(last == 0 ? run.HorizontalOffset : 0D), false));
            if (visible.Length < text.Length) pieces.Add((CreateStyledTextRun(text.Substring(visible.Length), run, run.StandardFont), false));
        }

        int spacedGlyphs = pieces.Where(piece => piece.Spaced).Sum(piece => CountTableCellFitTextGlyphs(piece.Run, baseFont, options));
        if (spacedGlyphs == 0) return null;
        // Wrapping can drop spaced glyphs (collapsed whitespace) and splitting off the final grapheme can end a
        // kerned pair, so calibrate against the wrapped measurement: the width grows linearly with the spacing.
        double trial = (target - natural) / spacedGlyphs;
        double unspaced = MeasureTableCellFitTextLine(ApplyTableCellFitTextSpacing(pieces, 0D), baseFont, fontSize, options);
        double perPoint = (MeasureTableCellFitTextLine(ApplyTableCellFitTextSpacing(pieces, trial), baseFont, fontSize, options) - unspaced) / trial;
        if (perPoint < 0.5D) return null;
        double extraSpacing = (target - unspaced) / perPoint;
        if (!(extraSpacing > 0D)) return line;
        PdfTextRun[] spread = ApplyTableCellFitTextSpacing(pieces, extraSpacing);
        double overshoot = MeasureTableCellFitTextLine(spread, baseFont, fontSize, options) - target;
        if (overshoot > 0D) {
            extraSpacing -= overshoot / perPoint;
            spread = extraSpacing > 0D ? ApplyTableCellFitTextSpacing(pieces, extraSpacing) : line.ToArray();
        }

        return spread;
    }

    private static PdfTextRun[] ApplyTableCellFitTextSpacing(
        System.Collections.Generic.IReadOnlyList<(PdfTextRun Run, bool Spaced)> pieces, double extraSpacing) =>
        pieces.Select(piece => piece.Spaced
            ? piece.Run.WithCharacterSpacing(OfficeIMO.OfficeLength.Points(extraSpacing))
            : piece.Run).ToArray();

    private static bool EndsWithTableCellFitTextShapingContext(System.Collections.Generic.IReadOnlyList<PdfTextRun> line) {
        PdfTextRun? last = line.LastOrDefault(run => TrimTableCellFitTextTrailingSpaces(run.Text).Length > 0);
        return last != null && UsesTableCellFitTextShapingContext(TrimTableCellFitTextTrailingSpaces(last.Text));
    }

    // Ordinary spaces at the end of a line are dropped by wrapping and take no visible width.
    private static string TrimTableCellFitTextTrailingSpaces(string? text) => (text ?? string.Empty).TrimEnd(' ');

    private static bool UsesTableCellFitTextShapingContext(string text) {
        foreach (char ch in text) {
            // Hebrew, Arabic, Syriac, Thaana, N'Ko, Indic, Southeast Asian scripts and Hebrew/Arabic presentation forms.
            int code = ch;
            if ((code >= 0x0590 && code <= 0x109F) || (code >= 0xFB1D && code <= 0xFEFC)) return true;
        }
        return false;
    }

    private static int CountTableCellFitTextGlyphs(PdfTextRun run, PdfStandardFont baseFont, PdfCreateOptions options) =>
        PrepareTableTextLineWidths(new[] { run }, baseFont, options)
            .SelectMany(line => line.Runs)
            .Sum(measurement => measurement.Command?.GlyphCount ?? 0);
}
