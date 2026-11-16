using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Keeps a fitted line a hair inside its frame so floating-point measurement never wraps it.
    private const double TableCellFitTextTolerance = 0.01D;

    /// <summary>
    /// Stretches or condenses each source line of a fit-text cell paragraph to the resolved line width.
    /// Long lines are condensed with horizontal glyph scaling; short lines are spread with character spacing.
    /// Font height is unchanged.
    /// </summary>
    private static System.Collections.Generic.IReadOnlyList<PdfTextRun> FitTableCellParagraphText(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, double firstLineWidth, double followingLineWidth,
        PdfStandardFont baseFont, double fontSize, PdfCreateOptions? options) {
        // Tabs and inline objects position themselves, so their paragraphs keep the authored layout.
        if (runs.Count == 0 || runs.Any(run => run.InlineElement != null || (run.Text?.IndexOf('\t') ?? -1) >= 0)) return runs;

        PdfCreateOptions measurementOptions = options ?? new PdfCreateOptions();
        var sourceLines = BuildPageTextLineRuns(NormalizeFallbackRuns(runs, baseFont, measurementOptions));
        var fitted = new System.Collections.Generic.List<PdfTextRun>();
        for (int lineIndex = 0; lineIndex < sourceLines.Count; lineIndex++) {
            if (lineIndex > 0) fitted.Add(PdfTextRun.LineBreak());
            System.Collections.Generic.IReadOnlyList<PdfTextRun> line = sourceLines[lineIndex];
            double target = (lineIndex == 0 ? firstLineWidth : followingLineWidth) - TableCellFitTextTolerance;
            int characters = line.Sum(run => StringInfo.ParseCombiningCharacters(run.Text ?? string.Empty).Length);
            double natural = MeasureTableCellFitTextLine(line, baseFont, fontSize, measurementOptions);
            if (characters == 0 || natural <= 0D || !(target > 0D) || double.IsInfinity(target)) {
                fitted.AddRange(line);
            } else if (natural > target) {
                fitted.AddRange(CondenseTableCellFitTextLine(line, natural, target, baseFont, fontSize, measurementOptions));
            } else if (characters > 1) {
                fitted.AddRange(SpreadTableCellFitTextLine(line, (target - natural) / (characters - 1), characters));
            } else {
                fitted.AddRange(line);
            }
        }

        return fitted.AsReadOnly();
    }

    private static double MeasureTableCellFitTextLine(System.Collections.Generic.IReadOnlyList<PdfTextRun> line,
        PdfStandardFont baseFont, double fontSize, PdfCreateOptions options) =>
        MeasurePreparedTableTextLineWidth(PrepareTableTextLineWidths(line, baseFont, options)[0], fontSize, 1D, 0D);

    private static System.Collections.Generic.IReadOnlyList<PdfTextRun> CondenseTableCellFitTextLine(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> line, double natural, double target,
        PdfStandardFont baseFont, double fontSize, PdfCreateOptions options) {
        // Existing character spacing does not scale with glyph width, so refine the factor against the measured result.
        System.Collections.Generic.IReadOnlyList<PdfTextRun> current = line;
        double width = natural;
        for (int attempt = 0; attempt < 3 && width > target; attempt++) {
            double factor = target / width;
            current = current.Select(run => run.WithHorizontalTextScaling(Math.Max(1D, run.HorizontalTextScaling * factor))).ToArray();
            width = MeasureTableCellFitTextLine(current, baseFont, fontSize, options);
        }

        return current;
    }

    private static System.Collections.Generic.IEnumerable<PdfTextRun> SpreadTableCellFitTextLine(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> line, double extraSpacing, int characters) {
        // Character spacing follows every glyph, including the last one; keep the final glyph unspaced
        // so the line ends at the cell edge.
        int remaining = characters;
        foreach (PdfTextRun run in line) {
            string text = run.Text ?? string.Empty;
            int[] boundaries = StringInfo.ParseCombiningCharacters(text);
            remaining -= boundaries.Length;
            var spaced = OfficeIMO.OfficeLength.Points(run.CharacterSpacing.ToPoints() + extraSpacing);
            if (remaining > 0 || boundaries.Length == 0) {
                yield return run.WithCharacterSpacing(spaced);
                continue;
            }

            int last = boundaries[boundaries.Length - 1];
            if (last > 0) {
                yield return CreateStyledTextRun(text.Substring(0, last), run, run.StandardFont).WithCharacterSpacing(spaced);
                yield return CreateStyledTextRun(text.Substring(last), run, run.StandardFont).WithHorizontalOffset(0D);
            } else {
                yield return run;
            }
        }
    }
}
