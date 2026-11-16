using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Keeps a fitted line just inside its frame so measurement and wrapping round-off never wrap it.
    private const double TableCellFitTextTolerance = 0.5D;
    private const double TableCellFitTextMinimumScaling = 1D;

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
        var sourceLines = SplitTableCellFitTextLines(runs, out var separators);
        var fitted = new System.Collections.Generic.List<PdfTextRun>();
        for (int lineIndex = 0; lineIndex < sourceLines.Count; lineIndex++) {
            if (lineIndex > 0) fitted.Add(separators[lineIndex - 1]);
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
                fitted.AddRange(SpreadTableCellFitTextLine(line, target - natural, target, baseFont, fontSize, measurementOptions));
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
        return MeasurePreparedTableTextLineWidth(PrepareTableTextLineWidths(line, baseFont, options)[0], fontSize, 1D, 0D);
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

    private static System.Collections.Generic.IReadOnlyList<PdfTextRun> SpreadTableCellFitTextLine(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> line, double extraWidth, double target,
        PdfStandardFont baseFont, double fontSize, PdfCreateOptions options) {
        // Character spacing follows every rendered glyph, so distribute the extra width over shaped glyphs rather
        // than source characters. The final grapheme is split off unspaced so the line ends at the cell edge.
        int lastTextRun = -1;
        for (int index = line.Count - 1; index >= 0 && lastTextRun < 0; index--) {
            if (!string.IsNullOrEmpty(line[index].Text)) lastTextRun = index;
        }
        if (lastTextRun < 0) return line;

        var pieces = new System.Collections.Generic.List<(PdfTextRun Run, bool Spaced)>();
        for (int index = 0; index < line.Count; index++) {
            PdfTextRun run = line[index];
            string text = run.Text ?? string.Empty;
            if (index != lastTextRun) {
                pieces.Add((run, text.Length > 0));
                continue;
            }

            int[] boundaries = StringInfo.ParseCombiningCharacters(text);
            int last = boundaries[boundaries.Length - 1];
            if (last > 0) pieces.Add((CreateStyledTextRun(text.Substring(0, last), run, run.StandardFont), true));
            pieces.Add((CreateStyledTextRun(text.Substring(last), run, run.StandardFont).WithHorizontalOffset(last == 0 ? run.HorizontalOffset : 0D), false));
        }

        int spacedGlyphs = pieces.Where(piece => piece.Spaced).Sum(piece => CountTableCellFitTextGlyphs(piece.Run, baseFont, options));
        if (spacedGlyphs == 0) return line;
        double extraSpacing = extraWidth / spacedGlyphs;
        PdfTextRun[] spread = ApplyTableCellFitTextSpacing(pieces, extraSpacing);
        // Splitting off the final grapheme can end a kerned pair, so check the spread line and take any
        // overshoot back out of the spacing.
        double overshoot = MeasureTableCellFitTextLine(spread, baseFont, fontSize, options) - target;
        if (overshoot > 0D) {
            extraSpacing -= overshoot / spacedGlyphs;
            spread = extraSpacing > 0D ? ApplyTableCellFitTextSpacing(pieces, extraSpacing) : line.ToArray();
        }

        return spread;
    }

    private static PdfTextRun[] ApplyTableCellFitTextSpacing(
        System.Collections.Generic.IReadOnlyList<(PdfTextRun Run, bool Spaced)> pieces, double extraSpacing) =>
        pieces.Select(piece => piece.Spaced
            ? piece.Run.WithCharacterSpacing(OfficeIMO.OfficeLength.Points(extraSpacing))
            : piece.Run).ToArray();

    private static int CountTableCellFitTextGlyphs(PdfTextRun run, PdfStandardFont baseFont, PdfCreateOptions options) =>
        PrepareTableTextLineWidths(new[] { run }, baseFont, options)
            .SelectMany(line => line.Runs)
            .Sum(measurement => measurement.Command?.GlyphCount ?? 0);
}
