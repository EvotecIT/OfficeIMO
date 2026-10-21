namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static RichSeg[] BuildPageTextSegments(IReadOnlyList<PdfTextRun> runs, PdfStandardFont baseFont, double fontSize, PdfOptions options) {
        var segments = new RichSeg[runs.Count];
        for (int index = 0; index < runs.Count; index++) {
            PdfTextRun run = runs[index];
            PdfStandardFont font = ResolvePageTextRunFont(run, baseFont);
            PdfNamedFontFace? named = options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace face) ? face : null;
            double size = run.FontSize ?? fontSize;
            double width = run.InlineElement?.Width ?? MeasureRichText(run.Text, font, named, size, run.Baseline, options,
                run.FeatureSettings, run.HorizontalTextScaling, run.CharacterSpacing, run.TextDirection,
                preservesPaintedGlyphs: run.PreservesPaintedGlyphs);
            segments[index] = new RichSeg(run.Text, run.Bold, run.Italic, run.Underline, run.Strike,
                run.Color, run.BackgroundColor, run.LinkUri, run.LinkDestinationName, run.LinkContents,
                font, size, run.Baseline, width, leadingSpace: run.HorizontalOffset != 0D,
                leadingAdvance: run.HorizontalOffset, inlineElement: run.InlineElement, namedFont: named,
                featureSettings: run.FeatureSettings, textDirection: run.TextDirection,
                horizontalTextScaling: run.HorizontalTextScaling, characterSpacing: run.CharacterSpacing,
                contributesLineBoxDescent: run.ContributesLineBoxDescent,
                preservesPaintedGlyphs: run.PreservesPaintedGlyphs, formField: run.FormField);
        }
        new RichRunKerning(options).Apply(segments);
        return segments;
    }
}
