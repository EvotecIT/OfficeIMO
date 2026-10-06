namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
        private static List<PdfTextRun> BuildTextRunsFromWrappedLines(
            IReadOnlyList<List<RichSeg>> lines,
            int start,
            int count) {
            var runs = new List<PdfTextRun>();
            for (int lineIndex = 0; lineIndex < count; lineIndex++) {
                IReadOnlyList<RichSeg> line = lines[start + lineIndex];
                for (int segmentIndex = 0; segmentIndex < line.Count; segmentIndex++) {
                    RichSeg segment = line[segmentIndex];
                    if (segment.LeadingIsTab) runs.Add(PdfTextRun.Tab(segment.LeadingTabLeader, segment.LeadingTabAlignment));
                    if (segment.InlineElement != null) {
                        if (segment.LeadingSpace && !segment.LeadingIsTab) {
                            runs.Add(BuildTextRunFromWrappedSegment(" ", segment));
                        }

                        runs.Add(PdfTextRun.Inline(segment.InlineElement));
                        continue;
                    }

                    string text = (segment.LeadingSpace && !segment.LeadingIsTab ? " " : string.Empty) + segment.Text;
                    if (text.Length == 0) {
                        continue;
                    }

                    runs.Add(BuildTextRunFromWrappedSegment(text, segment));
                }

                if (line.Count > 0 && line[line.Count - 1].EndsWithHardBreak) {
                    runs.Add(BuildTextRunFromWrappedSegment("\n", line[line.Count - 1]));
                } else if (lineIndex + 1 < count) {
                    if (line.Count == 0) {
                        runs.Add(PdfTextRun.LineBreak());
                    } else if (line[line.Count - 1].EndsWithTextSeparator) {
                        runs.Add(BuildTextRunFromWrappedSegment(" ", line[line.Count - 1].WithoutLink()));
                    }
                }
            }

            return runs;
        }

        private static PdfTextRun BuildTextRunFromWrappedSegment(string text, RichSeg segment) {
            var run = new PdfTextRun(
                text,
                segment.Bold,
                segment.Underline,
                segment.Color,
                segment.Italic,
                segment.Strike,
                segment.FontSize,
                segment.Font,
                segment.Uri,
                segment.Contents,
                segment.Baseline,
                segment.DestinationName,
                backgroundColor: segment.BackgroundColor,
                fontFamily: segment.NamedFont?.FamilyName,
                underlineStyle: segment.UnderlineStyle,
                strikeStyle: segment.StrikeStyle,
                decorationColor: segment.DecorationColor);
            if (!segment.FeatureSettings.Equals(OfficeIMO.Drawing.OfficeTextFeatureSettings.Default))
                run = run.WithFeatureSettings(segment.FeatureSettings);
            if (segment.TextDirection != OfficeIMO.Drawing.OfficeTextDirection.Auto)
                run = run.WithTextDirection(segment.TextDirection);
            return run;
        }

}
