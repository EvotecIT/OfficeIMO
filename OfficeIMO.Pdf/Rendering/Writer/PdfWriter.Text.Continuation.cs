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
                    // A soft interval boundary retains its source separator even though the
                    // next interval starts without a painted leading space.
                    if (segmentIndex > 0 && !segment.LeadingSpace &&
                        line[segmentIndex - 1].EndsWithTextSeparator &&
                        line is RichLine { Fragments: { } fragments } &&
                        fragments.Any(fragment => fragment.Start == segmentIndex))
                        runs.Add(BuildTextRunFromWrappedSegment(" ", line[segmentIndex - 1].WithoutLink(), preserveFootnote: false));
                    if (segment.LeadingIsTab) runs.Add(BuildTextRunFromWrappedSegment("\t", segment.WithoutLink(),
                        preserveFootnote: false, preserveTab: true));
                    if (segment.LeadingSpace && !segment.LeadingIsTab) {
                        AppendPreservedSpaceRuns(runs, segment);
                    }
                    if (segment.InlineElement != null) {
                        PdfTextRun inline = segment.InlineElement is PdfInlineBox { SourceRule: { } sourceRule }
                            ? PdfTextRun.Inline(sourceRule, BuildTextRunFromWrappedSegment(string.Empty, segment))
                            : PdfTextRun.Inline(segment.InlineElement);
                        runs.Add(segment.FormField == null ? inline : inline.WithFormField(segment.FormField));
                        continue;
                    }

                    string text = segment.Text;
                    if (text.Length == 0 && !segment.IsLineBoxStrut && segment.FootnoteReference == null) {
                        continue;
                    }

                    runs.Add(BuildTextRunFromWrappedSegment(text, segment));
                }

                if (line.Count > 0 && line[line.Count - 1].EndsWithHardBreak) {
                    runs.Add(BuildTextRunFromWrappedSegment("\n", line[line.Count - 1], preserveFootnote: false));
                } else if (lineIndex + 1 < count) {
                    if (line.Count == 0) {
                        runs.Add(PdfTextRun.LineBreak());
                    } else if (line[line.Count - 1].EndsWithTextSeparator) {
                        runs.Add(BuildTextRunFromWrappedSegment(" ", line[line.Count - 1].WithoutLink(), preserveFootnote: false));
                    }
                }
            }

            return runs;
        }

        private static PdfTextRun BuildTextRunFromWrappedSegment(string text, RichSeg segment, bool preserveFootnote = true, bool preserveTab = false) {
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
                tabLeader: preserveTab ? segment.SourceTabRun?.TabLeader ?? segment.LeadingTabLeader : PdfTabLeaderStyle.None,
                tabAlignment: preserveTab ? segment.SourceTabRun?.TabAlignment ?? segment.LeadingTabAlignment : PdfTabAlignment.Left,
                backgroundColor: segment.BackgroundColor,
                fontFamily: segment.NamedFont?.FamilyName,
                underlineStyle: segment.UnderlineStyle,
                strikeStyle: segment.StrikeStyle,
                decorationColor: segment.DecorationColor);
            if (preserveTab && segment.SourceTabRun != null) run = run.WithRelativeTabFrom(segment.SourceTabRun);
            if (HasRichTextSpacing(segment.HorizontalTextScaling, segment.CharacterSpacing))
                run = run.WithHorizontalTextScaling(segment.HorizontalTextScaling).WithCharacterSpacing(segment.CharacterSpacing);
            if (!segment.FeatureSettings.Equals(OfficeIMO.Drawing.OfficeTextFeatureSettings.Default))
                run = run.WithFeatureSettings(segment.FeatureSettings);
            if (segment.TextDirection != OfficeIMO.Drawing.OfficeTextDirection.Auto)
                run = run.WithTextDirection(segment.TextDirection);
            run.ContributesLineBoxDescent = segment.ContributesLineBoxDescent;
            run.IsLineBoxStrut = segment.IsLineBoxStrut;
            if (!preserveTab && segment.FormField != null) run = run.WithFormField(segment.FormField);
            if (preserveFootnote && segment.FootnoteReference != null && text == segment.Text) run = run.WithFootnoteReference(segment.FootnoteReference);
            return run;
        }

}
