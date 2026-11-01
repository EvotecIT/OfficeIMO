using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Compares projected text and visual style templates; explicit native attribute requests are separate from opaque source correspondence.</summary>
internal static class PublisherTextComparison {
    internal static IEqualityComparer<OfficeRichTextParagraph> Paragraphs { get; } = new ParagraphComparer();
    /// <summary>Matches editable visual formatting while allowing run text to change. Native language and other source-token differences are checked by template correspondence.</summary>
    internal static IEqualityComparer<OfficeRichTextParagraph> ParagraphFormats { get; } = new ParagraphComparer(includeText: false);
    /// <summary>Matches retained run text and fields the native formatting codecs cannot rewrite.</summary>
    internal static IEqualityComparer<OfficeRichTextParagraph> ParagraphTexts { get; } = new ParagraphComparer(includeFormatting: false);
    /// <summary>Matches nonempty paragraph text and preserved fields at each position, independently of run boundaries.</summary>
    internal static IEqualityComparer<OfficeRichTextParagraph> ParagraphContent { get; } = new ContentComparer();
    internal static bool StoryEqual(PublisherTextStory a, PublisherTextStory b) => ReferenceEquals(a, b)
        || a.Id == b.Id && a.Text == b.Text && a.HasFinalParagraphTerminator == b.HasFinalParagraphTerminator
        && a.Paragraphs.Count == b.Paragraphs.Count && a.Paragraphs.Zip(b.Paragraphs, (before, after) =>
            ParagraphEqual(before, after) && !before.Runs.Zip(after.Runs, PublisherQuillCharacterStyleWriter.HasExplicitAttributeChange).Any(value => value)
            && (before.Label == null || after.Label == null || !PublisherQuillCharacterStyleWriter.HasExplicitAttributeChange(before.Label.Run, after.Label.Run))).All(value => value);
    internal static bool ParagraphEqual(OfficeRichTextParagraph a, OfficeRichTextParagraph b) => ReferenceEquals(a, b)
        || ParagraphFormattingEqual(a, b)
        && a.Runs.Count == b.Runs.Count && a.Runs.Zip(b.Runs, RunEqual).All(value => value);
    internal static bool ParagraphFormattingEqual(OfficeRichTextParagraph a, OfficeRichTextParagraph b) =>
        !PublisherQuillParagraphStyleWriter.HasChanges(a, b) && LabelEqual(a.Label, b.Label);
    internal static bool RunEqual(OfficeRichTextRun a, OfficeRichTextRun b) => a.Text == b.Text && RunFormattingEqual(a, b);
    internal static bool RunFormattingEqual(OfficeRichTextRun a, OfficeRichTextRun b) => a.FontSize == b.FontSize
        && a.Bold == b.Bold && !PublisherQuillCharacterStyleWriter.HasExtendedChanges(a, b) && RunPreservedFormattingEqual(a, b);
    internal static bool RunPreservedFormattingEqual(OfficeRichTextRun a, OfficeRichTextRun b) =>
        a.BackgroundColor == b.BackgroundColor && a.ParagraphIndent == b.ParagraphIndent
        && a.LinkUri == b.LinkUri && a.LinkDestinationName == b.LinkDestinationName;
    private static bool LabelEqual(OfficeTextParagraphLabel? a, OfficeTextParagraphLabel? b) => ReferenceEquals(a, b)
        || a != null && b != null && a.Position == b.Position && a.Alignment == b.Alignment && a.MinimumWidth == b.MinimumWidth
        && a.MinimumDistance == b.MinimumDistance && a.FollowedBy == b.FollowedBy && a.TextPosition == b.TextPosition && RunEqual(a.Run, b.Run);

    private sealed class ContentComparer : IEqualityComparer<OfficeRichTextParagraph> {
        public bool Equals(OfficeRichTextParagraph? a, OfficeRichTextParagraph? b) {
            if (ReferenceEquals(a, b)) return true;
            if (a == null || b == null || !LabelEqual(a.Label, b.Label)
                || a.Runs.Count == 0 || b.Runs.Count == 0 || a.Runs.Any(r => r.Text.Length == 0) || b.Runs.Any(r => r.Text.Length == 0)) return false;
            int i = 0, j = 0, left = 0, right = 0;
            while (i < a.Runs.Count && j < b.Runs.Count) {
                OfficeRichTextRun x = a.Runs[i], y = b.Runs[j];
                int length = Math.Min(x.Text.Length - left, y.Text.Length - right);
                if (!RunPreservedFormattingEqual(x, y) || string.CompareOrdinal(x.Text, left, y.Text, right, length) != 0) return false;
                left += length; right += length;
                if (left == x.Text.Length) { i++; left = 0; }
                if (right == y.Text.Length) { j++; right = 0; }
            }
            return i == a.Runs.Count && j == b.Runs.Count;
        }

        public int GetHashCode(OfficeRichTextParagraph paragraph) {
            int hash = 17;
            foreach (var run in paragraph.Runs)
                foreach (char value in run.Text) hash = unchecked(hash * 31 + value);
            return hash;
        }
    }

    private sealed class ParagraphComparer : IEqualityComparer<OfficeRichTextParagraph> {
        private readonly bool _includeText, _includeFormatting;
        internal ParagraphComparer(bool includeText = true, bool includeFormatting = true) { _includeText = includeText; _includeFormatting = includeFormatting; }
        public bool Equals(OfficeRichTextParagraph? a, OfficeRichTextParagraph? b) => ReferenceEquals(a, b)
            || a != null && b != null && LabelEqual(a.Label, b.Label) && (!_includeFormatting || ParagraphFormattingEqual(a, b))
                && a.Runs.Count == b.Runs.Count && a.Runs.Zip(b.Runs, (x, y) => (!_includeText || x.Text == y.Text)
                    && (_includeFormatting ? RunFormattingEqual(x, y) : RunPreservedFormattingEqual(x, y))).All(value => value);

        public int GetHashCode(OfficeRichTextParagraph paragraph) {
            int hash = 17;
            if (_includeFormatting) {
                Add(paragraph.Alignment); Add(paragraph.LineHeight); Add(paragraph.LineHeightFactor);
                Add(paragraph.Margins); Add(paragraph.Indent);
            }
            if (_includeFormatting && paragraph.TabStops is { } tabs) {
                Add(tabs.DefaultInterval); Add(tabs.Origin); Add(tabs.AlignWithParagraph); Add(tabs.Stops.Count);
                foreach (var stop in tabs.Stops) {
                    Add(stop.Position); Add(stop.Alignment); Add(stop.Character); Add(stop.LeaderText);
                    Add(stop.LineLeader); Add(stop.LeaderStyle);
                }
            }
            Add(paragraph.Runs.Count);
            foreach (var run in paragraph.Runs) {
                if (_includeText) Add(run.Text);
                if (_includeFormatting) {
                    Add(run.FontSize); Add(run.FontFamily); Add(run.Color); Add(run.Bold); Add(run.Italic);
                    Add(run.UnderlineStyle); Add(run.StrikethroughStyle); Add(run.Baseline);
                }
                Add(run.BackgroundColor);
                Add(run.ParagraphIndent); Add(run.LinkUri); Add(run.LinkDestinationName);
            }
            return hash;

            void Add(object? value) { hash = unchecked(hash * 31 + (value?.GetHashCode() ?? 0)); }
        }
    }
}
