using OfficeIMO.Drawing;
using static OfficeIMO.Publisher.Internal.PublisherTextComparison;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillParagraphWriter {
    /// <summary>Admits semantic changes before source-bound planning, whose run correspondence requires these invariants.</summary>
    internal static bool CanRewriteStory(PublisherTextStory before, PublisherTextStory after, out bool extended) {
        extended = false;
        if (before.Id != after.Id || before.Paragraphs.Count == 0 || after.Paragraphs.Count == 0) return false;
        bool reshape = before.Paragraphs.Count != after.Paragraphs.Count;
        if (reshape && before.Paragraphs.Any(paragraph => paragraph.Runs.Count != 1 || paragraph.Label != null
            || !ParagraphFormattingEqual(before.Paragraphs[0], paragraph)
            || !RunFormattingEqual(before.Paragraphs[0].Runs[0], paragraph.Runs[0]))) {
            // Formatting matches admit text replacement; the native planner
            // separately requires complete, unambiguous source definitions.
            var retained = new HashSet<OfficeRichTextParagraph>(before.Paragraphs.Where(p => p.Label == null), PublisherTextComparison.Paragraphs);
            var formats = new HashSet<OfficeRichTextParagraph>(before.Paragraphs.Where(p => p.Label == null
                && p.Runs.Count > 0 && (p.Runs.Count == 1 || p.Runs.All(r => r.Text.Length > 0))), PublisherTextComparison.ParagraphFormats);
            var texts = new HashSet<OfficeRichTextParagraph>(before.Paragraphs.Where(p => p.Label == null
                && p.Runs.Count > 0 && p.Runs.All(r => r.Text.Length > 0)), PublisherTextComparison.ParagraphTexts);
            var content = new HashSet<OfficeRichTextParagraph>(before.Paragraphs.Where(p => p.Label == null
                && p.Runs.Count > 0 && p.Runs.All(r => r.Text.Length > 0)), PublisherTextComparison.ParagraphContent);
            extended = true;
            return after.Paragraphs.All(p => p.Label == null && (retained.Contains(p)
                || p.Runs.Count > 0 && p.Runs.All(r => r.Text.Length > 0) && (texts.Contains(p) || content.Contains(p) || formats.Contains(p))));
        }
        for (int paragraph = 0; paragraph < after.Paragraphs.Count; paragraph++) {
            OfficeRichTextParagraph a = before.Paragraphs[reshape ? 0 : paragraph], b = after.Paragraphs[paragraph];
            bool empty = b.Runs.Count == 0 || b.Runs.Count == 1 && b.Runs[0].Text.Length == 0;
            bool content = !empty && !PublisherTextComparison.ParagraphTexts.Equals(a, b) && PublisherTextComparison.ParagraphContent.Equals(a, b);
            if (!reshape && a.Runs.Count > 1 && a.Runs.Count != b.Runs.Count && CanReplaceText(a, b)) { extended = true; continue; }
            if (a.Runs.Count < 1 || !empty && !content && a.Runs.Count != 1 && a.Runs.Count != b.Runs.Count
                || a.Label != null || b.Label != null) return false;
            if (content && a.Runs.Count > 1) { extended = true; continue; }
            extended |= empty || PublisherQuillParagraphStyleWriter.HasChanges(a, b);
            if (empty) {
                if (b.Runs.Count == 1 && !RunPreservedFormattingEqual(a.Runs[a.Runs.Count - 1], b.Runs[0])) return false;
                continue;
            }
            for (int index = 0; index < b.Runs.Count; index++) {
                if (b.Runs[index].Text.Length == 0
                    || !RunPreservedFormattingEqual(a.Runs[a.Runs.Count == 1 ? 0 : index], b.Runs[index])) return false;
                extended |= PublisherQuillCharacterStyleWriter.HasExtendedChanges(a.Runs[a.Runs.Count == 1 ? 0 : index], b.Runs[index])
                    || PublisherQuillCharacterStyleWriter.HasExplicitAttributeChange(a.Runs[a.Runs.Count == 1 ? 0 : index], b.Runs[index]);
            }
        }
        return true;
    }

}
