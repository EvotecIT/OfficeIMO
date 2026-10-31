using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static class PublisherStoryEditor {
    private static readonly Encoding StrictUtf16 = new UnicodeEncoding(false, false, true);

    /// <summary>Creates an independent authored story without a native-only final terminator.</summary>
    internal static PublisherTextStory Create(uint storyId, IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        PublisherReadOptions options, CancellationToken token) =>
        Replace(new PublisherTextStory(storyId, string.Empty, Array.Empty<OfficeRichTextParagraph>(), false), paragraphs, options, token);

    internal static PublisherTextStory Replace(PublisherTextStory original, IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        PublisherReadOptions options, CancellationToken token) {
        if (paragraphs.Count > 4096) throw new ArgumentException("The story exceeds the paragraph limit.", nameof(paragraphs));
        var copied = new List<OfficeRichTextParagraph>(paragraphs.Count);
        var text = new StringBuilder();
        int runs = 0, labelCharacters = 0;
        foreach (OfficeRichTextParagraph paragraph in paragraphs) {
            token.ThrowIfCancellationRequested();
            if (paragraph == null) throw new ArgumentException("Paragraphs cannot contain null entries.", nameof(paragraphs));
            var values = new List<OfficeRichTextRun>();
            foreach (OfficeRichTextRun run in paragraph.Runs) {
                token.ThrowIfCancellationRequested();
                if (++runs > 4096) throw new ArgumentException("The story exceeds the styled-run limit.", nameof(paragraphs));
                string value = run.Text.Replace("\r\n", "\n").Replace('\r', '\n');
                ValidateRun(run, value);
                values.Add(new OfficeRichTextRun(value, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline,
                    run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle,
                    run.Baseline, run.ParagraphIndent) { LanguageId = run.LanguageId });
            }
            if (paragraph.Label != null) {
                OfficeRichTextRun label = paragraph.Label.Run;
                ValidateRun(label, label.Text);
                if (label.Text.Any(char.IsControl))
                    throw new ArgumentException("Story labels cannot contain control characters.", nameof(paragraphs));
                if (++runs > 4096) throw new ArgumentException("The story exceeds the styled-run limit, including labels.", nameof(paragraphs));
                labelCharacters += label.Text.Length;
            }
            OfficeRichTextParagraph copy = paragraph.Label == null
                ? new OfficeRichTextParagraph(values, paragraph.Alignment, paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor)
                : new OfficeRichTextParagraph(values, paragraph.Label, paragraph.Alignment, paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor);
            copy = copy.WithTabStops(paragraph.TabStops);
            // Immutable source paragraphs identify a deliberate native selection when
            // equal projected values carry different opaque formatting definitions.
            copied.Add(PublisherTextComparison.ParagraphEqual(paragraph, copy) ? paragraph : copy);
            if (copied.Count > 1) text.Append('\n');
            foreach (OfficeRichTextRun run in values) text.Append(run.Text);
            if ((long)text.Length + labelCharacters > 100000 || (long)text.Length + labelCharacters > options.Limits.MaxTextCharacters)
                throw new ArgumentException("Story text exceeds the configured character limit.", nameof(paragraphs));
        }
        if (copied.Count != 0 && original.HasFinalParagraphTerminator) text.Append('\n');
        if (text.Length > 100000 || text.Length > options.Limits.MaxTextCharacters || runs + Math.Max(0, copied.Count - 1) > 4096)
            throw new ArgumentException("The story exceeds the configured text or run limit.", nameof(paragraphs));
        return new PublisherTextStory(original.Id, text.ToString(), copied, original.HasFinalParagraphTerminator);
    }

    /// <summary>Counts semantic story characters and separately stored labels for edit admission.</summary>
    internal static long CharacterCount(PublisherTextStory story) => story.Text.Length
        + story.Paragraphs.Sum(paragraph => (long)(paragraph.Label?.Run.Text.Length ?? 0));

    internal static void ValidatePageName(string name, PublisherReadOptions options) {
        if (name.Length > options.Limits.MaxTextCharacters)
            throw new ArgumentException("The page name exceeds the configured character limit.", nameof(name));
        ValidateText(name, nameof(name));
    }

    private static void ValidateRun(OfficeRichTextRun run, string text) {
        ValidateText(text, "paragraphs");
        if (run.FontSize <= 0 || double.IsNaN(run.FontSize) || double.IsInfinity(run.FontSize))
            throw new ArgumentException("Story font sizes must be positive finite values.", "paragraphs");
    }

    private static void ValidateText(string text, string parameterName) {
        try { StrictUtf16.GetByteCount(text); }
        catch (EncoderFallbackException exception) { throw new ArgumentException("Publication text must contain well-formed UTF-16.", parameterName, exception); }
        if (text.IndexOf('\0') >= 0) throw new ArgumentException("Publication text cannot contain null characters.", parameterName);
    }
}
