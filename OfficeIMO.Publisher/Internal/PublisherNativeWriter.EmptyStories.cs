using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    /// <summary>Produces the observed native empty-CR representation without changing the public edited model.</summary>
    private static PublisherPublicationState NormalizeEmptyStories(PublisherPublicationState current, PublisherPublicationState baseline,
        long maximumCharacters, CancellationToken token, out List<OfficeConversionFidelityDiagnostic> diagnostics) {
        diagnostics = new(); long? characters = null;
        foreach (PublisherTextStory story in current.Text.Stories.Values) {
            token.ThrowIfCancellationRequested();
            if (story.Paragraphs.Count != 0 || current.Text.IsTableStory(story.Id)
                || !baseline.Text.Stories.TryGetValue(story.Id, out PublisherTextStory? source) || !source.HasFinalParagraphTerminator
                || source.Paragraphs.Count == 0 || source.Paragraphs.Any(p => p.Label != null)) continue;
            characters ??= current.Text.Stories.Values.Sum(s => (long)s.Text.Length);
            if (++characters > maximumCharacters) throw new InvalidDataException("Native empty-story paragraphs exceed the configured text-character limit.");
            OfficeRichTextParagraph last = source.Paragraphs[source.Paragraphs.Count - 1];
            OfficeRichTextParagraph empty = last;
            if (source.Paragraphs.Count != 1 || source.Text != "\n") {
                empty = new OfficeRichTextParagraph(Array.Empty<OfficeRichTextRun>(), last.Alignment, last.LineHeight,
                    last.Margins, last.Indent, last.LineHeightFactor).WithTabStops(last.TabStops);
                empty.SuppressAutomaticHyphenation = last.SuppressAutomaticHyphenation;
            }
            var native = new PublisherTextStory(story.Id, "\n", new[] { empty }, true, story.AutomaticHyphenation);
            current = current.WithText(current.Text.ReplaceStory(native));
            diagnostics.Add(Finding("PUB_NATIVE_EMPTY_STORY_PARAGRAPH",
                "A zero-paragraph public story is represented by one native empty CR paragraph retaining the source final paragraph and CR typing. The public edited model is unchanged; reloading the native file exposes that empty paragraph.",
                OfficeConversionLossKind.Approximation, "Quill/story/" + story.Id));
        }
        return current;
    }
}
