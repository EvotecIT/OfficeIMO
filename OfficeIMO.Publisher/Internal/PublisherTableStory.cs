using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>The ordered semantic cell segments of one complete story. Boundaries and terminator conventions survive edits independently of native TCD offsets.</summary>
internal sealed class PublisherTableStory {
    internal PublisherTableStory(IEnumerable<PublisherTextStory> segments) => Segments = Array.AsReadOnly(segments.ToArray());
    internal IReadOnlyList<PublisherTextStory> Segments { get; }

    internal bool Matches(PublisherTextStory story) {
        if (string.Concat(Segments.Select(segment => segment.Text)) != story.Text) return false;
        // A cell boundary inside an ordinary paragraph cannot be rebuilt by
        // concatenating paragraphs without changing the surrounding content.
        for (int i = 0; i < Segments.Count - 1; i++)
            if (!Segments[i].HasFinalParagraphTerminator) return false;
        return true;
    }

    internal PublisherTableStory Replace(int index, IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        PublisherReadOptions options, CancellationToken token) {
        PublisherTextStory original = Segments[index];
        // Clearing retains the final paragraph's formatting and cell boundary.
        // Otherwise the following cell would become part of this cell's paragraph.
        if (paragraphs.Count == 0) {
            OfficeRichTextParagraph? retained = original.Paragraphs.LastOrDefault();
            paragraphs = new[] { retained == null ? new OfficeRichTextParagraph(Array.Empty<OfficeRichTextRun>())
                : new OfficeRichTextParagraph(Array.Empty<OfficeRichTextRun>(), retained.Alignment, retained.LineHeight,
                    retained.Margins, retained.Indent, retained.LineHeightFactor) {
                        SuppressAutomaticHyphenation = retained.SuppressAutomaticHyphenation
                    }.WithTabStops(retained.TabStops) };
        }
        PublisherTextStory replacement = PublisherStoryEditor.Replace(original, paragraphs, options, token);
        var segments = Segments.ToArray(); segments[index] = replacement;
        return new PublisherTableStory(segments);
    }

    internal PublisherTextStory Complete(PublisherTextStory original, PublisherReadOptions options, CancellationToken token) {
        var paragraphs = Segments.SelectMany(segment => segment.Paragraphs).ToArray();
        PublisherTextStory validated = PublisherStoryEditor.Replace(original, paragraphs, options, token);
        string exact = string.Concat(Segments.Select(segment => segment.Text));
        if (validated.Text != exact)
            throw new NotSupportedException("The cell boundaries cannot be represented by the complete styled story without changing adjacent content.");
        return validated;
    }
}
