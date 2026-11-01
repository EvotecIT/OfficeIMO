using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Applies the qualified native flags to scene copies; original text, formatting requests and save admission stay unchanged.</summary>
internal static class PublisherHyphenationProjection {
    internal static IReadOnlyList<OfficeRichTextParagraph> Project(PublisherTextStory story,
        IReadOnlyList<OfficeRichTextParagraph> paragraphs, PublisherParseContext context) {
        if (story.AutomaticHyphenation != true) return paragraphs;
        var result = new List<OfficeRichTextParagraph>(paragraphs.Count); bool enabled = false;
        foreach (var paragraph in paragraphs) {
            context.Token.ThrowIfCancellationRequested();
            var plan = paragraph.SuppressAutomaticHyphenation == false
                ? OfficeRichTextHyphenation.Create(paragraph.Runs, 18D, context.Token) : null;
            if (plan == null) result.Add(paragraph);
            else {
                var copy = paragraph.WithTabStops(paragraph.TabStops); copy.Hyphenation = plan;
                result.Add(copy); enabled = true;
            }
        }
        if (enabled) context.Add("PUB_HYPHENATION_LAYOUT_APPROXIMATED",
            "Automatic display hyphens use the shared versioned language patterns and the observed 18-point application-zone profile. Native dictionaries and other application-zone profiles can select different breaks; stored story text remains unchanged.",
            OfficeConversionLossKind.Approximation, "Quill/story/" + story.Id);
        return enabled ? result.AsReadOnly() : paragraphs;
    }
}
