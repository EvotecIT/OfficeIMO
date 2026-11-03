using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Uses the canonical paragraph codec to encode authored overrides above empty native defaults.</summary>
internal static class PublisherNativeCreationParagraphFormatting {
    internal static byte[] Create(OfficeRichTextParagraph? paragraph, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        byte[] empty = PublisherOplWriter.Chunk();
        if (paragraph == null) return empty;
        if (paragraph.Label != null)
            throw new NotSupportedException("Paragraph labels require a separately qualified native creation profile.");
        var defaults = new OfficeRichTextParagraph(Array.Empty<OfficeRichTextRun>());
        var tabs = new OfficeTextTabStops(Array.Empty<OfficeTextTabStop>());
        defaults = defaults.WithTabStops(tabs);
        if (!PublisherQuillParagraphStyleWriter.HasChanges(defaults, paragraph)) {
            context.Token.ThrowIfCancellationRequested();
            return empty;
        }
        if (paragraph.TabStops == null) paragraph = paragraph.WithTabStops(tabs);
        byte[] result = PublisherQuillParagraphStyleWriter.Create(defaults, paragraph, context)
            ?? throw new NotSupportedException("The probe paragraph exceeds the qualified native paragraph codec.");
        context.Token.ThrowIfCancellationRequested();
        return result;
    }
}
