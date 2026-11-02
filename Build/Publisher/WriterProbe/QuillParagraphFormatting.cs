using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Uses the canonical paragraph codec to encode authored overrides above empty native defaults.</summary>
internal static class QuillParagraphFormatting {
    internal static byte[] Create(OfficeRichTextParagraph? paragraph) {
        byte[] empty = Chunk();
        if (paragraph == null) return empty;
        if (paragraph.Label != null)
            throw new NotSupportedException("Paragraph labels require a separately qualified native creation profile.");
        var defaults = new OfficeRichTextParagraph(Array.Empty<OfficeRichTextRun>());
        var tabs = new OfficeTextTabStops(Array.Empty<OfficeTextTabStop>());
        defaults = defaults.WithTabStops(tabs);
        if (!PublisherQuillParagraphStyleWriter.HasChanges(defaults, paragraph)) return empty;
        if (paragraph.TabStops == null) paragraph = paragraph.WithTabStops(tabs);
        return PublisherQuillParagraphStyleWriter.Rewrite(new PublisherBinaryData(empty, "authored paragraph"),
            new PublisherQuillSourceSpan(0, empty.Length), defaults, paragraph,
            new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None))
            ?? throw new NotSupportedException("The probe paragraph exceeds the qualified native paragraph codec.");
    }
}
