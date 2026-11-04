using OfficeIMO.Publisher.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Encodes the bounded Arial font, size and bold direct-formatting slice.</summary>
internal static class PublisherNativeCreationCharacterFormatting {
    private static readonly byte[] FontSelectorIds = { 0, 1, 2, 3, 4, 5, 0x2A, 0x2C };

    internal static IReadOnlyList<PublisherQuillFormattingRange> FromParagraph(OfficeRichTextParagraph paragraph, int absoluteTextStart,
        PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        if (paragraph.Runs.Count == 0) throw new ArgumentException("The styled probe paragraph requires runs.", nameof(paragraph));
        int characters = 0;
        var ranges = new List<PublisherQuillFormattingRange>();
        for (int index = 0; index < paragraph.Runs.Count; index++) {
            context.Token.ThrowIfCancellationRequested();
            OfficeRichTextRun run = paragraph.Runs[index];
            if (run.Italic || run.Color != OfficeColor.Black || run.Underline || run.Strikethrough
                || run.Baseline != OfficeTextBaseline.Normal || run.BackgroundColor.HasValue
                || run.AutomaticPairKerningThreshold.HasValue || run.ParagraphIndent.HasValue
                || run.LinkUri != null || run.LinkDestinationName != null)
                throw new NotSupportedException("The native character specimen supports black Arial text with size, bold and optional language only.");
            if (!string.Equals(run.FontFamily, "Arial", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("The current direct-formatting profile uses the generated Arial font at index zero.", nameof(paragraph));
            characters = checked(characters + run.Text.Length);
            if (index == paragraph.Runs.Count - 1) characters++; // Native paragraph terminator follows the last run's style.
            byte[] font = PublisherOplWriter.Nested(0x24, 0x8A, FontSelectorIds.Select(id => PublisherOplWriter.Nested(id, 0x88, PublisherOplWriter.UInt16(0, 0, 0x18))).ToArray());
            byte[] style = PublisherOplWriter.Chunk(new byte[] { 2, run.Bold ? (byte)0x0A : (byte)0x02 },
                PublisherOplWriter.UInt32(0x0C, checked((uint)Math.Round(run.FontSize * 12700)), 0x22),
                run.LanguageId.HasValue ? PublisherOplWriter.UInt32(0x12, run.LanguageId.Value, 0x22) : Array.Empty<byte>(), font);
            ranges.Add(new PublisherQuillFormattingRange(checked(absoluteTextStart + characters * 2), style));
        }
        context.Token.ThrowIfCancellationRequested();
        return ranges;
    }
}
