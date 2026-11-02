using OfficeIMO.Publisher.Internal;
using OfficeIMO.Drawing;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Encodes the bounded Arial font, size and bold direct-formatting slice.</summary>
internal static class QuillCharacterFormatting {
    private static readonly byte[] FontSelectorIds = { 0, 1, 2, 3, 4, 5, 0x2A, 0x2C };

    internal static IReadOnlyList<PublisherQuillFormattingRange> FromParagraph(OfficeRichTextParagraph paragraph, int absoluteTextStart) {
        if (paragraph.Runs.Count == 0) throw new ArgumentException("The styled probe paragraph requires runs.", nameof(paragraph));
        int characters = 0;
        var ranges = new List<PublisherQuillFormattingRange>();
        for (int index = 0; index < paragraph.Runs.Count; index++) {
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
            byte[] font = Variable(0x24, 0x8A, FontSelectorIds.Select(id => Variable(id, 0x88, ShortScalar(0, 0x18, 0))).ToArray());
            byte[] style = Chunk(Marker(2, run.Bold ? (byte)0x0A : (byte)0x02),
                Scalar(0x0C, 0x22, checked((uint)Math.Round(run.FontSize * 12700))),
                run.LanguageId.HasValue ? Scalar(0x12, 0x22, run.LanguageId.Value) : Array.Empty<byte>(), font);
            ranges.Add(new PublisherQuillFormattingRange(checked(absoluteTextStart + characters * 2), style));
        }
        return ranges;
    }
}
